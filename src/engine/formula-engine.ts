import {
  type CellArray,
  type CellData,
  type FormulaArgumentKind,
  type FormulaContext,
  type FormulaFunction,
  type FormulaFunctionOptions,
  type GridData,
  cellKey,
  colToLetter,
  parseKey,
  refToCoord,
} from '../types.js';
import { MAX_COLS, MAX_ROWS, type Node, type RefInfo, callsAny, parseFormula, sheetPrefix } from './parser.js';
import {
  type EngineSheet,
  type Scope,
  ERROR_CODES,
  LambdaValue,
  RangeValue,
  RefValue,
  arrayToText,
  cellArraysEqual,
  flattenDeep,
  isErrorCode,
  literalValue,
  numberToText,
  toBoolean,
  toCellArray,
  toErrorCode,
  toNumber,
  toText,
} from './values.js';
import { sameValue } from './functions-logical-info.js';
import { ArgumentError, MAX_ARRAY_CELLS, type Registrar, asArray, power, registerPureBuiltins, transpose } from './functions.js';
import { mapFormulaReferences } from '../utils/references.js';

export { RangeValue, RefValue, LambdaValue } from './values.js';

/**
 * FormulaEngine -- parser and evaluator for Excel-like formulas.
 *
 * Formulas are parsed once into an AST (`parser.ts`) and evaluated against a
 * workbook of one or more sheets. On top of scalar formulas it supports:
 *
 * - Arrays: operators and single-value functions broadcast over arrays, and
 *   a formula whose result is an array spills into the cells below and to
 *   the right of it (`#SPILL!` when something is in the way). The engine
 *   writes spilled values into the sheet data as cells with `spilledFrom` set.
 * - Several values in one cell: an array whose result is a single element
 *   that is itself an array (`={{1;2;3}}`, `={FILTER(...)}`) stays in the cell.
 * - References as values: whole columns/rows (`A:A`, `1:1`), `A1:INDEX(...)`,
 *   OFFSET, INDIRECT, the spill operator (`A1#`) and `@`.
 * - Defined names, LET and LAMBDA, and sheet-qualified references (`Sheet2!A1`).
 *
 * Dependencies are recorded while evaluating (cells, and range nodes for
 * larger ranges), so `recalculateAffected()` re-evaluates only what a change
 * can reach. Circular references are detected at evaluation time.
 */

/** Maximum nesting depth for formula evaluation to prevent stack overflow from deeply nested formulas */
const MAX_EVAL_DEPTH = 64;

/** Upper bound on deferred cells when resolving reference chains deeper than MAX_EVAL_DEPTH */
const MAX_DEFERRED_CELLS = 100_000;

/** Ranges up to this many cells are tracked cell by cell; larger ones as one range node */
const SMALL_RANGE = 64;

/**
 * How often a recalculation re-runs cells that read a range a spill grew
 * into or left. A chain of spills that each depend on the previous one's
 * size needs one round per link; the limit only stops shapes that never settle.
 */
const MAX_SPILL_ROUNDS = 2000;

/**
 * How often the spill of a formula that reads a whole column or row may grow
 * in one recalculation before the engine checks whether the growth has an end.
 */
const MAX_SPILL_GROWTHS = 3;

/** Parsed formulas kept for reuse */
const AST_CACHE_LIMIT = 20_000;

const MAX_LAMBDA_DEPTH = 256;

/**
 * Thrown when resolving a reference would exceed MAX_EVAL_DEPTH. `id` is the
 * cell that could not be entered; the top-level evaluation computes it on its
 * own first (memoizing the result) and then retries, so long reference chains
 * resolve without growing the JS call stack.
 */
class EvalDepthError extends Error {
  constructor(readonly id: number) {
    super('#ERROR!');
  }
}

type EvalResult = {
  displayValue: string;
  type: 'text' | 'number' | 'boolean' | 'error';
  /** Set when the value is several values in one cell */
  array?: CellArray;
};

/** What a formula cell evaluated to. */
interface CellResult {
  raw: string;
  /** The cell's own value: a scalar, or an array kept in the cell. Undefined when `error` is set (or blank). */
  value?: unknown;
  error?: string;
  /** The whole array, when the result spills over more than this cell */
  full?: RangeValue;
  volatile?: boolean;
}

type ArgKind = 'n' | 't' | 'v' | 'a' | 'b' | 'r' | 'x';

interface FunctionEntry {
  fn?: FormulaFunction;
  /** Kinds of the leading arguments */
  kinds: ArgKind[];
  /** Kinds that repeat for the arguments after `kinds` */
  repeat: ArgKind[];
  agg?: (values: unknown[]) => unknown;
  propagate?: boolean;
  arity?: [number, number];
  volatile?: boolean;
  /** Registered by the host, so it replaces a built-in of the same name */
  custom?: boolean;
}

interface EvalCtx {
  sheet: EngineSheet;
  /** The cell being evaluated; (0, 0) when there is none */
  row: number;
  col: number;
  scope: Scope | null;
  fctx?: FormulaContext;
}

interface RangeNode {
  key: string;
  sheet: EngineSheet;
  top: number;
  left: number;
  bottom: number;
  right: number;
  dependents: Set<number>;
}

interface Precedents {
  cells: Set<number>;
  nodes: Set<RangeNode>;
}

interface DefinedName {
  name: string;
  formula: string;
  /** Sheet that unqualified references in the formula point to */
  sheetId: number;
}

/**
 * State of one evaluation pass (one `evaluate()`, or a whole
 * `recalculate()` / `recalculateAffected()`).
 *
 * - `pure`: nothing is stored; every formula cell a reference reaches is
 *   computed afresh once (memoized), so the result never depends on stale caches.
 * - `root`: like `pure`, but the dependencies of everything reached are
 *   recorded for one cell, whose result is cached (`evaluate(raw, key)`).
 * - `commit`: a recalculation. Cells record their own dependencies, results
 *   are cached and spills are written to the sheets.
 */
interface Pass {
  mode: 'pure' | 'root' | 'commit';
  /** pure: read formula cells from their cached results (volatile ones too) instead of computing them */
  committed: boolean;
  memo: Map<number, CellResult>;
  /** Values of ranges read in this pass */
  ranges: Map<string, RangeValue>;
  /** commit: cells that must be recomputed rather than read from the cache */
  stale: Set<number> | null;
  /** commit: cells computed in this pass, whose display values are written at the end */
  touched: Set<number>;
  changed: Map<EngineSheet, Set<string>>;
  seq: number;
  started: Map<number, number>;
  /** Cells a spill started or stopped covering, with when */
  shapeChanges: Array<{ id: number; seq: number; released?: boolean }>;
  /** Cells a spill covered at some point in this pass → its anchor, so a formula reading its own spill range is circular */
  claimed: Map<number, number>;
  /** commit: how often the spill of a whole-column reader grew in this pass */
  growths: Map<number, number>;
  /**
   * commit: whole-column readers whose spill grows without end, because it
   * moves the end of what they read (`=A:A*2` below row 1). No size settles
   * that, so they are #SPILL! for the rest of the pass, as in Excel.
   */
  selfFeeding: Set<number>;
  /** commit: whole-column readers that grew but were found to have a size they stop at */
  bounded: Set<number>;
  /** The cell that is reading whole columns and rows in full, to see how large its result can get */
  untrimmed: number | null;
  /**
   * pure / root: the ranges that formulas computed in this pass would spill
   * over. Nothing is written to the sheet in these passes, so cells in them
   * are read from here.
   */
  pureSpills: Array<{ anchor: number; rows: number; cols: number }>;
}

const COMPARISONS = new Set(['=', '<>', '<', '>', '<=', '>=']);

/** Stands in for the sheet of a defined name whose sheet was removed: references without a sheet name are #REF! on it. */
const NO_SHEET: EngineSheet = { id: 0, name: '', data: new Map(), rows: MAX_ROWS, cols: MAX_COLS, used: null };

function cellId(sheet: EngineSheet, row: number, col: number): number {
  return (sheet.id * MAX_ROWS + row) * MAX_COLS + col;
}

function idCol(id: number): number {
  return id % MAX_COLS;
}

function idRow(id: number): number {
  return Math.floor(id / MAX_COLS) % MAX_ROWS;
}

function idSheet(id: number): number {
  return Math.floor(Math.floor(id / MAX_COLS) / MAX_ROWS);
}

function parseKinds(spec: string): { kinds: ArgKind[]; repeat: ArgKind[] } {
  const [head, tail] = spec.split('|');
  const kinds = [...head] as ArgKind[];
  const repeat = tail !== undefined ? ([...tail] as ArgKind[]) : kinds.length ? [kinds[kinds.length - 1]] : ['a' as ArgKind];
  return { kinds, repeat };
}

const PUBLIC_KINDS: Record<FormulaArgumentKind, ArgKind> = {
  number: 'n',
  text: 't',
  value: 'v',
  array: 'a',
  reference: 'r',
};

/** Whether a name can be a defined name: an identifier that is not a cell reference or a boolean. */
export function isValidName(name: string): boolean {
  if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(name)) return false;
  if (/^[A-Za-z]{1,3}[0-9]+$/.test(name) || /^[Rr][0-9]*[Cc][0-9]*$/.test(name)) return false;
  const upper = name.toUpperCase();
  return upper !== 'TRUE' && upper !== 'FALSE';
}

export class FormulaEngine {
  private _functions: Map<string, FunctionEntry> = new Map();
  private _builtinFunctionNames: Set<string> = new Set();
  private _volatileNames: Set<string> = new Set();

  private _sheets: EngineSheet[] = [];
  private _sheetsByName: Map<string, EngineSheet> = new Map();
  private _sheetsById: Map<number, EngineSheet> = new Map();
  private _active!: EngineSheet;
  private _nextSheetId = 1;

  private _names: Map<string, DefinedName> = new Map();
  private _nameDepth = 0;
  private _lambdaDepth = 0;

  private _evaluating: Set<number> = new Set(); // circular reference detection
  private _evalDepth = 0;

  // Dependency tracking for targeted recalculation
  private _precedents: Map<number, Precedents> = new Map();
  private _cellDependents: Map<number, Set<number>> = new Map();
  private _rangeNodes: Map<number, Map<string, RangeNode>> = new Map();
  /** The cell whose dependencies are being recorded */
  private _tracking: number | null = null;

  /** Results of formula cells, valid for the formula text they were computed from */
  private _results: Map<number, CellResult> = new Map();
  /**
   * Formula cells that read a whole column or row, with the used extent they
   * read up to. That extent depends on the used part of the workbook (and not
   * on the reader's own spill), so a reader is recomputed when its extent moves.
   */
  private _wholeReaders: Map<number, { rows: number; cols: number }> = new Map();
  /** Cached result of `_spillReach()`; null when spills changed */
  private _reach: { rows: number; rowsBy: number; rowsNext: number; cols: number; colsBy: number; colsNext: number } | null =
    null;
  /** The formula cell being computed (innermost), whatever the kind of pass */
  private _current: number | null = null;
  /** Formula cells whose result is volatile (NOW, TODAY, ...): recomputed by every recalculation */
  private _volatileCells: Set<number> = new Set();
  /** Spilled cell → the anchor cell whose formula spills into it */
  private _ghosts: Map<number, number> = new Map();
  /** Anchor → size of the range it currently spills over */
  private _spills: Map<number, { rows: number; cols: number }> = new Map();
  /** Anchor → size of the range it would spill over if nothing were in the way */
  private _blocked: Map<number, { rows: number; cols: number }> = new Map();

  private _pass: Pass | null = null;
  private _asts: Map<string, Node | Error> = new Map();

  /**
   * Cells whose display value changed in the last `recalculate()` /
   * `recalculateAffected()`, by sheet name. Includes cells that started or
   * stopped showing a spilled value.
   */
  changes: Map<string, Set<string>> = new Map();

  constructor() {
    this._active = this._createSheet('Sheet1', new Map());
    this._registerBuiltins();
    this._builtinFunctionNames = new Set(this._functions.keys());
  }

  // ─── Functions ──────────────────────────────────────

  /**
   * Register a user-defined formula function.
   * The name is normalized to uppercase for case-insensitive lookup.
   *
   * @param name - Function name (e.g. "MYFUNC"); stored as uppercase
   * @param fn - The function implementation
   * @param options - How arguments are passed, and whether the result is volatile
   */
  registerFunction(name: string, fn: FormulaFunction, options: FormulaFunctionOptions = {}): void {
    const upper = name.toUpperCase();
    const kinds = (options.args ?? []).map((k) => PUBLIC_KINDS[k] ?? 'a');
    this._functions.set(upper, {
      fn,
      kinds,
      repeat: [kinds.length ? kinds[kinds.length - 1] : 'a'],
      volatile: options.volatile,
      custom: true,
    });
    if (options.volatile) this._volatileNames.add(upper);
    else if (!this._builtinFunctionNames.has(upper)) this._volatileNames.delete(upper);
    this._asts.clear();
    // Which formulas are volatile depends on the functions
    this._volatility.clear();
    for (const result of this._results.values()) delete result.volatile;
    this._volatileCells.clear();
  }

  /**
   * Replace the entire set of user-defined functions while keeping built-ins.
   *
   * @param functions - New custom functions keyed by name
   */
  setFunctions(functions: Record<string, FormulaFunction>): void {
    for (const name of Array.from(this._functions.keys())) {
      if (!this._builtinFunctionNames.has(name)) {
        this._functions.delete(name);
        this._volatileNames.delete(name);
      }
    }

    for (const [name, fn] of Object.entries(functions)) {
      this.registerFunction(name, fn);
    }
    this._volatility.clear();
  }

  /** Names of all registered functions (upper case), built-in and custom. */
  functionNames(): string[] {
    return [...new Set([...this._functions.keys(), ...Object.keys(this._special)])].sort();
  }

  // ─── Sheets ─────────────────────────────────────────

  private _createSheet(name: string, data: GridData, rows = MAX_ROWS, cols = MAX_COLS): EngineSheet {
    const sheet: EngineSheet = { id: this._nextSheetId++, name, data, rows, cols, used: null };
    this._sheets.push(sheet);
    this._sheetsByName.set(name.toLowerCase(), sheet);
    this._sheetsById.set(sheet.id, sheet);
    return sheet;
  }

  private _sheet(name?: string): EngineSheet {
    if (name === undefined) return this._active;
    const sheet = this._sheetsByName.get(name.toLowerCase());
    if (!sheet) throw new Error(`Unknown sheet: ${name}`);
    return sheet;
  }

  /** Names of the sheets, in order. */
  sheetNames(): string[] {
    return this._sheets.map((s) => s.name);
  }

  hasSheet(name: string): boolean {
    return this._sheetsByName.has(name.toLowerCase());
  }

  /** The sheet that keys and unqualified formulas refer to when no sheet is named. */
  get activeSheet(): string {
    return this._active.name;
  }

  setActiveSheet(name: string): void {
    this._active = this._sheet(name);
  }

  /**
   * Add a sheet. Formulas that already refer to its name start resolving
   * after the next `recalculate()`.
   *
   * @param name - Sheet name; must be unique ignoring case
   * @param data - The sheet's cells (default: empty)
   */
  addSheet(name: string, data: GridData = new Map(), rows?: number, cols?: number): void {
    if (this.hasSheet(name)) throw new Error(`Sheet already exists: ${name}`);
    this._createSheet(name, data, rows, cols);
  }

  /**
   * Remove a sheet. Formulas referring to it evaluate to `#REF!` after the
   * next `recalculate()`. The last sheet cannot be removed.
   */
  removeSheet(name: string): void {
    const sheet = this._sheet(name);
    if (this._sheets.length === 1) throw new Error('A workbook needs at least one sheet');
    this._sheets = this._sheets.filter((s) => s !== sheet);
    this._sheetsByName.delete(name.toLowerCase());
    this._sheetsById.delete(sheet.id);
    if (this._active === sheet) this._active = this._sheets[0];
    this._reset();
  }

  /**
   * Rename a sheet and rewrite every formula (and defined name) that refers
   * to it, in all sheets.
   *
   * @returns The keys of the cells whose formula text changed, by (new) sheet name
   */
  renameSheet(oldName: string, newName: string): Map<string, string[]> {
    const sheet = this._sheet(oldName);
    const existing = this._sheetsByName.get(newName.toLowerCase());
    if (existing && existing !== sheet) throw new Error(`Sheet already exists: ${newName}`);
    const rewrite = (formula: string) => renameSheetInFormula(formula, oldName, newName);
    this._sheetsByName.delete(oldName.toLowerCase());
    sheet.name = newName;
    this._sheetsByName.set(newName.toLowerCase(), sheet);

    const changed = new Map<string, string[]>();
    for (const s of this._sheets) {
      const keys: string[] = [];
      for (const [key, cell] of s.data) {
        if (!cell.rawValue.startsWith('=')) continue;
        const next = rewrite(cell.rawValue);
        if (next !== cell.rawValue) {
          cell.rawValue = next;
          keys.push(key);
        }
      }
      if (keys.length) changed.set(s.name, keys);
    }
    for (const def of this._names.values()) def.formula = rewrite(def.formula);
    this._reset();
    return changed;
  }

  /** Move a sheet to a new position in the order. */
  moveSheet(name: string, index: number): void {
    const sheet = this._sheet(name);
    this._sheets = this._sheets.filter((s) => s !== sheet);
    this._sheets.splice(Math.max(0, Math.min(index, this._sheets.length)), 0, sheet);
  }

  /**
   * Set a sheet's grid size, which whole-column/row references and spills are bounded by.
   *
   * @returns Whether the size changed
   */
  setSheetSize(rows: number, cols: number, sheet?: string): boolean {
    const s = this._sheet(sheet);
    const nextRows = Math.min(MAX_ROWS, Math.max(1, rows));
    const nextCols = Math.min(MAX_COLS, Math.max(1, cols));
    if (s.rows === nextRows && s.cols === nextCols) return false;
    s.rows = nextRows;
    s.cols = nextCols;
    return true;
  }

  /** The data map of a sheet. */
  getSheetData(sheet?: string): GridData {
    return this._sheet(sheet).data;
  }

  /** Update the data reference for evaluation (full reset of dep graph and cache) */
  setData(data: GridData, sheet?: string): void {
    const s = this._sheet(sheet);
    s.data = data;
    s.used = null;
    this._reset();
  }

  /**
   * Update the data reference while preserving the dependency graph for
   * unchanged cells. Only clears deps and cache for the specified changed keys,
   * allowing recalculateAffected() to use the existing graph for BFS traversal.
   */
  updateData(data: GridData, changedKeys: string[], sheet?: string): void {
    const s = this._sheet(sheet);
    s.data = data;
    s.used = null;
    for (const key of changedKeys) {
      const { row, col } = parseKey(key);
      const id = cellId(s, row, col);
      this._clearDeps(id);
      this._results.delete(id);
    }
  }

  /** Forget every cached result, dependency and spill (the next recalculation rebuilds them). */
  private _reset(): void {
    this._precedents.clear();
    this._cellDependents.clear();
    this._rangeNodes.clear();
    this._results.clear();
    this._volatileCells.clear();
    this._wholeReaders.clear();
    this._reach = null;
    this._ghosts.clear();
    this._spills.clear();
    this._blocked.clear();
  }

  // ─── Defined names ──────────────────────────────────

  /**
   * Define (or redefine) a name for a formula: a reference (`=A1:A10`,
   * `=Sheet2!B2`), a constant (`=0.19`) or a LAMBDA. Names are workbook-wide
   * and case-insensitive. Call `recalculate()` afterwards.
   *
   * @param name - The name; must be an identifier that is not a cell reference
   * @param formula - What the name stands for, with or without the leading `=`
   * @param sheet - Sheet that unqualified references in the formula point to (default: the active sheet)
   */
  defineName(name: string, formula: string, sheet?: string): void {
    if (!isValidName(name)) throw new Error(`Invalid name: ${name}`);
    this._names.set(name.toUpperCase(), {
      name,
      formula: formula.startsWith('=') ? formula : '=' + formula,
      sheetId: this._sheet(sheet).id,
    });
  }

  /** Remove a defined name. Returns whether it existed. */
  removeName(name: string): boolean {
    return this._names.delete(name.toUpperCase());
  }

  /** All defined names with their formulas and the sheet their unqualified references point to. */
  getNames(): Array<{ name: string; formula: string; sheet: string }> {
    return [...this._names.values()].map((d) => ({
      name: d.name,
      formula: d.formula,
      sheet: this._sheetsById.get(d.sheetId)?.name ?? '',
    }));
  }

  /**
   * Change the formula of an existing name, keeping the sheet its
   * unqualified references point to (even one that no longer exists).
   *
   * @returns Whether the name exists
   */
  setNameFormula(name: string, formula: string): boolean {
    const def = this._names.get(name.toUpperCase());
    if (!def) return false;
    def.formula = formula.startsWith('=') ? formula : '=' + formula;
    return true;
  }

  /** Apply a rewrite to the formula of every defined name, in place. */
  mapNameFormulas(rewrite: (formula: string) => string): void {
    for (const def of this._names.values()) def.formula = rewrite(def.formula);
  }

  /**
   * Replace all defined names. A name without `sheet` is based on the active
   * sheet. One whose `sheet` is not a sheet of the workbook (as `getNames()`
   * reports, with `''`, a name whose sheet was removed) keeps having no
   * sheet: its references without a sheet name stay `#REF!`.
   */
  setNames(names: Array<{ name: string; formula: string; sheet?: string }>): void {
    this._names.clear();
    for (const n of names) {
      if (n.sheet === undefined || this.hasSheet(n.sheet)) {
        this.defineName(n.name, n.formula, n.sheet);
        continue;
      }
      if (!isValidName(n.name)) throw new Error(`Invalid name: ${n.name}`);
      this._names.set(n.name.toUpperCase(), {
        name: n.name,
        formula: n.formula.startsWith('=') ? n.formula : '=' + n.formula,
        sheetId: NO_SHEET.id,
      });
    }
  }

  // ─── Evaluation API ─────────────────────────────────

  /**
   * The unformatted result of a cell's raw value: for a formula, the result
   * cached by the last evaluation of that cell (`recalculate()`,
   * `recalculateAffected()` or `evaluate(rawValue, key)`), when it was
   * computed from the same formula text; for a cell showing a spilled value,
   * that value. Falls back to evaluating `rawValue` without recording
   * dependencies (cells not evaluated yet).
   *
   * Prefer this over `evaluate(rawValue)` for cells that were just
   * recalculated: a standalone evaluation re-walks every precedent, which
   * makes formatting a long formula chain quadratic.
   *
   * @param key - Cell key in `"row:col"` format
   * @param rawValue - The cell's raw value
   * @param sheet - Sheet name (default: the active sheet)
   * @returns The display value and resolved type
   */
  resultFor(key: string, rawValue: string, sheet?: string): EvalResult {
    const s = this._sheet(sheet);
    const { row, col } = parseKey(key);
    const id = cellId(s, row, col);
    if (rawValue.startsWith('=')) {
      // Also for a volatile formula: what the last recalculation gave is what its
      // dependents were computed from, and asking for it must not roll again
      const cached = this._results.get(id);
      if (cached && cached.raw === rawValue) return this._resultToDisplay(cached);
      return this.evaluate(rawValue, undefined, sheet, { row, col });
    }
    if (rawValue === '') {
      const anchor = this._ghosts.get(id);
      if (anchor !== undefined) {
        const full = this._results.get(anchor)?.full;
        if (full) return this._toDisplay(full.get(row - idRow(anchor), col - idCol(anchor)));
      }
    }
    return this.evaluate(rawValue);
  }

  /**
   * Evaluate a raw value. If it starts with `=`, parse and execute the formula.
   * Otherwise, return the value as-is (possibly coerced to number/boolean).
   * For a formula whose result is an array of several cells, this is the
   * top-left value; use `evaluateValue()` for the whole array.
   *
   * @param rawValue - The raw cell value (formula or literal)
   * @param forCellKey - If provided, registers dependencies for this cell key
   *   (side-effect: updates the dep graph and result cache)
   * @param sheet - Sheet the formula is on (default: the active sheet)
   * @param at - Cell the formula is evaluated as, for ROW() and COLUMN(), when no key is given
   * @returns The display value and resolved type
   */
  evaluate(rawValue: string, forCellKey?: string, sheet?: string, at?: { row: number; col: number }): EvalResult {
    if (!rawValue || rawValue.trim() === '' || !rawValue.startsWith('=')) {
      if (forCellKey) {
        const s = this._sheet(sheet);
        const { row, col } = parseKey(forCellKey);
        const id = cellId(s, row, col);
        this._clearDeps(id);
        this._results.delete(id);
      }
      if (!rawValue || rawValue.trim() === '') return { displayValue: '', type: 'text' };
      return this.coerceValue(rawValue);
    }
    return this._resultToDisplay(this._evaluateRaw(rawValue, forCellKey, sheet, at));
  }

  /**
   * Evaluate a formula and return its value rather than its display text: a
   * number, string or boolean (a blank result is `0`, as in a cell), a
   * `RangeValue` for an array result, or an error code string such as `'#DIV/0!'`.
   * Nothing is cached or recorded.
   *
   * @param formula - The formula, starting with `=`
   * @param sheet - Sheet the formula is on (default: the active sheet)
   * @param at - Cell the formula is evaluated as, for ROW() and COLUMN()
   */
  evaluateValue(formula: string, sheet?: string, at?: { row: number; col: number }): unknown {
    const result = this._evaluateRaw(formula.startsWith('=') ? formula : '=' + formula, undefined, sheet, at);
    // An array whose first element is an error is still the whole array
    if (result.full) return result.full;
    return result.error !== undefined ? result.error : result.value;
  }

  /**
   * Evaluate a formula against the sheet as the last recalculation left it:
   * formula cells it refers to give the results they have (a volatile one is
   * not rolled again), instead of being computed afresh as `evaluate()` does.
   * For formulas that test what the grid shows, such as conditional formats.
   * Only cells that have no result yet are computed. Nothing is cached or recorded.
   *
   * @param rawValue - The formula, starting with `=` (other text is returned as a literal)
   * @param sheet - Sheet the formula is on (default: the active sheet)
   * @param at - Cell the formula is evaluated as, for ROW() and COLUMN()
   */
  evaluateCommitted(rawValue: string, sheet?: string, at?: { row: number; col: number }): EvalResult {
    if (!rawValue || rawValue.trim() === '') return { displayValue: '', type: 'text' };
    if (!rawValue.startsWith('=')) return this.coerceValue(rawValue);
    const s = this._sheet(sheet);
    const result = this._inPass(
      'pure',
      () => this._deferring(null, () => this._computeExpression(rawValue, s, at?.row ?? 0, at?.col ?? 0)),
      true
    );
    return this._resultToDisplay(result);
  }

  private _evaluateRaw(
    rawValue: string,
    forCellKey: string | undefined,
    sheet: string | undefined,
    at: { row: number; col: number } | undefined
  ): CellResult {
    const s = this._sheet(sheet);
    if (!forCellKey) {
      return this._inPass('pure', () =>
        this._deferring(null, () => this._computeExpression(rawValue, s, at?.row ?? 0, at?.col ?? 0))
      );
    }

    const { row, col } = parseKey(forCellKey);
    const id = cellId(s, row, col);
    const cached = this._results.get(id);
    if (cached && cached.raw === rawValue && !cached.volatile) return cached;

    // Keep the cell marked in-progress across deferred retries so a long chain
    // leading back to it is reported as circular rather than reading its old value.
    return this._inPass('root', () => {
      this._clearDeps(id);
      const previous = this._tracking;
      this._tracking = id;
      try {
        const result = this._deferring(id, () => this._computeCell(id, s, row, col, rawValue));
        this._results.set(id, result);
        // A volatile formula registered this way is recomputed by every recalculation too
        if (result.volatile) this._volatileCells.add(id);
        else this._volatileCells.delete(id);
        return result;
      } finally {
        this._tracking = previous;
      }
    });
  }

  /**
   * Re-evaluate all formula cells in every sheet.
   * Clears and rebuilds the entire dependency graph and result cache.
   *
   * @returns Keys of the active sheet's cells whose display value or type changed
   *   (see `changes` for all sheets)
   */
  recalculate(): Set<string> {
    this._reset();
    return this._inPass('commit', (pass) => {
      const stale = new Set<number>();
      // Spilled values are recomputed from scratch. Formulas that were spilling
      // go first, so of two overlapping spills the same one keeps its range.
      for (const sheet of this._sheets) {
        sheet.used = null;
        for (const [key, cell] of sheet.data) {
          if (cell.spilledFrom !== undefined || (cell.rawValue === '' && (cell.spill || cell.array))) {
            this._releaseGhostCell(sheet, key, cell, pass);
          } else if (cell.spill && cell.rawValue.startsWith('=')) {
            const { row, col } = parseKey(key);
            stale.add(cellId(sheet, row, col));
          }
        }
      }
      for (const sheet of this._sheets) {
        for (const [key, cell] of sheet.data) {
          if (cell.rawValue.startsWith('=')) {
            const { row, col } = parseKey(key);
            stale.add(cellId(sheet, row, col));
          }
        }
      }
      this._runCommit(pass, stale);
      return this._finishCommit(pass);
    });
  }

  /**
   * Recalculate only formulas affected by the given changed cell keys.
   * Uses BFS traversal of the reverse dependency graph for targeted recalculation.
   * Falls back to full `recalculate()` if the dep graph is empty.
   *
   * @param changedKeys - Cell keys whose raw values changed
   * @param sheet - Sheet the keys are on (default: the active sheet)
   * @returns Keys of that sheet's cells whose display value or type changed
   *   (see `changes` for all sheets)
   */
  recalculateAffected(changedKeys: string[], sheet?: string): Set<string> {
    const s = this._sheet(sheet);
    // If the dependency graph is empty but data exists, fall back to a full
    // recalculate so we never silently skip dependents after a setData() call.
    if (this._precedents.size === 0 && this._results.size === 0 && this._sheets.some((sh) => sh.data.size > 0)) {
      const changed = this.recalculate();
      return sheet === undefined || s === this._active ? changed : this.changes.get(s.name) ?? new Set();
    }

    // Spills on any sheet can change with this, and with them how far whole columns reach
    for (const other of this._sheets) other.used = null;
    return this._inPass('commit', (pass) => {
      const roots: number[] = [];
      for (const key of changedKeys) {
        const { row, col } = parseKey(key);
        const id = cellId(s, row, col);
        roots.push(id);
        // A value written into a spill range blocks the spill; one removed from
        // a blocked range lets it through
        const anchor = this._ghosts.get(id);
        if (anchor !== undefined) roots.push(anchor);
        for (const [blockedId, size] of this._blocked) {
          if (idSheet(blockedId) !== s.id) continue;
          const r0 = idRow(blockedId);
          const c0 = idCol(blockedId);
          if (row >= r0 && row < r0 + size.rows && col >= c0 && col < c0 + size.cols) roots.push(blockedId);
        }
      }
      for (const id of this._volatileCells) roots.push(id);
      // The change may have moved the end of the used part of the workbook
      for (const id of this._movedWholeReaders(pass)) roots.push(id);
      this._runCommit(pass, this._closure(roots));
      return this._finishCommit(pass, s);
    });
  }

  /**
   * The cells a reference or a defined name stands for: `"B5"`, `"A1:C3"`,
   * `"Sheet2!A:A"`, `"sales"`. Null when the text is neither, or the name
   * does not refer to cells (a constant, a LAMBDA).
   *
   * @param text - A reference or name, with or without a leading `=`
   * @param sheet - Sheet that a reference without a sheet name points to (default: the active sheet)
   */
  referenceOf(
    text: string,
    sheet?: string
  ): { sheet: string; start: { row: number; col: number }; end: { row: number; col: number } } | null {
    const s = this._sheet(sheet);
    const formula = text.startsWith('=') ? text : '=' + text;
    return this._inPass('pure', () => {
      try {
        const ast = this._parse(formula);
        // Only something that names cells: no arithmetic, no function calls
        const names = (node: Node): boolean =>
          node.t === 'ref' || node.t === 'name' || (node.t === 'bin' && node.op === ':' && names(node.l) && names(node.r));
        if (!names(ast)) return null;
        const ref = this._eval(ast, { sheet: s, row: 0, col: 0, scope: null });
        if (!(ref instanceof RefValue)) return null;
        return {
          sheet: ref.sheet.name,
          start: { row: ref.top, col: ref.left },
          end: { row: ref.bottom, col: ref.right },
        };
      } catch {
        return null;
      }
    });
  }

  /** The size of the range a cell's formula spills over, or null when it doesn't spill. */
  spillRange(key: string, sheet?: string): { rows: number; cols: number } | null {
    const s = this._sheet(sheet);
    const { row, col } = parseKey(key);
    return this._spills.get(cellId(s, row, col)) ?? null;
  }

  /**
   * For a formula that shows `#SPILL!`: the size of the range its result
   * needs. Null when the cell's result is not a blocked spill.
   */
  blockedSpill(key: string, sheet?: string): { rows: number; cols: number } | null {
    const s = this._sheet(sheet);
    const { row, col } = parseKey(key);
    return this._blocked.get(cellId(s, row, col)) ?? null;
  }

  // ─── Passes ─────────────────────────────────────────

  private _inPass<T>(mode: Pass['mode'], fn: (pass: Pass) => T, committed = false): T {
    if (this._pass) {
      // Reentrant call (e.g. a custom function evaluating something): share the pass
      return fn(this._pass);
    }
    const pass: Pass = {
      mode,
      committed,
      memo: new Map(),
      ranges: new Map(),
      stale: null,
      touched: new Set(),
      changed: new Map(),
      seq: 0,
      started: new Map(),
      shapeChanges: [],
      claimed: new Map(),
      growths: new Map(),
      selfFeeding: new Set(),
      bounded: new Set(),
      untrimmed: null,
      pureSpills: [],
    };
    this._pass = pass;
    try {
      return fn(pass);
    } finally {
      this._pass = null;
    }
  }

  /** Compute every stale cell, then again whatever read a range that a spill grew into or left. */
  private _runCommit(pass: Pass, stale: Set<number>): void {
    for (let round = 0; round < MAX_SPILL_ROUNDS && stale.size > 0; round++) {
      pass.stale = stale;
      pass.ranges.clear();
      for (const id of stale) pass.memo.delete(id);
      pass.shapeChanges = [];

      for (const id of stale) {
        if (pass.memo.has(id)) continue;
        const sheet = this._sheetsById.get(idSheet(id));
        if (!sheet) continue;
        const row = idRow(id);
        const col = idCol(id);
        const cell = sheet.data.get(cellKey(row, col));
        if (!cell || !cell.rawValue.startsWith('=')) {
          // No longer a formula: drop what it left behind
          this._clearDeps(id);
          this._results.delete(id);
          this._volatileCells.delete(id);
          this._materialize(id, sheet, row, col, { raw: '' }, pass);
          if (cell && (cell.spill || cell.array) && cell.spilledFrom === undefined) {
            // A value written over the formula in place: it no longer spills or holds an array
            delete cell.spill;
            delete cell.array;
            this._markChanged(pass, sheet, cellKey(row, col));
          }
          continue;
        }
        this._deferring(id, () => this._computeCell(id, sheet, row, col, cell.rawValue));
      }

      // Cells that read a cell before a spill started (or stopped) covering it saw the wrong value
      const again: number[] = [];
      for (const change of pass.shapeChanges) {
        this._eachDependent(change.id, (dep) => {
          const started = pass.started.get(dep);
          if (started === undefined || started < change.seq) again.push(dep);
        });
      }
      // A spill that left a cell may have made room for one that was blocked there
      if (this._blocked.size > 0) {
        for (const [blockedId, size] of this._blocked) {
          const sid = idSheet(blockedId);
          const top = idRow(blockedId);
          const left = idCol(blockedId);
          for (const change of pass.shapeChanges) {
            if (!change.released || idSheet(change.id) !== sid) continue;
            const row = idRow(change.id);
            const col = idCol(change.id);
            if (row >= top && row < top + size.rows && col >= left && col < left + size.cols) {
              again.push(blockedId);
              break;
            }
          }
        }
      }
      // Spills written in this round may have moved the end of the used part of the workbook
      for (const id of this._movedWholeReaders(pass)) again.push(id);
      stale = again.length ? this._closure(again) : new Set();
    }
    pass.stale = null;
  }

  /** The whole-column readers whose extent is no longer the one they were computed with. */
  private _movedWholeReaders(pass: Pass): number[] {
    const moved: number[] = [];
    for (const [id, read] of this._wholeReaders) {
      const now = this._usedExtent(pass, id);
      if (now.rows !== read.rows || now.cols !== read.cols) moved.push(id);
    }
    return moved;
  }

  /** Write the display values of everything computed, and report what changed. */
  private _finishCommit(pass: Pass, primary: EngineSheet = this._active): Set<string> {
    for (const id of pass.touched) {
      const sheet = this._sheetsById.get(idSheet(id));
      const result = pass.memo.get(id);
      if (!sheet || !result) continue;
      const key = cellKey(idRow(id), idCol(id));
      const cell = sheet.data.get(key);
      if (!cell || cell.rawValue !== result.raw) continue;
      const display = this._resultToDisplay(result);
      const spill = this._spills.get(id);
      let changed = false;
      if (cell.displayValue !== display.displayValue || cell.type !== display.type) {
        cell.displayValue = display.displayValue;
        cell.type = display.type;
        changed = true;
      }
      if (!cellArraysEqual(cell.array, display.array)) {
        if (display.array) cell.array = display.array;
        else delete cell.array;
        changed = true;
      }
      if (cell.spill?.rows !== spill?.rows || cell.spill?.cols !== spill?.cols) {
        if (spill) cell.spill = { ...spill };
        else delete cell.spill;
        changed = true;
      }
      if (cell.spilledFrom !== undefined) {
        delete cell.spilledFrom;
        changed = true;
      }
      if (changed) this._markChanged(pass, sheet, key);
    }

    this.changes = new Map();
    for (const [sheet, keys] of pass.changed) this.changes.set(sheet.name, keys);
    return pass.changed.get(primary) ?? new Set();
  }

  private _markChanged(pass: Pass, sheet: EngineSheet, key: string): void {
    let keys = pass.changed.get(sheet);
    if (!keys) pass.changed.set(sheet, (keys = new Set()));
    keys.add(key);
  }

  /**
   * Run `compute`, which evaluates the cell `rootId` (or a standalone
   * expression when null). Cells whose evaluation hits the depth limit are
   * computed on their own first (innermost last), then `compute` is retried.
   */
  private _deferring(rootId: number | null, compute: () => CellResult): CellResult {
    const pass = this._pass!;
    const ownsRoot = rootId !== null && !this._evaluating.has(rootId);
    if (ownsRoot) this._evaluating.add(rootId!);
    try {
      const deferred: number[] = [];
      for (;;) {
        const target = deferred.length > 0 ? deferred[deferred.length - 1] : null;
        try {
          if (target === null) return compute();
          const sheet = this._sheetsById.get(idSheet(target));
          const cell = sheet?.data.get(cellKey(idRow(target), idCol(target)));
          if (sheet && cell?.rawValue.startsWith('=') && !pass.memo.has(target)) {
            this._computeCell(target, sheet, idRow(target), idCol(target), cell.rawValue);
          }
          deferred.pop();
        } catch (e) {
          if (!(e instanceof EvalDepthError)) throw e;
          const cycleStart = deferred.indexOf(e.id);
          if (cycleStart !== -1) {
            // deferred[cycleStart..] form a (long) cycle. Memoize the verdict
            // for its members so later evaluations in this pass don't walk the
            // cycle again (quadratic otherwise), then retry the cells that
            // merely depend on it so they can handle the error (e.g. IFERROR).
            for (const id of deferred.slice(cycleStart)) {
              const sheet = this._sheetsById.get(idSheet(id));
              const raw = sheet?.data.get(cellKey(idRow(id), idCol(id)))?.rawValue ?? '';
              this._memoize(id, { raw, error: '#CIRC!' });
            }
            deferred.length = cycleStart;
            continue;
          }
          if (deferred.length >= MAX_DEFERRED_CELLS) {
            return { raw: '', error: '#ERROR!' };
          }
          deferred.push(e.id);
        }
      }
    } finally {
      if (ownsRoot) this._evaluating.delete(rootId!);
    }
  }

  private _memoize(id: number, result: CellResult): void {
    const pass = this._pass!;
    pass.memo.set(id, result);
    if (pass.mode === 'commit') {
      this._results.set(id, result);
      if (result.volatile) this._volatileCells.add(id);
      else this._volatileCells.delete(id);
      pass.touched.add(id);
    }
  }

  /** Evaluate an expression that belongs to no cell. */
  private _computeExpression(rawValue: string, sheet: EngineSheet, row: number, col: number): CellResult {
    try {
      const ast = this._parse(rawValue);
      return this._makeResult(rawValue, this._value(ast, { sheet, row, col, scope: null }));
    } catch (e) {
      if (e instanceof EvalDepthError) throw e;
      return { raw: rawValue, error: toErrorCode(e) };
    }
  }

  /**
   * Evaluate a formula cell and memoize the result. Throws only
   * EvalDepthError; every other failure becomes an error result.
   */
  private _computeCell(id: number, sheet: EngineSheet, row: number, col: number, raw: string): CellResult {
    const pass = this._pass!;
    const commit = pass.mode === 'commit';
    const owns = !this._evaluating.has(id);
    if (owns) this._evaluating.add(id);
    const previousTracking = this._tracking;
    if (commit) {
      this._clearDeps(id);
      this._tracking = id;
      pass.started.set(id, ++pass.seq);
    }
    this._evalDepth++;
    const outerCurrent = this._current;
    this._current = id;
    const outerVolatile = this._usedVolatile;
    this._usedVolatile = false;
    let result: CellResult;
    // Known before evaluating, so a volatile formula that fails is still volatile
    let volatile = false;
    try {
      const ast = this._parse(raw);
      volatile = this._isVolatile(raw, ast);
      result = this._makeResult(raw, this._value(ast, { sheet, row, col, scope: null }));
      if (commit && result.full && this._wholeReaders.has(id) && this._growsWithoutEnd(id, sheet, row, col, ast, result.full, pass)) {
        result = { raw, error: '#SPILL!' };
      }
      if (volatile || this._usedVolatile) result.volatile = true;
    } catch (e) {
      if (e instanceof EvalDepthError) throw e;
      // Preserve specific error codes (#DIV/0!, #NAME?, #CIRC!)
      result = { raw, error: toErrorCode(e) };
      if (volatile || this._usedVolatile) result.volatile = true;
    } finally {
      this._usedVolatile = outerVolatile;
      this._current = outerCurrent;
      this._evalDepth--;
      this._tracking = previousTracking;
      if (owns) this._evaluating.delete(id);
    }
    if (commit) {
      this._materialize(id, sheet, row, col, result, pass);
    } else if (result.full) {
      // Nothing is written in this pass, but the result is what the sheet would show
      if (this._spillBlocked(id, sheet, row, col, result.full)) {
        result = { raw, error: '#SPILL!', volatile: result.volatile };
      } else {
        pass.pureSpills.push({ anchor: id, rows: result.full.rows, cols: result.full.cols });
      }
    }
    this._memoize(id, result);
    return result;
  }

  /** The result of the formula cell `id`, computing it if this pass has not yet. */
  private _pull(id: number, sheet: EngineSheet, row: number, col: number, raw: string): CellResult {
    const pass = this._pass!;
    const memoized = pass.memo.get(id);
    if (memoized) return memoized;
    if (this._evaluating.has(id)) throw new Error('#CIRC!');
    if (pass.committed) {
      // What the last recalculation gave, as the sheet shows it
      const cached = this._results.get(id);
      if (cached && cached.raw === raw) {
        pass.memo.set(id, cached);
        return cached;
      }
    } else if (pass.mode === 'commit' && !pass.stale?.has(id)) {
      const cached = this._results.get(id);
      if (cached && cached.raw === raw && !cached.volatile) {
        pass.memo.set(id, cached);
        return cached;
      }
    }
    if (this._evalDepth >= MAX_EVAL_DEPTH) throw new EvalDepthError(id);
    return this._computeCell(id, sheet, row, col, raw);
  }

  private _parse(raw: string): Node {
    let ast = this._asts.get(raw);
    if (ast === undefined) {
      try {
        ast = parseFormula(raw.substring(1));
      } catch (e) {
        ast = e instanceof Error ? e : new Error('#ERROR!');
      }
      if (this._asts.size >= AST_CACHE_LIMIT) this._asts.clear();
      this._asts.set(raw, ast);
    }
    if (ast instanceof Error) throw ast;
    return ast;
  }

  private _volatility: Map<string, boolean> = new Map();
  /** Set while a cell is evaluated, when a defined name it uses is volatile */
  private _usedVolatile = false;

  private _isVolatile(raw: string, ast: Node): boolean {
    let v = this._volatility.get(raw);
    if (v === undefined) {
      v = callsAny(ast, this._volatileNames);
      if (this._volatility.size >= AST_CACHE_LIMIT) this._volatility.clear();
      this._volatility.set(raw, v);
    }
    return v;
  }

  /** Turn the value of a formula into the cell's result: its own value, plus the array to spill. */
  private _makeResult(raw: string, value: unknown): CellResult {
    if (value instanceof LambdaValue) return { raw, error: '#CALC!' };
    if (value instanceof RangeValue) {
      if (value.rows * value.cols === 0) return { raw, error: '#CALC!' };
      const first = value.values[0];
      if (value.rows * value.cols === 1) {
        // A 1×1 array is its element; an element that is an array stays in the cell
        return isErrorCode(first) ? { raw, error: first } : { raw, value: first ?? 0 };
      }
      // A blank in a result is 0, for the cell that shows it and for whatever reads that cell
      const full = value.values.includes(undefined)
        ? new RangeValue(value.values.map((v) => v ?? 0), value.rows, value.cols)
        : value;
      const top = full.values[0];
      return isErrorCode(top) ? { raw, error: top, full } : { raw, value: top, full };
    }
    if (isErrorCode(value)) return { raw, error: value };
    return { raw, value: value ?? 0 };
  }

  // ─── Spilling ───────────────────────────────────────

  /**
   * Make the sheet show what the formula in `id` spills: write the spilled
   * values into the cells of its range, and clear the cells it no longer covers.
   * A spill that doesn't fit turns the result into #SPILL!.
   */
  private _materialize(id: number, sheet: EngineSheet, row: number, col: number, result: CellResult, pass: Pass): void {
    const previous = this._spills.get(id);
    let full = result.full;

    if (full) {
      if (this._spillBlocked(id, sheet, row, col, full)) {
        this._blocked.set(id, { rows: full.rows, cols: full.cols });
        result.error = '#SPILL!';
        result.value = undefined;
        result.full = undefined;
        full = undefined;
      } else {
        this._blocked.delete(id);
      }
    } else {
      this._blocked.delete(id);
    }

    if (!full && !previous) return;
    const anchorKey = cellKey(row, col);

    // Leave the cells the spill no longer covers
    if (previous) {
      for (let r = 0; r < previous.rows; r++) {
        for (let c = 0; c < previous.cols; c++) {
          if (r === 0 && c === 0) continue;
          if (full && r < full.rows && c < full.cols) continue;
          const ghost = cellId(sheet, row + r, col + c);
          if (this._ghosts.get(ghost) !== id) continue;
          this._ghosts.delete(ghost);
          const key = cellKey(row + r, col + c);
          const cell = sheet.data.get(key);
          if (cell) this._releaseGhostCell(sheet, key, cell, pass);
          pass.shapeChanges.push({ id: ghost, seq: ++pass.seq, released: true });
        }
      }
    }

    // How far the spills reach is worked out again when asked for
    this._reach = null;
    if (!full) {
      this._spills.delete(id);
      pass.ranges.clear();
      return;
    }

    for (let r = 0; r < full.rows; r++) {
      for (let c = 0; c < full.cols; c++) {
        if (r === 0 && c === 0) continue;
        const ghost = cellId(sheet, row + r, col + c);
        const key = cellKey(row + r, col + c);
        if (this._ghosts.get(ghost) !== id) {
          this._ghosts.set(ghost, id);
          pass.claimed.set(ghost, id);
          pass.shapeChanges.push({ id: ghost, seq: ++pass.seq });
        }
        const display = this._toDisplay(full.get(r, c));
        const cell = sheet.data.get(key);
        if (!cell) {
          const created: CellData = {
            rawValue: '',
            displayValue: display.displayValue,
            type: display.type,
            spilledFrom: anchorKey,
          };
          if (display.array) created.array = display.array;
          sheet.data.set(key, created);
          this._markChanged(pass, sheet, key);
        } else if (
          cell.spilledFrom !== anchorKey ||
          cell.displayValue !== display.displayValue ||
          cell.type !== display.type ||
          !cellArraysEqual(cell.array, display.array)
        ) {
          cell.spilledFrom = anchorKey;
          cell.displayValue = display.displayValue;
          cell.type = display.type;
          if (display.array) cell.array = display.array;
          else delete cell.array;
          this._markChanged(pass, sheet, key);
        }
      }
    }
    this._spills.set(id, { rows: full.rows, cols: full.cols });
    pass.ranges.clear();
  }

  /**
   * Whether the spill of a formula that reads a whole column or row can
   * never settle: the spill moves the end of what the formula reads, so each
   * recomputation makes it larger (`=A:A*2` below row 1).
   *
   * A spill that has grown a few times in this pass is put to the test: the
   * formula is evaluated reading its whole columns and rows in full. If even
   * that result fits on the sheet, the growth has an end (`=TAKE(A:A, 10)`,
   * or a formula that merely follows another one's growth) and is left to
   * settle. If it doesn't fit, nothing will.
   */
  private _growsWithoutEnd(
    id: number,
    sheet: EngineSheet,
    row: number,
    col: number,
    ast: Node,
    full: RangeValue,
    pass: Pass
  ): boolean {
    if (pass.selfFeeding.has(id)) return true;
    if (pass.bounded.has(id)) return false;
    const shown = this._spills.get(id);
    if (!shown || (full.rows <= shown.rows && full.cols <= shown.cols)) return false;
    const growths = (pass.growths.get(id) ?? 0) + 1;
    pass.growths.set(id, growths);
    if (growths < MAX_SPILL_GROWTHS) return false;

    let fits = true;
    const outer = pass.untrimmed;
    pass.untrimmed = id;
    try {
      const largest = this._value(ast, { sheet, row, col, scope: null });
      if (largest instanceof RangeValue) {
        fits = row + largest.rows <= sheet.rows && col + largest.cols <= sheet.cols;
      }
    } catch (e) {
      if (e instanceof EvalDepthError) throw e;
    } finally {
      pass.untrimmed = outer;
    }
    (fits ? pass.bounded : pass.selfFeeding).add(id);
    return !fits;
  }

  /** Whether an array can't spill from a cell: it would leave the sheet, or a value or another spill is in the way. */
  private _spillBlocked(id: number, sheet: EngineSheet, row: number, col: number, full: RangeValue): boolean {
    if (row + full.rows > sheet.rows || col + full.cols > sheet.cols) return true;
    for (let r = 0; r < full.rows; r++) {
      for (let c = 0; c < full.cols; c++) {
        if (r === 0 && c === 0) continue;
        const owner = this._ghosts.get(cellId(sheet, row + r, col + c));
        if (owner !== undefined && owner !== id) return true;
        const cell = sheet.data.get(cellKey(row + r, col + c));
        if (cell && cell.rawValue !== '') return true;
      }
    }
    return false;
  }

  /** Stop a cell from showing a spilled value: empty it, or remove it when it holds nothing else. */
  private _releaseGhostCell(sheet: EngineSheet, key: string, cell: CellData, pass: Pass): void {
    delete cell.spilledFrom;
    if (cell.rawValue !== '') return;
    delete cell.array;
    delete cell.spill;
    if (cell.format || cell.readOnly) {
      cell.displayValue = '';
      cell.type = 'text';
    } else {
      sheet.data.delete(key);
    }
    this._markChanged(pass, sheet, key);
  }

  // ─── Dependency Tracking ─────────────────────────────

  private _clearDeps(id: number): void {
    this._wholeReaders.delete(id);
    const prec = this._precedents.get(id);
    if (!prec) return;
    for (const dep of prec.cells) {
      const set = this._cellDependents.get(dep);
      if (set) {
        set.delete(id);
        if (set.size === 0) this._cellDependents.delete(dep);
      }
    }
    for (const node of prec.nodes) {
      node.dependents.delete(id);
      if (node.dependents.size === 0) this._rangeNodes.get(node.sheet.id)?.delete(node.key);
    }
    this._precedents.delete(id);
  }

  private _precedentsOf(id: number): Precedents {
    let prec = this._precedents.get(id);
    if (!prec) this._precedents.set(id, (prec = { cells: new Set(), nodes: new Set() }));
    return prec;
  }

  private _trackCell(id: number): void {
    const tracking = this._tracking;
    if (tracking === null) return;
    const prec = this._precedentsOf(tracking);
    if (prec.cells.has(id)) return;
    prec.cells.add(id);
    let rev = this._cellDependents.get(id);
    if (!rev) this._cellDependents.set(id, (rev = new Set()));
    rev.add(tracking);
  }

  private _trackRange(sheet: EngineSheet, top: number, left: number, bottom: number, right: number): void {
    const tracking = this._tracking;
    if (tracking === null) return;
    if ((bottom - top + 1) * (right - left + 1) <= SMALL_RANGE) {
      for (let r = top; r <= bottom; r++) {
        for (let c = left; c <= right; c++) this._trackCell(cellId(sheet, r, c));
      }
      return;
    }
    const key = `${top}:${left}:${bottom}:${right}`;
    let nodes = this._rangeNodes.get(sheet.id);
    if (!nodes) this._rangeNodes.set(sheet.id, (nodes = new Map()));
    let node = nodes.get(key);
    if (!node) nodes.set(key, (node = { key, sheet, top, left, bottom, right, dependents: new Set() }));
    node.dependents.add(tracking);
    this._precedentsOf(tracking).nodes.add(node);
  }

  /** Visit the formula cells that read cell `id` directly. */
  private _eachDependent(id: number, visit: (dep: number) => void): void {
    const direct = this._cellDependents.get(id);
    if (direct) for (const dep of direct) visit(dep);
    const nodes = this._rangeNodes.get(idSheet(id));
    if (nodes && nodes.size > 0) {
      const row = idRow(id);
      const col = idCol(id);
      for (const node of nodes.values()) {
        if (row >= node.top && row <= node.bottom && col >= node.left && col <= node.right) {
          for (const dep of node.dependents) visit(dep);
        }
      }
    }
  }

  /** The given cells plus everything that depends on them, directly or not (including through spilled values). */
  private _closure(roots: number[]): Set<number> {
    const seen = new Set<number>(roots);
    const queue = [...roots];
    const visit = (dep: number) => {
      if (!seen.has(dep)) {
        seen.add(dep);
        queue.push(dep);
      }
    };
    // Walk with an index: queue.shift() is O(n), quadratic for large batches
    for (let head = 0; head < queue.length; head++) {
      const id = queue[head];
      this._eachDependent(id, visit);
      // Whatever reads a cell this one spills into depends on it too
      const spill = this._spills.get(id);
      if (spill) {
        const sid = idSheet(id);
        const row = idRow(id);
        const col = idCol(id);
        const base = id - col;
        for (let r = 0; r < spill.rows; r++) {
          for (let c = 0; c < spill.cols; c++) {
            if (r === 0 && c === 0) continue;
            const direct = this._cellDependents.get(base + r * MAX_COLS + col + c);
            if (direct) for (const dep of direct) visit(dep);
          }
        }
        const nodes = this._rangeNodes.get(sid);
        if (nodes) {
          const bottom = row + spill.rows - 1;
          const right = col + spill.cols - 1;
          for (const node of nodes.values()) {
            if (node.top <= bottom && node.bottom >= row && node.left <= right && node.right >= col) {
              for (const dep of node.dependents) visit(dep);
            }
          }
        }
      }
    }
    return seen;
  }

  // ─── Value Coercion ──────────────────────────────────

  private coerceValue(val: string): EvalResult {
    if (ERROR_CODES.has(val)) {
      return { displayValue: val, type: 'error' };
    }

    const upper = val.toUpperCase();
    if (upper === 'TRUE' || upper === 'FALSE') {
      return { displayValue: upper, type: 'boolean' };
    }

    const num = Number(val);
    if (val.trim() !== '' && !isNaN(num)) {
      const clean = Number.isFinite(num) ? parseFloat(num.toPrecision(15)) : num;
      return { displayValue: String(clean), type: 'number' };
    }

    return { displayValue: val, type: 'text' };
  }

  private _resultToDisplay(result: CellResult): EvalResult {
    if (result.error !== undefined) return { displayValue: result.error, type: 'error' };
    return this._toDisplay(result.value);
  }

  /**
   * Convert a value to its display string and type based on its JS type, so
   * text results like "007" or TEXT() output stay text.
   */
  private _toDisplay(value: unknown): EvalResult {
    if (typeof value === 'number') {
      if (isNaN(value)) return { displayValue: '#VALUE!', type: 'error' };
      if (!Number.isFinite(value)) return { displayValue: '#NUM!', type: 'error' };
      return { displayValue: numberToText(value), type: 'number' };
    }
    if (typeof value === 'boolean') {
      return { displayValue: value ? 'TRUE' : 'FALSE', type: 'boolean' };
    }
    if (typeof value === 'string') {
      return isErrorCode(value) ? { displayValue: value, type: 'error' } : { displayValue: value, type: 'text' };
    }
    if (value instanceof RangeValue) {
      return { displayValue: arrayToText(value), type: 'text', array: toCellArray(value) };
    }
    if (value === undefined || value === null) {
      return { displayValue: '0', type: 'number' };
    }
    if (value instanceof RefValue || value instanceof LambdaValue) {
      return { displayValue: '#VALUE!', type: 'error' };
    }
    return this.coerceValue(String(value));
  }

  // ─── Reading cells ──────────────────────────────────

  /**
   * The value of one cell: a scalar, an array kept in the cell, or undefined
   * when it is blank. Throws the cell's error.
   */
  private _readCell(sheet: EngineSheet, row: number, col: number): unknown {
    const id = cellId(sheet, row, col);
    this._trackCell(id);
    // Circular reference detection
    if (this._evaluating.has(id)) throw new Error('#CIRC!');

    const cell = sheet.data.get(cellKey(row, col));
    const raw = cell?.rawValue ?? '';
    if (raw === '') return this._readGhost(id, row, col);
    if (raw.charCodeAt(0) === 61 /* = */) {
      const result = this._pull(id, sheet, row, col, raw);
      if (result.error !== undefined) throw new Error(result.error);
      return result.value;
    }
    const value = literalValue(raw);
    if (isErrorCode(value)) throw new Error(value);
    return value;
  }

  /** The value an empty cell shows because another cell's formula spills into it. */
  private _readGhost(id: number, row: number, col: number): unknown {
    const pass = this._pass;
    if (pass && pass.pureSpills.length > 0) {
      // A spill computed in this pass that the sheet doesn't show (yet)
      for (const spill of pass.pureSpills) {
        if (idSheet(spill.anchor) !== idSheet(id)) continue;
        const r = row - idRow(spill.anchor);
        const c = col - idCol(spill.anchor);
        if (r < 0 || c < 0 || r >= spill.rows || c >= spill.cols || (r === 0 && c === 0)) continue;
        const value = pass.memo.get(spill.anchor)?.full?.get(r, c);
        if (isErrorCode(value)) throw new Error(value);
        return value;
      }
    }
    const anchor = this._ghosts.get(id);
    if (anchor === undefined) {
      // A formula that reads a cell of its own (earlier) spill range is circular
      const claimedBy = pass?.claimed.get(id);
      if (claimedBy !== undefined && this._evaluating.has(claimedBy)) throw new Error('#CIRC!');
      return undefined;
    }
    const sheet = this._sheetsById.get(idSheet(anchor));
    const anchorRow = idRow(anchor);
    const anchorCol = idCol(anchor);
    const cell = sheet?.data.get(cellKey(anchorRow, anchorCol));
    if (!sheet || !cell || !cell.rawValue.startsWith('=')) return undefined;
    this._trackCell(anchor);
    const full = this._pull(anchor, sheet, anchorRow, anchorCol, cell.rawValue).full;
    const r = row - anchorRow;
    const c = col - anchorCol;
    if (!full || r >= full.rows || c >= full.cols) return undefined;
    const value = full.get(r, c);
    if (isErrorCode(value)) throw new Error(value);
    return value;
  }

  /** The extent of the cells of a sheet that hold a value of their own (spilled values are counted separately). */
  private _usedRange(sheet: EngineSheet): { rows: number; cols: number } {
    if (!sheet.used) {
      let rows = 0;
      let cols = 0;
      for (const [key, cell] of sheet.data) {
        if (cell.rawValue === '') continue;
        const sep = key.indexOf(':');
        const row = +key.slice(0, sep);
        const col = +key.slice(sep + 1);
        if (row >= rows) rows = row + 1;
        if (col >= cols) cols = col + 1;
      }
      sheet.used = { rows, cols };
    }
    return sheet.used;
  }

  /**
   * How far the spills on the sheets reach, down and to the right: the
   * furthest and the second furthest, so one formula's own spill can be left out.
   */
  private _spillReach(): NonNullable<FormulaEngine['_reach']> {
    if (!this._reach) {
      const reach = { rows: 0, rowsBy: -1, rowsNext: 0, cols: 0, colsBy: -1, colsNext: 0 };
      for (const [anchor, spill] of this._spills) {
        const rows = idRow(anchor) + spill.rows;
        const cols = idCol(anchor) + spill.cols;
        if (rows > reach.rows) {
          reach.rowsNext = reach.rows;
          reach.rows = rows;
          reach.rowsBy = anchor;
        } else if (rows > reach.rowsNext) {
          reach.rowsNext = rows;
        }
        if (cols > reach.cols) {
          reach.colsNext = reach.cols;
          reach.cols = cols;
          reach.colsBy = anchor;
        } else if (cols > reach.colsNext) {
          reach.colsNext = cols;
        }
      }
      this._reach = reach;
    }
    return this._reach;
  }

  /** The values of a reference to several cells. Errors become error-code values. */
  private _readRange(ref: RefValue): RangeValue {
    const { sheet, top, left } = ref;
    let { bottom, right } = ref;
    this._trackRange(sheet, top, left, bottom, right);
    const pass = this._pass!;
    if (ref.whole) {
      if (this._tracking !== null && pass.mode === 'commit') {
        this._wholeReaders.set(this._tracking, this._usedExtent(pass, this._current));
      }
      // Seeing how large the result can get: every row and column there is
      if (pass.untrimmed !== null && pass.untrimmed === this._tracking) {
        return this._readRect(sheet, top, left, bottom, right, pass);
      }
      // Whole columns/rows end where the sheet's content ends. Outside a
      // recalculation, a formula read on the way can turn out to spill further
      // than the sheet shows: read again until the extent holds still.
      for (let attempt = 0; ; attempt++) {
        const extent = this._usedExtent(pass, this._current);
        const b = ref.whole === 'cols' ? Math.max(top, Math.min(bottom, extent.rows - 1)) : bottom;
        const r = ref.whole === 'rows' ? Math.max(left, Math.min(right, extent.cols - 1)) : right;
        const range = this._readRect(sheet, top, left, b, r, pass);
        const after = this._usedExtent(pass, this._current);
        if (attempt >= 8 || (after.rows === extent.rows && after.cols === extent.cols)) return range;
      }
    }
    return this._readRect(sheet, top, left, bottom, right, pass);
  }

  /**
   * How far whole-column and whole-row references are read: the used part
   * of the workbook's sheets taken together (values and spills), so `A:A` and
   * `Other!A:A` have the same number of rows and can be paired up (SUMIFS,
   * XLOOKUP, ...). Includes spills computed in this pass that the sheets
   * don't show.
   *
   * @param reader - The formula cell that is reading. Its own spill doesn't
   *   count: otherwise `=A:A*2` below row 1 would lengthen what it reads with
   *   every result, and a spill would never get shorter again.
   */
  private _usedExtent(pass: Pass, reader: number | null = null): { rows: number; cols: number } {
    let rows = 0;
    let cols = 0;
    for (const sheet of this._sheets) {
      const used = this._usedRange(sheet);
      rows = Math.max(rows, used.rows);
      cols = Math.max(cols, used.cols);
    }
    const reach = this._spillReach();
    rows = Math.max(rows, reader !== null && reach.rowsBy === reader ? reach.rowsNext : reach.rows);
    cols = Math.max(cols, reader !== null && reach.colsBy === reader ? reach.colsNext : reach.cols);
    for (const spill of pass.pureSpills) {
      if (spill.anchor === reader) continue;
      rows = Math.max(rows, idRow(spill.anchor) + spill.rows);
      cols = Math.max(cols, idCol(spill.anchor) + spill.cols);
    }
    return { rows, cols };
  }

  private _readRect(sheet: EngineSheet, top: number, left: number, bottom: number, right: number, pass: Pass): RangeValue {
    const cacheKey = `${sheet.id}|${top}:${left}:${bottom}:${right}`;
    const cached = pass.ranges.get(cacheKey);
    if (cached) return cached;

    const numRows = bottom - top + 1;
    const numCols = right - left + 1;
    // Filled, not sparse: array methods skip holes, and a blank cell is a value (undefined)
    const values: unknown[] = new Array(numRows * numCols).fill(undefined);
    // A cell that is being evaluated reads as #CIRC! here, which is only true for this reader
    let cacheable = true;
    let i = 0;
    for (let r = top; r <= bottom; r++) {
      for (let c = left; c <= right; c++, i++) {
        const cell = sheet.data.get(cellKey(r, c));
        const raw = cell?.rawValue ?? '';
        if (raw === '') {
          // A formatted-but-empty cell is still empty, unless a spill shows in it
          // (checked per cell: reading an earlier cell of the range can start a spill)
          if (this._ghosts.size > 0 || pass.claimed.size > 0 || pass.pureSpills.length > 0) {
            try {
              values[i] = this._readGhost(cellId(sheet, r, c), r, c);
            } catch (e) {
              if (e instanceof EvalDepthError) throw e;
              values[i] = toErrorCode(e);
              if (values[i] === '#CIRC!') cacheable = false;
            }
          }
        } else if (raw.charCodeAt(0) === 61 /* = */) {
          // Errors become error-code values so functions like COUNTIF can
          // skip them while aggregates like SUM propagate them.
          const id = cellId(sheet, r, c);
          if (this._evaluating.has(id)) {
            values[i] = '#CIRC!';
            cacheable = false;
          } else {
            const result = this._pull(id, sheet, r, c, raw);
            values[i] = result.error !== undefined ? result.error : result.value;
          }
        } else {
          values[i] = literalValue(raw);
        }
      }
    }
    const range = new RangeValue(values, numRows, numCols);
    if (cacheable) pass.ranges.set(cacheKey, range);
    return range;
  }

  /** The value(s) a reference points to: one cell's value, or an array for several cells. */
  private _deref(value: unknown): unknown {
    if (!(value instanceof RefValue)) return value;
    if (value.isCell) return this._readCell(value.sheet, value.top, value.left);
    return this._readRange(value);
  }

  private _makeRef(
    sheet: EngineSheet,
    top: number,
    left: number,
    bottom: number,
    right: number,
    whole: RefValue['whole'] = null,
    area = false
  ): RefValue {
    if (top < 0 || left < 0 || bottom >= MAX_ROWS || right >= MAX_COLS || top > bottom || left > right) {
      throw new Error('#REF!');
    }
    return new RefValue(sheet, top, left, bottom, right, whole, area);
  }

  private _refFromInfo(info: RefInfo, ctx: EvalCtx): RefValue {
    let sheet = ctx.sheet;
    if (info.sheet !== undefined) {
      const named = this._sheetsByName.get(info.sheet.toLowerCase());
      if (!named) throw new Error('#REF!');
      sheet = named;
    } else if (sheet === NO_SHEET) {
      // A defined name whose sheet was removed
      throw new Error('#REF!');
    }
    const { start, end } = info;
    switch (info.kind) {
      case 'cell':
        return this._makeRef(sheet, start.row, start.col, start.row, start.col);
      case 'range':
        return this._makeRef(
          sheet,
          Math.min(start.row, end.row),
          Math.min(start.col, end.col),
          Math.max(start.row, end.row),
          Math.max(start.col, end.col),
          null,
          true
        );
      case 'cols':
        return this._makeRef(sheet, 0, Math.min(start.col, end.col), sheet.rows - 1, Math.max(start.col, end.col), 'cols', true);
      case 'rows':
        return this._makeRef(sheet, Math.min(start.row, end.row), 0, Math.max(start.row, end.row), sheet.cols - 1, 'rows', true);
    }
  }

  // ─── Evaluator ──────────────────────────────────────

  /** Evaluate a node and resolve a reference result to its values. */
  private _value(node: Node, ctx: EvalCtx): unknown {
    return this._deref(this._eval(node, ctx));
  }

  /** Evaluate a node. The result can be a reference or a lambda; errors are thrown. */
  private _eval(node: Node, ctx: EvalCtx): unknown {
    switch (node.t) {
      case 'num':
      case 'str':
      case 'bool':
        return node.v;
      case 'err':
        throw new Error(node.v);
      case 'ref':
        return this._refFromInfo(node.ref, ctx);
      case 'name':
        return this._lookupName(node.name, ctx);
      case 'unary':
        return this._unary(node, ctx);
      case 'bin': {
        if (node.op === ':') return this._rangeOf(this._eval(node.l, ctx), this._eval(node.r, ctx));
        const left = this._value(node.l, ctx);
        const right = this._value(node.r, ctx);
        return this._binary(node.op, left, right);
      }
      case 'call':
        return this._call(node.name, node.args, ctx);
      case 'invoke': {
        const callee = this._eval(node.callee, ctx);
        if (!(callee instanceof LambdaValue)) throw new Error('#VALUE!');
        return callee.call(node.args.map((a) => (a === null ? undefined : this._bindable(a, ctx))));
      }
      case 'array': {
        const values: unknown[] = [];
        for (const row of node.rows) {
          for (const element of row) values.push(this._elementValue(element, ctx));
        }
        return new RangeValue(values, node.rows.length, node.rows[0].length);
      }
    }
  }

  /** The value of an expression as an element of an array: errors become codes. */
  private _elementValue(node: Node, ctx: EvalCtx): unknown {
    try {
      const value = this._value(node, ctx);
      if (value instanceof LambdaValue) return '#CALC!';
      return value;
    } catch (e) {
      if (e instanceof EvalDepthError) throw e;
      return toErrorCode(e);
    }
  }

  private _lookupName(name: string, ctx: EvalCtx): unknown {
    const upper = name.toUpperCase();
    for (let scope = ctx.scope; scope; scope = scope.parent) {
      if (scope.name === upper) {
        // An error bound to the name surfaces where the name is used, so IFERROR(x, ...) can handle it
        if (scope.value instanceof BoundError) throw new Error(scope.value.code);
        return scope.value;
      }
    }
    const def = this._names.get(upper);
    if (!def) throw new Error('#NAME?');
    // The sheet only says where references without a sheet name point: a name
    // whose sheet is gone still works as long as it has no such reference
    const sheet = this._sheetsById.get(def.sheetId) ?? NO_SHEET;
    if (this._nameDepth >= MAX_EVAL_DEPTH) throw new Error('#CIRC!');
    this._nameDepth++;
    try {
      const ast = this._parse(def.formula);
      // A formula is as volatile as the names it uses
      if (this._isVolatile(def.formula, ast)) this._usedVolatile = true;
      return this._eval(ast, { sheet, row: ctx.row, col: ctx.col, scope: null });
    } finally {
      this._nameDepth--;
    }
  }

  private _rangeOf(left: unknown, right: unknown): RefValue {
    if (!(left instanceof RefValue) || !(right instanceof RefValue)) throw new Error('#VALUE!');
    if (left.sheet !== right.sheet) throw new Error('#VALUE!');
    return this._makeRef(
      left.sheet,
      Math.min(left.top, right.top),
      Math.min(left.left, right.left),
      Math.max(left.bottom, right.bottom),
      Math.max(left.right, right.right),
      left.whole && left.whole === right.whole ? left.whole : null,
      true
    );
  }

  private _unary(node: Extract<Node, { t: 'unary' }>, ctx: EvalCtx): unknown {
    switch (node.op) {
      case '-':
        return mapValue(this._value(node.x, ctx), (v) => -toNumber(v));
      case '+':
        return mapValue(this._value(node.x, ctx), (v) => toNumber(v));
      case '%':
        return mapValue(this._value(node.x, ctx), (v) => toNumber(v) / 100);
      case '@': {
        // Implicit intersection: the cell of the range on the formula's row or column
        let target = this._eval(node.x, ctx);
        // One cell: its value, or the first of the values it holds
        if (target instanceof RefValue && target.isCell) target = this._deref(target);
        if (target instanceof RefValue) {
          if (target.cols === 1 && ctx.row >= target.top && ctx.row <= target.bottom) {
            return this._makeRef(target.sheet, ctx.row, target.left, ctx.row, target.left);
          }
          if (target.rows === 1 && ctx.col >= target.left && ctx.col <= target.right) {
            return this._makeRef(target.sheet, target.top, ctx.col, target.top, ctx.col);
          }
          return this._makeRef(target.sheet, target.top, target.left, target.top, target.left);
        }
        if (target instanceof RangeValue) {
          const first = target.values[0];
          if (isErrorCode(first)) throw new Error(first);
          return first;
        }
        return target;
      }
      case '#': {
        // Spill operator: the whole range the formula in that cell spills over
        const anchor = this._eval(node.x, ctx);
        if (!(anchor instanceof RefValue) || !anchor.isCell) throw new Error('#REF!');
        const { sheet, top, left } = anchor;
        const id = cellId(sheet, top, left);
        this._trackCell(id);
        if (this._evaluating.has(id)) throw new Error('#CIRC!');
        const cell = sheet.data.get(cellKey(top, left));
        if (!cell || !cell.rawValue.startsWith('=')) throw new Error('#REF!');
        const result = this._pull(id, sheet, top, left, cell.rawValue);
        // An error in the first cell is one element's error, not the whole range's
        if (!result.full) {
          if (result.error !== undefined) throw new Error(result.error);
          return anchor;
        }
        return this._makeRef(sheet, top, left, top + result.full.rows - 1, left + result.full.cols - 1, null, true);
      }
    }
  }

  private _binary(op: string, left: unknown, right: unknown): unknown {
    if (left instanceof LambdaValue || right instanceof LambdaValue) throw new Error('#VALUE!');
    if (left instanceof RangeValue || right instanceof RangeValue) {
      return broadcast([left, right], ([a, b]) => this._binary(op, a, b));
    }
    return this._scalarBinary(op, left, right);
  }

  private _scalarBinary(op: string, left: unknown, right: unknown): unknown {
    switch (op) {
      case '+':
        return toNumber(left) + toNumber(right);
      case '-':
        return toNumber(left) - toNumber(right);
      case '*':
        return toNumber(left) * toNumber(right);
      case '/': {
        const dividend = toNumber(left);
        const divisor = toNumber(right);
        if (divisor === 0) throw new Error('#DIV/0!');
        return dividend / divisor;
      }
      case '^':
        return power(toNumber(left), toNumber(right));
      case '&':
        if (isErrorCode(left)) throw new Error(left);
        if (isErrorCode(right)) throw new Error(right);
        return toText(left) + toText(right);
      default:
        if (COMPARISONS.has(op)) return this._compareValues(left, right, op);
        throw new Error('#ERROR!');
    }
  }

  private _compareValues(left: unknown, right: unknown, op: string): boolean {
    if (isErrorCode(left)) throw new Error(left);
    if (isErrorCode(right)) throw new Error(right);
    // A blank cell compares as 0, as "" and as FALSE
    if (left === undefined || left === null) left = typeof right === 'string' ? '' : typeof right === 'boolean' ? false : 0;
    if (right === undefined || right === null) right = typeof left === 'string' ? '' : typeof left === 'boolean' ? false : 0;

    // Determine if both sides can be compared numerically.
    // Booleans are numeric (TRUE=1, FALSE=0). Strings that look like numbers
    // are numeric. Empty strings and non-numeric strings are NOT numeric.
    const isNumeric = (v: unknown): boolean => {
      if (typeof v === 'boolean') return true;
      if (typeof v === 'number') return true;
      const s = String(v).trim();
      return s !== '' && !isNaN(Number(s));
    };

    if (isNumeric(left) && isNumeric(right)) {
      const l = Number(left);
      const r = Number(right);
      switch (op) {
        case '=':  return l === r;
        case '<>': return l !== r;
        case '<':  return l < r;
        case '>':  return l > r;
        case '<=': return l <= r;
        case '>=': return l >= r;
        default:   return false;
      }
    }

    // Text sorts before booleans, whatever it says
    if ((typeof left === 'boolean') !== (typeof right === 'boolean')) {
      const less = typeof right === 'boolean';
      switch (op) {
        case '=':  return false;
        case '<>': return true;
        case '<':
        case '<=': return less;
        default:   return !less;
      }
    }

    // String comparison: = and <> are case-insensitive (Excel convention),
    // relational operators (<, >, <=, >=) use locale-unaware ordinal comparison.
    const ls = toText(left);
    const rs = toText(right);
    switch (op) {
      case '=':  return ls.toLowerCase() === rs.toLowerCase();
      case '<>': return ls.toLowerCase() !== rs.toLowerCase();
      case '<':  return ls < rs;
      case '>':  return ls > rs;
      case '<=': return ls <= rs;
      case '>=': return ls >= rs;
      default:   return false;
    }
  }

  // ─── Function calls ─────────────────────────────────

  private _call(name: string, argNodes: Array<Node | null>, ctx: EvalCtx): unknown {
    // A LAMBDA bound by LET or passed as a parameter
    for (let scope = ctx.scope; scope; scope = scope.parent) {
      if (scope.name === name) return this._invoke(scope.value, argNodes, ctx);
    }

    const entry = this._functions.get(name);
    const special = this._special[name];
    if (special && !entry?.custom) return special(argNodes, ctx);

    if (!entry) {
      // A defined name that holds a LAMBDA
      if (this._names.has(name)) return this._invoke(this._lookupName(name, ctx), argNodes, ctx);
      throw new Error('#NAME?');
    }
    if (entry.arity && (argNodes.length < entry.arity[0] || argNodes.length > entry.arity[1])) {
      throw new Error('#VALUE!');
    }

    if (entry.agg) return this._aggregate(entry, argNodes, ctx);

    // Evaluate the arguments the way the function wants them
    const args: unknown[] = new Array(argNodes.length);
    const kinds: ArgKind[] = new Array(argNodes.length);
    let lift = false;
    for (let i = 0; i < argNodes.length; i++) {
      const kind =
        i < entry.kinds.length ? entry.kinds[i] : entry.repeat[(i - entry.kinds.length) % entry.repeat.length];
      kinds[i] = kind;
      const node = argNodes[i];
      if (node === null) {
        // An omitted argument is undefined, which functions treat as "use the default"
        args[i] = undefined;
        kinds[i] = 'a';
        continue;
      }
      let value: unknown;
      if (kind === 'r') {
        value = this._eval(node, ctx);
      } else if (kind === 'v' || kind === 'x') {
        try {
          value = this._value(node, ctx);
        } catch (e) {
          if (e instanceof EvalDepthError) throw e;
          value = kind === 'x' ? new ArgumentError(toErrorCode(e)) : toErrorCode(e);
        }
      } else {
        value = this._value(node, ctx);
      }
      if (value instanceof LambdaValue && kind !== 'a') throw new Error('#VALUE!');
      if (value === undefined) {
        // An empty cell reads as 0 (or "" for text)
        if (kind === 'n' || kind === 'a' || kind === 'x') value = 0;
        else if (kind === 't') value = '';
      } else if (value instanceof RangeValue && (kind === 'n' || kind === 't' || kind === 'v')) {
        lift = true;
      }
      args[i] = value;
    }

    const fn = entry.fn!;
    const fctx = this._functionContext(ctx);
    if (!lift) {
      // An empty cell returned by e.g. VLOOKUP reads as 0, matching how it displays
      return fromJsArray(fn(fctx, ...args) ?? 0);
    }

    // Call once per element of the array arguments. Elements that are arrays
    // themselves (nested arrays) are gone into the same way.
    const single = (i: number) => kinds[i] === 'n' || kinds[i] === 't' || kinds[i] === 'v';
    const liftCall = (values: unknown[]): unknown => {
      if (!values.some((v, i) => single(i) && v instanceof RangeValue)) {
        const callArgs = values.map((v, i) =>
          v === undefined && single(i) ? (kinds[i] === 'n' ? 0 : kinds[i] === 't' ? '' : undefined) : v
        );
        const result = fromJsArray(fn(fctx, ...callArgs) ?? 0);
        return result instanceof RefValue ? this._deref(result) : result;
      }
      const wrapped = values.map((v, i) => (!single(i) ? new Scalar(v) : kinds[i] === 'v' ? new Tolerant(v) : v));
      return broadcast(wrapped, (elements) => liftCall(elements.map((e) => (e instanceof Scalar ? e.value : e))));
    };
    return liftCall(args);
  }

  private _aggregate(entry: FunctionEntry, argNodes: Array<Node | null>, ctx: EvalCtx): unknown {
    const flat: unknown[] = [];
    for (const node of argNodes) {
      if (node === null) continue;
      const value = this._value(node, ctx);
      if (value instanceof RangeValue) {
        // Empty cells are skipped
        for (const v of flattenDeep(value)) {
          if (v !== undefined) flat.push(v);
        }
      } else if (value instanceof LambdaValue) {
        throw new Error('#VALUE!');
      } else if (value !== undefined) {
        flat.push(value);
      }
    }
    if (entry.propagate) {
      for (const v of flat) {
        if (isErrorCode(v)) throw new Error(v);
      }
    }
    return entry.agg!(flat) ?? 0;
  }

  private _invoke(callee: unknown, argNodes: Array<Node | null>, ctx: EvalCtx): unknown {
    if (!(callee instanceof LambdaValue)) throw new Error('#VALUE!');
    return callee.call(argNodes.map((a) => (a === null ? undefined : this._bindable(a, ctx))));
  }

  /** The value an expression gives a LET name or LAMBDA parameter: an error is kept until the name is used. */
  private _bindable(node: Node, ctx: EvalCtx): unknown {
    try {
      return this._eval(node, ctx);
    } catch (e) {
      if (e instanceof EvalDepthError) throw e;
      return new BoundError(toErrorCode(e));
    }
  }

  private _makeLambda(argNodes: Array<Node | null>, ctx: EvalCtx): LambdaValue {
    if (argNodes.length === 0) throw new Error('#VALUE!');
    const body = argNodes[argNodes.length - 1];
    if (body === null) throw new Error('#VALUE!');
    const params = argNodes.slice(0, -1).map((p) => {
      if (!p || p.t !== 'name') throw new Error('#VALUE!');
      return p.name.toUpperCase();
    });
    const lambda: LambdaValue = new LambdaValue(params, body, ctx.scope, (args) => {
      if (args.length > params.length) throw new Error('#VALUE!');
      if (this._lambdaDepth >= MAX_LAMBDA_DEPTH) throw new Error('#NUM!');
      let scope = ctx.scope;
      params.forEach((name, i) => {
        scope = { name, value: args[i], parent: scope };
      });
      this._lambdaDepth++;
      try {
        return this._eval(body, { sheet: ctx.sheet, row: ctx.row, col: ctx.col, scope });
      } finally {
        this._lambdaDepth--;
      }
    });
    return lambda;
  }

  /** One-argument function that reads a reference range as its first cell and maps a computed array element by element. */
  private _firstCellFunction(args: Array<Node | null>, ctx: EvalCtx, fn: (v: unknown) => unknown): unknown {
    if (args.length !== 1 || args[0] === null) throw new Error('#VALUE!');
    const v = this._eval(args[0], ctx);
    if (v instanceof RefValue) {
      // Only the first cell is read (and depended on): the rest of the range may depend on this formula
      const first = this._deref(new RefValue(v.sheet, v.top, v.left, v.top, v.left));
      return fn(first instanceof RangeValue ? first.get(0, 0) : first);
    }
    // A computed array (an array constant, A1:A5>2) goes element by element
    if (v instanceof RangeValue) {
      return new RangeValue(
        v.values.map((x) => {
          try {
            return fn(x);
          } catch (e) {
            return toErrorCode(e);
          }
        }),
        v.rows,
        v.cols
      );
    }
    return fn(v);
  }

  /** Calls a registered value function with arguments that are already values (nodes are evaluated, an error in one is captured). */
  private _pureCall(name: string, parts: Array<unknown>, ctx: EvalCtx): unknown {
    const entry = this._functions.get(name)!;
    const values = parts.map((part) => {
      if (part instanceof Evaluated) return part.value ?? 0;
      if (part instanceof RangeValue) return part;
      if (part === null) return undefined;
      try {
        return this._value(part as Node, ctx) ?? 0; // a blank reads as 0
      } catch (e) {
        if (e instanceof EvalDepthError) throw e;
        return new ArgumentError(toErrorCode(e));
      }
    });
    return fromJsArray(entry.fn!(this._functionContext(ctx), ...values) ?? 0);
  }

  private _functionContext(ctx: EvalCtx): FormulaContext {
    if (ctx.fctx) return ctx.fctx;
    const sheet = ctx.sheet;
    const rawOf = (cell: CellData | undefined): unknown => {
      if (!cell) return undefined;
      const num = Number(cell.rawValue);
      return !isNaN(num) && cell.rawValue.trim() !== '' ? num : cell.rawValue;
    };
    ctx.fctx = {
      getCellValue: (ref: string) => {
        if (ref.includes(':') && !/[A-Z]/i.test(ref.charAt(0))) {
          // It's a key like "0:0"
          return rawOf(sheet.data.get(ref)) ?? 0;
        }
        const coord = refToCoord(ref.replace(/\$/g, ''));
        return this._inPass('pure', () => this._readCell(sheet, coord.row, coord.col)) ?? 0;
      },
      getRangeValues: (startRef: string, endRef: string) => {
        const start = refToCoord(startRef);
        const end = refToCoord(endRef);
        const values: unknown[] = [];
        for (let r = Math.min(start.row, end.row); r <= Math.max(start.row, end.row); r++) {
          for (let c = Math.min(start.col, end.col); c <= Math.max(start.col, end.col); c++) {
            const cell = sheet.data.get(cellKey(r, c));
            if (cell) values.push(rawOf(cell));
          }
        }
        return values;
      },
      cell: { row: ctx.row, col: ctx.col },
      sheet: sheet.name,
    };
    return ctx.fctx;
  }

  // ─── Special forms and reference functions ──────────

  /** Functions that evaluate their arguments themselves: lazily, as references, or as lambdas. */
  private _special: Record<string, (args: Array<Node | null>, ctx: EvalCtx) => unknown> = {
    // COUNTBLANK counts empty cells and "" results; a whole column or row is counted at the size of
    // the grid, not just up to the content, and an error is not a blank
    COUNTBLANK: (args, ctx) => {
      if (args.length !== 1 || args[0] === null) throw new Error('#VALUE!');
      const v = this._eval(args[0], ctx);
      const blank = (x: unknown) => x === undefined || x === '';
      if (!(v instanceof RefValue)) {
        return v instanceof RangeValue ? v.values.filter(blank).length : blank(v) ? 1 : 0;
      }
      let values: unknown;
      try {
        values = this._deref(v);
      } catch (e) {
        if (e instanceof EvalDepthError) throw e;
        return 0; // a cell holding an error is not blank
      }
      if (!(values instanceof RangeValue)) return blank(values) ? 1 : 0;
      let count = 0;
      for (const x of flattenDeep(values)) if (blank(x)) count++;
      return count + Math.max(0, v.rows * v.cols - values.rows * values.cols);
    },

    // CHOOSE and SWITCH return the chosen argument as it is, so a reference stays a reference
    // (OFFSET(CHOOSE(1, A1:A2, B1:B2), 0, 0)); an array index or expression goes element by element
    CHOOSE: (args, ctx) => {
      if (args.length < 2 || args[0] === null) throw new Error('#VALUE!');
      const index = this._value(args[0], ctx);
      if (index instanceof RangeValue) return this._pureCall('CHOOSE', [index, ...args.slice(1)], ctx);
      const i = Math.trunc(toNumber(index === undefined ? 0 : index));
      if (i < 1 || i >= args.length) throw new Error('#VALUE!');
      const chosen = args[i];
      return chosen === null ? 0 : this._eval(chosen, ctx);
    },
    SWITCH: (args, ctx) => {
      if (args.length < 3 || args[0] === null) throw new Error('#VALUE!');
      const value = this._value(args[0], ctx);
      if (value instanceof RangeValue) return this._pureCall('SWITCH', [new Evaluated(value), ...args.slice(1)], ctx);
      const hasDefault = (args.length - 1) % 2 === 1;
      const pairsEnd = args.length - (hasDefault ? 1 : 0);
      for (let i = 1; i < pairsEnd; i += 2) {
        const caseNode = args[i];
        const candidate = caseNode === null ? 0 : this._value(caseNode, ctx); // an omitted match is blank, i.e. 0
        if (candidate instanceof RangeValue) return this._pureCall('SWITCH', [new Evaluated(value), ...args.slice(1)], ctx);
        if (sameValue(value === undefined ? 0 : value, candidate === undefined ? 0 : candidate)) {
          const result = args[i + 1];
          return result === null ? 0 : this._eval(result, ctx);
        }
      }
      if (hasDefault) {
        const fallback = args[args.length - 1];
        return fallback === null ? 0 : this._eval(fallback, ctx);
      }
      throw new Error('#N/A');
    },

    // T(value) is the text itself or "" ; N(value) the number (TRUE is 1, anything else 0). Given a
    // reference range they read its first cell (not each cell); a computed array goes element by element
    T: (args, ctx) => this._firstCellFunction(args, ctx, (v) => {
      if (isErrorCode(v)) throw new Error(v);
      return typeof v === 'string' ? v : '';
    }),
    N: (args, ctx) => this._firstCellFunction(args, ctx, (v) => {
      if (isErrorCode(v)) throw new Error(v);
      if (typeof v === 'number') return v;
      return typeof v === 'boolean' ? (v ? 1 : 0) : 0;
    }),

    // ISREF sees the argument as a reference, and an invalid one (INDIRECT of a missing sheet,
    // 1/0) is just not a reference: FALSE, not an error
    ISREF: (args, ctx) => {
      if (args.length !== 1 || args[0] === null) throw new Error('#VALUE!');
      try {
        return this._eval(args[0], ctx) instanceof RefValue;
      } catch (e) {
        if (e instanceof EvalDepthError) throw e;
        return false;
      }
    },

    // IF only evaluates the branch it returns, so guards like
    // =IF(A1=0, 0, 1/A1) don't raise errors from the untaken branch.
    IF: (args, ctx) => {
      if (args.length === 0 || args[0] === null) throw new Error('#VALUE!');
      const condition = this._value(args[0], ctx);
      if (isErrorCode(condition)) throw new Error(condition);
      // Excel defaults: omitted true branch → 0, omitted false branch → FALSE
      if (!(condition instanceof RangeValue)) {
        if (condition) return args[1] ? this._eval(args[1], ctx) : 0;
        return args[2] ? this._eval(args[2], ctx) : false;
      }
      // An array condition picks element by element; both branches are evaluated
      const branch = (node: Node | null | undefined, fallback: unknown) =>
        node ? this._elementValue(node, ctx) : fallback;
      const whenTrue = branch(args[1], 0);
      const whenFalse = branch(args[2], false);
      return broadcast([condition, new Tolerant(whenTrue), new Tolerant(whenFalse)], ([c, t, f]) => {
        const picked = toBoolean(c) ? t : f;
        return picked === undefined ? 0 : picked;
      });
    },

    // IFS evaluates conditions in order until one is true; only that
    // condition's value is evaluated. No true condition → #N/A.
    IFS: (args, ctx) => {
      if (args.length === 0 || args.length % 2 !== 0) throw new Error('#VALUE!');
      for (let i = 0; i < args.length; i += 2) {
        const condNode = args[i];
        const valueNode = args[i + 1];
        if (!condNode) throw new Error('#VALUE!');
        const condition = this._value(condNode, ctx);
        if (condition instanceof RangeValue) {
          // An array condition: from here on, every element takes the value of its first true condition
          const rest: unknown[] = [new Tolerant(condition), new Tolerant(valueNode ? this._elementValue(valueNode, ctx) : 0)];
          for (let j = i + 2; j < args.length; j += 2) {
            const c = args[j];
            const v = args[j + 1];
            if (!c) throw new Error('#VALUE!');
            rest.push(new Tolerant(this._elementValue(c, ctx)), new Tolerant(v ? this._elementValue(v, ctx) : 0));
          }
          return broadcast(
            rest,
            (elements) => {
              for (let k = 0; k < elements.length; k += 2) {
                const c = elements[k];
                if (isErrorCode(c)) throw new Error(c);
                if (toBoolean(c)) {
                  const v = elements[k + 1];
                  if (isErrorCode(v)) throw new Error(v);
                  return v ?? 0;
                }
              }
              throw new Error('#N/A');
            },
            false
          );
        }
        if (isErrorCode(condition)) throw new Error(condition);
        // An omitted value is 0
        if (condition) return valueNode ? this._eval(valueNode, ctx) : 0;
      }
      throw new Error('#N/A');
    },

    // IFERROR traps errors thrown by its first argument. The fallback is
    // evaluated only when it is used, so an error inside an unused fallback
    // (e.g. a #REF! left by a deleted row) can't surface.
    IFERROR: (args, ctx) => this._ifError(args, ctx, (code) => isErrorCode(code)),
    IFNA: (args, ctx) => this._ifError(args, ctx, (code) => code === '#N/A'),

    LET: (args, ctx) => {
      if (args.length < 3 || args.length % 2 === 0) throw new Error('#VALUE!');
      let scope = ctx.scope;
      for (let i = 0; i < args.length - 1; i += 2) {
        const nameNode = args[i];
        const valueNode = args[i + 1];
        if (!nameNode || nameNode.t !== 'name' || !valueNode) throw new Error('#VALUE!');
        const value = this._bindable(valueNode, { ...ctx, scope, fctx: undefined });
        scope = { name: nameNode.name.toUpperCase(), value, parent: scope };
      }
      const body = args[args.length - 1];
      if (!body) throw new Error('#VALUE!');
      return this._eval(body, { ...ctx, scope, fctx: undefined });
    },

    LAMBDA: (args, ctx) => this._makeLambda(args, ctx),

    ROW: (args, ctx) => this._rowOrColumn(args, ctx, true),
    COLUMN: (args, ctx) => this._rowOrColumn(args, ctx, false),

    ROWS: (args, ctx) => this._shapeOf(args, ctx).rows,
    COLUMNS: (args, ctx) => this._shapeOf(args, ctx).cols,

    // INDEX(array, row, [column]): given a reference it returns a reference,
    // so A1:INDEX(A:A, 5) is a range.
    INDEX: (args, ctx) => {
      if (args.length < 2 || args.length > 3 || !args[0]) throw new Error('#VALUE!');
      let source = this._eval(args[0], ctx);
      const num = (node: Node | null | undefined): number | undefined => {
        if (!node) return undefined;
        const v = this._value(node, ctx);
        if (v instanceof RangeValue) throw new Error('#VALUE!');
        return Math.trunc(toNumber(v));
      };
      let rowNum = num(args[1]);
      let colNum = num(args[2]);
      // A cell that holds several values is indexed into
      if (source instanceof RefValue && source.isCell) {
        try {
          const held = this._deref(source);
          if (held instanceof RangeValue) source = held;
        } catch (e) {
          // The cell's error only matters if its value is used; as a reference it is fine
          if (e instanceof EvalDepthError) throw e;
        }
      }
      const array = source instanceof RefValue ? source : asArray(source);
      // A single index into a row (or column) picks along it
      if (args.length === 2 && array.rows === 1 && array.cols > 1) {
        colNum = rowNum;
        rowNum = 1;
      }
      const r = rowNum ?? 0;
      const c = colNum ?? (array.cols === 1 ? 1 : args.length === 2 ? 1 : 0);
      if (r < 0 || c < 0 || r > array.rows || c > array.cols) throw new Error('#REF!');
      if (source instanceof RefValue) {
        const top = r === 0 ? source.top : source.top + r - 1;
        const bottom = r === 0 ? source.bottom : top;
        const left = c === 0 ? source.left : source.left + c - 1;
        const right = c === 0 ? source.right : left;
        // All rows of a whole column are still a whole column (and are read like one)
        const whole =
          (r === 0 && source.whole === 'cols') || (c === 0 && source.whole === 'rows') ? source.whole : null;
        return this._makeRef(source.sheet, top, left, bottom, right, whole, whole !== null);
      }
      const arr = array as RangeValue;
      if (r === 0 && c === 0) return arr;
      if (r === 0) return RangeValue.column(arr.getCol(c - 1));
      if (c === 0) return RangeValue.row(arr.getRow(r - 1));
      const value = arr.get(r - 1, c - 1);
      if (isErrorCode(value)) throw new Error(value);
      return value;
    },

    // OFFSET(reference, rows, cols, [height], [width])
    OFFSET: (args, ctx) => {
      if (args.length < 3 || args.length > 5 || !args[0]) throw new Error('#VALUE!');
      const base = this._eval(args[0], ctx);
      if (!(base instanceof RefValue)) throw new Error('#VALUE!');
      const num = (node: Node | null | undefined, fallback: number): number => {
        if (!node) return fallback;
        const v = this._value(node, ctx);
        if (v instanceof RangeValue) throw new Error('#VALUE!');
        return Math.trunc(toNumber(v));
      };
      const top = base.top + num(args[1], 0);
      const left = base.left + num(args[2], 0);
      const height = num(args[3], base.rows);
      const width = num(args[4], base.cols);
      if (height === 0 || width === 0) throw new Error('#REF!');
      const bottom = top + height + (height > 0 ? -1 : 1);
      const right = left + width + (width > 0 ? -1 : 1);
      // A whole column moved sideways (or a whole row moved up or down) is still whole
      const whole =
        (base.whole === 'cols' && top === base.top && height === base.rows) ||
        (base.whole === 'rows' && left === base.left && width === base.cols)
          ? base.whole
          : null;
      return this._makeRef(
        base.sheet,
        Math.min(top, bottom),
        Math.min(left, right),
        Math.max(top, bottom),
        Math.max(left, right),
        whole,
        whole !== null
      );
    },

    // INDIRECT(text): the reference the text spells out
    INDIRECT: (args, ctx) => {
      if (args.length < 1 || args.length > 2 || !args[0]) throw new Error('#VALUE!');
      const text = this._value(args[0], ctx);
      if (text instanceof RangeValue || text === undefined) throw new Error('#REF!');
      let ast: Node;
      try {
        ast = this._parse('=' + toText(text));
      } catch {
        throw new Error('#REF!');
      }
      if (ast.t !== 'ref' && ast.t !== 'name' && !(ast.t === 'bin' && ast.op === ':')) throw new Error('#REF!');
      let ref: unknown;
      try {
        ref = this._eval(ast, { sheet: ctx.sheet, row: ctx.row, col: ctx.col, scope: null });
      } catch (e) {
        if (e instanceof EvalDepthError) throw e;
        throw new Error('#REF!');
      }
      if (!(ref instanceof RefValue)) throw new Error('#REF!');
      return ref;
    },

    // MAP(array1, [array2, ...], lambda)
    MAP: (args, ctx) => {
      if (args.length < 2) throw new Error('#VALUE!');
      const lambda = this._lambdaArg(args[args.length - 1], ctx);
      const arrays = args.slice(0, -1).map((a) => (a ? this._value(a, ctx) : undefined));
      return broadcast(arrays.map((a) => new Tolerant(a)), (elements) => this._deref(lambda.call(elements)), false);
    },

    // REDUCE(initial, array, lambda(accumulator, value))
    REDUCE: (args, ctx) => {
      if (args.length !== 3) throw new Error('#VALUE!');
      const lambda = this._lambdaArg(args[2], ctx);
      let acc = args[0] ? this._value(args[0], ctx) : undefined;
      const array = args[1] ? this._value(args[1], ctx) : undefined;
      for (const v of asArray(array).values) acc = this._deref(lambda.call([acc, v]));
      return acc;
    },

    BYROW: (args, ctx) => this._byVector(args, ctx, true),
    BYCOL: (args, ctx) => this._byVector(args, ctx, false),
  };

  private _lambdaArg(node: Node | null, ctx: EvalCtx): LambdaValue {
    const value = node ? this._eval(node, ctx) : undefined;
    if (!(value instanceof LambdaValue)) throw new Error('#VALUE!');
    return value;
  }

  private _byVector(args: Array<Node | null>, ctx: EvalCtx, byRow: boolean): unknown {
    if (args.length !== 2 || !args[0]) throw new Error('#VALUE!');
    const lambda = this._lambdaArg(args[1], ctx);
    // Over a reference every row/column handed to the lambda is itself a reference (so N(row) reads
    // its first cell and ISREF(row) is TRUE, as in Excel). The source is not read up front: the
    // whole range is registered as a dependency, and only what the lambda reads is evaluated.
    const first = this._eval(args[0], ctx);
    const ref = first instanceof RefValue && !first.whole ? first : undefined;
    let array: RangeValue | undefined;
    let count: number;
    if (ref) {
      this._trackRange(ref.sheet, ref.top, ref.left, ref.bottom, ref.right);
      count = byRow ? ref.rows : ref.cols;
    } else {
      array = asArray(this._deref(first));
      count = byRow ? array.rows : array.cols;
    }
    const source = array && (byRow ? array : transpose(array));
    const results: unknown[] = [];
    for (let r = 0; r < count; r++) {
      const vector: unknown = ref
        ? byRow
          ? new RefValue(ref.sheet, ref.top + r, ref.left, ref.top + r, ref.right, null, true)
          : new RefValue(ref.sheet, ref.top, ref.left + r, ref.bottom, ref.left + r, null, true)
        : byRow
          ? RangeValue.row(source!.getRow(r))
          : RangeValue.column(source!.getRow(r));
      try {
        let value = this._deref(lambda.call([vector]));
        // One value per row: a 1×1 array is that value
        if (value instanceof RangeValue && value.values.length === 1) value = value.values[0];
        results.push(value instanceof LambdaValue ? '#CALC!' : value);
      } catch (e) {
        if (e instanceof EvalDepthError) throw e;
        results.push(toErrorCode(e));
      }
    }
    return byRow ? RangeValue.column(results) : RangeValue.row(results);
  }

  private _ifError(args: Array<Node | null>, ctx: EvalCtx, handles: (code: string) => boolean): unknown {
    if (args.length < 1 || args.length > 2) throw new Error('#VALUE!');
    // An omitted value is 0
    if (!args[0]) return 0;
    const fallback = () => (args[1] ? this._eval(args[1], ctx) : '');
    let value: unknown;
    try {
      value = this._eval(args[0], ctx);
      // A reference to one cell: reading it is what can fail
      if (value instanceof RefValue && value.isCell) value = this._deref(value);
    } catch (e) {
      if (e instanceof EvalDepthError) throw e;
      if (!handles(toErrorCode(e))) throw e;
      return fallback();
    }
    if (value instanceof RefValue) value = this._deref(value);
    if (isErrorCode(value)) return handles(value) ? fallback() : value;
    if (value instanceof RangeValue && value.values.some((v) => isErrorCode(v) && handles(v))) {
      // A fallback that fails only fails the elements that needed it
      let replacement: unknown;
      try {
        replacement = this._deref(fallback());
      } catch (e) {
        if (e instanceof EvalDepthError) throw e;
        replacement = toErrorCode(e);
      }
      return broadcast([value, new Tolerant(replacement)], ([v, f]) => (isErrorCode(v) && handles(v) ? f : v), false);
    }
    return value;
  }

  private _rowOrColumn(args: Array<Node | null>, ctx: EvalCtx, row: boolean): unknown {
    if (args.length > 1) throw new Error('#VALUE!');
    if (args.length === 0 || !args[0]) return (row ? ctx.row : ctx.col) + 1;
    const ref = this._eval(args[0], ctx);
    if (!(ref instanceof RefValue)) throw new Error('#VALUE!');
    if (ref.isCell) return (row ? ref.top : ref.left) + 1;
    if (row) return RangeValue.column(Array.from({ length: ref.rows }, (_, i) => ref.top + i + 1));
    return RangeValue.row(Array.from({ length: ref.cols }, (_, i) => ref.left + i + 1));
  }

  private _shapeOf(args: Array<Node | null>, ctx: EvalCtx): { rows: number; cols: number } {
    if (args.length !== 1 || !args[0]) throw new Error('#VALUE!');
    const value = this._eval(args[0], ctx);
    if (value instanceof RefValue) {
      if (!value.isCell) return value;
      // One cell: an array kept in it counts with its own size
      const inner = this._deref(value);
      return inner instanceof RangeValue ? inner : { rows: 1, cols: 1 };
    }
    if (value instanceof RangeValue) return value;
    return { rows: 1, cols: 1 };
  }

  private _registerBuiltins(): void {
    const registrar: Registrar = {
      def: (name, spec, fn, opts) => {
        this._functions.set(name, { fn, ...parseKinds(spec), arity: opts?.arity, volatile: opts?.volatile });
        if (opts?.volatile) this._volatileNames.add(name);
      },
      agg: (name, fn, opts) => {
        this._functions.set(name, { agg: fn, kinds: [], repeat: ['a'], propagate: opts?.propagate });
      },
    };
    registerPureBuiltins(registrar);

    // ADDRESS(row, column, [abs], [a1], [sheet]): the text of a cell reference
    registrar.def('ADDRESS', 'nnnnt', (_ctx, row, col, abs?, _a1?, sheet?) => {
      const r = Math.trunc(toNumber(row));
      const c = Math.trunc(toNumber(col));
      if (r < 1 || c < 1 || r > MAX_ROWS || c > MAX_COLS) throw new Error('#VALUE!');
      const mode = abs === undefined ? 1 : Math.trunc(toNumber(abs));
      if (mode < 1 || mode > 4) throw new Error('#VALUE!');
      const rowAbs = mode === 1 || mode === 2 ? '$' : '';
      const colAbs = mode === 1 || mode === 3 ? '$' : '';
      const prefix = sheet === undefined || sheet === '' ? '' : sheetPrefix(toText(sheet));
      return `${prefix}${colAbs}${colToLetter(c - 1)}${rowAbs}${r}`;
    });

    // The engine's own IF / IFERROR handle laziness; these keep the names registered
    for (const name of Object.keys(this._special)) {
      if (!this._functions.has(name)) this._functions.set(name, { kinds: [], repeat: ['a'], fn: () => 0 });
    }
  }
}

/**
 * A plain JS array returned by a (custom) function as an array value: an
 * array of arrays is a table of rows, a flat array a column.
 */
function fromJsArray(value: unknown): unknown {
  if (!Array.isArray(value) || value.length === 0) return value;
  if (value.every((row) => Array.isArray(row))) {
    const rows = value as unknown[][];
    const cols = Math.max(1, ...rows.map((row) => row.length));
    const values: unknown[] = [];
    for (const row of rows) {
      for (let c = 0; c < cols; c++) values.push(c < row.length ? row[c] : undefined);
    }
    return new RangeValue(values, rows.length, cols);
  }
  return RangeValue.column(value);
}

/** An error bound to a LET name or LAMBDA parameter, thrown when the name is used. */
class BoundError {
  constructor(readonly code: string) {}
}

/** Marks a broadcast argument that is passed whole to every call instead of element by element. */
class Scalar {
  constructor(readonly value: unknown) {}
}

/** Marks a broadcast argument whose elements may be errors without failing the element. */
class Tolerant {
  constructor(readonly value: unknown) {}
}

/** Apply a function of one scalar to a value, element by element for arrays. */
function mapValue(value: unknown, fn: (v: unknown) => unknown): unknown {
  if (value instanceof LambdaValue) throw new Error('#VALUE!');
  if (!(value instanceof RangeValue)) return fn(value);
  return broadcast([value], ([v]) => mapValue(v, fn));
}

/**
 * Call `fn` for every element position of the array arguments and collect
 * the results into an array. Arguments with one row (or column) repeat along
 * that axis; where an argument has no element, the result is #N/A. An error
 * element makes that result the error (unless the argument is `Tolerant`),
 * and an error thrown by `fn` becomes that element's value.
 */
/** A value already computed, passed where an argument node is expected. */
class Evaluated {
  constructor(readonly value: unknown) {}
}

function broadcast(args: unknown[], fn: (elements: unknown[]) => unknown, propagateErrors = true): RangeValue {
  let rows = 1;
  let cols = 1;
  const sources = args.map((a) => {
    const tolerant = a instanceof Tolerant;
    const value = tolerant ? (a as Tolerant).value : a;
    if (value instanceof RangeValue) {
      rows = Math.max(rows, value.rows);
      cols = Math.max(cols, value.cols);
    }
    return { value, tolerant };
  });
  const values: unknown[] = new Array(rows * cols);
  const elements: unknown[] = new Array(args.length);
  let nestedCells = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      let error: string | undefined;
      for (let i = 0; i < sources.length; i++) {
        const { value, tolerant } = sources[i];
        let element: unknown = value;
        if (value instanceof RangeValue) {
          const rr = value.rows === 1 ? 0 : r;
          const cc = value.cols === 1 ? 0 : c;
          element = rr < value.rows && cc < value.cols ? value.get(rr, cc) : '#N/A';
        }
        if (propagateErrors && !tolerant && error === undefined && isErrorCode(element)) error = element;
        elements[i] = element;
      }
      if (error !== undefined) {
        values[r * cols + c] = error;
        continue;
      }
      try {
        const result = fn(elements.slice());
        values[r * cols + c] = result instanceof LambdaValue ? '#CALC!' : result;
        // A call per element may each return an array: bound what all of them hold together
        if (result instanceof RangeValue) nestedCells += result.rows * result.cols;
      } catch (e) {
        if (e instanceof EvalDepthError) throw e;
        values[r * cols + c] = toErrorCode(e);
      }
      if (nestedCells > MAX_ARRAY_CELLS) throw new Error('#NUM!');
    }
  }
  return new RangeValue(values, rows, cols);
}

/** Rewrite the sheet name of every reference to `oldName` in a formula. */
export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  if (!formula.startsWith('=')) return formula;
  const lower = oldName.toLowerCase();
  return mapFormulaReferences(formula, (ref, text) =>
    ref.sheet !== undefined && ref.sheet.toLowerCase() === lower ? sheetPrefix(newName) + text.ref : null
  );
}
