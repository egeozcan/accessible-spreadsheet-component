import {
  type GridData,
  type FormulaContext,
  type FormulaFunction,
  cellKey,
  refToCoord,
} from '../types.js';

/** Token types for the lexer */
type TokenType =
  | 'NUMBER'
  | 'STRING'
  | 'BOOLEAN'
  | 'REF'
  | 'RANGE'
  | 'FUNC'
  | 'OPERATOR'
  | 'LPAREN'
  | 'RPAREN'
  | 'COMMA'
  | 'EOF';

interface Token {
  type: TokenType;
  value: string;
}

/** Mutable parser state passed through the recursive descent chain */
interface ParserState {
  tokens: Token[];
  pos: number;
}

/**
 * FormulaEngine -- Recursive descent parser and evaluator for Excel-like formulas.
 *
 * Supports cell references (A1, $B$3), range references (A1:B5),
 * functions (SUM, IF, VLOOKUP, ...), arithmetic (+, -, *, /),
 * comparison (=, <>, <, >, <=, >=), string concatenation (&),
 * and boolean/string/numeric literals.
 *
 * Grammar:
 *   expression     = comparison
 *   comparison     = concat (("=" | "<>" | "<" | ">" | "<=" | ">=") concat)*
 *   concat         = additive ("&" additive)*
 *   additive       = multiplicative (("+" | "-") multiplicative)*
 *   multiplicative = unary (("*" | "/") unary)*
 *   unary          = ("-" unary) | primary
 *   primary        = NUMBER | STRING | BOOLEAN | REF | RANGE | FUNC "(" args ")" | "(" expression ")"
 *
 * Dependency tracking: forward/reverse dep graph enables targeted BFS recalculation.
 * Circular references are detected at evaluation time via a re-entrancy set.
 */

/** Maximum nesting depth for formula evaluation to prevent stack overflow from deeply nested formulas */
const MAX_EVAL_DEPTH = 64;

/** Set of aggregate function names that should have RangeValue flattened into individual args */
const AGGREGATE_FUNCTIONS = new Set([
  'SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'CONCAT',
]);

/** Set of volatile function names whose results should never be cached */
const VOLATILE_FUNCTIONS = new Set(['NOW']);

/** Aggregates that propagate an error found anywhere in their arguments (COUNT/COUNTA skip them) */
const ERROR_PROPAGATING_AGGREGATES = new Set(['SUM', 'AVERAGE', 'MIN', 'MAX', 'CONCAT']);

/** Upper bound on deferred cells when resolving reference chains deeper than MAX_EVAL_DEPTH */
const MAX_DEFERRED_CELLS = 100_000;

const ERROR_CODES = new Set([
  '#ERROR!', '#REF!', '#DIV/0!', '#NAME?', '#CIRC!', '#VALUE!', '#N/A', '#NUM!',
]);

function isErrorCode(v: unknown): v is string {
  return typeof v === 'string' && ERROR_CODES.has(v);
}

/** Map a thrown value to a spreadsheet error code. */
function toErrorCode(e: unknown): string {
  const msg = e instanceof Error ? e.message : '';
  return msg.startsWith('#') ? msg : '#ERROR!';
}

function isVolatileFormula(rawValue: string): boolean {
  const upper = rawValue.toUpperCase();
  return [...VOLATILE_FUNCTIONS].some((fn) => upper.includes(fn + '('));
}

/**
 * Coerce an operand of an arithmetic operator to a number.
 * Empty values are 0 and booleans are 1/0 (Excel semantics); anything else
 * that is not numeric is a #VALUE! error rather than a silent NaN.
 */
function toNumber(v: unknown): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v === undefined || v === null || v === '') return 0;
  if (typeof v === 'string') {
    if (ERROR_CODES.has(v)) throw new Error(v);
    const n = Number(v);
    if (v.trim() !== '' && !isNaN(n)) return n;
  }
  throw new Error('#VALUE!');
}

/**
 * Ordering used by approximate-match lookups: numbers compare numerically,
 * text compares case-insensitively, and values of different kinds are
 * incomparable (NaN), so they never match.
 */
function lookupCompare(a: unknown, b: unknown): number {
  const isNum = (v: unknown) =>
    typeof v === 'number' || (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v)));
  if (a === undefined || b === undefined) return NaN;
  if (isNum(a) && isNum(b)) return Number(a) - Number(b);
  if (isNum(a) || isNum(b) || typeof a === 'boolean' || typeof b === 'boolean') return NaN;
  const as = String(a).toLowerCase();
  const bs = String(b).toLowerCase();
  return as < bs ? -1 : as > bs ? 1 : 0;
}

/**
 * Index of the last value <= `lookup` in an ascending list, stopping at the
 * first larger value; values of a different kind are skipped. -1 if none.
 */
function approximateMatch(values: unknown[], lookup: unknown): number {
  let match = -1;
  for (let i = 0; i < values.length; i++) {
    const cmp = lookupCompare(values[i], lookup);
    if (isNaN(cmp)) continue;
    if (cmp > 0) break;
    match = i;
  }
  return match;
}

/**
 * Thrown when resolving a reference would exceed MAX_EVAL_DEPTH. `key` is the
 * cell that could not be entered; `evaluate()` evaluates it on its own first
 * (memoizing the result) and then retries, so long reference chains resolve
 * without growing the JS call stack.
 */
class EvalDepthError extends Error {
  constructor(readonly key: string) {
    super('#ERROR!');
  }
}

type EvalResult = { displayValue: string; type: 'text' | 'number' | 'boolean' | 'error' };

/**
 * Represents a 2D range of values with shape information.
 * Used by lookup functions (VLOOKUP, INDEX, etc.) that need row/col structure.
 */
export class RangeValue {
  readonly values: unknown[];
  readonly rows: number;
  readonly cols: number;

  constructor(values: unknown[], rows: number, cols: number) {
    this.values = values;
    this.rows = rows;
    this.cols = cols;
  }

  get(row: number, col: number): unknown {
    return this.values[row * this.cols + col];
  }

  getRow(row: number): unknown[] {
    const start = row * this.cols;
    return this.values.slice(start, start + this.cols);
  }

  getCol(col: number): unknown[] {
    const result: unknown[] = [];
    for (let r = 0; r < this.rows; r++) {
      result.push(this.values[r * this.cols + col]);
    }
    return result;
  }
}

/**
 * Helper to check if a value matches a criteria string.
 * Supports operator prefixes: <>, >=, <=, >, <, =
 * Without operator prefix, does exact match (case-insensitive for strings).
 */
function matchesCriteria(value: unknown, criteria: string): boolean {
  // Parse operator from criteria
  let op = '=';
  let target = criteria;

  if (criteria.startsWith('<>')) {
    op = '<>';
    target = criteria.substring(2);
  } else if (criteria.startsWith('>=')) {
    op = '>=';
    target = criteria.substring(2);
  } else if (criteria.startsWith('<=')) {
    op = '<=';
    target = criteria.substring(2);
  } else if (criteria.startsWith('>')) {
    op = '>';
    target = criteria.substring(1);
  } else if (criteria.startsWith('<')) {
    op = '<';
    target = criteria.substring(1);
  } else if (criteria.startsWith('=')) {
    op = '=';
    target = criteria.substring(1);
  }

  const numTarget = Number(target);
  const numValue = Number(value);
  const targetNumeric = !isNaN(numTarget) && target.trim() !== '';
  const valueNumeric = !isNaN(numValue) && String(value).trim() !== '';
  const bothNumeric = targetNumeric && valueNumeric;

  // Relational operators only compare like with like (Excel: ">5" never matches text)
  if (op !== '=' && op !== '<>' && targetNumeric !== valueNumeric) {
    return false;
  }

  const valLower = String(value).toLowerCase();
  const targetLower = target.toLowerCase();

  switch (op) {
    case '=':
      if (bothNumeric) return numValue === numTarget;
      return valLower === targetLower;
    case '<>':
      if (bothNumeric) return numValue !== numTarget;
      return valLower !== targetLower;
    case '>':
      return bothNumeric ? numValue > numTarget : valLower > targetLower;
    case '<':
      return bothNumeric ? numValue < numTarget : valLower < targetLower;
    case '>=':
      return bothNumeric ? numValue >= numTarget : valLower >= targetLower;
    case '<=':
      return bothNumeric ? numValue <= numTarget : valLower <= targetLower;
    default:
      return false;
  }
}

export class FormulaEngine {
  private _functions: Map<string, FormulaFunction> = new Map();
  private _builtinFunctionNames: Set<string> = new Set();
  private _data: GridData = new Map();
  private _evaluating: Set<string> = new Set(); // circular reference detection
  private _evalDepth = 0;

  // Dependency tracking for targeted recalculation
  private _deps: Map<string, Set<string>> = new Map();
  private _reverseDeps: Map<string, Set<string>> = new Map();
  private _trackingCellKey: string | null = null;

  // Formula result cache for performance. Entries remember the raw formula
  // they were computed from so an edited cell never returns a stale result.
  private _cache: Map<string, EvalResult & { rawValue: string }> = new Map();

  /**
   * Values of referenced formula cells computed during the current
   * evaluation pass (one top-level `evaluate()`, or a whole `recalculate()` /
   * `recalculateAffected()`). Avoids re-evaluating shared precedents, which is
   * otherwise exponential, and lets deep chains be resolved incrementally.
   * Null outside a pass, so standalone `evaluate()` calls always see fresh data.
   */
  private _memo: Map<string, { value?: unknown; error?: string }> | null = null;

  constructor() {
    this.registerBuiltins();
    this._builtinFunctionNames = new Set(this._functions.keys());
  }

  /**
   * Register a user-defined formula function.
   * The name is normalized to uppercase for case-insensitive lookup.
   *
   * @param name - Function name (e.g. "MYFUNC"); stored as uppercase
   * @param fn - The function implementation
   */
  registerFunction(name: string, fn: FormulaFunction): void {
    this._functions.set(name.toUpperCase(), fn);
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
      }
    }

    for (const [name, fn] of Object.entries(functions)) {
      this.registerFunction(name, fn);
    }
  }

  /** Update the data reference for evaluation (full reset of dep graph and cache) */
  setData(data: GridData): void {
    this._data = data;
    this._deps.clear();
    this._reverseDeps.clear();
    this._cache.clear();
  }

  /**
   * Update the data reference while preserving the dependency graph for
   * unchanged cells. Only clears deps and cache for the specified changed keys,
   * allowing recalculateAffected() to use the existing graph for BFS traversal.
   */
  updateData(data: GridData, changedKeys: string[]): void {
    this._data = data;
    for (const key of changedKeys) {
      this._clearDepsFor(key);
    }
  }

  /**
   * Evaluate a raw value. If it starts with `=`, parse and execute the formula.
   * Otherwise, return the value as-is (possibly coerced to number/boolean).
   *
   * @param rawValue - The raw cell value (formula or literal)
   * @param forCellKey - If provided, registers dependencies for this cell key
   *   (side-effect: updates the forward/reverse dep graph and result cache)
   * @returns The display value and resolved type
   */
  evaluate(rawValue: string, forCellKey?: string): EvalResult {
    // Keep the cell marked in-progress across deferred retries so a long chain
    // leading back to it is reported as circular rather than reading its old value.
    const ownsKey =
      !!forCellKey && rawValue.startsWith('=') && !this._evaluating.has(forCellKey);
    if (ownsKey) this._evaluating.add(forCellKey!);
    try {
      return this._evaluateDeferred(rawValue, forCellKey);
    } finally {
      if (ownsKey) this._evaluating.delete(forCellKey!);
    }
  }

  private _evaluateDeferred(rawValue: string, forCellKey?: string): EvalResult {
    return this._withMemo(() => {
      // Cells whose evaluation hit the depth limit, innermost last. Each is
      // evaluated on its own (memoizing its value) before retrying its parent.
      const deferred: string[] = [];
      for (;;) {
        const target = deferred.length > 0 ? deferred[deferred.length - 1] : null;
        try {
          if (target === null) {
            return this._evaluateTopLevel(rawValue, forCellKey);
          }
          this._evaluateFormulaCell(target, this._data.get(target)?.rawValue ?? '');
          deferred.pop();
        } catch (e) {
          if (!(e instanceof EvalDepthError)) {
            // A deferred cell failed with a normal error; it is memoized as
            // such, so its dependents will see the error on retry.
            deferred.pop();
            continue;
          }
          const cycleStart = deferred.indexOf(e.key);
          if (cycleStart !== -1) {
            // deferred[cycleStart..] form a (long) cycle. Memoize the verdict
            // for its members so later evaluations in this pass don't walk the
            // cycle again (quadratic otherwise), then retry the cells that
            // merely depend on it so they can handle the error (e.g. IFERROR).
            for (const key of deferred.slice(cycleStart)) {
              this._memo?.set(key, { error: '#CIRC!' });
            }
            deferred.length = cycleStart;
            continue;
          }
          if (deferred.length >= MAX_DEFERRED_CELLS) {
            return { displayValue: '#ERROR!', type: 'error' };
          }
          deferred.push(e.key);
        }
      }
    });
  }

  /** Run `fn` inside an evaluation pass, creating the pass memo if needed. */
  private _withMemo<T>(fn: () => T): T {
    if (this._memo) return fn();
    this._memo = new Map();
    try {
      return fn();
    } finally {
      this._memo = null;
    }
  }

  /**
   * Evaluate a raw value as the top-level formula of `forCellKey`.
   * Throws only EvalDepthError; every other failure becomes an error result.
   */
  private _evaluateTopLevel(rawValue: string, forCellKey?: string): EvalResult {
    if (!rawValue || rawValue.trim() === '') {
      if (forCellKey) {
        this._clearDepsFor(forCellKey);
        this._cache.delete(forCellKey);
      }
      return { displayValue: '', type: 'text' };
    }

    if (!rawValue.startsWith('=')) {
      if (forCellKey) {
        this._clearDepsFor(forCellKey);
        this._cache.delete(forCellKey);
      }
      return this.coerceValue(rawValue);
    }

    // Skip the cache for volatile functions
    const isVolatile = isVolatileFormula(rawValue);

    // Check cache for formula cells (only valid for the same formula text)
    if (forCellKey && !isVolatile) {
      const cached = this._cache.get(forCellKey);
      if (cached && cached.rawValue === rawValue) {
        return { displayValue: cached.displayValue, type: cached.type };
      }
    }

    const ownsEvaluatingKey = !!forCellKey && !this._evaluating.has(forCellKey);
    try {
      if (forCellKey) {
        this._clearDepsFor(forCellKey);
        this._trackingCellKey = forCellKey;
        // Mark the cell as in-progress so self-references (direct or through
        // a range) are reported as circular.
        if (ownsEvaluatingKey) this._evaluating.add(forCellKey);
      }
      const formula = rawValue.substring(1);
      const result = this.parseExpression(formula);
      const evaluated = this._resultToDisplay(result);

      if (forCellKey) {
        this._memo?.set(
          forCellKey,
          // Same shape as nested evaluation memoizes, so results don't depend on order
          { value: result }
        );
        if (!isVolatile) {
          this._cache.set(forCellKey, { ...evaluated, rawValue });
        }
      }

      return evaluated;
    } catch (e) {
      if (e instanceof EvalDepthError) throw e;
      // Preserve specific error codes (#DIV/0!, #NAME?, #CIRC!)
      const errorResult: EvalResult = { displayValue: toErrorCode(e), type: 'error' };

      if (forCellKey) {
        this._memo?.set(forCellKey, { error: errorResult.displayValue });
        if (!isVolatile) {
          this._cache.set(forCellKey, { ...errorResult, rawValue });
        }
      }

      return errorResult;
    } finally {
      this._trackingCellKey = null;
      if (ownsEvaluatingKey) this._evaluating.delete(forCellKey!);
    }
  }

  /**
   * Convert a formula result to its display string and type based on the
   * JS type of the result, so text results like "007" or TEXT() output stay text.
   */
  private _resultToDisplay(result: unknown): EvalResult {
    if (typeof result === 'number') {
      if (isNaN(result)) return { displayValue: '#VALUE!', type: 'error' };
      if (!Number.isFinite(result)) return { displayValue: '#NUM!', type: 'error' };
      return { displayValue: String(parseFloat(result.toPrecision(15))), type: 'number' };
    }
    if (typeof result === 'boolean') {
      return { displayValue: result ? 'TRUE' : 'FALSE', type: 'boolean' };
    }
    if (typeof result === 'string') {
      return isErrorCode(result)
        ? { displayValue: result, type: 'error' }
        : { displayValue: result, type: 'text' };
    }
    if (result instanceof RangeValue) {
      return { displayValue: '#VALUE!', type: 'error' };
    }
    if (result === undefined || result === null) {
      return { displayValue: '0', type: 'number' };
    }
    return this.coerceValue(String(result));
  }

  /**
   * Re-evaluate all formula cells in the grid.
   * Clears and rebuilds the entire dependency graph and result cache.
   *
   * @returns Set of cell keys whose display value or type changed
   */
  recalculate(): Set<string> {
    this._deps.clear();
    this._reverseDeps.clear();
    this._cache.clear();

    return this._withMemo(() => {
      const changed = new Set<string>();

      for (const [key, cell] of this._data) {
        if (cell.rawValue.startsWith('=')) {
          const result = this.evaluate(cell.rawValue, key);
          if (cell.displayValue !== result.displayValue || cell.type !== result.type) {
            cell.displayValue = result.displayValue;
            cell.type = result.type;
            changed.add(key);
          }
        }
      }

      return changed;
    });
  }

  /**
   * Recalculate only formulas affected by the given changed cell keys.
   * Uses BFS traversal of the reverse dependency graph for targeted recalculation.
   * Falls back to full `recalculate()` if the dep graph is empty.
   *
   * @param changedKeys - Cell keys whose raw values changed
   * @returns Set of cell keys whose display value or type changed
   */
  recalculateAffected(changedKeys: string[]): Set<string> {
    // If the dependency graph is empty but data exists, fall back to a full
    // recalculate so we never silently skip dependents after a setData() call.
    if (this._reverseDeps.size === 0 && this._data.size > 0) {
      return this.recalculate();
    }

    // Collect the changed cells plus all transitive dependents first, and
    // invalidate them all before evaluating anything, so no formula is ever
    // computed from a stale cached precedent.
    const affected = new Set<string>(changedKeys);
    const queue = [...changedKeys];
    while (queue.length > 0) {
      const key = queue.shift()!;
      const dependents = this._reverseDeps.get(key);
      if (dependents) {
        for (const dep of dependents) {
          if (!affected.has(dep)) {
            affected.add(dep);
            queue.push(dep);
          }
        }
      }
    }

    for (const key of affected) {
      this._cache.delete(key);
    }

    return this._withMemo(() => {
      const changed = new Set<string>();
      for (const key of affected) {
        const cell = this._data.get(key);
        if (cell?.rawValue.startsWith('=')) {
          const result = this.evaluate(cell.rawValue, key);
          if (cell.displayValue !== result.displayValue || cell.type !== result.type) {
            cell.displayValue = result.displayValue;
            cell.type = result.type;
            changed.add(key);
          }
        }
      }
      return changed;
    });
  }

  // ─── Dependency Tracking ─────────────────────────────

  private _clearDepsFor(targetKey: string): void {
    const oldDeps = this._deps.get(targetKey);
    if (oldDeps) {
      for (const dep of oldDeps) {
        this._reverseDeps.get(dep)?.delete(targetKey);
      }
      this._deps.delete(targetKey);
    }
    this._cache.delete(targetKey);
  }

  private _trackDep(referencedKey: string): void {
    if (!this._trackingCellKey) return;

    let deps = this._deps.get(this._trackingCellKey);
    if (!deps) {
      deps = new Set();
      this._deps.set(this._trackingCellKey, deps);
    }
    deps.add(referencedKey);

    let rev = this._reverseDeps.get(referencedKey);
    if (!rev) {
      rev = new Set();
      this._reverseDeps.set(referencedKey, rev);
    }
    rev.add(this._trackingCellKey);
  }

  // ─── Value Coercion ──────────────────────────────────

  private coerceValue(val: string): { displayValue: string; type: 'text' | 'number' | 'boolean' | 'error' } {
    if (val === '#ERROR!' || val === '#REF!' || val === '#DIV/0!' || val === '#NAME?' || val === '#CIRC!' || val === '#VALUE!' || val === '#N/A' || val === '#NUM!') {
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

  // ─── Lexer ───────────────────────────────────────────

  private tokenize(input: string): Token[] {
    const tokens: Token[] = [];
    let i = 0;

    while (i < input.length) {
      const ch = input[i];

      // Skip whitespace
      if (/\s/.test(ch)) {
        i++;
        continue;
      }

      // String literal (supports "" as escaped quote, matching Excel convention)
      if (ch === '"') {
        let str = '';
        i++; // skip opening quote
        while (i < input.length) {
          if (input[i] === '"') {
            // Doubled "" is an escaped literal quote
            if (i + 1 < input.length && input[i + 1] === '"') {
              str += '"';
              i += 2;
            } else {
              break; // closing quote
            }
          } else {
            str += input[i];
            i++;
          }
        }
        if (i >= input.length) {
          throw new Error('Unterminated string literal');
        }
        i++; // skip closing quote
        tokens.push({ type: 'STRING', value: str });
        continue;
      }

      // Number literal
      if (/\d/.test(ch) || (ch === '.' && i + 1 < input.length && /\d/.test(input[i + 1]))) {
        let num = '';
        while (i < input.length && (/\d/.test(input[i]) || input[i] === '.')) {
          num += input[i];
          i++;
        }
        tokens.push({ type: 'NUMBER', value: num });
        continue;
      }

      // Operators
      if (ch === '+' || ch === '-' || ch === '*' || ch === '/' || ch === '&') {
        tokens.push({ type: 'OPERATOR', value: ch });
        i++;
        continue;
      }

      // Comparison operators
      if (ch === '<' || ch === '>') {
        if (i + 1 < input.length && input[i + 1] === '=') {
          tokens.push({ type: 'OPERATOR', value: ch + '=' });
          i += 2;
        } else if (ch === '<' && i + 1 < input.length && input[i + 1] === '>') {
          tokens.push({ type: 'OPERATOR', value: '<>' });
          i += 2;
        } else {
          tokens.push({ type: 'OPERATOR', value: ch });
          i++;
        }
        continue;
      }

      if (ch === '=') {
        tokens.push({ type: 'OPERATOR', value: '=' });
        i++;
        continue;
      }

      // Parentheses
      if (ch === '(') {
        tokens.push({ type: 'LPAREN', value: '(' });
        i++;
        continue;
      }
      if (ch === ')') {
        tokens.push({ type: 'RPAREN', value: ')' });
        i++;
        continue;
      }

      // Comma
      if (ch === ',') {
        tokens.push({ type: 'COMMA', value: ',' });
        i++;
        continue;
      }

      // Identifiers: cell references, function names, booleans
      // Also match $ for absolute/mixed references like $A$1, $A1, A$1
      if (/[A-Za-z_$]/.test(ch)) {
        let ident = '';
        while (i < input.length && /[A-Za-z0-9_$]/.test(input[i])) {
          ident += input[i];
          i++;
        }

        // Strip $ for structure checks, but keep original for token value
        const stripped = ident.replace(/\$/g, '');
        const hasDollar = stripped !== ident;
        const upper = stripped.toUpperCase();

        // Check if boolean (only when no $ is present)
        if (!hasDollar && (upper === 'TRUE' || upper === 'FALSE')) {
          tokens.push({ type: 'BOOLEAN', value: upper });
          continue;
        }

        // Check for range (e.g. A1:B2 or $A$1:$B$2) - look ahead for colon
        if (i < input.length && input[i] === ':' && /^[A-Z]+\d+$/i.test(stripped)) {
          i++; // skip colon
          let end = '';
          while (i < input.length && /[A-Za-z0-9$]/.test(input[i])) {
            end += input[i];
            i++;
          }
          const endStripped = end.replace(/\$/g, '');
          if (/^[A-Z]+\d+$/i.test(endStripped)) {
            tokens.push({ type: 'RANGE', value: `${ident.toUpperCase()}:${end.toUpperCase()}` });
            continue;
          }
          // If the part after colon isn't a valid ref, treat as error
          throw new Error(`Invalid range: ${ident}:${end}`);
        }

        // Check if function call (next non-space is '(') - only when no $ is present
        if (!hasDollar) {
          let lookAhead = i;
          while (lookAhead < input.length && /\s/.test(input[lookAhead])) lookAhead++;
          if (lookAhead < input.length && input[lookAhead] === '(') {
            tokens.push({ type: 'FUNC', value: upper });
            continue;
          }
        }

        // Cell reference
        if (/^[A-Z]+\d+$/i.test(stripped)) {
          tokens.push({ type: 'REF', value: ident.toUpperCase() });
          continue;
        }

        // Unknown identifier - treat as function name or error (only without $)
        if (!hasDollar) {
          tokens.push({ type: 'FUNC', value: upper });
          continue;
        }

        throw new Error(`Unexpected identifier: ${ident}`);
      }

      throw new Error(`Unexpected character: ${ch}`);
    }

    tokens.push({ type: 'EOF', value: '' });
    return tokens;
  }

  // ─── Parser ──────────────────────────────────────────
  //
  // Parser state is passed as a mutable object through the recursive
  // descent chain so that nested evaluations (resolveRef / resolveRange)
  // each get their own independent state without save/restore.

  private _peek(s: ParserState): Token {
    return s.tokens[s.pos];
  }

  private _consume(s: ParserState, expectedType?: TokenType): Token {
    const token = s.tokens[s.pos];
    if (expectedType && token.type !== expectedType) {
      throw new Error(`Expected ${expectedType} but got ${token.type} (${token.value})`);
    }
    s.pos++;
    return token;
  }

  private parseExpression(input: string): unknown {
    if (++this._evalDepth > MAX_EVAL_DEPTH) {
      this._evalDepth--;
      throw new Error('#ERROR!');
    }
    try {
      const s: ParserState = { tokens: this.tokenize(input), pos: 0 };
      const result = this._parseComparison(s);
      if (this._peek(s).type !== 'EOF') {
        throw new Error(`Unexpected token: ${this._peek(s).value}`);
      }
      return result;
    } finally {
      this._evalDepth--;
    }
  }

  private _parseComparison(s: ParserState): unknown {
    let left = this._parseConcatenation(s);

    while (
      this._peek(s).type === 'OPERATOR' &&
      ['=', '<>', '<', '>', '<=', '>='].includes(this._peek(s).value)
    ) {
      const op = this._consume(s).value;
      const right = this._parseConcatenation(s);
      left = this._compareValues(left, right, op);
    }

    return left;
  }

  private _parseConcatenation(s: ParserState): unknown {
    let left = this._parseAddSub(s);

    while (this._peek(s).type === 'OPERATOR' && this._peek(s).value === '&') {
      this._consume(s); // &
      const right = this._parseAddSub(s);
      left = String(left) + String(right);
    }

    return left;
  }

  private _parseAddSub(s: ParserState): unknown {
    let left = this._parseMulDiv(s);

    while (
      this._peek(s).type === 'OPERATOR' &&
      (this._peek(s).value === '+' || this._peek(s).value === '-')
    ) {
      const op = this._consume(s).value;
      const right = this._parseMulDiv(s);
      if (op === '+') left = toNumber(left) + toNumber(right);
      else left = toNumber(left) - toNumber(right);
    }

    return left;
  }

  private _parseMulDiv(s: ParserState): unknown {
    let left = this._parseUnary(s);

    while (
      this._peek(s).type === 'OPERATOR' &&
      (this._peek(s).value === '*' || this._peek(s).value === '/')
    ) {
      const op = this._consume(s).value;
      const right = this._parseUnary(s);
      if (op === '*') left = toNumber(left) * toNumber(right);
      else {
        const dividend = toNumber(left);
        const divisor = toNumber(right);
        if (divisor === 0) throw new Error('#DIV/0!');
        left = dividend / divisor;
      }
    }

    return left;
  }

  private _parseUnary(s: ParserState): unknown {
    if (this._peek(s).type === 'OPERATOR' && this._peek(s).value === '-') {
      this._consume(s);
      return -toNumber(this._parseUnary(s));
    }
    if (this._peek(s).type === 'OPERATOR' && this._peek(s).value === '+') {
      this._consume(s);
      return toNumber(this._parseUnary(s));
    }
    return this._parsePrimary(s);
  }

  private _parsePrimary(s: ParserState): unknown {
    const token = this._peek(s);

    switch (token.type) {
      case 'NUMBER':
        this._consume(s);
        return parseFloat(token.value);

      case 'STRING':
        this._consume(s);
        return token.value;

      case 'BOOLEAN':
        this._consume(s);
        return token.value === 'TRUE';

      case 'REF':
        this._consume(s);
        return this._resolveRef(token.value);

      case 'RANGE':
        this._consume(s);
        return this._resolveRange(token.value);

      case 'FUNC':
        return this._parseFunction(s);

      case 'LPAREN':
        this._consume(s);
        const expr = this._parseComparison(s);
        this._consume(s, 'RPAREN');
        return expr;

      default:
        throw new Error(`Unexpected token: ${token.type} (${token.value})`);
    }
  }

  private _parseFunction(s: ParserState): unknown {
    const name = this._consume(s, 'FUNC').value;
    this._consume(s, 'LPAREN');

    // IFERROR needs to trap errors thrown by its first argument rather
    // than letting them propagate past the function call.
    if (name === 'IFERROR') {
      return this._parseIFERROR(s);
    }

    // IF only evaluates the branch it returns, so guards like
    // =IF(A1=0, 0, 1/A1) don't raise errors from the untaken branch.
    if (name === 'IF') {
      return this._parseIF(s);
    }

    const args: unknown[] = [];
    if (this._peek(s).type !== 'RPAREN') {
      args.push(this._parseComparison(s));
      while (this._peek(s).type === 'COMMA') {
        this._consume(s);
        args.push(this._parseComparison(s));
      }
    }

    this._consume(s, 'RPAREN');

    const fn = this._functions.get(name);
    if (!fn) throw new Error(`#NAME?`);

    const ctx = this._createContext();

    // For aggregate functions, flatten RangeValue into individual args
    // For other functions, pass RangeValue directly so they can access shape
    if (AGGREGATE_FUNCTIONS.has(name)) {
      const flatArgs: unknown[] = [];
      for (const arg of args) {
        if (arg instanceof RangeValue) {
          // Filter out undefined (empty cells) for aggregates
          for (const v of arg.values) {
            if (v !== undefined) flatArgs.push(v);
          }
        } else if (Array.isArray(arg)) {
          flatArgs.push(...arg);
        } else {
          flatArgs.push(arg);
        }
      }
      if (ERROR_PROPAGATING_AGGREGATES.has(name)) {
        const err = flatArgs.find(isErrorCode);
        if (err !== undefined) throw new Error(err);
      }
      return fn(ctx, ...flatArgs) ?? 0;
    }

    // An empty cell returned by e.g. VLOOKUP reads as 0, matching how it displays
    return fn(ctx, ...args) ?? 0;
  }

  /**
   * Parse IFERROR(value, fallback) with error trapping on the first argument.
   * Called after LPAREN has already been consumed.
   */
  private _parseIFERROR(s: ParserState): unknown {
    const savedPos = s.pos;
    let value: unknown;
    let caught = false;

    try {
      value = this._parseComparison(s);
    } catch (e) {
      // Running out of evaluation depth is not a formula error; let evaluate() retry.
      if (e instanceof EvalDepthError) throw e;
      caught = true;
      // The first argument threw — scan forward to the comma or closing paren
      // so we can parse the fallback argument from a known position.
      s.pos = savedPos;
      this._skipArgument(s);
    }

    // Parse the fallback argument (if present)
    let fallback: unknown = '';
    if (this._peek(s).type === 'COMMA') {
      this._consume(s); // consume comma
      fallback = this._parseComparison(s);
    }
    this._consume(s, 'RPAREN');

    if (caught) return fallback;

    // No throw — but the value itself might be an error string (e.g. from a
    // cell whose displayValue is already an error code).
    if (isErrorCode(value)) {
      return fallback;
    }
    return value;
  }

  /**
   * Parse IF(condition, then, else), evaluating only the selected branch.
   * Called after LPAREN has already been consumed.
   */
  private _parseIF(s: ParserState): unknown {
    const condition = this._parseComparison(s);
    // Excel defaults: omitted true branch → 0, omitted false branch → FALSE
    let result: unknown = condition ? 0 : false;

    if (this._peek(s).type === 'COMMA') {
      this._consume(s);
      if (condition) result = this._parseComparison(s);
      else this._skipArgument(s);

      if (this._peek(s).type === 'COMMA') {
        this._consume(s);
        if (condition) this._skipArgument(s);
        else result = this._parseComparison(s);
      }
    }

    this._consume(s, 'RPAREN');
    return result;
  }

  /** Advance past one function argument without evaluating it. */
  private _skipArgument(s: ParserState): void {
    let depth = 0;
    while (this._peek(s).type !== 'EOF') {
      const t = this._peek(s);
      if (t.type === 'LPAREN') { depth++; s.pos++; }
      else if (t.type === 'RPAREN') {
        if (depth === 0) break;
        depth--;
        s.pos++;
      }
      else if (t.type === 'COMMA' && depth === 0) break;
      else s.pos++;
    }
  }

  // ─── Comparison helper ────────────────────────────────

  private _compareValues(left: unknown, right: unknown, op: string): boolean {
    // Determine if both sides can be compared numerically.
    // Booleans are numeric (TRUE=1, FALSE=0). Strings that look like numbers
    // are numeric. Empty strings and non-numeric strings are NOT numeric.
    const isNumeric = (v: unknown): boolean => {
      if (typeof v === 'boolean') return true;
      if (typeof v === 'number') return true;
      const s = String(v).trim();
      return s !== '' && !isNaN(Number(s));
    };

    const bothNumeric = isNumeric(left) && isNumeric(right);

    if (bothNumeric) {
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

    // String comparison: = and <> are case-insensitive (Excel convention),
    // relational operators (<, >, <=, >=) use locale-unaware ordinal comparison.
    const ls = String(left ?? '');
    const rs = String(right ?? '');
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

  // ─── Reference resolution ───────────────────────────
  //
  // Each nested evaluation calls parseExpression which creates its
  // own ParserState, so no save/restore is needed.

  private _resolveRef(ref: string): unknown {
    const coord = refToCoord(ref.replace(/\$/g, ''));
    const key = cellKey(coord.row, coord.col);

    this._trackDep(key);

    // Circular reference detection
    if (this._evaluating.has(key)) {
      throw new Error('#CIRC!');
    }

    const cell = this._data.get(key);
    if (!cell) return 0; // empty cells are 0

    if (cell.rawValue.startsWith('=')) {
      return this._evaluateFormulaCell(key, cell.rawValue);
    }

    // Return the value, coerced to number if possible
    const num = Number(cell.rawValue);
    if (!isNaN(num) && cell.rawValue.trim() !== '') return num;
    if (cell.rawValue.toUpperCase() === 'TRUE') return true;
    if (cell.rawValue.toUpperCase() === 'FALSE') return false;
    return cell.rawValue;
  }

  /**
   * Evaluate a referenced formula cell, reusing its value if it was already
   * computed in this evaluation pass. Keeps `_trackingCellKey` so transitive
   * dependencies are recorded (e.g. C1→B1→A1 means C1 depends on A1 too).
   */
  private _evaluateFormulaCell(key: string, rawValue: string): unknown {
    if (this._evaluating.has(key)) {
      throw new Error('#CIRC!');
    }

    const memoized = this._memo?.get(key);
    if (memoized) {
      if (memoized.error !== undefined) throw new Error(memoized.error);
      return memoized.value;
    }

    if (this._evalDepth >= MAX_EVAL_DEPTH) {
      throw new EvalDepthError(key);
    }

    this._evaluating.add(key);
    try {
      const value = this.parseExpression(rawValue.substring(1));
      this._memo?.set(key, { value });
      return value;
    } catch (e) {
      if (e instanceof EvalDepthError) throw e;
      // Preserve specific error codes (#DIV/0!, #N/A, etc.)
      const code = toErrorCode(e);
      this._memo?.set(key, { error: code });
      throw new Error(code);
    } finally {
      this._evaluating.delete(key);
    }
  }

  private _resolveRange(rangeStr: string): RangeValue {
    const [startRef, endRef] = rangeStr.split(':');
    const start = refToCoord(startRef.replace(/\$/g, ''));
    const end = refToCoord(endRef.replace(/\$/g, ''));

    const minRow = Math.min(start.row, end.row);
    const maxRow = Math.max(start.row, end.row);
    const minCol = Math.min(start.col, end.col);
    const maxCol = Math.max(start.col, end.col);

    const numRows = maxRow - minRow + 1;
    const numCols = maxCol - minCol + 1;

    const values: unknown[] = [];
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const key = cellKey(r, c);

        this._trackDep(key);

        const cell = this._data.get(key);
        if (cell) {
          if (cell.rawValue.startsWith('=')) {
            // Errors become error-code values so functions like COUNTIF can
            // skip them while aggregates like SUM propagate them.
            try {
              values.push(this._evaluateFormulaCell(key, cell.rawValue));
            } catch (e) {
              if (e instanceof EvalDepthError) throw e;
              values.push(toErrorCode(e));
            }
          } else if (cell.rawValue === '') {
            // A formatted-but-empty cell is still empty
            values.push(undefined);
          } else {
            // Match _resolveRef's coercion logic for consistency
            const num = Number(cell.rawValue);
            if (!isNaN(num) && cell.rawValue.trim() !== '') {
              values.push(num);
            } else if (cell.rawValue.toUpperCase() === 'TRUE') {
              values.push(true);
            } else if (cell.rawValue.toUpperCase() === 'FALSE') {
              values.push(false);
            } else {
              values.push(cell.rawValue);
            }
          }
        } else {
          // Empty cells contribute to shape
          values.push(undefined);
        }
      }
    }

    return new RangeValue(values, numRows, numCols);
  }

  private _createContext(): FormulaContext {
    return {
      getCellValue: (ref: string) => {
        if (ref.includes(':') && !/[A-Z]/i.test(ref.charAt(0))) {
          // It's a key like "0:0"
          const cell = this._data.get(ref);
          if (!cell) return 0;
          const num = Number(cell.rawValue);
          return !isNaN(num) && cell.rawValue.trim() !== '' ? num : cell.rawValue;
        }
        return this._resolveRef(ref);
      },
      getRangeValues: (startRef: string, endRef: string) => {
        const start = refToCoord(startRef);
        const end = refToCoord(endRef);

        const minRow = Math.min(start.row, end.row);
        const maxRow = Math.max(start.row, end.row);
        const minCol = Math.min(start.col, end.col);
        const maxCol = Math.max(start.col, end.col);

        const values: unknown[] = [];
        for (let r = minRow; r <= maxRow; r++) {
          for (let c = minCol; c <= maxCol; c++) {
            const key = cellKey(r, c);
            const cell = this._data.get(key);
            if (cell) {
              const num = Number(cell.rawValue);
              values.push(!isNaN(num) && cell.rawValue.trim() !== '' ? num : cell.rawValue);
            }
          }
        }
        return values;
      },
    };
  }

  // ─── Built-in Functions ─────────────────────────────

  private registerBuiltins(): void {
    this.registerFunction('SUM', (_ctx, ...args) => {
      return args.reduce((sum: number, v) => sum + (Number(v) || 0), 0);
    });

    this.registerFunction('AVERAGE', (_ctx, ...args) => {
      const nums = args.filter((v) => typeof v === 'number' || !isNaN(Number(v)));
      if (nums.length === 0) throw new Error('#DIV/0!');
      const sum = nums.reduce((s: number, v) => s + Number(v), 0);
      return sum / nums.length;
    });

    this.registerFunction('MIN', (_ctx, ...args) => {
      const nums = args.map(Number).filter((n) => !isNaN(n));
      return nums.length ? Math.min(...nums) : 0;
    });

    this.registerFunction('MAX', (_ctx, ...args) => {
      const nums = args.map(Number).filter((n) => !isNaN(n));
      return nums.length ? Math.max(...nums) : 0;
    });

    this.registerFunction('COUNT', (_ctx, ...args) => {
      return args.filter((v) => typeof v === 'number' || (!isNaN(Number(v)) && String(v).trim() !== '')).length;
    });

    this.registerFunction('COUNTA', (_ctx, ...args) => {
      return args.filter((v) => v !== '' && v !== null && v !== undefined).length;
    });

    this.registerFunction('IF', (_ctx, condition, trueVal, falseVal) => {
      // Excel defaults: omitted true branch → 0, omitted false branch → FALSE
      if (condition) return trueVal !== undefined ? trueVal : 0;
      return falseVal !== undefined ? falseVal : false;
    });

    this.registerFunction('CONCAT', (_ctx, ...args) => {
      return args.map(String).join('');
    });

    this.registerFunction('ABS', (_ctx, val) => {
      return Math.abs(Number(val));
    });

    this.registerFunction('ROUND', (_ctx, val, digits) => {
      const d = Number(digits) || 0;
      const n = Number(val);
      const factor = Math.pow(10, d);
      // Neutralise IEEE 754 drift (e.g. 1.005*100 → 100.4999… → 100.5)
      // then round half away from zero (Excel convention).
      const shifted = parseFloat((Math.abs(n) * factor).toPrecision(15));
      return Math.sign(n) * Math.round(shifted) / factor;
    });

    this.registerFunction('UPPER', (_ctx, val) => {
      return String(val).toUpperCase();
    });

    this.registerFunction('LOWER', (_ctx, val) => {
      return String(val).toLowerCase();
    });

    this.registerFunction('LEN', (_ctx, val) => {
      return String(val).length;
    });

    this.registerFunction('TRIM', (_ctx, val) => {
      return String(val).trim().replace(/ +/g, ' ');
    });

    // ─── Logic/Conditional ────────────────────────────────

    this.registerFunction('IFERROR', (_ctx, value, fallback) => {
      const s = String(value);
      if (
        s === '#ERROR!' || s === '#REF!' || s === '#DIV/0!' ||
        s === '#NAME?' || s === '#CIRC!' || s === '#VALUE!' || s === '#N/A'
      ) {
        return fallback;
      }
      return value;
    });

    this.registerFunction('AND', (_ctx, ...args) => {
      for (const arg of args) {
        if (arg instanceof RangeValue) {
          for (const v of arg.values) {
            if (v !== undefined && !v) return false;
          }
        } else {
          if (!arg) return false;
        }
      }
      return true;
    });

    this.registerFunction('OR', (_ctx, ...args) => {
      for (const arg of args) {
        if (arg instanceof RangeValue) {
          for (const v of arg.values) {
            if (v !== undefined && v) return true;
          }
        } else {
          if (arg) return true;
        }
      }
      return false;
    });

    this.registerFunction('NOT', (_ctx, val) => {
      return !val;
    });

    // ─── Conditional Aggregation ─────────────────────────

    this.registerFunction('SUMIF', (_ctx, range, criteria, sumRange?) => {
      const criteriaStr = String(criteria);
      const rangeArr: unknown[] = range instanceof RangeValue ? range.values : (Array.isArray(range) ? range : [range]);
      const sumArr: unknown[] | undefined = sumRange instanceof RangeValue ? sumRange.values : (Array.isArray(sumRange) ? sumRange : undefined);

      let total = 0;
      for (let i = 0; i < rangeArr.length; i++) {
        const val = rangeArr[i];
        const matches = val === undefined
          ? (criteriaStr === '' || criteriaStr === '=')
          : matchesCriteria(val, criteriaStr);
        if (matches) {
          if (sumArr) {
            total += Number(sumArr[i]) || 0;
          } else {
            total += Number(val) || 0;
          }
        }
      }
      return total;
    });

    this.registerFunction('COUNTIF', (_ctx, range, criteria) => {
      const criteriaStr = String(criteria);
      const rangeArr: unknown[] = range instanceof RangeValue ? range.values : (Array.isArray(range) ? range : [range]);

      let count = 0;
      for (const val of rangeArr) {
        if (val === undefined) {
          // Empty cells match empty-string criteria ("" or "=")
          if (criteriaStr === '' || criteriaStr === '=') count++;
        } else if (matchesCriteria(val, criteriaStr)) {
          count++;
        }
      }
      return count;
    });

    this.registerFunction('AVERAGEIF', (_ctx, range, criteria, avgRange?) => {
      const criteriaStr = String(criteria);
      const rangeArr: unknown[] = range instanceof RangeValue ? range.values : (Array.isArray(range) ? range : [range]);
      const avgArr: unknown[] | undefined = avgRange instanceof RangeValue ? avgRange.values : (Array.isArray(avgRange) ? avgRange : undefined);

      let total = 0;
      let count = 0;
      for (let i = 0; i < rangeArr.length; i++) {
        const val = rangeArr[i];
        if (val !== undefined && matchesCriteria(val, criteriaStr)) {
          const numVal = avgArr ? Number(avgArr[i]) : Number(val);
          if (!isNaN(numVal)) {
            total += numVal;
            count++;
          }
        }
      }
      if (count === 0) throw new Error('#DIV/0!');
      return total / count;
    });

    // ─── Lookup ─────────────────────────────────────────

    this.registerFunction('VLOOKUP', (_ctx, lookupValue, tableRange, colIndex, exactMatch?) => {
      if (!(tableRange instanceof RangeValue)) {
        throw new Error('#VALUE!');
      }
      const colIdx = Math.trunc(Number(colIndex));
      if (colIdx < 1 || colIdx > tableRange.cols) {
        throw new Error('#REF!');
      }
      // Excel convention: omitted or TRUE/1 = approximate match, FALSE/0 = exact match
      const isExact = exactMatch === false || exactMatch === 0;

      // Search first column
      const firstCol = tableRange.getCol(0);
      for (let r = 0; r < firstCol.length; r++) {
        if (isExact) {
          const numLookup = Number(lookupValue);
          const numCell = Number(firstCol[r]);
          const bothNum = !isNaN(numLookup) && !isNaN(numCell)
            && String(lookupValue).trim() !== '' && String(firstCol[r]).trim() !== '';
          if (bothNum ? numLookup === numCell : String(firstCol[r]).toLowerCase() === String(lookupValue).toLowerCase()) {
            return tableRange.get(r, colIdx - 1);
          }
        }
      }
      if (!isExact) {
        // Approximate match: largest value <= lookupValue (data assumed sorted ascending)
        const r = approximateMatch(firstCol, lookupValue);
        if (r !== -1) return tableRange.get(r, colIdx - 1);
      }
      throw new Error('#N/A');
    });

    this.registerFunction('HLOOKUP', (_ctx, lookupValue, tableRange, rowIndex, exactMatch?) => {
      if (!(tableRange instanceof RangeValue)) {
        throw new Error('#VALUE!');
      }
      const rowIdx = Math.trunc(Number(rowIndex));
      if (rowIdx < 1 || rowIdx > tableRange.rows) {
        throw new Error('#REF!');
      }
      // Excel convention: omitted or TRUE/1 = approximate match, FALSE/0 = exact match
      const isExact = exactMatch === false || exactMatch === 0;

      // Search first row
      const firstRow = tableRange.getRow(0);
      for (let c = 0; c < firstRow.length; c++) {
        if (isExact) {
          const numLookup = Number(lookupValue);
          const numCell = Number(firstRow[c]);
          const bothNum = !isNaN(numLookup) && !isNaN(numCell)
            && String(lookupValue).trim() !== '' && String(firstRow[c]).trim() !== '';
          if (bothNum ? numLookup === numCell : String(firstRow[c]).toLowerCase() === String(lookupValue).toLowerCase()) {
            return tableRange.get(rowIdx - 1, c);
          }
        }
      }
      if (!isExact) {
        const c = approximateMatch(firstRow, lookupValue);
        if (c !== -1) return tableRange.get(rowIdx - 1, c);
      }
      throw new Error('#N/A');
    });

    this.registerFunction('INDEX', (_ctx, rangeArg, rowNum, colNum?) => {
      if (rangeArg instanceof RangeValue) {
        const r = Math.trunc(Number(rowNum)) - 1;
        const c = colNum !== undefined ? Math.trunc(Number(colNum)) - 1 : 0;
        if (r < 0 || r >= rangeArg.rows || c < 0 || c >= rangeArg.cols) {
          throw new Error('#REF!');
        }
        const val = rangeArg.get(r, c);
        return val !== undefined ? val : 0;
      }
      throw new Error('#VALUE!');
    });

    this.registerFunction('MATCH', (_ctx, lookupValue, rangeArg, matchType?) => {
      let arr: unknown[];
      if (rangeArg instanceof RangeValue) {
        // Use flat values array for 1D lookup
        arr = rangeArg.values;
      } else if (Array.isArray(rangeArg)) {
        arr = rangeArg;
      } else {
        throw new Error('#VALUE!');
      }

      const mt = matchType !== undefined ? Number(matchType) : 1;

      if (mt === 0) {
        // Exact match
        for (let i = 0; i < arr.length; i++) {
          const numLookup = Number(lookupValue);
          const numCell = Number(arr[i]);
          const bothNum = !isNaN(numLookup) && !isNaN(numCell)
            && String(lookupValue).trim() !== '' && String(arr[i]).trim() !== '';
          if (bothNum ? numLookup === numCell : String(arr[i]).toLowerCase() === String(lookupValue).toLowerCase()) {
            return i + 1; // 1-indexed
          }
        }
        throw new Error('#N/A');
      } else if (mt === 1) {
        // Largest value <= lookupValue (data assumed sorted ascending)
        const match = approximateMatch(arr, lookupValue);
        if (match === -1) throw new Error('#N/A');
        return match + 1;
      } else {
        // mt === -1: Smallest value >= lookupValue (data assumed sorted descending)
        let lastMatch = -1;
        for (let i = 0; i < arr.length; i++) {
          if (lookupCompare(arr[i], lookupValue) >= 0) {
            lastMatch = i;
          }
        }
        if (lastMatch === -1) throw new Error('#N/A');
        return lastMatch + 1;
      }
    });

    // ─── Math ───────────────────────────────────────────

    this.registerFunction('MOD', (_ctx, num, divisor) => {
      const n = Number(num);
      const d = Number(divisor);
      if (d === 0) throw new Error('#DIV/0!');
      // Excel MOD: n - d * FLOOR(n/d). Result sign follows the divisor,
      // unlike JavaScript's % which follows the dividend.
      return n - d * Math.floor(n / d);
    });

    this.registerFunction('POWER', (_ctx, base, exp) => {
      const b = Number(base);
      const e = Number(exp);
      const result = Math.pow(b, e);
      if (!Number.isFinite(result) && Number.isFinite(b) && Number.isFinite(e)) {
        // 0^(-n) → Infinity → #DIV/0!, negative^fraction → NaN → #NUM!
        if (isNaN(result)) throw new Error('#NUM!');
        throw new Error('#DIV/0!');
      }
      return result;
    });

    this.registerFunction('CEILING', (_ctx, num, significance?) => {
      const n = Number(num);
      const sig = significance !== undefined ? Number(significance) : 1;
      if (sig === 0) return 0;
      // Round the quotient to 10 significant digits to neutralise IEEE 754
      // drift (e.g. 0.07/0.01 → 7.000000000000001 → 7) before ceiling.
      const q = parseFloat((n / sig).toPrecision(10));
      return Math.ceil(q) * sig;
    });

    this.registerFunction('FLOOR', (_ctx, num, significance?) => {
      const n = Number(num);
      const sig = significance !== undefined ? Number(significance) : 1;
      if (sig === 0) return 0;
      const q = parseFloat((n / sig).toPrecision(10));
      return Math.floor(q) * sig;
    });

    // ─── String ─────────────────────────────────────────

    this.registerFunction('LEFT', (_ctx, text, n?) => {
      const count = n !== undefined ? Number(n) : 1;
      if (!(count >= 0)) throw new Error('#VALUE!');
      return String(text).substring(0, count);
    });

    this.registerFunction('RIGHT', (_ctx, text, n?) => {
      const s = String(text);
      const count = n !== undefined ? Number(n) : 1;
      if (!(count >= 0)) throw new Error('#VALUE!');
      return s.substring(s.length - count);
    });

    this.registerFunction('MID', (_ctx, text, start, n) => {
      const s = String(text);
      const startNum = Number(start);
      const count = Number(n);
      if (!(startNum >= 1) || !(count >= 0)) throw new Error('#VALUE!');
      return s.substring(startNum - 1, startNum - 1 + count);
    });

    this.registerFunction('SUBSTITUTE', (_ctx, text, oldStr, newStr, instance?) => {
      const s = String(text);
      const old = String(oldStr);
      const replacement = String(newStr);

      // Excel: empty search string → return text unchanged
      if (old === '') return s;

      if (instance !== undefined) {
        const nth = Number(instance);
        if (!(nth >= 1)) throw new Error('#VALUE!');
        let count = 0;
        let idx = -1;
        let searchFrom = 0;
        while (searchFrom < s.length) {
          idx = s.indexOf(old, searchFrom);
          if (idx === -1) break;
          count++;
          if (count === nth) {
            return s.substring(0, idx) + replacement + s.substring(idx + old.length);
          }
          searchFrom = idx + 1;
        }
        return s; // nth occurrence not found, return unchanged
      }

      // Replace all occurrences
      return s.split(old).join(replacement);
    });

    this.registerFunction('FIND', (_ctx, search, text, start?) => {
      const s = String(text);
      const searchStr = String(search);
      const startNum = start !== undefined ? Number(start) : 1;
      if (!(startNum >= 1)) throw new Error('#VALUE!');
      const idx = s.indexOf(searchStr, startNum - 1);
      if (idx === -1) throw new Error('#VALUE!');
      return idx + 1; // 1-indexed
    });

    // ─── Conversion ─────────────────────────────────────

    this.registerFunction('TEXT', (_ctx, value, format) => {
      const num = Number(value);
      const fmt = String(format);

      if (isNaN(num)) return String(value);

      // Support common number formats
      if (fmt === '0') {
        // Round half away from zero to match Excel (Math.round rounds toward +Infinity)
        return (Math.sign(num) * Math.round(Math.abs(num))).toString();
      }
      if (fmt === '0.00' || fmt === '0.0') {
        const decimals = (fmt.split('.')[1] || '').length;
        return num.toFixed(decimals);
      }
      if (fmt === '#,##0' || fmt === '#,##0.00') {
        const decimals = fmt.includes('.') ? (fmt.split('.')[1] || '').length : 0;
        return num.toLocaleString('en-US', {
          minimumFractionDigits: decimals,
          maximumFractionDigits: decimals,
        });
      }

      return num.toString();
    });

    this.registerFunction('VALUE', (_ctx, text) => {
      if (typeof text === 'boolean') return text ? 1 : 0;
      const num = Number(String(text));
      if (isNaN(num)) throw new Error('#VALUE!');
      return num;
    });

    // ─── Date ───────────────────────────────────────────

    this.registerFunction('DATE', (_ctx, year, month, day) => {
      // Excel serial number: days since 1899-12-30
      let y = Number(year);
      const m = Number(month);
      const d = Number(day);
      // Excel two-digit year windowing: 0-29 → 2000+, 30-99 → 1900+
      if (y >= 0 && y <= 29) y += 2000;
      else if (y >= 30 && y <= 99) y += 1900;
      // Use setFullYear to bypass JS Date's own 0-99 → 1900+ mapping
      const date = new Date(0);
      date.setFullYear(y, m - 1, d);
      date.setHours(0, 0, 0, 0);
      const epoch = new Date(0);
      epoch.setFullYear(1899, 11, 30);
      epoch.setHours(0, 0, 0, 0);
      const diff = date.getTime() - epoch.getTime();
      return Math.round(diff / (1000 * 60 * 60 * 24));
    });

    this.registerFunction('NOW', () => {
      const now = new Date();
      const epoch = new Date(1899, 11, 30); // 1899-12-30
      const diff = now.getTime() - epoch.getTime();
      return diff / (1000 * 60 * 60 * 24);
    });
  }
}
