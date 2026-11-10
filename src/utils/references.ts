/**
 * Rewriting of cell references inside formula text, used when formulas are
 * copied or filled and when rows or columns are inserted or deleted. Built
 * on the formula tokenizer, so it sees exactly the references the engine does:
 * `A1`, `$A$1:B2`, whole columns and rows (`A:C`, `1:3`) and sheet-qualified
 * references (`Sheet2!A1`, `'My sheet'!A:A`).
 */
import { colToLetter } from '../types.js';
import { type Node, type RefEnd, type RefInfo, MAX_COLS, MAX_ROWS, parseFormula, tokenizeFormula } from '../engine/parser.js';

/** A row/column insertion or deletion. */
export interface StructureChange {
  /** Which axis changes: `'row'` or `'col'` */
  axis: 'row' | 'col';
  /** Whether rows/columns are inserted or deleted */
  kind: 'insert' | 'delete';
  /** 0-based index of the first inserted/deleted row or column */
  index: number;
  /** How many rows/columns are inserted or deleted (>= 1) */
  count: number;
}

/** Which sheet a rewrite applies to, for formulas in a workbook with several sheets. */
export interface ReferenceScope {
  /** Name of the sheet the formula is on (what its unqualified references point to) */
  formulaSheet?: string;
  /**
   * Name of the sheet whose rows/columns change. References to other sheets
   * are left alone. Default: the formula's own sheet.
   */
  changedSheet?: string;
}

/**
 * Replace every reference in a formula, leaving everything else (including
 * string literals) exactly as written. `replace` receives the parsed
 * reference and its source text split into the sheet prefix (`"Sheet2!"`,
 * or `""`) and the reference itself, and returns the replacement for the
 * whole reference including the prefix, or `null` to keep the original.
 * A formula that doesn't tokenize is returned unchanged.
 *
 * @param formula - The raw cell value, starting with `=`
 */
export function mapFormulaReferences(
  formula: string,
  replace: (ref: RefInfo, text: { sheet: string; ref: string }) => string | null
): string {
  if (!formula.startsWith('=')) return formula;
  const body = formula.slice(1);
  let tokens;
  try {
    tokens = tokenizeFormula(body);
  } catch {
    return formula;
  }
  let out = '=';
  let last = 0;
  for (const token of tokens) {
    if (token.type !== 'REF' || !token.ref) continue;
    const refStart = token.refStart ?? token.start;
    const replacement = replace(token.ref, {
      sheet: body.slice(token.start, refStart),
      ref: body.slice(refStart, token.end),
    });
    if (replacement === null) continue;
    out += body.slice(last, token.start) + replacement;
    last = token.end;
  }
  return last === 0 ? formula : out + body.slice(last);
}

/** The names a formula binds itself, with LET and LAMBDA (lower case). */
function localNames(node: Node, out: Set<string> = new Set()): Set<string> {
  switch (node.t) {
    case 'call': {
      const bound =
        node.name === 'LET'
          ? node.args.slice(0, -1).filter((_, i) => i % 2 === 0)
          : node.name === 'LAMBDA'
            ? node.args.slice(0, -1)
            : [];
      for (const arg of bound) if (arg?.t === 'name') out.add(arg.name.toLowerCase());
      for (const arg of node.args) if (arg) localNames(arg, out);
      break;
    }
    case 'invoke':
      localNames(node.callee, out);
      for (const arg of node.args) if (arg) localNames(arg, out);
      break;
    case 'unary':
      localNames(node.x, out);
      break;
    case 'bin':
      localNames(node.l, out);
      localNames(node.r, out);
      break;
    case 'array':
      for (const row of node.rows) for (const n of row) localNames(n, out);
      break;
  }
  return out;
}

/**
 * Whether renaming a defined name can't be carried through a formula
 * without changing what it means:
 *
 * - the formula mentions the name and binds the old or the new name itself
 *   with LET or LAMBDA, so some of those mentions are (or would become) the
 *   formula's own variable; or
 * - the formula calls the name (`=double(4)`) and the new name is a
 *   function, which the rewritten call would run instead (`=SUM(4)`).
 *
 * @param isFunction - Whether a (upper-case) name is a function, built-in or custom
 */
export function renameIsCaptured(
  formula: string,
  oldName: string,
  newName: string,
  isFunction: (name: string) => boolean = () => false
): boolean {
  if (!formula.startsWith('=')) return false;
  const body = formula.slice(1);
  const lower = oldName.toLowerCase();
  if (!body.toLowerCase().includes(lower)) return false;
  try {
    const uses = tokenizeFormula(body).filter(
      (t) => (t.type === 'NAME' || t.type === 'FUNC') && body.slice(t.start, t.end).toLowerCase() === lower
    );
    if (uses.length === 0) return false;
    // A call of the name would become a call of the function of the new name
    if (!isFunction(oldName.toUpperCase()) && isFunction(newName.toUpperCase()) && uses.some((t) => t.type === 'FUNC')) {
      return true;
    }
    const locals = localNames(parseFormula(body));
    return locals.has(newName.toLowerCase()) || locals.has(lower);
  } catch {
    return false;
  }
}

/**
 * Rewrite the uses of a defined name in a formula, whether it is read
 * (`=SUM(sales)`) or called (`=double(4)`). String literals are left alone.
 *
 * A formula is left unchanged when the rewrite could change what it means:
 * when it binds the old or the new name itself (`=LET(sales, ...)`), since
 * those uses are the formula's own variable. A call is only rewritten when
 * `isFunction` says no function of that name exists (`SUM(` stays `SUM(`
 * even when a name `Sum` is renamed).
 *
 * @param formula - The raw cell value; non-formulas are returned unchanged
 * @param isFunction - Whether a (upper-case) name is a function, built-in or custom
 */
export function renameNameInFormula(
  formula: string,
  oldName: string,
  newName: string,
  isFunction: (name: string) => boolean = () => false
): string {
  if (!formula.startsWith('=')) return formula;
  const body = formula.slice(1);
  const lower = oldName.toLowerCase();
  if (!body.toLowerCase().includes(lower)) return formula;
  let tokens;
  let locals: Set<string>;
  try {
    tokens = tokenizeFormula(body);
    locals = localNames(parseFormula(body));
  } catch {
    return formula;
  }
  if (locals.has(lower) || locals.has(newName.toLowerCase())) return formula;
  const callable = !isFunction(oldName.toUpperCase());
  let out = '=';
  let last = 0;
  for (const token of tokens) {
    if (token.type !== 'NAME' && !(token.type === 'FUNC' && callable)) continue;
    if (body.slice(token.start, token.end).toLowerCase() !== lower) continue;
    out += body.slice(last, token.start) + newName;
    last = token.end;
  }
  return last === 0 ? formula : out + body.slice(last);
}

function formatEnd(end: RefEnd): string {
  const col = end.col >= 0 ? `${end.colAbs ? '$' : ''}${colToLetter(end.col)}` : '';
  const row = end.row >= 0 ? `${end.rowAbs ? '$' : ''}${end.row + 1}` : '';
  return col + row;
}

/** The text of a reference (without any sheet prefix). */
export function formatReference(ref: RefInfo): string {
  return ref.kind === 'cell' ? formatEnd(ref.start) : `${formatEnd(ref.start)}:${formatEnd(ref.end)}`;
}

/**
 * Shift the relative references of a formula by an offset, as when a formula
 * is copied, pasted or filled into another cell. `$`-anchored parts stay put.
 * A reference (or either end of a range) pushed above row 1 or left of
 * column A becomes `#REF!`; a range with an invalid end becomes a single
 * `#REF!`, so the result always parses. String literals are left alone.
 *
 * @param formula - The raw cell value; non-formulas are returned unchanged
 * @param rowOffset - Rows to move relative references by
 * @param colOffset - Columns to move relative references by
 * @returns The shifted formula
 */
export function shiftFormulaReferences(formula: string, rowOffset: number, colOffset: number): string {
  if (!formula.startsWith('=')) return formula;
  const shift = (end: RefEnd): RefEnd => ({
    ...end,
    row: end.row < 0 || end.rowAbs ? end.row : end.row + rowOffset,
    col: end.col < 0 || end.colAbs ? end.col : end.col + colOffset,
  });
  // A missing part (-1) stays -1; a present part must stay on the sheet
  const valid = (before: RefEnd, after: RefEnd) =>
    (before.row < 0 || (after.row >= 0 && after.row < MAX_ROWS)) &&
    (before.col < 0 || (after.col >= 0 && after.col < MAX_COLS));
  return mapFormulaReferences(formula, (ref, text) => {
    const start = shift(ref.start);
    const end = shift(ref.end);
    if (!valid(ref.start, start) || !valid(ref.end, end)) return '#REF!';
    return text.sheet + formatReference({ ...ref, start, end });
  });
}

/**
 * New position of a single row/column index, or `null` when it was deleted.
 */
function shiftIndex(p: number, change: StructureChange): number | null {
  const { kind, index, count } = change;
  if (kind === 'insert') return p >= index ? p + count : p;
  if (p < index) return p;
  if (p >= index + count) return p - count;
  return null;
}

/**
 * Rewrite the references in a formula after rows or columns are inserted or
 * deleted, the way Excel does:
 *
 * - References at or after an insertion point move by `count`. A range that
 *   spans the insertion point grows.
 * - References after a deletion move back by `count`. A range that loses
 *   some of its rows/columns shrinks.
 * - A reference to a deleted cell, or a range that is deleted entirely,
 *   becomes `#REF!`.
 *
 * Absolute (`$`) references move too: they are fixed against copying, not
 * against the cells they point to moving. Text inside string literals is
 * left alone. Non-formulas (not starting with `=`) are returned unchanged.
 *
 * With several sheets, pass `scope`: only references to the sheet that
 * changed are rewritten (unqualified ones count as the formula's own sheet).
 *
 * @param formula - The raw cell value
 * @param change - The insertion or deletion
 * @param scope - The formula's sheet and the sheet that changed
 * @returns The rewritten formula
 */
export function adjustFormulaForStructureChange(
  formula: string,
  change: StructureChange,
  scope: ReferenceScope = {}
): string {
  if (!formula.startsWith('=')) return formula;
  const key = change.axis === 'row' ? 'row' : 'col';
  const own = scope.formulaSheet?.toLowerCase();
  const target = (scope.changedSheet ?? scope.formulaSheet)?.toLowerCase();

  return mapFormulaReferences(formula, (ref, text) => {
    const refSheet = ref.sheet !== undefined ? ref.sheet.toLowerCase() : own;
    if (refSheet !== target) return null;
    // A whole-column reference has no rows to move (and vice versa)
    if (ref.start[key] < 0) return null;

    if (ref.kind === 'cell') {
      const moved = shiftIndex(ref.start[key], change);
      if (moved === null) return '#REF!';
      if (moved === ref.start[key]) return null;
      const start = { ...ref.start, [key]: moved };
      return text.sheet + formatReference({ ...ref, start, end: start });
    }

    // Ranges: move each edge, clipping edges that fall inside a deletion.
    const loIsStart = ref.start[key] <= ref.end[key];
    const lo = loIsStart ? ref.start[key] : ref.end[key];
    const hi = loIsStart ? ref.end[key] : ref.start[key];
    let newLo: number;
    let newHi: number;
    if (change.kind === 'insert') {
      newLo = lo >= change.index ? lo + change.count : lo;
      newHi = hi >= change.index ? hi + change.count : hi;
    } else {
      const last = change.index + change.count - 1;
      if (lo >= change.index && hi <= last) return '#REF!';
      newLo = lo < change.index ? lo : lo > last ? lo - change.count : change.index;
      newHi = hi < change.index ? hi : hi > last ? hi - change.count : change.index - 1;
    }
    if (newLo === lo && newHi === hi) return null;
    const start = { ...ref.start, [key]: loIsStart ? newLo : newHi };
    const end = { ...ref.end, [key]: loIsStart ? newHi : newLo };
    return text.sheet + formatReference({ ...ref, start, end });
  });
}
