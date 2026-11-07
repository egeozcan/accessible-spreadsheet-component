/**
 * Rewriting of cell references inside formula text, used when rows or
 * columns are inserted or deleted.
 */
import { colToLetter, letterToCol } from '../types.js';

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

interface RefParts {
  colAbs: string;
  col: number;
  rowAbs: string;
  row: number;
}

const REF = String.raw`(\$?)([A-Za-z]+)(\$?)(\d+)`;
// Lookarounds keep identifiers that merely contain a ref-like substring (e.g.
// a LOG10() call or a name like ABC_1) from being treated as references.
const REF_PATTERN = new RegExp(
  String.raw`(?<![A-Za-z0-9_.$])${REF}(?::${REF})?(?![A-Za-z0-9_.]|\s*\()`,
  'g'
);

/**
 * Replace every cell reference and range in a formula, skipping string
 * literals. `replace` receives the parsed reference (and the range end, for
 * `A1:B2`) and returns the replacement text, or `null` to keep the original.
 */
function mapReferences(
  formula: string,
  replace: (start: RefParts, end: RefParts | null) => string | null
): string {
  let out = '';
  let segment = '';
  const flush = () => {
    out += segment.replace(REF_PATTERN, (match, ...g: string[]) => {
      const start: RefParts = {
        colAbs: g[0],
        col: letterToCol(g[1].toUpperCase()),
        rowAbs: g[2],
        row: parseInt(g[3], 10) - 1,
      };
      const end: RefParts | null =
        g[5] === undefined
          ? null
          : {
              colAbs: g[4],
              col: letterToCol(g[5].toUpperCase()),
              rowAbs: g[6],
              row: parseInt(g[7], 10) - 1,
            };
      return replace(start, end) ?? match;
    });
    segment = '';
  };

  let i = 0;
  while (i < formula.length) {
    if (formula[i] === '"') {
      flush();
      // Copy the string literal verbatim; "" is an escaped quote inside it
      let j = i + 1;
      while (j < formula.length) {
        if (formula[j] === '"') {
          if (formula[j + 1] === '"') {
            j += 2;
            continue;
          }
          break;
        }
        j++;
      }
      out += formula.slice(i, j + 1);
      i = j + 1;
    } else {
      segment += formula[i];
      i++;
    }
  }
  flush();
  return out;
}

function formatRef(ref: RefParts): string {
  return `${ref.colAbs}${colToLetter(ref.col)}${ref.rowAbs}${ref.row + 1}`;
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
  const shift = (ref: RefParts): RefParts => ({
    ...ref,
    row: ref.rowAbs ? ref.row : ref.row + rowOffset,
    col: ref.colAbs ? ref.col : ref.col + colOffset,
  });
  const valid = (ref: RefParts) => ref.row >= 0 && ref.col >= 0;
  return mapReferences(formula, (start, end) => {
    const s = shift(start);
    if (!end) return valid(s) ? formatRef(s) : '#REF!';
    const e = shift(end);
    return valid(s) && valid(e) ? `${formatRef(s)}:${formatRef(e)}` : '#REF!';
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
 * @param formula - The raw cell value
 * @param change - The insertion or deletion
 * @returns The rewritten formula
 */
export function adjustFormulaForStructureChange(formula: string, change: StructureChange): string {
  if (!formula.startsWith('=')) return formula;
  const key = change.axis === 'row' ? 'row' : 'col';

  return mapReferences(formula, (start, end) => {
    if (!end) {
      const moved = shiftIndex(start[key], change);
      if (moved === null) return '#REF!';
      if (moved === start[key]) return null;
      return formatRef({ ...start, [key]: moved });
    }

    // Ranges: move each edge, clipping edges that fall inside a deletion.
    const loIsStart = start[key] <= end[key];
    const lo = loIsStart ? start[key] : end[key];
    const hi = loIsStart ? end[key] : start[key];
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
    const newStart = { ...start, [key]: loIsStart ? newLo : newHi };
    const newEnd = { ...end, [key]: loIsStart ? newHi : newLo };
    return `${formatRef(newStart)}:${formatRef(newEnd)}`;
  });
}
