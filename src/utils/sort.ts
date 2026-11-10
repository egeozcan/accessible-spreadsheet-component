/**
 * Ordering of cell values for sorting, following Excel: ascending puts
 * numbers first (smallest to largest), then text (A to Z, ignoring case),
 * then FALSE, TRUE, then errors (kept in their original order). Descending
 * reverses that. Blank cells always go last, in either direction.
 */
import type { CellData, SortDirection } from '../types.js';

/** A cell value reduced to what sorting compares. */
export type SortValue =
  | { kind: 'number'; value: number }
  | { kind: 'text'; value: string }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'error' }
  | { kind: 'blank' };

const RANK: Record<SortValue['kind'], number> = {
  number: 0,
  text: 1,
  boolean: 2,
  error: 3,
  blank: 4,
};

// Case-insensitive but accent-sensitive, like Excel's default sort
const collator = new Intl.Collator(undefined, { sensitivity: 'accent' });

/**
 * The value a cell sorts by. `numeric` supplies the number behind a number
 * cell, since its `displayValue` may be formatted (`$1,234.50`).
 */
export function sortValueOf(cell: CellData | undefined, numeric: (cell: CellData) => number): SortValue {
  // A cell with no value of its own can still show one spilled into it by a formula
  if (!cell || (cell.rawValue === '' && cell.spilledFrom === undefined)) return { kind: 'blank' };
  switch (cell.type) {
    case 'number': {
      const value = numeric(cell);
      return Number.isNaN(value) ? { kind: 'text', value: cell.displayValue } : { kind: 'number', value };
    }
    case 'boolean':
      return { kind: 'boolean', value: cell.displayValue.toUpperCase() === 'TRUE' };
    case 'error':
      return { kind: 'error' };
    default:
      // A formula that evaluates to an empty string still counts as blank
      return cell.displayValue === '' ? { kind: 'blank' } : { kind: 'text', value: cell.displayValue };
  }
}

/**
 * Compare two sort values. Negative when `a` goes first. Blanks go last in
 * both directions; values that compare equal return 0, so a stable sort keeps
 * their original order.
 */
export function compareSortValues(a: SortValue, b: SortValue, direction: SortDirection = 'ascending'): number {
  if (a.kind === 'blank' || b.kind === 'blank') {
    return (a.kind === 'blank' ? 1 : 0) - (b.kind === 'blank' ? 1 : 0);
  }
  let result = RANK[a.kind] - RANK[b.kind];
  if (result === 0) {
    if (a.kind === 'number' && b.kind === 'number') result = a.value - b.value;
    else if (a.kind === 'text' && b.kind === 'text') result = collator.compare(a.value, b.value);
    else if (a.kind === 'boolean' && b.kind === 'boolean') result = Number(a.value) - Number(b.value);
  }
  return direction === 'descending' ? -result : result;
}
