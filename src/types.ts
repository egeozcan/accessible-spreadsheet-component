/**
 * A zero-indexed cell coordinate.
 * (0, 0) corresponds to cell A1.
 */
export interface CellCoord {
  /** 0-indexed row number (0 = row 1) */
  row: number;
  /** 0-indexed column number (0 = column A) */
  col: number;
}

/**
 * Cell-level formatting options.
 * All properties are optional; only set properties are applied.
 */
export interface CellFormat {
  /** Whether the cell text is bold */
  bold?: boolean;
  /** Whether the cell text is italic */
  italic?: boolean;
  /** Whether the cell text is underlined */
  underline?: boolean;
  /** Whether the cell text has a strikethrough */
  strikethrough?: boolean;
  /** CSS color string for text (e.g. "#ff0000", "red") */
  textColor?: string;
  /** CSS color string for cell background */
  backgroundColor?: string;
  /** Horizontal text alignment */
  textAlign?: 'left' | 'center' | 'right';
  /** Vertical text alignment within the cell (default: middle) */
  verticalAlign?: 'top' | 'middle' | 'bottom';
  /**
   * Wrap long text onto multiple lines instead of truncating it with an
   * ellipsis. Rows keep their fixed height (`--ls-cell-height`), so only the
   * lines that fit are visible.
   */
  wrap?: boolean;
  /** Font size in pixels */
  fontSize?: number;
  /** Numeric display format (e.g. currency, percent) */
  numberFormat?: NumberFormatOptions;
}

/**
 * Available preset number format types. `date` renders a date serial number
 * (days since 1899-12-30, as returned by `DATE()`, `TODAY()` and `NOW()`)
 * using `datePattern`.
 */
export type NumberFormatType = 'number' | 'currency' | 'percent' | 'scientific' | 'date';

/** Options for number formatting */
export interface NumberFormatOptions {
  type: NumberFormatType;
  /** Number of decimal places (default: 2) */
  decimals?: number;
  /** Currency symbol (default: '$'), only used when type is 'currency' */
  currencySymbol?: string;
  /** Whether to use thousands separators (default: true for 'number' and 'currency') */
  thousandsSep?: boolean;
  /**
   * Pattern for the 'date' type (default: `'YYYY-MM-DD'`). Tokens: `YYYY`, `YY`,
   * `MM`, `M`, `DD`, `D`, `HH`, `H`, `mm`, `ss`. Other characters are literal,
   * e.g. `'DD/MM/YYYY'` or `'YYYY-MM-DD HH:mm'`.
   */
  datePattern?: string;
}

/**
 * Data for a single cell in the grid.
 */
export interface CellData {
  /** The raw user-entered value (may be a formula starting with "=") */
  rawValue: string;
  /** The computed display string shown in the cell */
  displayValue: string;
  /** The resolved data type after evaluation */
  type: 'text' | 'number' | 'boolean' | 'error';
  /** Optional visual formatting applied to the cell */
  format?: CellFormat;
  /**
   * When true, the user cannot edit, clear, paste into, fill or format this
   * cell, and it is exposed with `aria-readonly="true"`. Programmatic API
   * calls (`setCellFormat`, `fromCSV`, `setData`, ...) are not restricted.
   * The flag moves with the cell when rows or columns are inserted or deleted.
   */
  readOnly?: boolean;
}

/**
 * Sparse data map for the entire grid.
 * Keys use the format `"${row}:${col}"` (e.g. `"0:0"` for A1, `"2:3"` for D3).
 * Only cells with content are stored; absent keys represent empty cells.
 */
export type GridData = Map<string, CellData>;

/**
 * A normalized rectangular selection range.
 * `start` is always the top-left corner; `end` is always the bottom-right.
 */
export interface SelectionRange {
  /** Top-left corner of the selection (minimum row and column) */
  start: CellCoord;
  /** Bottom-right corner of the selection (maximum row and column) */
  end: CellCoord;
}

/**
 * Context object passed to custom formula functions, providing access to grid data.
 */
export interface FormulaContext {
  /** Retrieve the evaluated value of a single cell by reference (e.g. "A1"). */
  getCellValue: (ref: string) => unknown;
  /** Retrieve a flat array of values for a rectangular range (e.g. "A1", "B5"). */
  getRangeValues: (start: string, end: string) => unknown[];
}

/**
 * Signature for user-defined formula functions.
 *
 * @param ctx - Provides access to cell/range values from the grid.
 * @param args - Variadic evaluated arguments passed to the function in the formula.
 * @returns The computed result (number, string, boolean, or error string).
 */
export type FormulaFunction = (ctx: FormulaContext, ...args: unknown[]) => unknown;

/**
 * Detail for the `cell-change` event, fired when a single cell is edited.
 */
export interface CellChangeDetail {
  /** Cell key in `"row:col"` format */
  cellId: string;
  /** New raw value of the cell */
  value: string;
  /** Previous raw value of the cell */
  oldValue: string;
}

/**
 * Detail for the `selection-change` event, fired when the selected range changes.
 */
export interface SelectionChangeDetail {
  /** The current normalized selection range */
  range: SelectionRange;
}

/**
 * Detail for the `data-change` event, fired on bulk updates (paste, cut, clear, undo/redo).
 */
export interface DataChangeDetail {
  /** Array of cell updates, each with a cell key and new raw value */
  updates: Array<{ id: string; value: string }>;
  /** What triggered the change */
  source?: 'user' | 'undo' | 'redo' | 'programmatic';
  /** The type of operation that caused the change */
  operation?: DataChangeOperation;
}

/** Operations reported in `data-change` and `before-change` events. */
export type DataChangeOperation =
  | 'edit'
  | 'cut'
  | 'paste'
  | 'clear'
  | 'format'
  | 'fill'
  | 'import'
  | 'insert-rows'
  | 'delete-rows'
  | 'insert-columns'
  | 'delete-columns'
  | 'sort';

/**
 * Detail for the cancelable `cell-edit-start` event, fired before the inline
 * editor opens. Call `preventDefault()` to keep the cell from being edited.
 */
export interface CellEditStartDetail {
  /** Cell key in `"row:col"` format */
  cellId: string;
  /** The cell's current raw value */
  value: string;
}

/** One proposed cell write in a `before-change` event. */
export interface ProposedChange {
  /** Cell key in `"row:col"` format */
  cellId: string;
  /** The value about to be written. Assign to it to write something else. */
  value: string;
  /** The cell's current raw value */
  oldValue: string;
}

/**
 * Detail for the cancelable `before-change` event, fired before a user edit,
 * paste, cut, clear or fill is applied (not for undo/redo or API calls).
 *
 * - `preventDefault()` cancels the whole change.
 * - Assigning `changes[i].value` writes a different value to that cell.
 * - Removing an entry from `changes` (e.g. with `splice`) skips that cell.
 */
export interface BeforeChangeDetail {
  changes: ProposedChange[];
  operation: DataChangeOperation;
}

/**
 * Validator for user-entered values. Return an error message to reject the
 * value, or `null`/`undefined` to accept it. Empty values (clearing a cell)
 * are not validated.
 */
export type CellValidator = (cellId: string, value: string) => string | null | undefined;

/**
 * Detail for the `validation-error` event, fired when `validateCell` rejects
 * one or more values.
 */
export interface ValidationErrorDetail {
  errors: Array<{ cellId: string; value: string; message: string }>;
  operation: DataChangeOperation;
}

/**
 * Detail for `structure-change` (after rows/columns were inserted or deleted)
 * and the cancelable `before-structure-change` (before a user-initiated one).
 */
export interface StructureChangeDetail {
  /** `'row'` or `'col'` */
  axis: 'row' | 'col';
  /** `'insert'` or `'delete'` */
  kind: 'insert' | 'delete';
  /** 0-based index of the first inserted/deleted row or column */
  index: number;
  /** Number of rows/columns inserted or deleted */
  count: number;
  /** What triggered the change */
  source: 'user' | 'undo' | 'redo' | 'programmatic';
}

/** Sort order: `'ascending'` (A to Z, smallest first) or `'descending'`. */
export type SortDirection = 'ascending' | 'descending';

/** One sort key: the column to sort by and its direction. */
export interface SortKey {
  /** 0-based column index; must be inside the sorted range */
  col: number;
  /** Default: `'ascending'` */
  direction?: SortDirection;
}

/** Options for `sort()`. */
export interface SortOptions {
  /** Keep the range's first row in place as a header row. Default: false */
  hasHeader?: boolean;
}

/**
 * Detail for `sort-change` (after a range was sorted, and on undo/redo of a
 * sort) and the cancelable `before-sort` (before a user-initiated sort).
 * On undo, it describes the sort that was undone.
 */
export interface SortDetail {
  /** The range that was sorted, including the header row when `hasHeader` */
  range: SelectionRange;
  /** Sort keys in priority order */
  keys: Array<{ col: number; direction: SortDirection }>;
  /** Whether the range's first row was kept in place as a header */
  hasHeader: boolean;
  /** What triggered the sort */
  source: 'user' | 'undo' | 'redo' | 'programmatic';
}

/**
 * Detail for `column-resize` (after column widths changed, also on undo/redo)
 * and the cancelable `before-column-resize` (before a user resize: a drag
 * of a header border, Alt+Shift+Left/Right, or auto-fit).
 */
export interface ColumnResizeDetail {
  /** The columns whose width changes, with their new width in px */
  columns: Array<{ col: number; width: number }>;
  /** What triggered the resize */
  source: 'user' | 'undo' | 'redo' | 'programmatic';
}

/**
 * Detail for the `freeze-change` event, fired when the frozen rows/columns
 * change through `freezePanes()`, Alt+Shift+F, or rows/columns inserted or
 * deleted inside the frozen area (not when the properties are set directly).
 */
export interface FreezeChangeDetail {
  /** Number of frozen rows at the top */
  rows: number;
  /** Number of frozen columns at the left */
  cols: number;
  /** What triggered the change */
  source: 'user' | 'undo' | 'redo' | 'programmatic';
}

/** One column's filter: the displayed values whose rows stay visible. */
export interface FilterColumn {
  /** 0-based column index, inside the filter range */
  col: number;
  /** Displayed values (as shown, number formats included) to keep; `''` keeps blank cells */
  values: string[];
}

/** The current filter, as returned by `getFilter()`. */
export interface FilterInfo {
  /** The filtered range; its first row is the header row, which is never hidden */
  range: SelectionRange;
  /** Columns with criteria, in column order */
  columns: FilterColumn[];
  /** Rows the filter hides, ascending */
  hiddenRows: number[];
}

/**
 * Detail for `filter-change` (after a filter was added, changed, reapplied
 * or removed, also on undo/redo and when rows/columns move it) and the
 * cancelable `before-filter` (before a user change, describing the result).
 */
export interface FilterChangeDetail {
  /** The filtered range, or `null` when the filter was removed */
  range: SelectionRange | null;
  /** Columns with criteria, in column order */
  columns: FilterColumn[];
  /** Rows the filter hides, ascending */
  hiddenRows: number[];
  /** What triggered the change */
  source: 'user' | 'undo' | 'redo' | 'programmatic';
}

/**
 * Detail for the `format-change` event, fired when cell formatting is modified.
 */
export interface FormatChangeDetail {
  /** Cell keys that were affected */
  cellIds: string[];
  /** The format properties that were applied (empty object for clear) */
  format: CellFormat;
  /** What triggered the format change */
  source: 'user' | 'undo' | 'redo' | 'programmatic';
}

/**
 * Compare two CellFormat objects for deep equality (key-order independent).
 * Deep-compares the `numberFormat` sub-object.
 *
 * @param a - First format (or undefined)
 * @param b - Second format (or undefined)
 * @returns `true` if both formats are structurally equal
 */
export function formatsEqual(a: CellFormat | undefined, b: CellFormat | undefined): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  const keysA = Object.keys(a) as (keyof CellFormat)[];
  const keysB = Object.keys(b) as (keyof CellFormat)[];
  if (keysA.length !== keysB.length) return false;
  return keysA.every((k) => {
    if (k === 'numberFormat') {
      return numberFormatEqual(a.numberFormat, b.numberFormat);
    }
    return a[k] === b[k];
  });
}

function numberFormatEqual(a: NumberFormatOptions | undefined, b: NumberFormatOptions | undefined): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return a.type === b.type
    && a.decimals === b.decimals
    && a.currencySymbol === b.currencySymbol
    && a.thousandsSep === b.thousandsSep
    && a.datePattern === b.datePattern;
}

/**
 * Format a numeric value according to the given number format options.
 *
 * @param value - The numeric value to format
 * @param opts - Formatting options (type, decimals, currency symbol, thousands separator)
 * @returns The formatted string representation
 */
export function formatNumber(value: number, opts: NumberFormatOptions): string {
  const decimals = opts.decimals ?? 2;
  switch (opts.type) {
    case 'percent':
      return stripNegZero((value * 100).toFixed(decimals)) + '%';
    case 'scientific':
      return value.toExponential(decimals);
    case 'currency': {
      const symbol = opts.currencySymbol ?? '$';
      const useSep = opts.thousandsSep !== false;
      const abs = Math.abs(value);
      const formatted = useSep ? addThousandsSep(abs.toFixed(decimals)) : abs.toFixed(decimals);
      // Suppress negative sign when the formatted value rounds to zero
      return value < 0 && parseFloat(formatted) !== 0 ? `-${symbol}${formatted}` : `${symbol}${formatted}`;
    }
    case 'number': {
      const useSep = opts.thousandsSep !== false;
      const fixed = stripNegZero(value.toFixed(decimals));
      return useSep ? addThousandsSep(fixed) : fixed;
    }
    case 'date':
      return formatDateSerial(value, opts.datePattern ?? 'YYYY-MM-DD');
  }
}

/** Milliseconds timestamp of the date-serial epoch (serial 0 = 1899-12-30), in UTC. */
const DATE_SERIAL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

/**
 * Convert a date serial number (days since 1899-12-30; the fraction is the
 * time of day) to a `Date` whose UTC fields hold the wall-clock date and time.
 * Read it with `getUTCFullYear()`, `getUTCHours()`, etc.
 *
 * @param serial - Date serial number
 * @returns The corresponding date, truncated to the second
 */
export function dateSerialToDate(serial: number): Date {
  // Round to the millisecond to absorb float drift (0.75 → exactly 18:00),
  // then truncate to the second so 23:59:59.9 never displays as the next day.
  const ms = Math.round(serial * MS_PER_DAY);
  return new Date(DATE_SERIAL_EPOCH_MS + Math.floor(ms / 1000) * 1000);
}

/**
 * Convert wall-clock date/time fields to a date serial number. Uses UTC
 * arithmetic so daylight-saving and historical timezone offsets never shift
 * the result.
 *
 * @param year - Full year (e.g. 2024; 0-99 are not remapped)
 * @param month - 0-indexed month
 * @param day - Day of month (overflow rolls into the next month)
 * @param hours - Hours (default 0)
 * @param minutes - Minutes (default 0)
 * @param seconds - Seconds (default 0)
 * @returns The date serial number
 */
export function dateToSerial(
  year: number,
  month: number,
  day: number,
  hours = 0,
  minutes = 0,
  seconds = 0
): number {
  const d = new Date(0);
  d.setUTCFullYear(year, month, day);
  d.setUTCHours(hours, minutes, seconds, 0);
  return (d.getTime() - DATE_SERIAL_EPOCH_MS) / MS_PER_DAY;
}

/** Render a date serial number with a `YYYY-MM-DD`-style pattern. */
function formatDateSerial(serial: number, pattern: string): string {
  if (!Number.isFinite(serial)) return String(serial);
  const d = dateSerialToDate(serial);
  const pad = (n: number, len = 2) => String(n).padStart(len, '0');
  const tokens: Record<string, string> = {
    YYYY: pad(d.getUTCFullYear(), 4),
    YY: pad(d.getUTCFullYear() % 100),
    MM: pad(d.getUTCMonth() + 1),
    M: String(d.getUTCMonth() + 1),
    DD: pad(d.getUTCDate()),
    D: String(d.getUTCDate()),
    HH: pad(d.getUTCHours()),
    H: String(d.getUTCHours()),
    mm: pad(d.getUTCMinutes()),
    ss: pad(d.getUTCSeconds()),
  };
  return pattern.replace(/YYYY|YY|MM|M|DD|D|HH|H|mm|ss/g, (t) => tokens[t]);
}

/** Remove negative sign from "-0", "-0.00", etc. produced by toFixed */
function stripNegZero(s: string): string {
  return s === '-0' || parseFloat(s) === 0 && s.startsWith('-') ? s.slice(1) : s;
}

function addThousandsSep(numStr: string): string {
  const [intPart, decPart] = numStr.split('.');
  const withCommas = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return decPart !== undefined ? `${withCommas}.${decPart}` : withCommas;
}

/**
 * Creates a cell key string from row and column indices.
 *
 * @param row - 0-indexed row number
 * @param col - 0-indexed column number
 * @returns A key in `"row:col"` format (e.g. `"0:0"`)
 */
export function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

/**
 * Parse a cell key string back to row/col coordinates.
 *
 * @param key - A key in `"row:col"` format (e.g. `"0:0"`)
 * @returns The parsed coordinate
 */
export function parseKey(key: string): CellCoord {
  const [row, col] = key.split(':').map(Number);
  return { row, col };
}

/**
 * Convert a 0-based column index to a column letter (A, B, ... Z, AA, AB...).
 *
 * @param col - 0-indexed column number
 * @returns The column letter(s) (e.g. 0 -> "A", 25 -> "Z", 26 -> "AA")
 */
export function colToLetter(col: number): string {
  let letter = '';
  let c = col;
  while (c >= 0) {
    letter = String.fromCharCode((c % 26) + 65) + letter;
    c = Math.floor(c / 26) - 1;
  }
  return letter;
}

/**
 * Convert a column letter (A, B, ... Z, AA, AB...) to a 0-based index.
 *
 * @param letter - Column letter(s) (e.g. "A", "Z", "AA")
 * @returns 0-indexed column number (e.g. "A" -> 0, "Z" -> 25, "AA" -> 26)
 */
export function letterToCol(letter: string): number {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col - 1;
}

/**
 * Convert a cell reference string to a CellCoord.
 * Supports absolute/mixed references: `$A$1`, `$A1`, `A$1`.
 *
 * @param ref - Cell reference (e.g. "A1", "$B$3")
 * @returns The 0-indexed coordinate
 * @throws {Error} If the reference format is invalid
 */
export function refToCoord(ref: string): CellCoord {
  const stripped = ref.replace(/\$/g, '');
  const match = stripped.match(/^([A-Z]+)(\d+)$/i);
  if (!match) throw new Error(`Invalid cell reference: ${ref}`);
  const col = letterToCol(match[1].toUpperCase());
  const row = parseInt(match[2], 10) - 1; // 1-indexed to 0-indexed
  return { row, col };
}

/**
 * Convert a CellCoord to a cell reference string like "A1".
 *
 * @param coord - The 0-indexed coordinate
 * @returns Cell reference string (e.g. `{row: 0, col: 0}` -> "A1")
 */
export function coordToRef(coord: CellCoord): string {
  return `${colToLetter(coord.col)}${coord.row + 1}`;
}
