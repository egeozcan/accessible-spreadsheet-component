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
  /** Font size in pixels */
  fontSize?: number;
  /** Numeric display format (e.g. currency, percent) */
  numberFormat?: NumberFormatOptions;
}

/** Available preset number format types */
export type NumberFormatType = 'number' | 'currency' | 'percent' | 'scientific';

/** Options for number formatting */
export interface NumberFormatOptions {
  type: NumberFormatType;
  /** Number of decimal places (default: 2) */
  decimals?: number;
  /** Currency symbol (default: '$'), only used when type is 'currency' */
  currencySymbol?: string;
  /** Whether to use thousands separators (default: true for 'number' and 'currency') */
  thousandsSep?: boolean;
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
  operation?: 'edit' | 'cut' | 'paste' | 'clear' | 'format';
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
    && a.thousandsSep === b.thousandsSep;
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
      return (value * 100).toFixed(decimals) + '%';
    case 'scientific':
      return value.toExponential(decimals);
    case 'currency': {
      const symbol = opts.currencySymbol ?? '$';
      const useSep = opts.thousandsSep !== false;
      const abs = Math.abs(value);
      const formatted = useSep ? addThousandsSep(abs.toFixed(decimals)) : abs.toFixed(decimals);
      return value < 0 ? `-${symbol}${formatted}` : `${symbol}${formatted}`;
    }
    case 'number': {
      const useSep = opts.thousandsSep !== false;
      const fixed = value.toFixed(decimals);
      return useSep ? addThousandsSep(fixed) : fixed;
    }
  }
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
