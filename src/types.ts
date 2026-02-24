/**
 * Coordinate System
 * Zero-indexed (0,0 is A1)
 */
export interface CellCoord {
  row: number;
  col: number;
}

/**
 * Data Structure for a Single Cell
 */
export interface CellData {
  rawValue: string;
  displayValue: string;
  type: 'text' | 'number' | 'boolean' | 'error';
  style?: Partial<CSSStyleDeclaration>;
}

/**
 * The Sparse Data Map
 * Key format: "${row}:${col}" (e.g., "0:0")
 */
export type GridData = Map<string, CellData>;

/**
 * Selection Range (normalized bounding box)
 */
export interface SelectionRange {
  start: CellCoord;
  end: CellCoord;
}

/**
 * Formula Context
 * Passed to custom functions to allow them to access the grid
 */
export interface FormulaContext {
  getCellValue: (ref: string) => unknown;
  getRangeValues: (start: string, end: string) => unknown[];
}

/**
 * A formula function that receives a context and arguments
 */
export type FormulaFunction = (ctx: FormulaContext, ...args: unknown[]) => unknown;

/**
 * Cell change event detail
 */
export interface CellChangeDetail {
  cellId: string;
  value: string;
  oldValue: string;
}

/**
 * Selection change event detail
 */
export interface SelectionChangeDetail {
  range: SelectionRange;
}

/**
 * Data change event detail (bulk updates)
 */
export interface DataChangeDetail {
  updates: Array<{ id: string; value: string }>;
  source?: 'user' | 'undo' | 'redo' | 'programmatic';
  operation?: 'edit' | 'cut' | 'paste' | 'clear';
}

/**
 * Creates a cell key from row and col indices
 */
export function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

/**
 * Parse a cell key back to coordinates
 */
export function parseKey(key: string): CellCoord {
  const [row, col] = key.split(':').map(Number);
  return { row, col };
}

/**
 * Convert a column index (0-based) to a letter (A, B, ... Z, AA, AB...)
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
 * Convert a column letter (A, B, ... Z, AA, AB...) to 0-based index
 */
export function letterToCol(letter: string): number {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col - 1;
}

/**
 * Convert a cell reference like "A1" to a CellCoord
 * Supports absolute/mixed references: $A$1, $A1, A$1
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
 * Convert a CellCoord to a cell reference like "A1"
 */
export function coordToRef(coord: CellCoord): string {
  return `${colToLetter(coord.col)}${coord.row + 1}`;
}
