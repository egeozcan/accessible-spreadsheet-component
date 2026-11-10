import {
  type CellData,
  type CellFormat,
  type GridData,
  type NumberFormatOptions,
  type SelectionRange,
  cellKey,
} from '../types.js';
import { parseDelimited, quoteDelimitedField } from '../utils/delimited.js';
import { shiftFormulaReferences } from '../utils/references.js';
import { arrayToFormula } from '../utils/list.js';
import { ERROR_CODE_LIST } from '../engine/parser.js';

/** Maps the index of a pasted row to the grid row it goes to (-1 drops it). */
export type RowMapper = (index: number) => number;

/** Every row of a range, ascending. */
function rangeRows(range: SelectionRange): number[] {
  const rows: number[] = [];
  for (let r = range.start.row; r <= range.end.row; r++) rows.push(r);
  return rows;
}

/** A single cell update produced by a paste operation. */
export interface PasteUpdate {
  /** Cell key in `"row:col"` format */
  id: string;
  /** Raw value to write to the cell */
  value: string;
  /** Optional format to apply (from HTML or internal paste) */
  format?: CellFormat;
}

/** Options for `ClipboardManager.paste()`. */
export interface PasteOptions {
  /**
   * Paste values instead of formulas: formulas copied from this grid paste
   * as the values they had when copied. Formats are still returned; callers
   * that paste values only drop them.
   */
  values?: boolean;
  /** Swap rows and columns: the first copied row becomes the first column. */
  transpose?: boolean;
}

/** Swap the rows and columns of a ragged 2-D array, padding short rows with `empty`. */
function transposeRows<T>(rows: T[][], empty: () => T): T[][] {
  const width = rows.reduce((w, r) => Math.max(w, r.length), 0);
  return Array.from({ length: width }, (_, c) => rows.map((r) => (c < r.length ? r[c] : empty())));
}

/**
 * ClipboardManager - Handles copy, cut, and paste operations
 * using the system clipboard API with TSV and HTML table formats.
 */
/**
 * What to write into a cell so it holds a computed value as a value. Text
 * that would be read back as something else (`00456` as a number, `TRUE`,
 * `#N/A`, `=A1`) is written as a text formula, `="00456"`, so it stays text.
 */
function literalFor(value: string, type: CellData['type']): string {
  if (type !== 'text' || value === '') return value;
  const readsAsOther =
    value.startsWith('=') ||
    (value.trim() !== '' && !isNaN(Number(value))) ||
    /^(true|false)$/i.test(value) ||
    ERROR_CODE_LIST.includes(value);
  return readsAsOther ? `="${value.replace(/"/g, '""')}"` : value;
}

/**
 * What to write into a cell so it holds what `cell` computed (or what was
 * spilled into it): several values as an array constant, anything else as a literal.
 */
export function computedLiteral(
  key: string,
  cell: CellData,
  valueOf?: (key: string, cell: CellData) => string
): string {
  if (cell.array) return arrayToFormula(cell.array);
  return literalFor(valueOf ? valueOf(key, cell) : cell.displayValue, cell.type);
}

export class ClipboardManager {
  /** Source range and data from the last internal copy/cut, used for reference adjustment on paste */
  private _copySourceRange: SelectionRange | null = null;
  /** The rows of `_copySourceRange` that were copied; null for all of them */
  private _copySourceRows: number[] | null = null;
  private _copySourceData: GridData | null = null;
  /** Computed values of the copied formula cells, for pasting values */
  private _copySourceValues: Map<string, string> | null = null;
  /** The TSV text we last wrote to the clipboard, used to detect external overwrites */
  private _lastWrittenText: string | null = null;

  /**
   * Copy the selected range to the system clipboard.
   * Writes two formats: `text/html` (table with inline styles and data-format attributes)
   * and `text/plain` (TSV). Falls back to text-only if ClipboardItem API is unavailable.
   *
   * @param data - The grid data to read cell values from
   * @param range - The rectangular range to copy
   * @param rows - The rows of the range to copy, ascending (default: all of
   *   them). Rows left out (e.g. hidden by a filter) are skipped, so the copy
   *   holds the remaining rows next to each other.
   * @param valueOf - The computed, unformatted value of a formula cell, kept
   *   so a later values-only paste can paste it (default: its display value)
   */
  async copy(
    data: GridData,
    range: SelectionRange,
    rows?: number[],
    valueOf?: (key: string, cell: CellData) => string
  ): Promise<void> {
    this._copySourceRange = range;
    const copied = rows ?? rangeRows(range);
    this._copySourceRows = rows ?? null;
    // Snapshot only the cells in the range so later mutations (e.g. cut) don't invalidate the data
    const snapshot: GridData = new Map();
    const values = new Map<string, string>();
    for (const r of copied) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        const key = cellKey(r, c);
        const cell = data.get(key);
        if (cell && cell.spilledFrom !== undefined && cell.rawValue === '') {
          const { spilledFrom: _spilledFrom, ...rest } = cell;
          const value = computedLiteral(key, cell, valueOf);
          const [anchorRow, anchorCol] = cell.spilledFrom.split(':').map(Number);
          if (copied.includes(anchorRow) && anchorCol >= range.start.col && anchorCol <= range.end.col) {
            // Copied together with its formula: a paste leaves the cell to the formula, which
            // spills again; a values-only paste writes the value it had
            snapshot.set(key, { ...rest, rawValue: '' });
            values.set(key, value);
          } else {
            // On its own, a value spilled into the cell by another cell's formula is copied as a plain value
            snapshot.set(key, { ...rest, rawValue: value });
            values.set(key, value);
          }
        } else if (cell) {
          snapshot.set(key, { ...cell });
          if (cell.rawValue.startsWith('=')) {
            values.set(key, computedLiteral(key, cell, valueOf));
          }
        }
      }
    }
    this._copySourceData = snapshot;
    this._copySourceValues = values;
    const tsv = this.serializeRange(data, range, copied);
    this._lastWrittenText = tsv;
    const htmlStr = this._serializeRangeAsHTML(data, range, copied, valueOf);

    try {
      // Try writing both HTML and plain text using ClipboardItem
      const htmlBlob = new Blob([htmlStr], { type: 'text/html' });
      const textBlob = new Blob([tsv], { type: 'text/plain' });
      const item = new ClipboardItem({
        'text/html': htmlBlob,
        'text/plain': textBlob,
      });
      await navigator.clipboard.write([item]);
    } catch {
      // Fall back to text-only clipboard write
      try {
        await navigator.clipboard.writeText(tsv);
      } catch {
        this.fallbackCopy(tsv);
      }
    }
  }

  /**
   * Cut: copy to clipboard and return the cell keys that should be cleared by the caller.
   *
   * @param data - The grid data to read cell values from
   * @param range - The rectangular range to cut
   * @param rows - The rows of the range to cut (default: all), as for `copy()`
   * @param valueOf - As for `copy()`
   * @returns Array of cell keys (in `"row:col"` format) that were included in the cut
   */
  async cut(
    data: GridData,
    range: SelectionRange,
    rows?: number[],
    valueOf?: (key: string, cell: CellData) => string
  ): Promise<string[]> {
    const cutRows = rows ?? rangeRows(range);
    await this.copy(data, range, cutRows, valueOf);
    const keys: string[] = [];
    for (const r of cutRows) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        keys.push(cellKey(r, c));
      }
    }
    return keys;
  }

  /**
   * Read from the system clipboard and return cell updates for pasting.
   * Priority chain: internal copy data (with formula reference adjustment)
   * -> HTML table -> plain text TSV.
   *
   * @param targetRow - 0-indexed row of the top-left paste destination
   * @param targetCol - 0-indexed column of the top-left paste destination
   * @param maxRows - Grid row count (for bounds clamping)
   * @param maxCols - Grid column count (for bounds clamping)
   * @param rowFor - The grid row the `i`-th pasted row goes to, or -1 to drop
   *   it (default: `targetRow + i`). Lets a paste skip rows hidden by a filter.
   *   With `transpose`, `i` counts the pasted (transposed) rows.
   * @param options - Paste values instead of formulas, or transpose
   * @returns Array of cell updates to apply, or null if clipboard is empty/inaccessible
   */
  async paste(
    targetRow: number,
    targetCol: number,
    maxRows: number,
    maxCols: number,
    rowFor?: RowMapper,
    options: PasteOptions = {}
  ): Promise<PasteUpdate[] | null> {
    const toRow = rowFor ?? ((i: number) => targetRow + i);
    // Try internal paste with formula reference adjustment, but only if
    // the system clipboard still contains what we wrote (i.e. the user
    // hasn't copied something from an external app since the last copy).
    if (this._copySourceRange && this._copySourceData) {
      let clipboardMatchesInternal = false;
      try {
        const currentText = await navigator.clipboard.readText();
        // Compare with normalized line endings: some platforms rewrite \n as \r\n
        const normalize = (t: string | null) => t?.replace(/\r\n?/g, '\n');
        clipboardMatchesInternal = normalize(currentText) === normalize(this._lastWrittenText);
      } catch {
        // Clipboard read denied — assume internal data is still valid
        clipboardMatchesInternal = true;
      }
      if (clipboardMatchesInternal) {
        const internalResult = this._pasteInternal(targetRow, targetCol, maxRows, maxCols, toRow, options);
        if (internalResult) return internalResult;
      }
    }

    // Try reading HTML from clipboard first
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (item.types.includes('text/html')) {
          const blob = await item.getType('text/html');
          const htmlText = await blob.text();
          const parsed = this.parseHTMLTableWithFormat(htmlText, options.values);
          if (parsed) {
            const rows = options.transpose ? transposeRows(parsed, () => ({ value: '' })) : parsed;
            return this._applyParsedRowsWithFormat(rows, targetCol, maxRows, maxCols, toRow);
          }
        }
      }
    } catch {
      // clipboard.read() not available or denied; fall through to text
    }

    // Fall back to plain text / TSV
    let text: string;
    try {
      text = await navigator.clipboard.readText();
    } catch {
      return null;
    }

    if (options.transpose) {
      const rows = parseDelimited(text, '\t').map((r) => r.map((value) => ({ value })));
      return this._applyParsedRowsWithFormat(transposeRows(rows, () => ({ value: '' })), targetCol, maxRows, maxCols, toRow);
    }
    return this.parseTSV(text, targetRow, targetCol, maxRows, maxCols, toRow);
  }

  /**
   * Paste using internally stored copy data with formula reference adjustment.
   * Returns null if no internal data is available.
   */
  private _pasteInternal(
    targetRow: number,
    targetCol: number,
    maxRows: number,
    maxCols: number,
    rowFor: RowMapper = (i) => targetRow + i,
    options: PasteOptions = {}
  ): PasteUpdate[] | null {
    const srcRange = this._copySourceRange;
    const srcData = this._copySourceData;
    if (!srcRange || !srcData) return null;

    const srcRows = this._copySourceRows ?? rangeRows(srcRange);
    const updates: PasteUpdate[] = [];

    srcRows.forEach((r, i) => {
      for (let c = srcRange.start.col; c <= srcRange.end.col; c++) {
        const j = c - srcRange.start.col;
        // Transposed, the i-th copied row becomes the i-th pasted column
        const newRow = rowFor(options.transpose ? j : i);
        const newCol = targetCol + (options.transpose ? i : j);
        if (newRow < 0 || newRow >= maxRows || newCol >= maxCols) continue;

        const key = cellKey(r, c);
        const cell = srcData.get(key);
        let value = cell?.rawValue ?? '';

        if (options.values && this._copySourceValues?.has(key)) {
          // What a formula showed (or spilled into this cell) when it was copied
          value = this._copySourceValues.get(key)!;
        } else if (value.startsWith('=')) {
          if (options.values) {
            value = cell?.displayValue ?? '';
          } else {
            // Each cell shifts its references by how far it actually moved,
            // which differs per row when copied or pasted rows skip hidden ones
            value = this.adjustFormulaReferences(value, newRow - r, newCol - c);
          }
        }

        const update: PasteUpdate = { id: cellKey(newRow, newCol), value };
        if (cell?.format) {
          update.format = { ...cell.format };
        }
        updates.push(update);
      }
    });

    return updates;
  }

  /**
   * Parse an HTML string and extract table data with format information.
   * Prefers data-format (lossless) over inline styles (best-effort).
   *
   * `colspan`/`rowspan` are honoured the way Excel and Google Sheets write
   * merged ranges: the value lands in the merge's top-left slot and the other
   * slots it covers become empty cells, so later values keep their columns.
   */
  parseHTMLTableWithFormat(
    html: string,
    values = false
  ): Array<Array<{ value: string; format?: CellFormat }>> | null {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const table = doc.querySelector('table');
    if (!table) return null;

    type ParsedCell = { value: string; format?: CellFormat };
    // `table.rows` / `tr.cells` skip rows and cells of nested tables
    const trElements = Array.from(table.rows);
    const grid: ParsedCell[][] = trElements.map(() => []);
    let cellCount = 0;

    for (let r = 0; r < trElements.length; r++) {
      const row = grid[r];
      let col = 0;
      for (const el of Array.from(trElements[r].cells)) {
        // Skip slots already covered by a rowspan from a row above
        while (row[col]) col++;

        // data-raw holds the cell's formula; pasting values takes the shown text instead
        const raw = el.getAttribute('data-raw');
        const text = (el.textContent ?? '').trim();
        // data-spilled marks a value its (also copied) formula spills again; only a values paste writes it
        const spilled = el.hasAttribute('data-spilled') && !values;
        // data-value is the unformatted value behind a formula or a spilled cell
        const plain = el.getAttribute('data-value') ?? text;
        const value = spilled ? '' : raw !== null && !(values && raw.startsWith('=')) ? raw : plain;

        // Try lossless data-format first, then parse inline styles
        let format: CellFormat | undefined;
        const dataFormat = el.getAttribute('data-format');
        if (dataFormat) {
          try {
            format = this._sanitizeFormat(JSON.parse(dataFormat));
          } catch {
            // Invalid JSON, ignore
          }
        }
        if (!format) {
          const style = el.getAttribute('style');
          if (style) {
            format = this._parseStyleToFormat(style);
          }
        }

        // The DOM clamps colSpan to 1..1000 and rowSpan to 0..65534;
        // rowspan="0" means "to the end of the table".
        const colSpan = Math.max(1, el.colSpan);
        const rowSpan = el.rowSpan === 0 ? trElements.length - r : Math.min(el.rowSpan, trElements.length - r);
        for (let dr = 0; dr < Math.max(1, rowSpan); dr++) {
          for (let dc = 0; dc < colSpan; dc++) {
            grid[r + dr][col + dc] = dr === 0 && dc === 0 ? { value, format } : { value: '' };
          }
        }
        col += colSpan;
        cellCount++;
      }
    }

    if (cellCount === 0) return null;
    // Fill holes left by ragged rows next to a rowspan, so every slot is a cell
    return grid.map((row) => Array.from(row, (cell) => cell ?? { value: '' }));
  }

  /**
   * Convert format-aware parsed rows into PasteUpdate array.
   */
  private _applyParsedRowsWithFormat(
    rows: Array<Array<{ value: string; format?: CellFormat }>>,
    targetCol: number,
    maxRows: number,
    maxCols: number,
    rowFor: RowMapper
  ): PasteUpdate[] {
    const updates: PasteUpdate[] = [];
    for (let r = 0; r < rows.length; r++) {
      const row = rowFor(r);
      for (let c = 0; c < rows[r].length; c++) {
        const col = targetCol + c;
        if (row >= 0 && row < maxRows && col < maxCols) {
          const { value, format } = rows[r][c];
          const update: PasteUpdate = { id: cellKey(row, col), value };
          if (format) update.format = format;
          updates.push(update);
        }
      }
    }
    return updates;
  }

  /**
   * Parse a TSV string and generate cell updates.
   * Handles RFC 4180-style quoting: fields containing tabs, newlines, or
   * double-quotes are wrapped in double-quotes, with internal quotes escaped
   * as `""`. This is the format Excel and Google Sheets use on the clipboard.
   *
   * @param text - The TSV text to parse
   * @param targetRow - 0-indexed row of the top-left paste destination
   * @param targetCol - 0-indexed column of the top-left paste destination
   * @param maxRows - Grid row count (for bounds clamping)
   * @param maxCols - Grid column count (for bounds clamping)
   * @param rowFor - The grid row the `i`-th pasted row goes to, or -1 to drop it
   * @returns Array of cell updates with keys and values
   */
  parseTSV(
    text: string,
    targetRow: number,
    targetCol: number,
    maxRows: number,
    maxCols: number,
    rowFor: RowMapper = (i) => targetRow + i
  ): Array<{ id: string; value: string }> {
    const rows = parseDelimited(text, '\t');
    const updates: Array<{ id: string; value: string }> = [];

    for (let r = 0; r < rows.length; r++) {
      const cells = rows[r];
      const row = rowFor(r);
      for (let c = 0; c < cells.length; c++) {
        const col = targetCol + c;
        if (row >= 0 && row < maxRows && col < maxCols) {
          updates.push({ id: cellKey(row, col), value: cells[c] });
        }
      }
    }

    return updates;
  }

  // ─── Reference Adjustment ─────────────────────────

  /**
   * Adjust cell references in a formula by the given row/col offset.
   * Absolute references (`$A$1`) are not adjusted.
   * Mixed references (`$A1`, `A$1`) only adjust the non-absolute component.
   * References inside string literals are left untouched.
   *
   * @param formula - The formula string (must start with "=")
   * @param rowOffset - Number of rows to shift relative references
   * @param colOffset - Number of columns to shift relative references
   * @returns The adjusted formula string, with `#REF!` for out-of-bounds refs
   */
  adjustFormulaReferences(formula: string, rowOffset: number, colOffset: number): string {
    return shiftFormulaReferences(formula, rowOffset, colOffset);
  }

  // ─── Serialization ──────────────────────────────────

  private serializeRange(data: GridData, range: SelectionRange, rows = rangeRows(range)): string {
    const lines: string[] = [];

    for (const r of rows) {
      const cols: string[] = [];
      for (let c = range.start.col; c <= range.end.col; c++) {
        const cell = data.get(cellKey(r, c));
        const val = cell?.displayValue ?? '';
        cols.push(quoteDelimitedField(val, '\t'));
      }
      lines.push(cols.join('\t'));
    }

    return lines.join('\n');
  }

  /**
   * Convert a selection range to an HTML table string.
   * Includes inline styles for visual formatting and data-format for lossless round-trip.
   */
  private _serializeRangeAsHTML(
    data: GridData,
    range: SelectionRange,
    rowList = rangeRows(range),
    valueOf?: (key: string, cell: CellData) => string
  ): string {
    const rows: string[] = [];

    for (const r of rowList) {
      const cells: string[] = [];
      for (let c = range.start.col; c <= range.end.col; c++) {
        const key = cellKey(r, c);
        const cell = data.get(key);
        const value = cell?.displayValue ?? '';
        const rawValue = cell?.rawValue ?? '';
        const attrs: string[] = [];

        if (rawValue && rawValue !== value) {
          attrs.push(`data-raw="${this._escapeHTML(rawValue)}"`);
        }
        // What a formula computed, or what was spilled into the cell, without its
        // number format: what a paste of values (or of a spilled value alone) writes
        if (cell && valueOf && (rawValue.startsWith('=') || (rawValue === '' && cell.spilledFrom !== undefined))) {
          const plain = computedLiteral(key, cell, valueOf);
          if (plain !== value) attrs.push(`data-value="${this._escapeHTML(plain)}"`);
        }
        // A value spilled by a formula that is copied too: a paste into another
        // spreadsheet leaves the cell to that formula
        if (cell?.spilledFrom !== undefined && rawValue === '') {
          const [anchorRow, anchorCol] = cell.spilledFrom.split(':').map(Number);
          if (rowList.includes(anchorRow) && anchorCol >= range.start.col && anchorCol <= range.end.col) {
            attrs.push('data-spilled="true"');
          }
        }

        if (cell?.format) {
          const inlineStyle = this._formatToInlineStyle(cell.format);
          if (inlineStyle) {
            attrs.push(`style="${this._escapeHTML(inlineStyle)}"`);
          }
          attrs.push(`data-format="${this._escapeHTML(JSON.stringify(cell.format))}"`);
        }

        const attrStr = attrs.length > 0 ? ' ' + attrs.join(' ') : '';
        cells.push(`<td${attrStr}>${this._escapeHTML(value)}</td>`);
      }
      rows.push(`<tr>${cells.join('')}</tr>`);
    }

    return `<table>${rows.join('')}</table>`;
  }

  /** Convert CellFormat to a CSS inline style string */
  private _formatToInlineStyle(format: CellFormat): string {
    const parts: string[] = [];
    if (format.bold) parts.push('font-weight: bold');
    if (format.italic) parts.push('font-style: italic');
    const decorations: string[] = [];
    if (format.underline) decorations.push('underline');
    if (format.strikethrough) decorations.push('line-through');
    if (decorations.length > 0) parts.push(`text-decoration: ${decorations.join(' ')}`);
    if (format.textColor) parts.push(`color: ${format.textColor}`);
    if (format.backgroundColor) parts.push(`background-color: ${format.backgroundColor}`);
    if (format.textAlign) parts.push(`text-align: ${format.textAlign}`);
    if (format.verticalAlign) parts.push(`vertical-align: ${format.verticalAlign}`);
    if (format.wrap) parts.push('white-space: normal');
    if (format.fontSize) parts.push(`font-size: ${format.fontSize}px`);
    return parts.join('; ');
  }

  /** Parse a CSS inline style string into a CellFormat (best-effort for external sources) */
  private _parseStyleToFormat(style: string): CellFormat | undefined {
    if (!style) return undefined;
    const fmt: CellFormat = {};

    const fontWeight = this._extractStyleProp(style, 'font-weight');
    if (fontWeight === 'bold' || (fontWeight && parseInt(fontWeight, 10) >= 600)) fmt.bold = true;

    const fontStyle = this._extractStyleProp(style, 'font-style');
    if (fontStyle === 'italic') fmt.italic = true;

    const textDecoration = this._extractStyleProp(style, 'text-decoration');
    if (textDecoration) {
      if (textDecoration.includes('underline')) fmt.underline = true;
      if (textDecoration.includes('line-through')) fmt.strikethrough = true;
    }

    const validColorPattern = /^[a-zA-Z0-9#(),.\s%/-]+$/;

    const color = this._extractStyleProp(style, 'color');
    if (color && color.length <= 100 && validColorPattern.test(color)) fmt.textColor = color;

    const bgColor = this._extractStyleProp(style, 'background-color');
    if (bgColor && bgColor.length <= 100 && validColorPattern.test(bgColor)) fmt.backgroundColor = bgColor;

    const textAlign = this._extractStyleProp(style, 'text-align');
    if (textAlign === 'left' || textAlign === 'center' || textAlign === 'right') {
      fmt.textAlign = textAlign;
    }

    // vertical-align and white-space are deliberately not read back: Google
    // Sheets and Excel emit them on every cell as their defaults, which would
    // stamp a format onto every pasted cell. They round-trip via data-format.

    const fontSize = this._extractStyleProp(style, 'font-size');
    if (fontSize) {
      const px = parseInt(fontSize, 10);
      if (!isNaN(px) && px > 0 && px <= 200) fmt.fontSize = px;
    }

    return Object.keys(fmt).length > 0 ? fmt : undefined;
  }

  /** Validate and sanitize a parsed CellFormat from an external source */
  private _sanitizeFormat(raw: unknown): CellFormat | undefined {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return undefined;

    const obj = raw as Record<string, unknown>;
    const fmt: CellFormat = {};
    const validColorPattern = /^[a-zA-Z0-9#(),.\s%/-]+$/;
    const validAligns = new Set(['left', 'center', 'right']);
    const validVerticalAligns = new Set(['top', 'middle', 'bottom']);

    if (typeof obj.bold === 'boolean') fmt.bold = obj.bold;
    if (typeof obj.italic === 'boolean') fmt.italic = obj.italic;
    if (typeof obj.underline === 'boolean') fmt.underline = obj.underline;
    if (typeof obj.strikethrough === 'boolean') fmt.strikethrough = obj.strikethrough;

    if (typeof obj.textColor === 'string' && obj.textColor.length <= 100 && validColorPattern.test(obj.textColor)) {
      fmt.textColor = obj.textColor;
    }
    if (typeof obj.backgroundColor === 'string' && obj.backgroundColor.length <= 100 && validColorPattern.test(obj.backgroundColor)) {
      fmt.backgroundColor = obj.backgroundColor;
    }
    if (typeof obj.textAlign === 'string' && validAligns.has(obj.textAlign)) {
      fmt.textAlign = obj.textAlign as CellFormat['textAlign'];
    }
    if (typeof obj.verticalAlign === 'string' && validVerticalAligns.has(obj.verticalAlign)) {
      fmt.verticalAlign = obj.verticalAlign as CellFormat['verticalAlign'];
    }
    if (typeof obj.wrap === 'boolean') fmt.wrap = obj.wrap;
    if (typeof obj.fontSize === 'number' && obj.fontSize > 0 && obj.fontSize <= 200) {
      fmt.fontSize = obj.fontSize;
    }
    const numberFormat = this._sanitizeNumberFormat(obj.numberFormat);
    if (numberFormat) fmt.numberFormat = numberFormat;

    return Object.keys(fmt).length > 0 ? fmt : undefined;
  }

  private _sanitizeNumberFormat(raw: unknown): NumberFormatOptions | undefined {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return undefined;
    const obj = raw as Record<string, unknown>;
    const validTypes = new Set(['number', 'currency', 'percent', 'scientific', 'date']);
    if (typeof obj.type !== 'string' || !validTypes.has(obj.type)) return undefined;

    const nf: NumberFormatOptions = { type: obj.type as NumberFormatOptions['type'] };
    // toFixed/toExponential throw outside 0..20 (0..100 in newer engines)
    if (typeof obj.decimals === 'number' && Number.isInteger(obj.decimals) && obj.decimals >= 0 && obj.decimals <= 20) {
      nf.decimals = obj.decimals;
    }
    if (typeof obj.currencySymbol === 'string' && obj.currencySymbol.length <= 8) {
      nf.currencySymbol = obj.currencySymbol;
    }
    if (typeof obj.thousandsSep === 'boolean') nf.thousandsSep = obj.thousandsSep;
    if (typeof obj.datePattern === 'string' && obj.datePattern.length <= 64) {
      nf.datePattern = obj.datePattern;
    }
    return nf;
  }

  private _extractStyleProp(style: string, prop: string): string | undefined {
    // Match "prop: value" stopping at ; or end of string.
    // Use negative lookbehind to avoid matching "color" inside "background-color".
    const regex = new RegExp(`(?<![\\w-])${prop}\\s*:\\s*([^;]+)`, 'i');
    const match = style.match(regex);
    return match ? match[1].trim() : undefined;
  }

  /** Escape special HTML characters */
  private _escapeHTML(str: string): string {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ─── Fallback ───────────────────────────────────────

  /** @deprecated Uses document.execCommand('copy') which is deprecated. Only used as a fallback when the Clipboard API is unavailable. */
  private fallbackCopy(text: string): void {
    console.warn('y11n-spreadsheet: Clipboard API unavailable, using deprecated execCommand fallback.');
    const textarea = document.createElement('textarea');
    textarea.value = text;
    textarea.style.position = 'fixed';
    textarea.style.left = '-9999px';
    document.body.appendChild(textarea);
    textarea.select();
    try {
      document.execCommand('copy');
    } catch {
      // Silent fail
    }
    document.body.removeChild(textarea);
  }
}
