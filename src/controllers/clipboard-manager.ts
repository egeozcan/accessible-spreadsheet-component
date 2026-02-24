import {
  type CellFormat,
  type GridData,
  type SelectionRange,
  cellKey,
  colToLetter,
  letterToCol,
} from '../types.js';

export interface PasteUpdate {
  id: string;
  value: string;
  format?: CellFormat;
}

/**
 * ClipboardManager - Handles copy, cut, and paste operations
 * using the system clipboard API with TSV and HTML table formats.
 */
export class ClipboardManager {
  /** Source range and data from the last internal copy/cut, used for reference adjustment on paste */
  private _copySourceRange: SelectionRange | null = null;
  private _copySourceData: GridData | null = null;

  /**
   * Copy the selected range to the clipboard as both HTML table and TSV.
   * Falls back to text-only if ClipboardItem API is unavailable.
   */
  async copy(data: GridData, range: SelectionRange): Promise<void> {
    this._copySourceRange = range;
    this._copySourceData = data;
    const tsv = this.serializeRange(data, range);
    const htmlStr = this._serializeRangeAsHTML(data, range);

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
   * Cut: copy to clipboard and return the keys that should be cleared.
   */
  async cut(data: GridData, range: SelectionRange): Promise<string[]> {
    await this.copy(data, range);
    const keys: string[] = [];
    for (let r = range.start.row; r <= range.end.row; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        keys.push(cellKey(r, c));
      }
    }
    return keys;
  }

  /**
   * Parse clipboard content and return the updates to apply.
   * If data was copied internally, uses raw values with reference adjustment.
   * Otherwise tries HTML table format first, then falls back to TSV text.
   * The target is the top-left cell where paste begins.
   */
  async paste(
    targetRow: number,
    targetCol: number,
    maxRows: number,
    maxCols: number
  ): Promise<PasteUpdate[] | null> {
    // Try internal paste with formula reference adjustment first
    if (this._copySourceRange && this._copySourceData) {
      const internalResult = this._pasteInternal(targetRow, targetCol, maxRows, maxCols);
      if (internalResult) return internalResult;
    }

    // Try reading HTML from clipboard first
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (item.types.includes('text/html')) {
          const blob = await item.getType('text/html');
          const htmlText = await blob.text();
          const parsed = this.parseHTMLTableWithFormat(htmlText);
          if (parsed) {
            return this._applyParsedRowsWithFormat(parsed, targetRow, targetCol, maxRows, maxCols);
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

    return this.parseTSV(text, targetRow, targetCol, maxRows, maxCols);
  }

  /**
   * Paste using internally stored copy data with formula reference adjustment.
   * Returns null if no internal data is available.
   */
  private _pasteInternal(
    targetRow: number,
    targetCol: number,
    maxRows: number,
    maxCols: number
  ): PasteUpdate[] | null {
    const srcRange = this._copySourceRange;
    const srcData = this._copySourceData;
    if (!srcRange || !srcData) return null;

    const rowOffset = targetRow - srcRange.start.row;
    const colOffset = targetCol - srcRange.start.col;
    const updates: PasteUpdate[] = [];

    for (let r = srcRange.start.row; r <= srcRange.end.row; r++) {
      for (let c = srcRange.start.col; c <= srcRange.end.col; c++) {
        const newRow = r + rowOffset;
        const newCol = c + colOffset;
        if (newRow >= maxRows || newCol >= maxCols) continue;

        const cell = srcData.get(cellKey(r, c));
        let value = cell?.rawValue ?? '';

        // Adjust formula references
        if (value.startsWith('=')) {
          value = this.adjustFormulaReferences(value, rowOffset, colOffset);
        }

        const update: PasteUpdate = { id: cellKey(newRow, newCol), value };
        if (cell?.format) {
          update.format = { ...cell.format };
        }
        updates.push(update);
      }
    }

    return updates;
  }

  /**
   * Parse an HTML string and extract table data as a 2D string array.
   * Returns null if no <table> element is found.
   */
  parseHTMLTable(html: string): string[][] | null {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const table = doc.querySelector('table');
    if (!table) return null;

    const rows: string[][] = [];
    const trElements = table.querySelectorAll('tr');

    for (const tr of trElements) {
      const cells: string[] = [];
      const cellElements = tr.querySelectorAll('td, th');
      for (const cell of cellElements) {
        const raw = (cell as Element).getAttribute('data-raw');
        cells.push(raw ?? (cell.textContent ?? '').trim());
      }
      if (cells.length > 0) {
        rows.push(cells);
      }
    }

    return rows.length > 0 ? rows : null;
  }

  /**
   * Parse an HTML string and extract table data with format information.
   * Prefers data-format (lossless) over inline styles (best-effort).
   */
  parseHTMLTableWithFormat(html: string): Array<Array<{ value: string; format?: CellFormat }>> | null {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const table = doc.querySelector('table');
    if (!table) return null;

    const rows: Array<Array<{ value: string; format?: CellFormat }>> = [];
    const trElements = table.querySelectorAll('tr');

    for (const tr of trElements) {
      const cells: Array<{ value: string; format?: CellFormat }> = [];
      const cellElements = tr.querySelectorAll('td, th');
      for (const cellEl of cellElements) {
        const el = cellEl as Element;
        const raw = el.getAttribute('data-raw');
        const value = raw ?? (cellEl.textContent ?? '').trim();

        // Try lossless data-format first, then parse inline styles
        let format: CellFormat | undefined;
        const dataFormat = el.getAttribute('data-format');
        if (dataFormat) {
          try {
            format = JSON.parse(dataFormat) as CellFormat;
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

        cells.push({ value, format });
      }
      if (cells.length > 0) {
        rows.push(cells);
      }
    }

    return rows.length > 0 ? rows : null;
  }

  /**
   * Convert format-aware parsed rows into PasteUpdate array.
   */
  private _applyParsedRowsWithFormat(
    rows: Array<Array<{ value: string; format?: CellFormat }>>,
    targetRow: number,
    targetCol: number,
    maxRows: number,
    maxCols: number
  ): PasteUpdate[] {
    const updates: PasteUpdate[] = [];
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        const row = targetRow + r;
        const col = targetCol + c;
        if (row < maxRows && col < maxCols) {
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
   * as "". This is the format Excel and Google Sheets use on the clipboard.
   *
   * Exported for testability with paste events where text is already available.
   */
  parseTSV(
    text: string,
    targetRow: number,
    targetCol: number,
    maxRows: number,
    maxCols: number
  ): Array<{ id: string; value: string }> {
    const rows = this._parseTSVRows(text);
    const updates: Array<{ id: string; value: string }> = [];

    for (let r = 0; r < rows.length; r++) {
      const cells = rows[r];
      for (let c = 0; c < cells.length; c++) {
        const row = targetRow + r;
        const col = targetCol + c;
        if (row < maxRows && col < maxCols) {
          updates.push({ id: cellKey(row, col), value: cells[c] });
        }
      }
    }

    return updates;
  }

  /**
   * Parse TSV text into a 2D array of strings, respecting quoted fields.
   * Quoted fields may contain tabs, newlines, and escaped quotes ("").
   */
  private _parseTSVRows(text: string): string[][] {
    const rows: string[][] = [];
    let currentRow: string[] = [];
    let currentField = '';
    let inQuotes = false;
    let i = 0;

    while (i < text.length) {
      const ch = text[i];

      if (inQuotes) {
        if (ch === '"') {
          // Check for escaped quote ""
          if (i + 1 < text.length && text[i + 1] === '"') {
            currentField += '"';
            i += 2;
          } else {
            // End of quoted field
            inQuotes = false;
            i++;
          }
        } else {
          currentField += ch;
          i++;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
          i++;
        } else if (ch === '\t') {
          currentRow.push(currentField);
          currentField = '';
          i++;
        } else if (ch === '\r') {
          // Handle \r\n and bare \r
          currentRow.push(currentField);
          currentField = '';
          if (i + 1 < text.length && text[i + 1] === '\n') {
            i += 2;
          } else {
            i++;
          }
          rows.push(currentRow);
          currentRow = [];
        } else if (ch === '\n') {
          currentRow.push(currentField);
          currentField = '';
          i++;
          rows.push(currentRow);
          currentRow = [];
        } else {
          currentField += ch;
          i++;
        }
      }
    }

    // Push the last field and row
    currentRow.push(currentField);
    // Only add the last row if it's not a single empty field (trailing newline)
    if (currentRow.length > 1 || currentRow[0] !== '') {
      rows.push(currentRow);
    }

    return rows;
  }

  // ─── Reference Adjustment ─────────────────────────

  /**
   * Adjust cell references in a formula by the given row/col offset.
   * Absolute references ($A$1) are not adjusted.
   * Mixed references ($A1, A$1) only adjust the non-absolute part.
   */
  adjustFormulaReferences(formula: string, rowOffset: number, colOffset: number): string {
    if (!formula.startsWith('=')) return formula;

    // Match cell references including optional $ markers
    // Pattern: optional $ + column letters + optional $ + row digits
    // Handles ranges like $A$1:$B$2 by matching each ref separately
    const refPattern = /(\$?)([A-Z]+)(\$?)(\d+)/gi;

    const adjusted = formula.replace(refPattern, (_match, colDollar: string, colLetters: string, rowDollar: string, rowDigits: string) => {
      const col = letterToCol(colLetters.toUpperCase());
      const row = parseInt(rowDigits, 10) - 1; // 1-indexed to 0-indexed

      const newCol = colDollar === '$' ? col : col + colOffset;
      const newRow = rowDollar === '$' ? row : row + rowOffset;

      // Clamp to valid range (don't produce negative indices)
      if (newCol < 0 || newRow < 0) {
        return '#REF!';
      }

      return `${colDollar}${colToLetter(newCol)}${rowDollar}${newRow + 1}`;
    });

    return adjusted;
  }

  // ─── Serialization ──────────────────────────────────

  private serializeRange(data: GridData, range: SelectionRange): string {
    const lines: string[] = [];

    for (let r = range.start.row; r <= range.end.row; r++) {
      const cols: string[] = [];
      for (let c = range.start.col; c <= range.end.col; c++) {
        const cell = data.get(cellKey(r, c));
        const val = cell?.displayValue ?? '';
        cols.push(this._quoteTSVField(val));
      }
      lines.push(cols.join('\t'));
    }

    return lines.join('\n');
  }

  /** Quote a TSV field if it contains tabs, newlines, or double-quotes */
  private _quoteTSVField(value: string): string {
    if (value.includes('\t') || value.includes('\n') || value.includes('\r') || value.includes('"')) {
      return `"${value.replace(/"/g, '""')}"`;
    }
    return value;
  }

  /**
   * Convert a selection range to an HTML table string.
   * Includes inline styles for visual formatting and data-format for lossless round-trip.
   */
  private _serializeRangeAsHTML(data: GridData, range: SelectionRange): string {
    const rows: string[] = [];

    for (let r = range.start.row; r <= range.end.row; r++) {
      const cells: string[] = [];
      for (let c = range.start.col; c <= range.end.col; c++) {
        const cell = data.get(cellKey(r, c));
        const value = cell?.displayValue ?? '';
        const rawValue = cell?.rawValue ?? '';
        const attrs: string[] = [];

        if (rawValue && rawValue !== value) {
          attrs.push(`data-raw="${this._escapeHTML(rawValue)}"`);
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
  _formatToInlineStyle(format: CellFormat): string {
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
    if (format.fontSize) parts.push(`font-size: ${format.fontSize}px`);
    return parts.join('; ');
  }

  /** Parse a CSS inline style string into a CellFormat (best-effort for external sources) */
  _parseStyleToFormat(style: string): CellFormat | undefined {
    if (!style) return undefined;
    const fmt: CellFormat = {};

    const fontWeight = this._extractStyleProp(style, 'font-weight');
    if (fontWeight === 'bold' || fontWeight === '700') fmt.bold = true;

    const fontStyle = this._extractStyleProp(style, 'font-style');
    if (fontStyle === 'italic') fmt.italic = true;

    const textDecoration = this._extractStyleProp(style, 'text-decoration');
    if (textDecoration) {
      if (textDecoration.includes('underline')) fmt.underline = true;
      if (textDecoration.includes('line-through')) fmt.strikethrough = true;
    }

    const color = this._extractStyleProp(style, 'color');
    if (color) fmt.textColor = color;

    const bgColor = this._extractStyleProp(style, 'background-color');
    if (bgColor) fmt.backgroundColor = bgColor;

    const textAlign = this._extractStyleProp(style, 'text-align');
    if (textAlign === 'left' || textAlign === 'center' || textAlign === 'right') {
      fmt.textAlign = textAlign;
    }

    const fontSize = this._extractStyleProp(style, 'font-size');
    if (fontSize) {
      const px = parseInt(fontSize, 10);
      if (!isNaN(px)) fmt.fontSize = px;
    }

    return Object.keys(fmt).length > 0 ? fmt : undefined;
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
