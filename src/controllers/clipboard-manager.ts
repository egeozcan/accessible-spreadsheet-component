import {
  type GridData,
  type SelectionRange,
  cellKey,
} from '../types.js';

/**
 * ClipboardManager - Handles copy, cut, and paste operations
 * using the system clipboard API with TSV and HTML table formats.
 */
export class ClipboardManager {
  /**
   * Copy the selected range to the clipboard as both HTML table and TSV.
   * Falls back to text-only if ClipboardItem API is unavailable.
   */
  async copy(data: GridData, range: SelectionRange): Promise<void> {
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
   * Tries HTML table format first, then falls back to TSV text.
   * The target is the top-left cell where paste begins.
   */
  async paste(
    targetRow: number,
    targetCol: number,
    maxRows: number,
    maxCols: number
  ): Promise<Array<{ id: string; value: string }> | null> {
    // Try reading HTML from clipboard first
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (item.types.includes('text/html')) {
          const blob = await item.getType('text/html');
          const htmlText = await blob.text();
          const parsed = this.parseHTMLTable(htmlText);
          if (parsed) {
            return this._applyParsedRows(parsed, targetRow, targetCol, maxRows, maxCols);
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
        cells.push((cell.textContent ?? '').trim());
      }
      if (cells.length > 0) {
        rows.push(cells);
      }
    }

    return rows.length > 0 ? rows : null;
  }

  /**
   * Parse a TSV string and generate cell updates.
   * Exported for testability with paste events where text is already available.
   */
  parseTSV(
    text: string,
    targetRow: number,
    targetCol: number,
    maxRows: number,
    maxCols: number
  ): Array<{ id: string; value: string }> {
    const rows = text.split(/\r?\n/);
    // Remove trailing empty row (common in TSV)
    if (rows.length > 0 && rows[rows.length - 1] === '') {
      rows.pop();
    }

    const parsed: string[][] = rows.map((row) => row.split('\t'));
    return this._applyParsedRows(parsed, targetRow, targetCol, maxRows, maxCols);
  }

  // ─── Internal Helpers ─────────────────────────────────

  /**
   * Convert a 2D string array of parsed rows into cell updates
   * starting at the given target position, clipped to grid bounds.
   */
  private _applyParsedRows(
    rows: string[][],
    targetRow: number,
    targetCol: number,
    maxRows: number,
    maxCols: number
  ): Array<{ id: string; value: string }> {
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

  // ─── Serialization ──────────────────────────────────

  private serializeRange(data: GridData, range: SelectionRange): string {
    const lines: string[] = [];

    for (let r = range.start.row; r <= range.end.row; r++) {
      const cols: string[] = [];
      for (let c = range.start.col; c <= range.end.col; c++) {
        const cell = data.get(cellKey(r, c));
        cols.push(cell?.displayValue ?? '');
      }
      lines.push(cols.join('\t'));
    }

    return lines.join('\n');
  }

  /**
   * Convert a selection range to an HTML table string.
   * Escapes &, <, > for safe HTML embedding.
   */
  private _serializeRangeAsHTML(data: GridData, range: SelectionRange): string {
    const rows: string[] = [];

    for (let r = range.start.row; r <= range.end.row; r++) {
      const cells: string[] = [];
      for (let c = range.start.col; c <= range.end.col; c++) {
        const cell = data.get(cellKey(r, c));
        const value = cell?.displayValue ?? '';
        cells.push(`<td>${this._escapeHTML(value)}</td>`);
      }
      rows.push(`<tr>${cells.join('')}</tr>`);
    }

    return `<table>${rows.join('')}</table>`;
  }

  /** Escape special HTML characters */
  private _escapeHTML(str: string): string {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
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
