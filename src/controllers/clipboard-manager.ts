import {
  type GridData,
  type SelectionRange,
  cellKey,
} from '../types.js';

/**
 * ClipboardManager - Handles copy, cut, and paste operations
 * using the system clipboard API with TSV (tab-separated values) format.
 */
export class ClipboardManager {
  /**
   * Copy the selected range to the clipboard as TSV.
   */
  async copy(data: GridData, range: SelectionRange): Promise<void> {
    const tsv = this.serializeRange(data, range);
    try {
      await navigator.clipboard.writeText(tsv);
    } catch {
      // Fallback for environments without clipboard API
      this.fallbackCopy(tsv);
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
   * The target is the top-left cell where paste begins.
   */
  async paste(
    targetRow: number,
    targetCol: number,
    maxRows: number,
    maxCols: number
  ): Promise<Array<{ id: string; value: string }> | null> {
    let text: string;
    try {
      text = await navigator.clipboard.readText();
    } catch {
      return null;
    }

    return this.parseTSV(text, targetRow, targetCol, maxRows, maxCols);
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
