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

    const updates: Array<{ id: string; value: string }> = [];

    for (let r = 0; r < rows.length; r++) {
      const cells = rows[r].split('\t');
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
