import { cellKey } from 'y11n-spreadsheet';
import type { GridData } from 'y11n-spreadsheet';

/**
 * Turns rows of text into the sparse `Map<"row:col", CellData>` the grid
 * stores. Empty strings are left out, and text starting with `=` is a formula.
 */
export function gridFromRows(rows: string[][]): GridData {
  const data: GridData = new Map();
  rows.forEach((values, row) => {
    values.forEach((value, col) => {
      if (value === '') return;
      const isNumber = !value.startsWith('=') && !isNaN(Number(value));
      data.set(cellKey(row, col), {
        rawValue: value,
        // For a formula, these two are computed when the data is loaded
        displayValue: value,
        type: isNumber ? 'number' : 'text',
      });
    });
  });
  return data;
}
