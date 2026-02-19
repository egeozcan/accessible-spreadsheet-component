import type { CellData, GridData } from '../src/types.js';

/** Convenience to build a CellData object */
function cell(rawValue: string, type: CellData['type'] = 'text', displayValue?: string): CellData {
  return {
    rawValue,
    displayValue: displayValue ?? rawValue,
    type,
  };
}

/** Build a GridData map from a 2D array (row-major). Empty strings are skipped. */
export function gridFromRows(rows: string[][]): GridData {
  const data: GridData = new Map();
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const val = rows[r][c];
      if (val === '') continue;
      const isNum = val !== '' && !isNaN(Number(val));
      const isFormula = val.startsWith('=');
      data.set(`${r}:${c}`, cell(val, isFormula ? 'text' : isNum ? 'number' : 'text'));
    }
  }
  return data;
}

/** Wrap the spreadsheet element in a container with fixed height for stories */
export function wrapInContainer(el: Element, height = '500px'): HTMLDivElement {
  const container = document.createElement('div');
  container.style.height = height;
  container.style.border = '1px solid #ccc';
  container.style.borderRadius = '4px';
  container.style.overflow = 'hidden';
  container.appendChild(el);
  return container;
}
