/**
 * Changing formats from code. Cells are addressed by their key (`"0:0"` is
 * A1) or by a range of `{ row, col }` corners.
 *
 * - `setCellFormat(key, format)` and `setRangeFormat(range, format)` merge
 *   the given properties into what the cells already have
 * - `getCellFormat(key)` returns a cell's format, or `undefined`
 * - `clearCellFormat(key)` and `clearRangeFormat(range)` remove it
 *
 * Each call is one undo step: press a button, then Ctrl+Z in the grid.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function formatApi() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 10;
  sheet.cols = 8;
  sheet.setData(
    gridFromRows([
      ['Name', 'Value', 'Notes'],
      ['Alpha', '100', 'First item'],
      ['Beta', '200', 'Second item'],
      ['Gamma', '300', 'Third item'],
    ])
  );

  const output = document.createElement('pre');
  output.tabIndex = 0;
  output.textContent = 'Press a button.';

  const headerRow = { start: { row: 0, col: 0 }, end: { row: 0, col: 2 } };
  const values = { start: { row: 1, col: 1 }, end: { row: 3, col: 1 } };
  const notes = { start: { row: 1, col: 2 }, end: { row: 3, col: 2 } };
  const everything = { start: { row: 0, col: 0 }, end: { row: 9, col: 7 } };

  const actions: Record<string, () => void> = {
    'Style the header row'() {
      sheet.setRangeFormat(headerRow, { bold: true, backgroundColor: '#e3f2fd' });
    },
    'Make B2 red'() {
      sheet.setCellFormat('1:1', { textColor: '#b3261e', bold: true });
    },
    'Notes in italics'() {
      sheet.setRangeFormat(notes, { italic: true });
    },
    'Values as currency'() {
      sheet.setRangeFormat(values, {
        numberFormat: { type: 'currency', currencySymbol: '$', decimals: 2 },
      });
    },
    'Read the format of A1'() {
      const format = sheet.getCellFormat('0:0');
      output.textContent = `getCellFormat('0:0') → ${JSON.stringify(format)}`;
    },
    'Clear all formats'() {
      sheet.clearRangeFormat(everything);
    },
  };

  const buttons = document.createElement('div');
  for (const [label, run] of Object.entries(actions)) {
    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = label;
    button.addEventListener('click', () => {
      output.textContent = label;
      run();
    });
    buttons.append(button);
  }

  // Layout: buttons and output above the grid
  const page = document.createElement('div');
  page.style.cssText = 'display: flex; flex-direction: column; height: 100%;';
  buttons.style.cssText =
    'display: flex; flex-wrap: wrap; gap: 8px; padding: 8px 12px;';
  output.style.cssText = `
    margin: 0; padding: 8px 12px; border-block: 1px solid #d0d7de;
    background: #f6f8fa; color: #1f2328;
    font: 12px/1.5 ui-monospace, monospace; white-space: pre-wrap;
  `;
  sheet.style.cssText = 'flex: 1; min-height: 0;';
  page.append(buttons, output, sheet);
  return page;
}
