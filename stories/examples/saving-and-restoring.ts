/**
 * `getData()` returns a copy of the cells and `setData()` replaces them, so
 * saving a grid is turning that `Map` into text and back. `toCSV()` and
 * `fromCSV()` do the same for delimited text.
 *
 * Try it: change some cells, press Save, press Clear, then Restore.
 */
import 'y11n-spreadsheet';
import type { GridData } from 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function savingAndRestoring() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 12;
  sheet.cols = 6;
  sheet.setData(
    gridFromRows([
      ['City', 'Population (M)'],
      ['Tokyo', '37.4'],
      ['Delhi', '30.3'],
      ['Shanghai', '27.1'],
      ['Total', '=SUM(B2:B4)'],
    ])
  );

  const output = document.createElement('pre');
  output.tabIndex = 0;
  output.textContent = 'Press a button.';

  let saved = '[]';

  const actions: Record<string, () => void> = {
    Save() {
      // A Map of plain objects: its entries are JSON as they are
      saved = JSON.stringify([...sheet.getData()]);
      output.textContent = saved;
    },
    Clear() {
      sheet.setData(new Map());
      output.textContent = 'setData(new Map())';
    },
    Restore() {
      const data: GridData = new Map(JSON.parse(saved));
      sheet.setData(data);
      output.textContent = `setData() with ${data.size} cells`;
    },
    'Export CSV'() {
      // Display values; toCSV({ values: 'raw' }) exports the formulas instead
      output.textContent = sheet.toCSV();
    },
    'Import CSV'() {
      // Written at A1 as one undoable step; values starting with = are formulas
      sheet.fromCSV(
        [
          'Fruit,Count,Price,Total',
          'Apple,10,1.20,=B2*C2',
          'Banana,25,0.50,=B3*C3',
          'Cherry,100,0.10,=B4*C4',
          'Sum,,,=SUM(D2:D4)',
        ].join('\n')
      );
      output.textContent = 'fromCSV(text)';
    },
  };

  const buttons = document.createElement('div');
  for (const [label, run] of Object.entries(actions)) {
    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = label;
    button.addEventListener('click', run);
    buttons.append(button);
  }

  // Layout: buttons and output above the grid
  const page = document.createElement('div');
  page.style.cssText = 'display: flex; flex-direction: column; height: 100%;';
  buttons.style.cssText =
    'display: flex; flex-wrap: wrap; gap: 8px; padding: 8px 12px;';
  output.style.cssText = `
    height: 64px; margin: 0; padding: 8px 12px; overflow: auto;
    border-block: 1px solid #d0d7de; background: #f6f8fa; color: #1f2328;
    font: 12px/1.5 ui-monospace, monospace; white-space: pre-wrap;
  `;
  sheet.style.cssText = 'flex: 1; min-height: 0;';
  page.append(buttons, output, sheet);
  return page;
}
