/**
 * The grid reports what happens in it with DOM events. The three you will use
 * most:
 *
 * - `cell-change`: one cell was edited (`cellId`, `value`, `oldValue`)
 * - `data-change`: several cells changed at once, by a paste, cut, clear,
 *   fill, undo or redo (`updates`, `operation`, `source`)
 * - `selection-change`: the selection moved (`range`)
 *
 * Try it: edit a cell, select a range and press Delete, then Ctrl+Z. The
 * panel under the grid lists each event with its `detail`.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function events() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 10;
  sheet.cols = 6;
  sheet.setData(
    gridFromRows([
      ['Name', 'Value'],
      ['Alpha', '100'],
      ['Beta', '200'],
      ['Gamma', '300'],
      ['Total', '=SUM(B2:B4)'],
    ])
  );

  const log = document.createElement('pre');
  log.tabIndex = 0;
  log.setAttribute('role', 'log');
  log.setAttribute('aria-label', 'Events');
  log.textContent = 'Events appear here.\n';

  for (const name of ['cell-change', 'data-change', 'selection-change']) {
    sheet.addEventListener(name, (event) => {
      const { detail } = event as CustomEvent;
      log.textContent += `${name} ${JSON.stringify(detail)}\n`;
      log.scrollTop = log.scrollHeight;
    });
  }

  // Layout: the grid above, the log below
  const page = document.createElement('div');
  page.style.cssText = 'display: flex; flex-direction: column; height: 100%;';
  sheet.style.cssText = 'flex: 1; min-height: 0;';
  log.style.cssText = `
    height: 140px; margin: 0; padding: 8px 12px; overflow: auto;
    border-top: 2px solid #d0d7de; background: #f6f8fa; color: #1f2328;
    font: 12px/1.5 ui-monospace, monospace; white-space: pre-wrap;
  `;
  page.append(sheet, log);
  return page;
}
