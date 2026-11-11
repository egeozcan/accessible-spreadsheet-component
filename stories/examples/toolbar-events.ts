/**
 * Every control fires an event before it acts, and `preventDefault()` stops
 * it, so your code can do something else instead. There are three events:
 *
 * - `format-action` for the formatting controls: `{ action, value }`
 * - `conditional-format-action` for the Conditional formatting menu
 * - `toolbar-action` for everything else: `{ action }`
 *
 * Here Paste is replaced by a paste of the values only, so nothing copied
 * from elsewhere brings its own formats, and the events are listed under
 * the grid.
 *
 * A toolbar without a `target` only fires these events and shows what its
 * properties say (see "Toolbar states"). That is the way to drive something
 * other than a grid with it.
 */
import 'y11n-spreadsheet';
import type { ToolbarActionDetail } from 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';
import { stack } from './stack.js';

export default function toolbarEvents() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 20;
  sheet.cols = 8;
  sheet.setData(
    gridFromRows([
      ['Item', 'Qty', 'Price'],
      ['Notebook', '4', '3.5'],
      ['Pen', '10', '1.2'],
    ])
  );

  const toolbar = document.createElement('y11n-toolbar');
  toolbar.target = sheet;
  toolbar.items = 'history clipboard | bold italic textAlign';

  toolbar.addEventListener('toolbar-action', (event) => {
    const { action } = (event as CustomEvent<ToolbarActionDetail>).detail;
    if (action === 'paste') {
      event.preventDefault();
      void sheet.paste({ mode: 'values' });
    }
  });

  // List the events
  const log = document.createElement('output');
  log.style.cssText = `
    display: block; height: 4.5em; overflow: auto; padding: 4px 12px;
    font: 12px ui-monospace, monospace; white-space: pre;
  `;
  log.textContent = 'Use the toolbar to see its events here.';
  const events = ['format-action', 'conditional-format-action', 'toolbar-action'];
  for (const type of events) {
    toolbar.addEventListener(type, (event) => {
      const detail = JSON.stringify((event as CustomEvent).detail);
      log.textContent = `${type} ${detail}\n${log.textContent}`;
    });
  }

  return stack(toolbar, sheet, log);
}
