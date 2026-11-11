/**
 * Columns start at the width of `--ls-cell-width` and can be given their own
 * width in px. `getColumnWidths()` returns the custom widths, ready to save,
 * and `setColumnWidths()` restores them; `column-resize` fires on every
 * change, which is the moment to save.
 *
 * Try it: drag the right border of a column header, or double-click it to
 * fit the column to its content. Column B is too narrow here: select a cell
 * in it and press Alt+Shift+0 (or the toolbar's button) to fit, or
 * Alt+Shift+Left / Right to resize.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';
import { stack } from './stack.js';

export default function columnWidths() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 20;
  sheet.cols = 8;
  // Column index -> width in px, the shape getColumnWidths() returns. Set
  // before the data is loaded, the widths are not an undo step: setData()
  // starts a fresh undo history.
  sheet.setColumnWidths({ 0: 190, 1: 70, 2: 60, 3: 220 });

  sheet.setData(
    gridFromRows([
      ['Customer', 'Country', 'Orders', 'Notes'],
      ['Northwind Traders', 'United Kingdom', '42', 'Prefers delivery before noon'],
      ['Contoso', 'NL', '7'],
      ['Fabrikam Incorporated', 'United States of America', '128', 'Key account'],
    ])
  );

  sheet.addEventListener('column-resize', () => {
    // The place to save the layout, e.g. to localStorage
    console.log('column widths', sheet.getColumnWidths());
  });

  const toolbar = document.createElement('y11n-toolbar');
  toolbar.target = sheet;
  toolbar.items = 'autoFit | undo redo';
  return stack(toolbar, sheet);
}
