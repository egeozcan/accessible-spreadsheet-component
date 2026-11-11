/**
 * Text that starts with `=` is a formula. Column D multiplies each quantity by
 * its price, and the rows under the table summarise the columns.
 *
 * Try it: change a quantity or a price and the totals follow. Type `=` in an
 * empty cell and use the arrow keys or the mouse to pick the cells to refer to.
 * Or select the cell under a column of numbers and press the toolbar's
 * AutoSum button (Alt+=).
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';
import { stack } from './stack.js';

export default function formulas() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 20;
  sheet.cols = 8;
  sheet.setData(
    gridFromRows([
      ['Item', 'Qty', 'Price', 'Total'],
      ['Widget A', '10', '5.99', '=B2*C2'],
      ['Widget B', '25', '3.50', '=B3*C3'],
      ['Widget C', '7', '12.00', '=B4*C4'],
      ['Gadget X', '15', '8.25', '=B5*C5'],
      ['Gadget Y', '3', '45.00', '=B6*C6'],
      [],
      ['Subtotal', '', '', '=SUM(D2:D6)'],
      ['Items', '', '', '=COUNT(B2:B6)'],
      ['Average price', '', '', '=AVERAGE(C2:C6)'],
      ['Lowest price', '', '', '=MIN(C2:C6)'],
      ['Highest price', '', '', '=MAX(C2:C6)'],
    ])
  );

  const toolbar = document.createElement('y11n-toolbar');
  toolbar.target = sheet;
  toolbar.items = 'autoSum | undo redo';
  return stack(toolbar, sheet);
}
