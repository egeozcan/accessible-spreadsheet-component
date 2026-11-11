/**
 * A workbook with two sheets and three defined names. Summary adds up the
 * Orders sheet with sheet-qualified references (`Orders!A:A`). The names are
 * a constant (`tax`), a range on the other sheet (`orders`) and a function
 * (`withTax`, a `LAMBDA`).
 *
 * `sheetTabs` shows the tabs even with one sheet, so users can add more.
 *
 * Try it: switch to Orders with the tabs (or Ctrl+PageDown), change an
 * amount and come back. Type `orders` in the Name Box at the left of the
 * formula bar, or press Ctrl+F3 for the name manager.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function sheetsAndNames() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 20;
  sheet.cols = 8;
  sheet.sheetTabs = true;

  // The first sheet exists from the start and is called Sheet1
  sheet.renameSheet('Sheet1', 'Summary');
  sheet.setData(
    gridFromRows([
      ['Region', 'Sales', 'With tax'],
      ['North', '=SUMIF(Orders!A:A, A2, Orders!C:C)', '=B2*(1+tax)'],
      ['South', '=SUMIF(Orders!A:A, A3, Orders!C:C)', '=B3*(1+tax)'],
      ['East', '=SUMIF(Orders!A:A, A4, Orders!C:C)', '=B4*(1+tax)'],
      ['Total', '=SUM(B2:B4)', '=withTax(B5)'],
      [],
      ['Orders', '=ROWS(orders)'],
    ])
  );

  sheet.addSheet('Orders', {
    data: gridFromRows([
      ['Region', 'Product', 'Amount'],
      ['North', 'Apples', '120'],
      ['South', 'Pears', '80'],
      ['North', 'Plums', '45'],
      ['East', 'Cherries', '200'],
      ['South', 'Melons', '60'],
    ]),
  });

  sheet.defineName('tax', '=0.19');
  sheet.defineName('orders', '=Orders!A2:C6');
  sheet.defineName('withTax', '=LAMBDA(amount, amount * (1 + tax))');
  return sheet;
}
