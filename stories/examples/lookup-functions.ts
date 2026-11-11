/**
 * Looking values up in a table. `XLOOKUP` finds a value in one column and
 * returns the matching cell of another; `VLOOKUP` does the same by column
 * number; `INDEX` and `MATCH` are the two halves, usable on their own.
 *
 * Try it: type another product ID (P001 to P008) into A12 or A13.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function lookupFunctions() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 25;
  sheet.cols = 8;
  sheet.style.setProperty('--ls-cell-width', '140px');
  sheet.setData(
    gridFromRows([
      ['Product ID', 'Product name', 'Category', 'Price', 'Stock'],
      ['P001', 'Wireless Mouse', 'Electronics', '29.99', '150'],
      ['P002', 'Mechanical KB', 'Electronics', '89.99', '75'],
      ['P003', 'USB-C Hub', 'Electronics', '45.00', '200'],
      ['P004', 'Desk Lamp', 'Office', '34.50', '120'],
      ['P005', 'Monitor Stand', 'Office', '59.99', '90'],
      ['P006', 'Webcam HD', 'Electronics', '69.99', '60'],
      ['P007', 'Notebook A5', 'Stationery', '4.99', '500'],
      ['P008', 'Pen Set', 'Stationery', '12.50', '300'],
      [],
      ['Look up', 'Name (XLOOKUP)', 'Price (VLOOKUP)', 'Row (MATCH)'],
      ['P003', '=XLOOKUP(A12, A2:A9, B2:B9, "not found")', '=VLOOKUP(A12, A2:E9, 4, FALSE)', '=MATCH(A12, A2:A9, 0)'],
      ['P006', '=XLOOKUP(A13, A2:A9, B2:B9, "not found")', '=VLOOKUP(A13, A2:E9, 4, FALSE)', '=MATCH(A13, A2:A9, 0)'],
      [],
      ['Third product', '=INDEX(A2:E9, 3, 2)'],
      ['First office item', '=INDEX(B2:B9, MATCH("Office", C2:C9, 0))'],
    ])
  );
  return sheet;
}
