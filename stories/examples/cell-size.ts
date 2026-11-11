/**
 * `--ls-cell-width` and `--ls-cell-height` set the size of every cell, and
 * `--ls-font-size` the text. Any CSS length works (`7rem`, `2.5em`).
 *
 * `--ls-cell-width` is the default for all columns; single columns get their
 * own width with `setColumnWidth()` (see Layout / Column widths).
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function cellSize() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 10;
  sheet.cols = 4;
  sheet.style.setProperty('--ls-cell-width', '250px');
  sheet.style.setProperty('--ls-cell-height', '36px');
  sheet.style.setProperty('--ls-font-size', '15px');
  sheet.setData(
    gridFromRows([
      ['Description', 'Amount'],
      ['Monthly subscription fee', '29.99'],
      ['Annual support contract', '299.00'],
      ['One-time setup charge', '150.00'],
      ['Total', '=SUM(B2:B4)'],
    ])
  );
  return sheet;
}
