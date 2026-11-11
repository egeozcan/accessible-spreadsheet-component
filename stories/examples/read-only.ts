/**
 * `readOnly` (the `read-only` attribute in HTML) turns off every way of
 * changing the cells: typing, paste, fill, delete and formatting. Moving
 * around, selecting, copying, filtering and resizing columns still work.
 *
 * To lock only some cells, see "Validation and locked cells" under Editing
 * and Tables.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function readOnly() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 10;
  sheet.cols = 8;
  sheet.readOnly = true;
  sheet.setData(
    gridFromRows([
      ['Metric', 'Q1', 'Q2', 'Q3', 'Q4'],
      ['Revenue ($K)', '1250', '1480', '1320', '1690'],
      ['Costs ($K)', '890', '920', '870', '950'],
      ['Profit ($K)', '=B2-B3', '=C2-C3', '=D2-D3', '=E2-E3'],
      ['Margin (%)', '=ROUND(B4/B2*100, 1)', '=ROUND(C4/C2*100, 1)', '=ROUND(D4/D2*100, 1)', '=ROUND(E4/E2*100, 1)'],
    ])
  );
  return sheet;
}
