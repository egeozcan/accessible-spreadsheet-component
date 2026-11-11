/**
 * 1,000 rows of data. The grid only renders the cells in view plus a small
 * buffer around them, so scrolling stays smooth whatever the size of the
 * data. Nothing needs to be turned on for this.
 *
 * Try it: scroll, or press Ctrl+Down to jump to the last row of a column.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function largeDataset() {
  const firstNames = 'Alice Bob Charlie Diana Eve Frank Grace Hank Ivy Jack'.split(' ');
  const lastNames = 'Smith Johnson Williams Brown Jones Garcia Miller Davis Wilson Moore'.split(' ');
  const statuses = ['Active', 'Inactive', 'Pending'];

  const rows = [['ID', 'First name', 'Last name', 'Email', 'Score', 'Status']];
  for (let i = 1; i <= 1000; i++) {
    const first = firstNames[i % firstNames.length];
    const last = lastNames[(i * 3) % lastNames.length];
    const email = `${first}.${last}@example.com`.toLowerCase();
    const score = String(50 + ((i * 7 + 13) % 51));
    rows.push([String(i), first, last, email, score, statuses[i % statuses.length]]);
  }
  rows.push(['', '', '', 'Average', '=ROUND(AVERAGE(E2:E1001), 1)']);

  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 1010;
  sheet.cols = 8;
  sheet.frozenRows = 1;
  sheet.setColumnWidth(3, 220);
  sheet.setData(gridFromRows(rows));
  return sheet;
}
