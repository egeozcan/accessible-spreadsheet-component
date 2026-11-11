/**
 * Three ways to control what users enter:
 *
 * - `readOnly: true` on a cell locks that cell (here the header row, the
 *   Total column and the grand total)
 * - `validateCell` returns a message to reject a value, or `null` to accept it
 * - the cancelable `before-change` event sees every user change before it is
 *   applied and can rewrite or drop it
 *
 * Try it: type "abc" into a Qty cell and press Enter. The message stays until
 * the value is fixed or the edit is cancelled with Escape.
 */
import 'y11n-spreadsheet';
import { parseKey } from 'y11n-spreadsheet';
import type { BeforeChangeDetail } from 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function validation() {
  const data = gridFromRows([
    ['Item', 'Qty', 'Unit price', 'Total'],
    ['Widget', '4', '2.50', '=B2*C2'],
    ['Gadget', '1', '10', '=B3*C3'],
    ['', '', '', '=B4*C4'],
    ['Grand total', '', '', '=SUM(D2:D4)'],
  ]);
  for (const [key, cell] of data) {
    const { row, col } = parseKey(key);
    if (row === 0 || col === 3 || row === 4) cell.readOnly = true;
  }

  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 12;
  sheet.cols = 6;
  sheet.setData(data);

  sheet.validateCell = (cellId, value) => {
    const { col } = parseKey(cellId);
    if (col === 1 && !/^\d+$/.test(value)) return 'Qty must be a whole number';
    if (col === 2 && isNaN(Number(value))) return 'Unit price must be a number';
    return null;
  };

  sheet.addEventListener('before-change', (event) => {
    const { changes } = (event as CustomEvent<BeforeChangeDetail>).detail;
    for (const change of changes) change.value = change.value.trim();
  });

  return sheet;
}
