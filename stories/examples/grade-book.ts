/**
 * Formulas across rows and down columns. Each student's row has the average,
 * best and worst of their four tests; the rows under the table do the same
 * per test. The class average in F8 reads other formula cells.
 *
 * The header row and the Student column are frozen, and the averages are
 * rounded with a number format, so the full value is still there to
 * calculate with.
 */
import 'y11n-spreadsheet';
import { parseKey } from 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function gradeBook() {
  const students = [
    ['Alice', '92', '88', '95', '90'],
    ['Bob', '78', '82', '75', '88'],
    ['Charlie', '95', '91', '98', '94'],
    ['Diana', '85', '79', '82', '86'],
    ['Eve', '88', '92', '87', '91'],
  ];

  const data = gridFromRows([
    ['Student', 'Test 1', 'Test 2', 'Test 3', 'Test 4', 'Average', 'Best', 'Worst'],
    ...students.map((student, i) => {
      const row = i + 2;
      const tests = `B${row}:E${row}`;
      return [...student, `=AVERAGE(${tests})`, `=MAX(${tests})`, `=MIN(${tests})`];
    }),
    [],
    ['Class average', '=AVERAGE(B2:B6)', '=AVERAGE(C2:C6)', '=AVERAGE(D2:D6)', '=AVERAGE(E2:E6)', '=AVERAGE(F2:F6)'],
    ['Best', '=MAX(B2:B6)', '=MAX(C2:C6)', '=MAX(D2:D6)', '=MAX(E2:E6)'],
  ]);

  // One decimal for every average
  const oneDecimal = { numberFormat: { type: 'number' as const, decimals: 1 } };
  for (const [key, cell] of data) {
    const { row, col } = parseKey(key);
    // The Average column, and the Class average row
    if (row > 0 && (col === 5 || row === 7)) cell.format = oneDecimal;
  }

  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 15;
  sheet.cols = 10;
  sheet.frozenRows = 1;
  sheet.frozenCols = 1;
  sheet.style.setProperty('--ls-cell-width', '110px');
  sheet.setData(data);
  return sheet;
}
