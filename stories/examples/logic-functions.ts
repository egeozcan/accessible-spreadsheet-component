/**
 * Conditions in formulas. `IF` picks one of two values, `AND`, `OR` and `NOT`
 * combine conditions, and `IFERROR` replaces an error with a fallback.
 *
 * Try it: lower one of Alice's scores below 70 and row 2 changes.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function logicFunctions() {
  const students = [
    ['Alice', '92', '88', '95', '98'],
    ['Bob', '65', '72', '58', '85'],
    ['Charlie', '78', '81', '74', '92'],
    ['Diana', '45', '52', '48', '0'],
  ];

  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 15;
  sheet.cols = 10;
  sheet.style.setProperty('--ls-cell-width', '120px');
  sheet.setData(
    gridFromRows([
      ['Student', 'Math', 'Science', 'English', 'Days present', 'Passed all?', 'Passed any?', 'Result', 'Points per day'],
      // The same formulas on every row, written for that row
      ...students.map((student, i) => {
        const row = i + 2;
        return [
          ...student,
          `=AND(B${row}>=70, C${row}>=70, D${row}>=70)`,
          `=OR(B${row}>=70, C${row}>=70, D${row}>=70)`,
          `=IF(F${row}, "Pass", IF(NOT(G${row}), "Fail", "Resit"))`,
          `=IFERROR(ROUND(SUM(B${row}:D${row})/E${row}, 2), "no days")`,
        ];
      }),
    ])
  );
  return sheet;
}
