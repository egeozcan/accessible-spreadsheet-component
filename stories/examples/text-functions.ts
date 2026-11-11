/**
 * Working with text: `&` joins values, `LEFT`, `MID` and `RIGHT` cut pieces
 * out, `FIND` gives a position, `SUBSTITUTE` replaces, and `UPPER`, `TRIM` and
 * `LEN` tidy and measure.
 *
 * The same formulas are written once per row below; in the grid you would
 * type the first one and drag the fill handle down.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function textFunctions() {
  const people = [
    ['  Alice Johnson ', 'alice.johnson@example.com', '(555) 123-4567'],
    ['Bob Smith', 'bob.smith@example.com', '(555) 987-6543'],
    ['Charlie Brown', 'charlie.brown@example.com', '(555) 456-7890'],
  ];

  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 12;
  sheet.cols = 10;
  sheet.style.setProperty('--ls-cell-width', '150px');
  sheet.setData(
    gridFromRows([
      ['Full name', 'Email', 'Phone', 'Trimmed', 'First name', 'Last name', 'Domain', 'Digits only', 'Label'],
      ...people.map((person, i) => {
        const row = i + 2;
        return [
          ...person,
          `=TRIM(A${row})`,
          `=LEFT(D${row}, FIND(" ", D${row})-1)`,
          `=MID(D${row}, FIND(" ", D${row})+1, 100)`,
          `=MID(B${row}, FIND("@", B${row})+1, 100)`,
          `=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(C${row}, "(", ""), ")", ""), " ", ""), "-", "")`,
          `=UPPER(F${row}) & ", " & E${row} & " (" & LEN(D${row}) & " characters)"`,
        ];
      }),
    ])
  );
  return sheet;
}
