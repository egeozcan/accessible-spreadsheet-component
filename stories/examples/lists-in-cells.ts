/**
 * A cell can hold several values. A list is stored as a formula of constants,
 * `={{"Ann";"Bob"}}`, which `listToFormula()` writes for you. Wrapping any
 * formula in braces keeps its whole result in one cell instead of spilling:
 * `={TEXTSPLIT(...)}`.
 *
 * `HAS`, `HASANY` and `HASALL` test the items, `COUNTA` counts them and
 * `FLATTEN` lays nested values out in the grid.
 *
 * Try it: press Alt+Down on a list to see its values, F2 to edit its items,
 * or select a cell with comma-separated text and press Ctrl+J. The filter on
 * the Owners header lists people, not combinations.
 */
import 'y11n-spreadsheet';
import { listToFormula } from 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function listsInCells() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 16;
  sheet.cols = 8;
  sheet.setColumnWidths({ 1: 130, 2: 170, 3: 120 });
  sheet.setData(
    gridFromRows([
      ['Project', 'Owners', 'Milestones', 'Owned by Bob?', 'Owner count', 'Ann and Bob?'],
      [
        'Website',
        listToFormula(['Ann', 'Bob']),
        '={TEXTSPLIT("Design,Build,Launch", ",")}',
        '=HAS(B2, "Bob")',
        '=COUNTA(B2)',
        '=HASALL(B2, {"Ann";"Bob"})',
      ],
      [
        'Mobile app',
        listToFormula(['Bob', 'Cy', 'Di']),
        '={TEXTSPLIT("Spec,Beta", ",")}',
        '=HAS(B3, "Bob")',
        '=COUNTA(B3)',
        '=HASALL(B3, {"Ann";"Bob"})',
      ],
      [
        'Data import',
        listToFormula(['Ann']),
        '={TEXTSPLIT("Plan", ",")}',
        '=HAS(B4, "Bob")',
        '=COUNTA(B4)',
        '=HASALL(B4, {"Ann";"Bob"})',
      ],
      [],
      ['Everyone', '=UNIQUE(FLATTEN(B2:B4))', '', 'All milestones', '=FLATTEN(C2:C4, "")'],
    ])
  );
  // Filter buttons on the table's header row, nothing hidden yet
  sheet.setFilter({ start: { row: 0, col: 0 }, end: { row: 3, col: 5 } });
  return sheet;
}
