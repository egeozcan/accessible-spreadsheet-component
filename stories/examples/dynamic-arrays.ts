/**
 * A formula whose result is several values spills: the values fill the cells
 * below and to the right of it. E2 lists the regions with `UNIQUE`, G2
 * filters and sorts three columns at once, and E7 doubles a range and lays
 * the result out in a row.
 *
 * `G2#` means "everything G2 spills", and `C:C` the whole column. A value in
 * the way turns the formula into `#SPILL!`, as in E11: delete E12 to let it
 * through.
 *
 * Try it: move onto a spilled value. Its range is outlined, the formula bar
 * shows where it comes from, and a screen reader hears "spilled from G2".
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function dynamicArrays() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 20;
  sheet.cols = 10;
  sheet.setData(
    gridFromRows([
      ['Product', 'Region', 'Units', '', 'Unique regions', '', 'North, by units'],
      ['Apples', 'North', '120', '', '=UNIQUE(B2:B8)', '', '=SORT(FILTER(A2:C8, B2:B8="North"), 3, -1)'],
      ['Pears', 'South', '80'],
      ['Plums', 'North', '45'],
      ['Cherries', 'East', '200'],
      ['Grapes', 'North', '150', '', 'Units x 2'],
      ['Melons', 'South', '60', '', '=TRANSPOSE(C2:C4*2)'],
      ['Kiwis', 'East', '30'],
      [],
      ['Total units', '=SUM(C:C)', '', '', 'Blocked', '', 'Rows in the spill'],
      ['', '', '', '', '=SEQUENCE(3)', '', '=ROWS(G2#)'],
      ['', '', '', '', 'in the way'],
    ])
  );
  return sheet;
}
