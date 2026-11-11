/**
 * `setFilter(range, columns)` turns a range into a filtered table. Its first
 * row is the header row: each header cell gets a filter button, and a row
 * stays visible only if every filtered column shows one of its values.
 *
 * Here the table is filtered to two teams. The row numbers skip the hidden
 * rows and turn blue.
 *
 * Try it: press Alt+Down on a header cell (or click its button) to open the
 * filter menu. Ctrl+Shift+L removes the filter, Ctrl+Z brings it back.
 * Arrow keys, copy, paste, fill and Delete all skip the hidden rows.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function filter() {
  const teams = ['Ops', 'Sales', 'Support'];
  const cities = ['Berlin', 'Lisbon', 'Oslo', ''];
  const rows = [['Name', 'Team', 'City', 'Score']];
  for (let i = 1; i <= 60; i++) {
    const score = String((i * 37) % 100);
    rows.push([`Person ${i}`, teams[(i * 7) % 3], cities[(i * 5) % 4], score]);
  }

  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 80;
  sheet.cols = 6;
  sheet.frozenRows = 1;
  sheet.setData(gridFromRows(rows));

  // Setting a filter is an undo step, like changing one from the menu
  const table = { start: { row: 0, col: 0 }, end: { row: 60, col: 3 } };
  sheet.setFilter(table, [{ col: 1, values: ['Sales', 'Support'] }]);
  return sheet;
}
