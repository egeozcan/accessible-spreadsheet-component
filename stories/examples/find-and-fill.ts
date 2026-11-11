/**
 * Editing helpers that need no code: the find bar, the fill handle, and
 * inserting or deleting rows. `openFind()` opens the find bar from code, as
 * Ctrl+F does.
 *
 * Try it:
 * - Type "jan" in the find bar; Enter and Shift+Enter step through matches.
 * - Select B2:B3 and drag the small square at the bottom-right corner of the
 *   selection down: the week numbers continue as a series.
 * - Press Shift+Space to select a row, then Ctrl+Shift+= to insert a row
 *   above it or Ctrl+- to delete it. The SUM in C7 follows.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function findAndFill() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 30;
  sheet.cols = 6;
  sheet.setData(
    gridFromRows([
      ['Month', 'Week', 'Sales'],
      ['January', '1', '120'],
      ['January', '2', '135'],
      ['February', '', '150'],
      ['March', '', '110'],
      [],
      ['Total', '', '=SUM(C2:C5)'],
    ])
  );

  // The find bar is part of what the element renders, so wait for that
  sheet.updateComplete.then(() => sheet.openFind());
  return sheet;
}
