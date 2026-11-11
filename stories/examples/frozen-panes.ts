/**
 * `frozenRows` and `frozenCols` (`frozen-rows` / `frozen-cols` in HTML) keep
 * the first rows and columns in view while the rest scrolls. Frozen cells
 * keep their place in the grid, so screen readers report the same row and
 * column numbers as without freezing.
 *
 * Try it: scroll down and to the right. Alt+Shift+F unfreezes; pressed again
 * it freezes the rows above and the columns left of the active cell. The
 * toolbar's Freeze panes menu does the same through `freezePanes()`.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';
import { stack } from './stack.js';

export default function frozenPanes() {
  const months = 'Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec'.split(' ');
  const rows = [['Region', ...months, 'Total']];
  for (let r = 1; r <= 80; r++) {
    const values = months.map((_, m) => String(((r * 37 + m * 53) % 900) + 100));
    rows.push([`Region ${r}`, ...values, `=SUM(B${r + 1}:M${r + 1})`]);
  }

  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 100;
  sheet.cols = 16;
  sheet.frozenRows = 1;
  sheet.frozenCols = 1;
  sheet.setData(gridFromRows(rows));

  const toolbar = document.createElement('y11n-toolbar');
  toolbar.target = sheet;
  toolbar.items = 'freeze';
  return stack(toolbar, sheet);
}
