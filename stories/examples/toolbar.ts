/**
 * `<y11n-toolbar>` is an element of its own, so it goes wherever your
 * layout wants it. Set its `target` to a grid (in HTML: its `for` attribute
 * to the grid's id) and it wires itself. Its controls show the format of the
 * selected cells, follow undo, the filter and read-only mode, and act on the
 * selection.
 *
 * Without `items` it shows everything it has: undo and redo, the clipboard,
 * fonts, alignment, number formats, conditional formatting, rows and
 * columns, sorting and filtering, lists, AutoSum, frozen panes, find and
 * the name manager.
 *
 * Try it: select the salaries and choose Currency, or click a cell of the
 * table and press Filter. Left and Right move along the toolbar, Tab leaves
 * it.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';
import { stack } from './stack.js';

export default function toolbar() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 30;
  sheet.cols = 10;
  sheet.setData(
    gridFromRows([
      ['Name', 'Department', 'Salary', 'Bonus'],
      ['Alice', 'Engineering', '95000', '0.1'],
      ['Bob', 'Design', '82000', '0.08'],
      ['Charlie', 'Marketing', '78000', '0.05'],
      ['Dana', 'Engineering', '101000', '0.12'],
    ])
  );

  const toolbar = document.createElement('y11n-toolbar');
  toolbar.target = sheet;

  return stack(toolbar, sheet);
}
