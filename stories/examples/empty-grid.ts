/**
 * The smallest setup: import the package once, which registers the
 * `<y11n-spreadsheet>` element, and put the element on the page. In HTML it is
 * `<y11n-spreadsheet rows="50" cols="10"></y11n-spreadsheet>`.
 *
 * The grid fills its parent, so give the parent a height.
 *
 * Try it: click a cell and type, press Enter to commit, move with the arrow
 * keys, select with Shift+Arrow, undo with Ctrl+Z.
 */
import 'y11n-spreadsheet';

export default function emptyGrid() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 50;
  sheet.cols = 10;
  return sheet;
}
