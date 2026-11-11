/**
 * `items` chooses the toolbar's controls and their order. It takes the names
 * of single controls (`bold`, `filter`), the names of groups (`history`,
 * `font`: see `TOOLBAR_GROUPS`) and `|` for a separator, as text or as an
 * array, which TypeScript checks. In HTML it is an attribute:
 *
 *     <y11n-toolbar for="sheet" items="history | bold italic">
 *
 * A grid can have several toolbars. Here one above the grid has what this
 * small table needs, and one below it has the view controls. Each has its
 * own `label`, which is what screen readers call it.
 *
 * `TOOLBAR_ITEMS` lists every control, to start from everything and leave
 * some out: `TOOLBAR_ITEMS.filter((item) => item !== 'paste')`.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';
import { stack } from './stack.js';

export default function toolbarItems() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 30;
  sheet.cols = 10;
  sheet.setData(
    gridFromRows([
      ['Task', 'Owner', 'Hours', 'Done'],
      ['Wireframes', 'Ann', '12', '1'],
      ['Copy', 'Bob', '5', '0.4'],
      ['Build', 'Cy', '30', '0.25'],
      ['Review', 'Ann', '3', '0'],
    ])
  );

  const editing = document.createElement('y11n-toolbar');
  editing.target = sheet;
  editing.label = 'Editing';
  editing.items = [
    'history', // a group: undo and redo
    '|',
    'bold',
    'italic',
    'backgroundColor',
    '|',
    'numberFormat',
    '|',
    'sortAscending',
    'sortDescending',
    'filter',
  ];

  const view = document.createElement('y11n-toolbar');
  view.target = sheet;
  view.label = 'View';
  view.items = 'view autoFit'; // the same, as text

  return stack(editing, sheet, view);
}
