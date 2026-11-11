/**
 * `sort(range, keys, options)` reorders the rows of a range by one or more
 * columns. `sortSelection(direction)` does what Alt+Shift+Up / Down do: it
 * sorts the table around the active cell by that cell's column and keeps a
 * header row in place.
 *
 * Formulas move with their rows, so each bonus still belongs to its score.
 * "Top score" is below a blank row, so it is not part of the table. Blank
 * scores sort last either way, and Ctrl+Z undoes a sort.
 *
 * The toolbar's two sort buttons call `sortSelection()`; the button under
 * it calls `sort()` with two keys.
 *
 * Try it: click a cell in the Score column, then a sort button.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';
import { stack } from './stack.js';

export default function sorting() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 20;
  sheet.cols = 6;
  sheet.setData(
    gridFromRows([
      ['Name', 'Team', 'Score', 'Bonus'],
      ['Dana', 'Blue', '82', '=C2*0.1'],
      ['Ali', 'Red', '95', '=C3*0.1'],
      ['Chen', 'Blue', '78', '=C4*0.1'],
      ['Bea', 'Red', '95', '=C5*0.1'],
      ['Emil', 'Green', '', '=C6*0.1'],
      [],
      ['Top score', '', '=MAX(C2:C6)'],
    ])
  );

  const table = { start: { row: 0, col: 0 }, end: { row: 5, col: 3 } };

  const toolbar = document.createElement('y11n-toolbar');
  toolbar.target = sheet;
  toolbar.items = 'sortAscending sortDescending | undo redo';

  const button = document.createElement('button');
  button.type = 'button';
  button.textContent = 'By team, then score (high to low)';
  button.addEventListener('click', () => {
    const keys = [{ col: 1 }, { col: 2, direction: 'descending' as const }];
    sheet.sort(table, keys, { hasHeader: true });
  });
  const buttons = document.createElement('div');
  buttons.style.cssText = 'padding: 8px 12px;';
  buttons.append(button);

  return stack(toolbar, buttons, sheet);
}
