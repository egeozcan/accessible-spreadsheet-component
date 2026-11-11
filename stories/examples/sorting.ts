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
 * Try it: click a cell in the Score column, then a button.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

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

  const actions: Record<string, () => void> = {
    'Sort A to Z': () => sheet.sortSelection('ascending'),
    'Sort Z to A': () => sheet.sortSelection('descending'),
    'By team, then score (high to low)': () => {
      const keys = [{ col: 1 }, { col: 2, direction: 'descending' as const }];
      sheet.sort(table, keys, { hasHeader: true });
    },
  };

  const buttons = document.createElement('div');
  for (const [label, run] of Object.entries(actions)) {
    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = label;
    button.addEventListener('click', run);
    buttons.append(button);
  }

  // Layout: buttons above the grid
  const page = document.createElement('div');
  page.style.cssText = 'display: flex; flex-direction: column; height: 100%;';
  buttons.style.cssText =
    'display: flex; flex-wrap: wrap; gap: 8px; padding: 8px 12px;';
  sheet.style.cssText = 'flex: 1; min-height: 0;';
  page.append(buttons, sheet);
  return page;
}
