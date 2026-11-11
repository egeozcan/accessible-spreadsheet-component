/**
 * The same properties as the dark theme with other values: a serif font,
 * pink headers and borders, and a pink selection.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function pastelTheme() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 12;
  sheet.cols = 6;
  sheet.style.cssText = `
    --ls-font-family: Georgia, serif;
    --ls-text-color: #4a148c;
    --ls-cell-width: 140px;
    --ls-cell-height: 34px;
    --ls-header-bg: #fce4ec;
    --ls-header-selected-bg: #f8bbd0;
    --ls-border-color: #f8bbd0;
    --ls-selection-border: 2px solid #c2185b;
    --ls-selection-bg: rgba(233, 30, 99, 0.08);
    --ls-focus-ring: 2px solid #c2185b;
    --ls-fill-handle-color: #c2185b;
    --ls-formula-bar-bg: #fce4ec;
    --ls-formula-mode-active-bg: #f8bbd0;
    --ls-status-bg: #fce4ec;
    --ls-status-color: #4a148c;
  `;
  sheet.setData(
    gridFromRows([
      ['Task', 'Assignee', 'Priority', 'Status'],
      ['Design mockups', 'Alice', 'High', 'Done'],
      ['API integration', 'Bob', 'High', 'In progress'],
      ['Unit tests', 'Charlie', 'Medium', 'To do'],
      ['Documentation', 'Diana', 'Low', 'To do'],
      ['Code review', 'Eve', 'Medium', 'In progress'],
    ])
  );
  return sheet;
}
