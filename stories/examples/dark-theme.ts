/**
 * The grid is styled with CSS custom properties, all prefixed `--ls-`. They
 * cross the shadow DOM boundary, so a rule in your own stylesheet is enough:
 *
 *     y11n-spreadsheet.dark { --ls-header-bg: #2d2d2d; ... }
 *
 * This example sets them on the elements instead, to keep everything in one
 * file. The toolbar is an element of its own, so it gets the properties it
 * shares with the grid and its own `--ls-toolbar-` ones. The full list of
 * properties is under Reference / Theming.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';
import { stack } from './stack.js';

export default function darkTheme() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 15;
  sheet.cols = 8;
  sheet.style.cssText = `
    background: #1e1e1e;
    --ls-font-family: ui-monospace, 'JetBrains Mono', 'Fira Code', monospace;
    --ls-text-color: #e0e0e0;
    --ls-cell-bg: #1e1e1e;
    --ls-cell-width: 160px;
    --ls-cell-height: 32px;
    --ls-header-bg: #2d2d2d;
    --ls-header-selected-bg: #4a3f63;
    --ls-border-color: #444;
    --ls-selection-border: 2px solid #bb86fc;
    --ls-selection-bg: rgba(187, 134, 252, 0.15);
    --ls-focus-ring: 2px solid #bb86fc;
    --ls-fill-handle-color: #bb86fc;
    --ls-editor-bg: #1e1e1e;
    --ls-editor-shadow: 0 2px 8px rgba(0, 0, 0, 0.5);
    --ls-formula-bar-bg: #2d2d2d;
    --ls-formula-ref-bg: #1e1e1e;
    --ls-formula-input-bg: #1e1e1e;
    --ls-formula-mode-active-bg: #4a3f63;
    --ls-status-bg: #2d2d2d;
    --ls-status-color: #c0c0c0;
  `;
  sheet.setData(
    gridFromRows([
      ['Planet', 'Diameter (km)', 'Distance from Sun (M km)', 'Moons'],
      ['Mercury', '4879', '57.9', '0'],
      ['Venus', '12104', '108.2', '0'],
      ['Earth', '12756', '149.6', '1'],
      ['Mars', '6792', '227.9', '2'],
      ['Jupiter', '142984', '778.6', '95'],
      ['Saturn', '120536', '1433.5', '146'],
      ['Uranus', '51118', '2872.5', '28'],
      ['Neptune', '49528', '4495.1', '16'],
    ])
  );

  const toolbar = document.createElement('y11n-toolbar');
  toolbar.target = sheet;
  toolbar.items = 'history font alignment | sortAscending sortDescending filter';
  toolbar.style.cssText = `
    --ls-font-family: ui-monospace, 'JetBrains Mono', 'Fira Code', monospace;
    --ls-text-color: #e0e0e0;
    --ls-border-color: #444;
    --ls-focus-ring: 2px solid #bb86fc;
    --ls-toolbar-bg: #2d2d2d;
    --ls-toolbar-hover-bg: #3d3d3d;
    --ls-toolbar-active-bg: #4a3f63;
    --ls-toolbar-active-border: #bb86fc;
    --ls-toolbar-input-bg: #1e1e1e;
  `;
  return stack(toolbar, sheet);
}
