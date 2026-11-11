/**
 * Without a `target` the toolbar shows whatever its properties say. From
 * top to bottom: the defaults, a toolbar describing a bold, italic, centered
 * 16px percent cell with red text on yellow in a filtered table, and a
 * `readOnly` toolbar. Read-only disables the controls that change cells;
 * copying, filtering, freezing, finding and fitting columns stay available,
 * as they do in a read-only grid.
 */
import 'y11n-spreadsheet';

export default function toolbarStates() {
  const plain = document.createElement('y11n-toolbar');
  plain.label = 'Defaults';

  const active = document.createElement('y11n-toolbar');
  active.label = 'A formatted cell';
  active.bold = true;
  active.italic = true;
  active.textColor = '#d32f2f';
  active.backgroundColor = '#fff9c4';
  active.textAlign = 'center';
  active.verticalAlign = 'top';
  active.wrap = true;
  active.fontSize = 16;
  active.numberFormat = { type: 'percent', decimals: 0 };
  active.filtered = true;
  active.canUndo = true;
  active.canRedo = false;

  const readOnly = document.createElement('y11n-toolbar');
  readOnly.label = 'Read-only';
  readOnly.readOnly = true;
  readOnly.bold = true;

  const page = document.createElement('div');
  page.style.cssText =
    'display: flex; flex-direction: column; gap: 16px; padding: 16px;';
  page.append(plain, active, readOnly);
  return page;
}
