/**
 * The toolbar shows whatever its properties say. From top to bottom: the
 * defaults, a toolbar describing a bold, italic, centered 16px cell with
 * red text on yellow, and a `readOnly` toolbar, whose controls are disabled.
 */
import 'y11n-spreadsheet';

export default function toolbarStates() {
  const plain = document.createElement('y11n-format-toolbar');

  const active = document.createElement('y11n-format-toolbar');
  active.bold = true;
  active.italic = true;
  active.textColor = '#d32f2f';
  active.backgroundColor = '#fff9c4';
  active.textAlign = 'center';
  active.fontSize = 16;

  const readOnly = document.createElement('y11n-format-toolbar');
  readOnly.readOnly = true;
  readOnly.bold = true;

  const page = document.createElement('div');
  page.style.cssText =
    'display: flex; flex-direction: column; gap: 16px; padding: 16px;';
  page.append(plain, active, readOnly);
  return page;
}
