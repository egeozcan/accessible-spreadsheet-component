/**
 * `<y11n-format-toolbar>` is a separate element that holds no state of its
 * own. You wire it to a grid in two directions:
 *
 * 1. When the selection or a format changes, set the toolbar's properties so
 *    its buttons show the format of the selected cells.
 * 2. When the toolbar fires `format-action`, apply that format to the grid.
 *
 * `conditional-format-action` comes from the toolbar's Conditional
 * formatting menu.
 *
 * Try it: select cells and use the toolbar. Ctrl+B / I / U work in the grid
 * and the toolbar follows.
 */
import 'y11n-spreadsheet';
import { computeSelectionFormat } from 'y11n-spreadsheet';
import type {
  ConditionalFormatActionDetail,
  FormatActionDetail,
  SelectionChangeDetail,
  SelectionRange,
} from 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function formatToolbar() {
  const toolbar = document.createElement('y11n-format-toolbar');
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 20;
  sheet.cols = 10;
  sheet.setData(
    gridFromRows([
      ['Name', 'Department', 'Salary'],
      ['Alice', 'Engineering', '95000'],
      ['Bob', 'Design', '82000'],
      ['Charlie', 'Marketing', '78000'],
    ])
  );

  // The grid starts with A1 selected
  let selection: SelectionRange = {
    start: { row: 0, col: 0 },
    end: { row: 0, col: 0 },
  };

  // 1. Grid -> toolbar
  const showSelectionFormat = () => {
    // The properties that all selected cells share; mixed ones are undefined
    const format = computeSelectionFormat(
      (cellId) => ({ format: sheet.getCellFormat(cellId) }),
      selection
    );
    toolbar.bold = format.bold ?? false;
    toolbar.italic = format.italic ?? false;
    toolbar.underline = format.underline ?? false;
    toolbar.strikethrough = format.strikethrough ?? false;
    toolbar.textColor = format.textColor ?? '#000000';
    toolbar.backgroundColor = format.backgroundColor ?? '#ffffff';
    toolbar.textAlign = format.textAlign ?? 'left';
    toolbar.fontSize = format.fontSize ?? 13;
  };

  sheet.addEventListener('selection-change', (event) => {
    selection = (event as CustomEvent<SelectionChangeDetail>).detail.range;
    showSelectionFormat();
  });
  sheet.addEventListener('format-change', showSelectionFormat);

  // 2. Toolbar -> grid
  toolbar.addEventListener('format-action', (event) => {
    const { action, value } = (event as CustomEvent<FormatActionDetail>).detail;
    switch (action) {
      case 'bold':
      case 'italic':
      case 'underline':
      case 'strikethrough':
        sheet.toggleFormat(action); // on the selection, as Ctrl+B does
        break;
      case 'clearFormat':
        sheet.clearRangeFormat(selection);
        break;
      default:
        // textColor, backgroundColor, textAlign, fontSize
        sheet.setRangeFormat(selection, { [action]: value });
    }
  });

  toolbar.addEventListener('conditional-format-action', (event) => {
    const { detail } = event as CustomEvent<ConditionalFormatActionDetail>;
    const { action } = detail;
    if (action === 'manage') void sheet.openConditionalFormatManager();
    else if (action === 'clear') sheet.clearSelectionConditionalFormats();
    else sheet.applyConditionalFormatPreset(action);
  });

  // Layout: the toolbar above the grid
  const page = document.createElement('div');
  page.style.cssText = 'display: flex; flex-direction: column; height: 100%;';
  sheet.style.cssText = 'flex: 1; min-height: 0;';
  page.append(toolbar, sheet);
  return page;
}
