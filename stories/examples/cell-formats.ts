/**
 * A cell's `format` is part of the cell, next to its value: text styling,
 * colors, alignment and a number format. Number formats change what a number
 * looks like, not what it is: the dates here are date serial numbers, and the
 * Share column holds fractions such as 0.5.
 *
 * Formats set in the data are there from the start. To change them later,
 * use `setCellFormat()` and `setRangeFormat()` (see "Format API").
 *
 * Try it: select some cells and use the toolbar, or press Ctrl+B, Ctrl+I or
 * Ctrl+U, Ctrl+Shift+4 for currency and Ctrl+Shift+5 for percent. The
 * toolbar shows the format of the cells you select.
 */
import 'y11n-spreadsheet';
import { cellKey } from 'y11n-spreadsheet';
import type { CellFormat } from 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';
import { stack } from './stack.js';

export default function cellFormats() {
  const data = gridFromRows([
    ['Invoice', 'Date', 'Amount', 'Paid', 'Share'],
    ['INV-001', '=DATE(2026, 1, 5)', '1200', '1200', '=D2/C2'],
    ['INV-002', '=DATE(2026, 1, 19)', '860.5', '430.25', '=D3/C3'],
    ['INV-003', '=DATE(2026, 2, 2)', '2300', '0', '=D4/C4'],
    ['Total', '', '=SUM(C2:C4)', '=SUM(D2:D4)', '=D5/C5'],
  ]);

  const header: CellFormat = { bold: true, backgroundColor: '#e3f2fd' };
  const date: CellFormat = {
    numberFormat: { type: 'date', datePattern: 'DD/MM/YYYY' },
  };
  const money: CellFormat = {
    numberFormat: { type: 'currency', currencySymbol: '€', decimals: 2 },
    textAlign: 'right',
  };
  const percent: CellFormat = {
    numberFormat: { type: 'percent', decimals: 0 },
    textAlign: 'right',
  };

  for (let col = 0; col <= 4; col++) {
    data.get(cellKey(0, col))!.format = header;
  }
  for (let row = 1; row <= 4; row++) {
    if (row < 4) data.get(cellKey(row, 1))!.format = date;
    data.get(cellKey(row, 2))!.format = money;
    data.get(cellKey(row, 3))!.format = money;
    data.get(cellKey(row, 4))!.format = percent;
  }
  // Nothing paid on INV-003
  data.get('3:3')!.format = { ...money, textColor: '#b3261e', bold: true };
  data.get('4:0')!.format = { bold: true };

  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 12;
  sheet.cols = 6;
  sheet.setData(data);

  const toolbar = document.createElement('y11n-toolbar');
  toolbar.target = sheet;
  toolbar.items = 'font alignment number clearFormat';
  return stack(toolbar, sheet);
}
