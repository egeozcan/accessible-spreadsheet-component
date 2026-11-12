/**
 * A budget with formulas, number formats and conditional formats working
 * together. Difference is Budget minus Actual; a conditional-format rule
 * colors it red when it is negative and green otherwise, so the colors
 * follow the numbers.
 *
 * Try it: raise an Actual value above its Budget, or restyle the table with
 * the toolbar.
 */
import 'y11n-spreadsheet';
import { cellKey, parseKey } from 'y11n-spreadsheet';
import type { CellFormat } from 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';
import { stack } from './stack.js';

export default function budgetTracker() {
  const categories = [
    ['Housing', '2000', '1950'],
    ['Food', '800', '920'],
    ['Transport', '400', '380'],
    ['Entertainment', '300', '450'],
    ['Utilities', '250', '230'],
    ['Healthcare', '200', '150'],
    ['Savings', '500', '500'],
  ];
  // Rows count from 1 in formulas and from 0 in keys and ranges
  // In formulas: the row of the last category
  const last = categories.length + 1;
  // In keys: the header, the categories and a blank row come before the totals
  const totals = categories.length + 2;

  const data = gridFromRows([
    ['Category', 'Budget', 'Actual', 'Difference', 'Status'],
    ...categories.map((category, i) => {
      const row = i + 2;
      return [...category, `=B${row}-C${row}`, `=IF(D${row}>=0, "Under", "Over")`];
    }),
    [],
    ['Total', `=SUM(B2:B${last})`, `=SUM(C2:C${last})`, `=SUM(D2:D${last})`],
  ]);

  // Formats that never change are part of the data
  const header: CellFormat = {
    bold: true,
    backgroundColor: '#1565c0',
    textColor: '#ffffff',
  };
  const money: CellFormat = {
    numberFormat: { type: 'currency', currencySymbol: '$', decimals: 0 },
    textAlign: 'right',
  };
  for (const [key, cell] of data) {
    const { row, col } = parseKey(key);
    if (row === 0) cell.format = header;
    else if (col >= 1 && col <= 3) {
      cell.format = row === totals ? { ...money, bold: true } : money;
    }
  }
  data.get(cellKey(totals, 0))!.format = { bold: true };

  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 18;
  sheet.cols = 8;

  // Formats that depend on the values are conditional-format rules
  const differences = {
    start: { row: 1, col: 3 },
    end: { row: totals, col: 4 },
  };
  sheet.setConditionalFormats([
    {
      range: differences,
      condition: { type: 'formula', formula: '=$D2<0' },
      format: { textColor: '#b3261e', bold: true },
      label: 'over budget',
    },
    {
      range: differences,
      condition: { type: 'formula', formula: '=AND($D2<>"", $D2>=0)' },
      format: { textColor: '#1b5e20' },
      label: 'within budget',
    },
  ]);

  sheet.setData(data);

  const toolbar = document.createElement('y11n-toolbar');
  toolbar.target = sheet;
  toolbar.items = 'history font number styles';
  return stack(toolbar, sheet);
}
