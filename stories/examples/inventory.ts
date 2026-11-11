/**
 * A reorder list. Status compares the stock with the reorder point, Order
 * qty works out how much to buy, and `COUNTIF` counts the items to reorder.
 * A conditional format marks the rows that need attention.
 *
 * Try it: type 20 into B2.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function inventory() {
  const items = [
    ['Bolts M6', '150', '100'],
    ['Nuts M6', '80', '100'],
    ['Washers M6', '200', '100'],
    ['Screws M4', '45', '100'],
    ['Rivets', '300', '200'],
    ['Springs', '90', '150'],
  ];

  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 15;
  sheet.cols = 8;
  sheet.style.setProperty('--ls-cell-width', '120px');

  sheet.setConditionalFormats([
    {
      range: { start: { row: 1, col: 0 }, end: { row: items.length, col: 4 } },
      condition: { type: 'formula', formula: '=$D2="REORDER"' },
      format: { backgroundColor: '#fff4ce', bold: true },
      label: 'reorder',
    },
  ]);

  sheet.setData(
    gridFromRows([
      ['Product', 'Stock', 'Reorder point', 'Status', 'Order qty'],
      ...items.map((item, i) => {
        const row = i + 2;
        return [
          ...item,
          `=IF(B${row}>C${row}, "OK", "REORDER")`,
          `=IF(B${row}>C${row}, 0, C${row}-B${row}+50)`,
        ];
      }),
      [],
      ['Items to reorder', '', '', `=COUNTIF(D2:D${items.length + 1}, "REORDER")`],
    ])
  );
  return sheet;
}
