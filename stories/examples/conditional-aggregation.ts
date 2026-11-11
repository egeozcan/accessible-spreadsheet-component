/**
 * Totals for the rows that match a condition: `SUMIF`, `COUNTIF` and
 * `AVERAGEIF` take a range to test, a criterion and the range to add up.
 * A criterion can be a value, a cell, or a comparison such as `">100"`.
 *
 * Try it: change a category in column C, or an amount in column D.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function conditionalAggregation() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 25;
  sheet.cols = 6;
  sheet.style.setProperty('--ls-cell-width', '140px');
  sheet.setData(
    gridFromRows([
      ['Date', 'Description', 'Category', 'Amount'],
      ['2024-01-05', 'Office supplies', 'Office', '125.00'],
      ['2024-01-08', 'Cloud hosting', 'Technology', '450.00'],
      ['2024-01-12', 'Team lunch', 'Food', '89.50'],
      ['2024-01-15', 'Software license', 'Technology', '299.00'],
      ['2024-01-18', 'Printer paper', 'Office', '35.00'],
      ['2024-01-20', 'Coffee beans', 'Food', '42.00'],
      ['2024-01-22', 'Domain renewal', 'Technology', '15.00'],
      ['2024-01-25', 'Desk organizer', 'Office', '28.50'],
      ['2024-01-28', 'Client dinner', 'Food', '156.00'],
      ['2024-01-30', 'Antivirus renewal', 'Technology', '79.99'],
      [],
      ['Category', 'Total', 'Count', 'Average'],
      ['Technology', '=SUMIF(C2:C11, A14, D2:D11)', '=COUNTIF(C2:C11, A14)', '=AVERAGEIF(C2:C11, A14, D2:D11)'],
      ['Office', '=SUMIF(C2:C11, A15, D2:D11)', '=COUNTIF(C2:C11, A15)', '=AVERAGEIF(C2:C11, A15, D2:D11)'],
      ['Food', '=SUMIF(C2:C11, A16, D2:D11)', '=COUNTIF(C2:C11, A16)', '=AVERAGEIF(C2:C11, A16, D2:D11)'],
      [],
      ['Over 100', '=SUMIF(D2:D11, ">100")', '=COUNTIF(D2:D11, ">100")'],
      ['All', '=SUM(D2:D11)', '=COUNTA(D2:D11)', '=AVERAGE(D2:D11)'],
    ])
  );
  return sheet;
}
