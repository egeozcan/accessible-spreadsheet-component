/**
 * `registerFunction(name, fn)` adds a function to the formula language. The
 * function receives a context (for reading other cells) followed by its
 * arguments, and returns the result.
 *
 * With `args`, the grid passes single values of that kind and calls the
 * function once per cell when it is given a range: `=NET(A2:A5, 0.19)` in E2
 * returns four values, which spill down the column.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function customFunctions() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 15;
  sheet.cols = 8;
  sheet.style.setProperty('--ls-cell-width', '130px');

  sheet.registerFunction('TAX', (_ctx, price, rate) => {
    return Number(price) * (1 + Number(rate));
  });

  sheet.registerFunction('DISCOUNT', (_ctx, price, percent) => {
    return Number(price) * (1 - Number(percent) / 100);
  });

  sheet.registerFunction(
    'NET',
    (_ctx, gross, rate) => {
      const net = Number(gross) / (1 + Number(rate));
      return Math.round(net * 100) / 100;
    },
    { args: ['number', 'number'] }
  );

  sheet.setData(
    gridFromRows([
      ['Price', 'Tax rate', 'With tax', '20% off', 'Net of 19%'],
      ['100', '0.08', '=TAX(A2, B2)', '=DISCOUNT(A2, 20)', '=NET(A2:A5, 0.19)'],
      ['250', '0.10', '=TAX(A3, B3)', '=DISCOUNT(A3, 20)'],
      ['49.99', '0.065', '=TAX(A4, B4)', '=DISCOUNT(A4, 20)'],
      ['1200', '0.12', '=TAX(A5, B5)', '=DISCOUNT(A5, 20)'],
    ])
  );
  return sheet;
}
