/**
 * A reference like `B2` is relative: copied one column to the right it
 * becomes `C2`. A `$` pins the part after it, so `$B$5` stays on the tax rate
 * wherever the formula goes.
 *
 * Try it: change the tax rate in B5. Then select B8, copy it and paste it
 * into F8: `B2` has moved to `F2`, `$B$5` has not.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function absoluteReferences() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 18;
  sheet.cols = 8;
  sheet.style.setProperty('--ls-cell-width', '130px');
  sheet.setData(
    gridFromRows([
      ['', 'Jan', 'Feb', 'Mar', 'Apr'],
      ['Sales', '10000', '12000', '11500', '13000'],
      ['Costs', '7000', '7500', '7200', '7800'],
      [],
      ['Tax rate', '0.08'],
      [],
      ['', 'Jan', 'Feb', 'Mar', 'Apr'],
      ['Sales after tax', '=B2*(1-$B$5)', '=C2*(1-$B$5)', '=D2*(1-$B$5)', '=E2*(1-$B$5)'],
      ['Profit', '=B2-B3', '=C2-C3', '=D2-D3', '=E2-E3'],
      ['Profit after tax', '=B9*(1-$B$5)', '=C9*(1-$B$5)', '=D9*(1-$B$5)', '=E9*(1-$B$5)'],
      ['Share of Jan sales', '=B2/$B$2', '=C2/$B$2', '=D2/$B$2', '=E2/$B$2'],
    ])
  );
  return sheet;
}
