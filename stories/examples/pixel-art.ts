/**
 * Not a spreadsheet at all: a 16 x 16 picture drawn by setting the
 * background color of one row of cells at a time with `setCellFormat()`.
 * It shows that formats can be changed at any moment, and how small cells
 * can get.
 */
import 'y11n-spreadsheet';
import { cellKey } from 'y11n-spreadsheet';

export default function pixelArt() {
  const colors: Record<string, string> = {
    _: '#e3f2fd', // sky
    Y: '#fdd835', // duck
    O: '#ff8f00', // beak
    B: '#333333', // eye
    W: '#81d4fa', // water
  };
  const picture = [
    '________________',
    '________________',
    '____YY__________',
    '___YYYY_________',
    '___YBYY_________',
    '___YYYY_________',
    '____YOOO________',
    '____YY__________',
    '___YYYY_________',
    '__YYYYYY________',
    '_YYYYYYYY_______',
    '_YYYYYYYY_______',
    'WWYWWWWYWWWWWWWW',
    'WWWWWWWWWWWWWWWW',
    'WWWWWWWWWWWWWWWW',
    'WWWWWWWWWWWWWWWW',
  ];

  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = picture.length;
  sheet.cols = picture[0].length;
  sheet.gridLabel = 'A duck on the water, drawn with cell colors';
  sheet.style.setProperty('--ls-cell-width', '28px');
  sheet.style.setProperty('--ls-cell-height', '28px');
  sheet.style.setProperty('--ls-border-color', 'transparent');

  // One row every 150ms
  let row = 0;
  const timer = setInterval(() => {
    if (row === picture.length) {
      clearInterval(timer);
      return;
    }
    [...picture[row]].forEach((pixel, col) => {
      sheet.setCellFormat(cellKey(row, col), { backgroundColor: colors[pixel] });
    });
    row++;
  }, 150);

  return sheet;
}
