/**
 * A conditional-format rule is a range, a condition and a format. Rules live
 * beside the data and are applied when the cells are drawn, so the cells'
 * own formats stay as they are for copy, sort and export.
 *
 * Sales highlights duplicates and values above average, Growth has data bars
 * (negative ones grow left from the zero axis), Stock a color scale, Score an
 * icon set, and a formula rule makes the overdue rows bold. Rules are in
 * priority order: where two set the same property, the first one wins.
 *
 * Color never carries the meaning alone: every rule has a label ("above
 * average", "data bar 64%", "up arrow") that screen readers hear with the
 * cell.
 *
 * Try it: edit a number and the whole range restyles. Alt+Shift+C opens the
 * rule manager.
 */
import 'y11n-spreadsheet';
import { gridFromRows } from './grid-from-rows.js';

export default function conditionalFormatting() {
  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 20;
  sheet.cols = 8;

  /** Rows 2 to 7 of a column */
  const column = (col: number) => ({ start: { row: 1, col }, end: { row: 6, col } });

  // Set before the data is loaded, the rules are not an undo step: setData()
  // starts a fresh undo history
  sheet.setConditionalFormats([
    {
      range: { start: { row: 1, col: 0 }, end: { row: 6, col: 5 } },
      // Written for the first cell of the range and shifted for the others
      condition: { type: 'formula', formula: '=$F2="overdue"' },
      format: { bold: true, textColor: '#9c0006' },
      label: 'overdue',
    },
    {
      range: column(1),
      condition: { type: 'duplicates' },
      format: { backgroundColor: '#ffc7ce', textColor: '#9c0006' },
    },
    {
      range: column(1),
      condition: { type: 'aboveAverage' },
      format: { backgroundColor: '#c6efce', textColor: '#006100' },
    },
    { range: column(2), condition: { type: 'dataBar' } },
    {
      range: column(3),
      condition: {
        type: 'colorScale',
        min: { type: 'min', color: '#f8696b' },
        mid: { type: 'percentile', value: 50, color: '#ffeb84' },
        max: { type: 'max', color: '#63be7b' },
      },
    },
    { range: column(4), condition: { type: 'iconSet', icons: 'arrows5' } },
  ]);

  sheet.setData(
    gridFromRows([
      ['Region', 'Sales', 'Growth', 'Stock', 'Score', 'Status'],
      ['North', '120', '12', '40', '88', 'ok'],
      ['South', '80', '-5', '15', '42', 'overdue'],
      ['East', '200', '30', '70', '95', 'ok'],
      ['West', '45', '-20', '5', '12', 'overdue'],
      ['Central', '150', '8', '55', '67', 'ok'],
      ['Islands', '80', '0', '30', '51', 'pending'],
    ])
  );
  return sheet;
}
