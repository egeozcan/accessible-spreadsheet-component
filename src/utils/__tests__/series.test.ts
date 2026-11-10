import { describe, expect, it } from 'vitest';
import { seriesFor, type SeriesCell } from '../series.js';

const text = (...raws: string[]): SeriesCell[] => raws.map((raw) => ({ raw, number: null, isDate: false }));
const nums = (...ns: number[]): SeriesCell[] => ns.map((n) => ({ raw: String(n), number: n, isDate: false }));
const run = (cells: SeriesCell[], offsets: number[]) => {
  const series = seriesFor(cells);
  return series ? offsets.map(series) : null;
};

describe('seriesFor', () => {
  it('continues numbers by their linear trend and repeats a single number', () => {
    expect(run(nums(1, 2), [2, 3, -1])).toEqual(['3', '4', '0']);
    expect(run(nums(0.1, 0.2), [2])).toEqual(['0.3']);
    expect(seriesFor(nums(5))).toBeNull();
  });

  it('steps a single date by one day', () => {
    expect(run([{ raw: '46000', number: 46000, isDate: true }], [1, 2, -1])).toEqual(['46001', '46002', '45999']);
  });

  it('continues day and month names, wrapping and keeping the letter case', () => {
    expect(run(text('Mon'), [1, 6, 7])).toEqual(['Tue', 'Sun', 'Mon']);
    expect(run(text('Friday'), [1, 2])).toEqual(['Saturday', 'Sunday']);
    expect(run(text('NOV', 'DEC'), [2, 3])).toEqual(['JAN', 'FEB']);
    expect(run(text('january', 'march'), [2])).toEqual(['may']);
    expect(run(text('Apr', 'May'), [2])).toEqual(['Jun']);
    expect(run(text('Dec'), [-1])).toEqual(['Nov']);
  });

  it('counts up text ending in a number, keeping zero padding', () => {
    expect(run(text('Item 1'), [1, 2])).toEqual(['Item 2', 'Item 3']);
    expect(run(text('Q09'), [1])).toEqual(['Q10']);
    expect(run(text('Row 1', 'Row 3'), [2, 3])).toEqual(['Row 5', 'Row 7']);
  });

  it('repeats mixed or unrelated content', () => {
    expect(seriesFor(text('Item 1', 'Thing 2'))).toBeNull();
    expect(seriesFor(text('Row 1', 'Row 2', 'Row 4'))).toBeNull();
    expect(seriesFor(text('hello'))).toBeNull();
    expect(seriesFor(text('=A1'))).toBeNull();
    expect(seriesFor(text('42'))).toBeNull();
    expect(seriesFor([...nums(1), ...text('Mon')])).toBeNull();
    expect(seriesFor(text('Mon', 'Jan'))).toBeNull();
  });
});
