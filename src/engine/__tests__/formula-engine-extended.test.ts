import { describe, it, expect, beforeEach, vi, afterEach } from 'vitest';
import { FormulaEngine } from '../formula-engine.js';
import type { GridData } from '../../types.js';
import { dateToSerial } from '../../types.js';

function makeData(entries: Record<string, string>): GridData {
  const data: GridData = new Map();
  for (const [key, raw] of Object.entries(entries)) {
    const isNum = !isNaN(Number(raw)) && raw.trim() !== '';
    data.set(key, { rawValue: raw, displayValue: raw, type: isNum ? 'number' : 'text' });
  }
  return data;
}

const num = (n: number | string) => ({ displayValue: String(n), type: 'number' });
const text = (s: string) => ({ displayValue: s, type: 'text' });
const err = (code: string) => ({ displayValue: code, type: 'error' });

describe('FormulaEngine – extended functions', () => {
  let engine: FormulaEngine;

  beforeEach(() => {
    engine = new FormulaEngine();
  });

  // Region | Product | Amount
  const sales = () =>
    makeData({
      '0:0': 'east', '0:1': 'apple', '0:2': '10',
      '1:0': 'west', '1:1': 'apple', '1:2': '20',
      '2:0': 'east', '2:1': 'pear', '2:2': '30',
      '3:0': 'east', '3:1': 'apple', '3:2': '40',
    });

  describe('SUMIFS / COUNTIFS / AVERAGEIFS', () => {
    beforeEach(() => engine.setData(sales()));

    it('SUMIFS sums rows matching every criterion', () => {
      expect(engine.evaluate('=SUMIFS(C1:C4, A1:A4, "east", B1:B4, "apple")')).toEqual(num(50));
    });

    it('SUMIFS supports operator criteria', () => {
      expect(engine.evaluate('=SUMIFS(C1:C4, C1:C4, ">15", A1:A4, "east")')).toEqual(num(70));
    });

    it('COUNTIFS counts rows matching every criterion', () => {
      expect(engine.evaluate('=COUNTIFS(A1:A4, "east", B1:B4, "apple")')).toEqual(num(2));
      expect(engine.evaluate('=COUNTIFS(A1:A4, "north")')).toEqual(num(0));
    });

    it('AVERAGEIFS averages matching rows', () => {
      expect(engine.evaluate('=AVERAGEIFS(C1:C4, A1:A4, "east")')).toEqual(num('26.6666666666667'));
    });

    it('AVERAGEIFS with no match is #DIV/0!', () => {
      expect(engine.evaluate('=AVERAGEIFS(C1:C4, A1:A4, "north")')).toEqual(err('#DIV/0!'));
    });

    it('mismatched range sizes are #VALUE!', () => {
      expect(engine.evaluate('=SUMIFS(C1:C4, A1:A3, "east")')).toEqual(err('#VALUE!'));
      expect(engine.evaluate('=COUNTIFS(A1:A4, "east", B1:B3, "apple")')).toEqual(err('#VALUE!'));
    });

    it('ranges with the same cell count but a different shape are #VALUE!', () => {
      engine.setData(makeData({ '0:0': 'yes', '1:0': 'yes', '0:2': '10', '0:3': '20' }));
      expect(engine.evaluate('=SUMIFS(C1:D1, A1:A2, "yes")')).toEqual(err('#VALUE!'));
    });

    it('errors in matching sum/average cells propagate; non-matching errors are ignored', () => {
      engine.setData(makeData({
        '0:0': 'yes', '0:1': '10',
        '1:0': 'yes', '1:1': '=1/0',
        '2:0': 'no', '2:1': '=1/0',
      }));
      expect(engine.evaluate('=SUMIFS(B1:B2, A1:A2, "yes")')).toEqual(err('#DIV/0!'));
      expect(engine.evaluate('=AVERAGEIFS(B1:B2, A1:A2, "yes")')).toEqual(err('#DIV/0!'));
      expect(engine.evaluate('=SUMIFS(B1:B3, A1:A3, "no", B1:B3, "<>5")')).toEqual(err('#DIV/0!'));
      engine.setData(makeData({ '0:0': 'yes', '0:1': '10', '1:0': 'no', '1:1': '=1/0' }));
      expect(engine.evaluate('=SUMIFS(B1:B2, A1:A2, "yes")')).toEqual(num(10));
    });

    it('supports * ? and ~ wildcards in criteria', () => {
      engine.setData(makeData({ '0:0': 'east', '1:0': 'eastside', '2:0': 'west', '3:0': 'e*st' }));
      expect(engine.evaluate('=COUNTIFS(A1:A4, "east*")')).toEqual(num(2));
      expect(engine.evaluate('=COUNTIFS(A1:A4, "?ast")')).toEqual(num(1));
      expect(engine.evaluate('=COUNTIFS(A1:A4, "<>*st")')).toEqual(num(1));
      expect(engine.evaluate('=COUNTIFS(A1:A4, "e~*st")')).toEqual(num(1));
      expect(engine.evaluate('=COUNTIF(A1:A4, "*SIDE")')).toEqual(num(1));
    });

    it('wildcards only match text, never numbers or booleans', () => {
      engine.setData(makeData({ '0:0': '123', '1:0': 'abc', '2:0': 'TRUE', '3:0': '="456"' }));
      expect(engine.evaluate('=COUNTIF(A1:A4, "*")')).toEqual(num(2));
      expect(engine.evaluate('=COUNTIFS(A1:A4, "1*")')).toEqual(num(0));
      expect(engine.evaluate('=COUNTIFS(A1:A4, "<>*")')).toEqual(num(2));
    });

    it('SUMIFS skips text values in the sum range', () => {
      engine.setData(makeData({ '0:0': 'yes', '0:1': '="10"', '1:0': 'yes', '1:1': '5' }));
      expect(engine.evaluate('=SUMIFS(B1:B2, A1:A2, "yes")')).toEqual(num(5));
    });

    it('wildcards never match error values', () => {
      engine.setData(makeData({ '0:0': '=1/0', '1:0': 'ok', '0:1': '100', '1:1': '5' }));
      expect(engine.evaluate('=SUMIFS(B1:B2, A1:A2, "*")')).toEqual(num(5));
      expect(engine.evaluate('=COUNTIF(A1:A2, "*")')).toEqual(num(1));
      expect(engine.evaluate('=COUNTIFS(A1:A2, "<>*")')).toEqual(num(1));
    });

    it('? matches one character even when its lowercase form is longer', () => {
      engine.setData(makeData({ '0:0': 'İ', '1:0': 'İx' }));
      expect(engine.evaluate('=COUNTIF(A1:A2, "?")')).toEqual(num(1));
      expect(engine.evaluate('=COUNTIF(A1:A2, "?X")')).toEqual(num(1));
    });

    it('pathological wildcard patterns run in bounded time', () => {
      engine.setData(makeData({ '0:0': 'a'.repeat(2000) }));
      const t0 = performance.now();
      expect(engine.evaluate('=COUNTIF(A1:A1, "*a*a*a*a*a*a*a*a*a*a*a*a*ab")')).toEqual(num(0));
      expect(performance.now() - t0).toBeLessThan(500);
    });

    it('consecutive stars and trailing stars match like a single star', () => {
      engine.setData(makeData({ '0:0': 'abc', '1:0': 'ab' }));
      expect(engine.evaluate('=COUNTIF(A1:A2, "a**c")')).toEqual(num(1));
      expect(engine.evaluate('=COUNTIF(A1:A2, "ab*")')).toEqual(num(2));
      expect(engine.evaluate('=COUNTIF(A1:A2, "?b?")')).toEqual(num(1));
    });

    it('wildcard criteria treat regex metacharacters literally', () => {
      engine.setData(makeData({ '0:0': 'a.b(c)', '1:0': 'axb(c)' }));
      expect(engine.evaluate('=COUNTIFS(A1:A2, "a.b(*")')).toEqual(num(1));
    });

    it('a criteria range without criteria is #VALUE!', () => {
      expect(engine.evaluate('=SUMIFS(C1:C4, A1:A4)')).toEqual(err('#VALUE!'));
    });
  });

  describe('IFS', () => {
    it('returns the value of the first true condition', () => {
      engine.setData(makeData({ '0:0': '75' }));
      expect(engine.evaluate('=IFS(A1>=90, "A", A1>=70, "C", TRUE, "F")')).toEqual(text('C'));
    });

    it('is #N/A when no condition is true', () => {
      expect(engine.evaluate('=IFS(1>2, "x", 2>3, "y")')).toEqual(err('#N/A'));
    });

    it('does not evaluate values of untaken branches', () => {
      engine.setData(makeData({ '0:0': '0' }));
      expect(engine.evaluate('=IFS(A1=0, 0, TRUE, 1/A1)')).toEqual(num(0));
    });

    it('does not evaluate conditions after the first match', () => {
      expect(engine.evaluate('=IFS(TRUE, "ok", 1/0, "never")')).toEqual(text('ok'));
    });

    it('a condition without a value is #VALUE!', () => {
      expect(engine.evaluate('=IFS(FALSE, 1, TRUE)')).toEqual(err('#VALUE!'));
    });
  });

  describe('XLOOKUP', () => {
    beforeEach(() =>
      engine.setData(
        makeData({
          '0:0': 'apple', '0:1': '1.5',
          '1:0': 'banana', '1:1': '0.5',
          '2:0': 'cherry', '2:1': '4',
          // Thresholds in E1:E3 → labels in F1:F3
          '0:4': '0', '0:5': 'low',
          '1:4': '50', '1:5': 'mid',
          '2:4': '100', '2:5': 'high',
          // Horizontal table in A6:C7
          '5:0': 'x', '5:1': 'y', '5:2': 'z',
          '6:0': '1', '6:1': '2', '6:2': '3',
        })
      )
    );

    it('finds an exact match (case-insensitive)', () => {
      expect(engine.evaluate('=XLOOKUP("Banana", A1:A3, B1:B3)')).toEqual(num(0.5));
    });

    it('returns if_not_found or #N/A', () => {
      expect(engine.evaluate('=XLOOKUP("kiwi", A1:A3, B1:B3, "none")')).toEqual(text('none'));
      expect(engine.evaluate('=XLOOKUP("kiwi", A1:A3, B1:B3)')).toEqual(err('#N/A'));
    });

    it('match_mode -1 returns the next smaller item', () => {
      expect(engine.evaluate('=XLOOKUP(75, E1:E3, F1:F3, , -1)')).toEqual(text('mid'));
    });

    it('match_mode 1 returns the next larger item', () => {
      expect(engine.evaluate('=XLOOKUP(75, E1:E3, F1:F3, "?", 1)')).toEqual(text('high'));
    });

    it('works on horizontal ranges', () => {
      expect(engine.evaluate('=XLOOKUP("y", A6:C6, A7:C7)')).toEqual(num(2));
    });

    it('search_mode -1 searches from the end', () => {
      engine.setData(makeData({ '0:0': 'a', '1:0': 'a', '0:1': '1', '1:1': '2' }));
      expect(engine.evaluate('=XLOOKUP("a", A1:A2, B1:B2)')).toEqual(num(1));
      expect(engine.evaluate('=XLOOKUP("a", A1:A2, B1:B2, "", 0, -1)')).toEqual(num(2));
    });

    it('finds boolean keys', () => {
      engine.setData(makeData({ '0:0': 'TRUE', '1:0': 'FALSE', '0:1': 'yes', '1:1': 'no' }));
      expect(engine.evaluate('=XLOOKUP(FALSE, A1:A2, B1:B2)')).toEqual(text('no'));
      expect(engine.evaluate('=XLOOKUP(TRUE, A1:A2, B1:B2)')).toEqual(text('yes'));
    });

    it('an error in an unused if_not_found does not break a successful lookup', () => {
      expect(engine.evaluate('=XLOOKUP("apple", A1:A3, B1:B3, 1/0)')).toEqual(num(1.5));
      expect(engine.evaluate('=XLOOKUP("apple", A1:A3, B1:B3, XLOOKUP("zzz", A1:A3, B1:B3))')).toEqual(num(1.5));
    });

    it('an error in if_not_found propagates when the fallback is used', () => {
      expect(engine.evaluate('=XLOOKUP("kiwi", A1:A3, B1:B3, 1/0)')).toEqual(err('#DIV/0!'));
      expect(engine.evaluate('=XLOOKUP("kiwi", A1:A3, B1:B3, XLOOKUP("banana", A1:A3, B1:B3))')).toEqual(num(0.5));
    });

    it('mismatched return range is #VALUE!', () => {
      expect(engine.evaluate('=XLOOKUP("apple", A1:A3, B1:B2)')).toEqual(err('#VALUE!'));
    });

    it('unsupported wildcard mode is #VALUE!', () => {
      expect(engine.evaluate('=XLOOKUP("a*", A1:A3, B1:B3, "", 2)')).toEqual(err('#VALUE!'));
    });
  });

  describe('omitted arguments', () => {
    it('an empty argument slot means "use the default"', () => {
      expect(engine.evaluate('=ROUND(2.5, )')).toEqual(num(3));
      expect(engine.evaluate('=SUM(1, , 2)')).toEqual(num(3));
    });
  });

  describe('math', () => {
    it('ROUNDUP rounds away from zero', () => {
      expect(engine.evaluate('=ROUNDUP(3.141, 2)')).toEqual(num(3.15));
      expect(engine.evaluate('=ROUNDUP(-3.141, 2)')).toEqual(num(-3.15));
      expect(engine.evaluate('=ROUNDUP(1234, -2)')).toEqual(num(1300));
    });

    it('ROUNDDOWN rounds toward zero', () => {
      expect(engine.evaluate('=ROUNDDOWN(3.149, 2)')).toEqual(num(3.14));
      expect(engine.evaluate('=ROUNDDOWN(-3.149, 2)')).toEqual(num(-3.14));
      expect(engine.evaluate('=ROUNDDOWN(1.1, 0)')).toEqual(num(1));
    });

    it('ROUNDUP / ROUNDDOWN reject non-numeric precision', () => {
      expect(engine.evaluate('=ROUNDUP(1.23, "oops")')).toEqual(err('#VALUE!'));
      expect(engine.evaluate('=ROUNDDOWN(1.23, "oops")')).toEqual(err('#VALUE!'));
      expect(engine.evaluate('=ROUNDUP(1.23)')).toEqual(num(2));
    });

    it('ROUNDUP does not overshoot on float drift', () => {
      expect(engine.evaluate('=ROUNDUP(0.1*3, 1)')).toEqual(num(0.3));
    });

    it('INT floors toward negative infinity', () => {
      expect(engine.evaluate('=INT(8.9)')).toEqual(num(8));
      expect(engine.evaluate('=INT(-8.1)')).toEqual(num(-9));
    });

    it('SQRT', () => {
      expect(engine.evaluate('=SQRT(16)')).toEqual(num(4));
      expect(engine.evaluate('=SQRT(-1)')).toEqual(err('#NUM!'));
    });

    it('PRODUCT multiplies numbers in ranges, skipping text', () => {
      engine.setData(makeData({ '0:0': '2', '1:0': 'x', '2:0': '5' }));
      expect(engine.evaluate('=PRODUCT(A1:A3, 3)')).toEqual(num(30));
      expect(engine.evaluate('=PRODUCT(B1:B3)')).toEqual(num(0));
    });

    it('MEDIAN of odd and even counts', () => {
      engine.setData(makeData({ '0:0': '3', '1:0': '1', '2:0': '2', '3:0': '10' }));
      expect(engine.evaluate('=MEDIAN(A1:A3)')).toEqual(num(2));
      expect(engine.evaluate('=MEDIAN(A1:A4)')).toEqual(num(2.5));
      expect(engine.evaluate('=MEDIAN(B1:B2)')).toEqual(err('#NUM!'));
    });

    it('PRODUCT and MEDIAN propagate errors from ranges', () => {
      engine.setData(makeData({ '0:0': '2', '1:0': '=1/0' }));
      expect(engine.evaluate('=PRODUCT(A1:A2)')).toEqual(err('#DIV/0!'));
      expect(engine.evaluate('=MEDIAN(A1:A2)')).toEqual(err('#DIV/0!'));
    });
  });

  describe('TEXTJOIN', () => {
    beforeEach(() =>
      engine.setData(makeData({ '0:0': 'a', '2:0': 'c', '3:0': 'TRUE' }))
    );

    it('joins values, skipping empties when asked', () => {
      expect(engine.evaluate('=TEXTJOIN(", ", TRUE, A1:A3)')).toEqual(text('a, c'));
    });

    it('keeps empties when ignore_empty is FALSE', () => {
      expect(engine.evaluate('=TEXTJOIN("-", FALSE, A1:A3)')).toEqual(text('a--c'));
    });

    it('accepts a mix of ranges and scalars, rendering booleans as TRUE/FALSE', () => {
      expect(engine.evaluate('=TEXTJOIN("/", TRUE, "x", A3:A4, 1)')).toEqual(text('x/c/TRUE/1'));
    });

    it('propagates errors', () => {
      engine.setData(makeData({ '0:0': '=1/0' }));
      expect(engine.evaluate('=TEXTJOIN(",", TRUE, A1)')).toEqual(err('#DIV/0!'));
    });
  });

  describe('dates', () => {
    afterEach(() => {
      vi.useRealTimers();
    });

    it('YEAR / MONTH / DAY of a DATE serial', () => {
      expect(engine.evaluate('=YEAR(DATE(2024, 2, 29))')).toEqual(num(2024));
      expect(engine.evaluate('=MONTH(DATE(2024, 2, 29))')).toEqual(num(2));
      expect(engine.evaluate('=DAY(DATE(2024, 2, 29))')).toEqual(num(29));
    });

    it('YEAR ignores the time-of-day fraction', () => {
      // 2023-12-31 23:59:59.9 must still be 2023
      expect(engine.evaluate('=YEAR(DATE(2023, 12, 31) + 0.99999)')).toEqual(num(2023));
    });

    it('rejects negative serials', () => {
      expect(engine.evaluate('=YEAR(-1)')).toEqual(err('#NUM!'));
    });

    it('WEEKDAY return types', () => {
      // 2024-01-01 was a Monday
      expect(engine.evaluate('=WEEKDAY(DATE(2024, 1, 1))')).toEqual(num(2));
      expect(engine.evaluate('=WEEKDAY(DATE(2024, 1, 1), 2)')).toEqual(num(1));
      expect(engine.evaluate('=WEEKDAY(DATE(2024, 1, 1), 3)')).toEqual(num(0));
      expect(engine.evaluate('=WEEKDAY(DATE(2024, 1, 7), 2)')).toEqual(num(7));
      expect(engine.evaluate('=WEEKDAY(1, 9)')).toEqual(err('#NUM!'));
    });

    it('TODAY is the local date as a whole serial', () => {
      vi.useFakeTimers();
      vi.setSystemTime(new Date(2024, 4, 17, 15, 30));
      expect(engine.evaluate('=TODAY()')).toEqual(num(dateToSerial(2024, 4, 17)));
      expect(engine.evaluate('=TODAY()=DATE(2024, 5, 17)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('NOW includes the local time of day', () => {
      vi.useFakeTimers();
      vi.setSystemTime(new Date(2024, 4, 17, 18, 0));
      const value = Number(engine.evaluate('=NOW()').displayValue);
      expect(value).toBeCloseTo(dateToSerial(2024, 4, 17) + 0.75, 6);
    });

    it('TODAY is volatile (re-evaluated, not cached)', () => {
      vi.useFakeTimers();
      vi.setSystemTime(new Date(2024, 0, 1, 12));
      const data = makeData({ '0:0': '=TODAY()' });
      engine.setData(data);
      const first = engine.evaluate('=TODAY()', '0:0');
      vi.setSystemTime(new Date(2024, 0, 2, 12));
      const second = engine.evaluate('=TODAY()', '0:0');
      expect(Number(second.displayValue) - Number(first.displayValue)).toBe(1);
    });
  });

  describe('error literals', () => {
    it('evaluates #REF! and other error codes typed into formulas', () => {
      engine.setData(makeData({}));
      expect(engine.evaluate('=#REF!')).toMatchObject(err('#REF!'));
      expect(engine.evaluate('=#REF!+1')).toMatchObject(err('#REF!'));
      expect(engine.evaluate('=SUM(#REF!)')).toMatchObject(err('#REF!'));
      expect(engine.evaluate('=#n/a')).toMatchObject(err('#N/A'));
      expect(engine.evaluate('=#DIV/0!')).toMatchObject(err('#DIV/0!'));
    });

    it('lets IFERROR catch them', () => {
      engine.setData(makeData({}));
      expect(engine.evaluate('=IFERROR(#REF!, 0)')).toMatchObject(num(0));
      // An unused fallback is not evaluated, so its error can't surface
      expect(engine.evaluate('=IFERROR(1, #REF!)')).toMatchObject(num(1));
      expect(engine.evaluate('=IFERROR(1, 1/0)')).toMatchObject(num(1));
      expect(engine.evaluate('=IFERROR(1/0, #N/A)')).toMatchObject(err('#N/A'));
    });

    it('still rejects an unknown #name', () => {
      engine.setData(makeData({}));
      expect(engine.evaluate('=#FOO').type).toBe('error');
    });
  });
});
