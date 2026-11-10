import { describe, it, expect, beforeEach } from 'vitest';
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
const bool = (b: boolean) => ({ displayValue: b ? 'TRUE' : 'FALSE', type: 'boolean' });
const err = (code: string) => ({ displayValue: code, type: 'error' });
const value = (engine: FormulaEngine, f: string) => Number(engine.evaluate(f).displayValue);

describe('FormulaEngine – information functions', () => {
  let engine: FormulaEngine;
  beforeEach(() => {
    engine = new FormulaEngine();
    engine.setData(makeData({ '0:0': '5', '0:1': 'hi', '0:2': 'TRUE', '0:3': '=1/0', '0:4': '=""' }));
  });

  it('ISBLANK is true only for an empty cell', () => {
    expect(engine.evaluate('=ISBLANK(Z9)')).toEqual(bool(true));
    expect(engine.evaluate('=ISBLANK(A1)')).toEqual(bool(false));
    expect(engine.evaluate('=ISBLANK(E1)')).toEqual(bool(false)); // formula returning ""
    expect(engine.evaluate('=ISBLANK("")')).toEqual(bool(false));
  });

  it('ISNUMBER / ISTEXT / ISLOGICAL classify values', () => {
    expect(engine.evaluate('=ISNUMBER(A1)')).toEqual(bool(true));
    expect(engine.evaluate('=ISNUMBER(Z9)')).toEqual(bool(false));
    expect(engine.evaluate('=ISNUMBER("5")')).toEqual(bool(false));
    expect(engine.evaluate('=ISNUMBER(A1*2)')).toEqual(bool(true));
    expect(engine.evaluate('=ISTEXT(B1)')).toEqual(bool(true));
    expect(engine.evaluate('=ISTEXT(Z9)')).toEqual(bool(false));
    expect(engine.evaluate('=ISTEXT(D1)')).toEqual(bool(false));
    expect(engine.evaluate('=ISLOGICAL(C1)')).toEqual(bool(true));
    expect(engine.evaluate('=ISLOGICAL(A1)')).toEqual(bool(false));
  });

  it('ISERROR / ISERR / ISNA receive errors instead of propagating them', () => {
    expect(engine.evaluate('=ISERROR(D1)')).toEqual(bool(true));
    expect(engine.evaluate('=ISERROR(1/0)')).toEqual(bool(true));
    expect(engine.evaluate('=ISERROR(A1)')).toEqual(bool(false));
    expect(engine.evaluate('=ISERR(D1)')).toEqual(bool(true));
    expect(engine.evaluate('=ISNA(MATCH(99, A1:A1, 0))')).toEqual(bool(true));
    expect(engine.evaluate('=ISERR(MATCH(99, A1:A1, 0))')).toEqual(bool(false));
    expect(engine.evaluate('=ISNA(D1)')).toEqual(bool(false));
    expect(engine.evaluate('=IF(ISERROR(D1), "bad", "ok")')).toEqual({ displayValue: 'bad', type: 'text' });
  });

  it('IS functions are usable inside larger expressions', () => {
    expect(engine.evaluate('=AND(ISBLANK(Z9), ISNUMBER(A1))')).toEqual(bool(true));
  });

  it('IS functions recalculate when the referenced cell changes', () => {
    const data = makeData({ '0:0': '=ISBLANK(B1)' });
    engine.setData(data);
    engine.recalculate();
    data.set('0:1', { rawValue: '1', displayValue: '1', type: 'number' });
    engine.updateData(data, ['0:1']);
    expect(engine.recalculateAffected(['0:1']).has('0:0')).toBe(true);
  });
});

describe('FormulaEngine – statistics', () => {
  let engine: FormulaEngine;
  beforeEach(() => {
    engine = new FormulaEngine();
    // 2, 4, 4, 4, 5, 5, 7, 9 (population sd 2, sample var 32/7) plus text and a boolean
    engine.setData(makeData({
      '0:0': '2', '1:0': '4', '2:0': '4', '3:0': '4', '4:0': '5', '5:0': '5', '6:0': '7', '7:0': '9',
      '8:0': 'text', '9:0': 'TRUE',
    }));
  });

  it('computes sample and population variance and deviation', () => {
    expect(value(engine, '=VAR.P(A1:A10)')).toBeCloseTo(4);
    expect(value(engine, '=STDEV.P(A1:A10)')).toBeCloseTo(2);
    expect(value(engine, '=VAR(A1:A10)')).toBeCloseTo(32 / 7);
    expect(value(engine, '=VAR.S(A1:A10)')).toBeCloseTo(32 / 7);
    expect(value(engine, '=STDEV(A1:A10)')).toBeCloseTo(Math.sqrt(32 / 7));
    expect(value(engine, '=STDEV.S(A1:A8)')).toBeCloseTo(Math.sqrt(32 / 7));
  });

  it('returns #DIV/0! with too few values', () => {
    expect(engine.evaluate('=STDEV(A1)')).toEqual(err('#DIV/0!'));
    expect(engine.evaluate('=VAR.P(A9)')).toEqual(err('#DIV/0!'));
    expect(value(engine, '=VAR.P(A1)')).toBe(0);
  });

  it('propagates errors in the values', () => {
    engine.setData(makeData({ '0:0': '2', '1:0': '4', '0:1': '=1/0' }));
    expect(engine.evaluate('=STDEV(A1:B2)')).toEqual(err('#DIV/0!'));
  });

  it('dotted names still leave plain references alone', () => {
    expect(engine.evaluate('=A1+A2')).toEqual(num(6));
  });
});

describe('FormulaEngine – SUMPRODUCT', () => {
  let engine: FormulaEngine;
  beforeEach(() => {
    engine = new FormulaEngine();
    engine.setData(makeData({
      '0:0': '1', '1:0': '2', '2:0': '3',
      '0:1': '4', '1:1': 'x', '2:1': '6',
    }));
  });

  it('multiplies equal-sized ranges element-wise, treating non-numbers as 0', () => {
    expect(engine.evaluate('=SUMPRODUCT(A1:A3, B1:B3)')).toEqual(num(22));
    expect(engine.evaluate('=SUMPRODUCT(A1:A3)')).toEqual(num(6));
    expect(engine.evaluate('=SUMPRODUCT(A1:A4, B1:B4)')).toEqual(num(22)); // blanks are 0
  });

  it('returns #VALUE! for ranges of different sizes', () => {
    expect(engine.evaluate('=SUMPRODUCT(A1:A3, B1:B2)')).toEqual(err('#VALUE!'));
    expect(engine.evaluate('=SUMPRODUCT(A1:A3, A1:C1)')).toEqual(err('#VALUE!'));
  });
});

describe('FormulaEngine – EDATE / EOMONTH', () => {
  const engine = new FormulaEngine();
  const d = (y: number, m: number, day: number) => dateToSerial(y, m - 1, day);

  it('EDATE moves by whole months and clamps to the month end', () => {
    expect(value(engine, `=EDATE(${d(2024, 1, 15)}, 1)`)).toBe(d(2024, 2, 15));
    expect(value(engine, `=EDATE(${d(2024, 1, 31)}, 1)`)).toBe(d(2024, 2, 29));
    expect(value(engine, `=EDATE(${d(2023, 3, 31)}, -1)`)).toBe(d(2023, 2, 28));
    expect(value(engine, `=EDATE(${d(2024, 1, 15)}, 13.9)`)).toBe(d(2025, 2, 15));
    expect(value(engine, `=EDATE(${d(2024, 1, 15)}, -2)`)).toBe(d(2023, 11, 15));
  });

  it('EOMONTH returns the last day of the target month', () => {
    expect(value(engine, `=EOMONTH(${d(2024, 1, 15)}, 0)`)).toBe(d(2024, 1, 31));
    expect(value(engine, `=EOMONTH(${d(2024, 1, 15)}, 1)`)).toBe(d(2024, 2, 29));
    expect(value(engine, `=EOMONTH(${d(2024, 1, 15)}, -1)`)).toBe(d(2023, 12, 31));
    expect(value(engine, '=EOMONTH(DATE(2024,5,10), 1.5)')).toBe(d(2024, 6, 30));
  });

  it('rejects negative dates', () => {
    expect(engine.evaluate('=EDATE(-1, 1)')).toEqual(err('#NUM!'));
    expect(engine.evaluate('=EOMONTH(1, -2)')).toEqual(err('#NUM!'));
  });
});

describe('FormulaEngine – financial', () => {
  const engine = new FormulaEngine();

  it('PMT matches Excel', () => {
    expect(value(engine, '=PMT(0.08/12, 10, 10000)')).toBeCloseTo(-1037.0320893, 6);
    expect(value(engine, '=PMT(0.08/12, 10, 10000, 0, 1)')).toBeCloseTo(-1030.1643271, 6);
    expect(value(engine, '=PMT(0.06/12, 18*12, 0, 50000)')).toBeCloseTo(-129.0811609, 6);
    expect(value(engine, '=PMT(0, 10, 1000)')).toBeCloseTo(-100);
    expect(engine.evaluate('=PMT(0.1, 0, 1000)')).toEqual(err('#NUM!'));
  });

  it('FV matches Excel', () => {
    expect(value(engine, '=FV(0.06/12, 10, -200, -500, 1)')).toBeCloseTo(2581.4033741, 6);
    expect(value(engine, '=FV(0.12/12, 12, -1000)')).toBeCloseTo(12682.503013, 5);
    expect(value(engine, '=FV(0, 12, -100, -1000)')).toBeCloseTo(2200);
  });

  it('PV matches Excel', () => {
    expect(value(engine, '=PV(0.08/12, 12*20, 500)')).toBeCloseTo(-59777.145851, 5);
    expect(value(engine, '=PV(0, 10, -100)')).toBeCloseTo(1000);
    expect(value(engine, '=PV(0.05, 10, 0, 1000, 1)')).toBeCloseTo(-613.9132535, 6);
  });

  it('NPV discounts values from the first period and skips text', () => {
    engine.setData(makeData({ '0:0': '-10000', '1:0': '3000', '2:0': '4200', '3:0': 'n/a', '4:0': '6800' }));
    expect(value(engine, '=NPV(0.1, A1:A5)')).toBeCloseTo(1188.4434123, 6);
    expect(value(engine, '=NPV(0.1, A1, A2, A3, A5)')).toBeCloseTo(1188.4434123, 6);
    expect(engine.evaluate('=NPV(-1, A1:A5)')).toEqual(err('#DIV/0!'));
  });
});
