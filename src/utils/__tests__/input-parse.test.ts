import { describe, expect, it } from 'vitest';
import { parseInput } from '../input-parse.js';
import { dateToSerial } from '../../types.js';

describe('parseInput', () => {
  it('leaves plain numbers, text and formulas alone', () => {
    for (const text of ['', '42', '-3.5', 'hello', '=A1', '12abc', '1,23', '$', '%', '1/5/2026', '$$5', '-$-5']) {
      expect(parseInput(text), text).toBeNull();
    }
  });

  it('reads percentages without float noise', () => {
    expect(parseInput('15%')).toEqual({ value: '0.15', numberFormat: { type: 'percent', decimals: 0 } });
    expect(parseInput('1.1%')).toEqual({ value: '0.011', numberFormat: { type: 'percent', decimals: 1 } });
    expect(parseInput('-2.50 %')).toEqual({ value: '-0.025', numberFormat: { type: 'percent', decimals: 2 } });
    expect(parseInput('1,200%')?.value).toBe('12');
    expect(parseInput('1,20%')).toBeNull();
  });

  it('reads currency with the symbol and sign in either order', () => {
    expect(parseInput('$1,200')).toEqual({
      value: '1200',
      numberFormat: { type: 'currency', decimals: 0, currencySymbol: '$' },
    });
    expect(parseInput('-€3.5')).toEqual({
      value: '-3.5',
      numberFormat: { type: 'currency', decimals: 2, currencySymbol: '€' },
    });
    expect(parseInput('£-3')?.value).toBe('-3');
    expect(parseInput(' ¥ 500 ')?.numberFormat.currencySymbol).toBe('¥');
  });

  it('reads numbers with thousands separators', () => {
    expect(parseInput('1,234,567.80')).toEqual({
      value: '1234567.8',
      numberFormat: { type: 'number', decimals: 2, thousandsSep: true },
    });
    expect(parseInput('-1,000')?.value).toBe('-1000');
    expect(parseInput('12,34')).toBeNull();
  });

  it('reads year-first dates and rejects days that do not exist', () => {
    expect(parseInput('2026-01-05')).toEqual({
      value: String(dateToSerial(2026, 0, 5)),
      numberFormat: { type: 'date', datePattern: 'YYYY-MM-DD' },
    });
    expect(parseInput('2026/1/5')?.value).toBe(String(dateToSerial(2026, 0, 5)));
    expect(parseInput('2024-02-29')).not.toBeNull();
    expect(parseInput('2026-02-29')).toBeNull();
    expect(parseInput('2026-13-01')).toBeNull();
    expect(parseInput('2026-00-10')).toBeNull();
  });

  it('reads times and date-times as fractions of a day', () => {
    expect(parseInput('18:00')).toEqual({ value: '0.75', numberFormat: { type: 'date', datePattern: 'HH:mm' } });
    expect(parseInput('6:00:30')?.numberFormat.datePattern).toBe('HH:mm:ss');
    expect(parseInput('24:00')).toBeNull();
    expect(parseInput('12:60')).toBeNull();
    const dt = parseInput('2026-01-05 12:00');
    expect(Number(dt?.value)).toBeCloseTo(dateToSerial(2026, 0, 5) + 0.5, 9);
    expect(dt?.numberFormat.datePattern).toBe('YYYY-MM-DD HH:mm');
  });
});
