import { describe, expect, it } from 'vitest';
import type { CellData } from '../../types.js';
import { compareSortValues, sortValueOf, type SortValue } from '../sort.js';

const cell = (rawValue: string, type: CellData['type'], displayValue = rawValue): CellData => ({
  rawValue,
  displayValue,
  type,
});
const value = (c: CellData | undefined) => sortValueOf(c, (x) => Number(x.displayValue));

describe('sortValueOf', () => {
  it('classifies cells by type', () => {
    expect(value(undefined)).toEqual({ kind: 'blank' });
    expect(value(cell('', 'text'))).toEqual({ kind: 'blank' });
    expect(value(cell('5', 'number'))).toEqual({ kind: 'number', value: 5 });
    expect(value(cell('abc', 'text'))).toEqual({ kind: 'text', value: 'abc' });
    expect(value(cell('TRUE', 'boolean'))).toEqual({ kind: 'boolean', value: true });
    expect(value(cell('=1/0', 'error', '#DIV/0!'))).toEqual({ kind: 'error' });
  });

  it('treats a formula that evaluates to an empty string as blank', () => {
    expect(value(cell('=""', 'text', ''))).toEqual({ kind: 'blank' });
  });

  it('uses the supplied number for formatted number cells', () => {
    const formatted = cell('1234.5', 'number', '$1,234.50');
    expect(sortValueOf(formatted, () => 1234.5)).toEqual({ kind: 'number', value: 1234.5 });
  });
});

describe('compareSortValues', () => {
  const mixed: SortValue[] = [
    { kind: 'blank' },
    { kind: 'error' },
    { kind: 'boolean', value: true },
    { kind: 'text', value: 'b' },
    { kind: 'boolean', value: false },
    { kind: 'number', value: 10 },
    { kind: 'text', value: 'A' },
    { kind: 'number', value: -2 },
  ];
  const label = (v: SortValue) => ('value' in v ? String(v.value) : v.kind);

  it('orders numbers, text, FALSE, TRUE, errors, then blanks when ascending', () => {
    const sorted = [...mixed].sort((a, b) => compareSortValues(a, b, 'ascending'));
    expect(sorted.map(label)).toEqual(['-2', '10', 'A', 'b', 'false', 'true', 'error', 'blank']);
  });

  it('reverses the order when descending but keeps blanks last', () => {
    const sorted = [...mixed].sort((a, b) => compareSortValues(a, b, 'descending'));
    expect(sorted.map(label)).toEqual(['error', 'true', 'false', 'b', 'A', '10', '-2', 'blank']);
  });

  it('compares text case-insensitively, so a stable sort keeps the original order', () => {
    const a: SortValue = { kind: 'text', value: 'apple' };
    const b: SortValue = { kind: 'text', value: 'APPLE' };
    expect(compareSortValues(a, b)).toBe(0);
    expect(Math.abs(compareSortValues({ kind: 'error' }, { kind: 'error' }, 'descending'))).toBe(0);
  });
});
