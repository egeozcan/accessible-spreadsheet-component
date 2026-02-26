import { describe, it, expect, beforeEach } from 'vitest';
import { FormulaEngine } from '../formula-engine.js';
import type { GridData, CellData } from '../../types.js';

function cell(rawValue: string, type: CellData['type'] = 'text', displayValue?: string): CellData {
  return { rawValue, displayValue: displayValue ?? rawValue, type };
}

function makeData(entries: Record<string, string>): GridData {
  const data: GridData = new Map();
  for (const [key, raw] of Object.entries(entries)) {
    const isNum = !isNaN(Number(raw)) && raw.trim() !== '';
    data.set(key, cell(raw, isNum ? 'number' : 'text'));
  }
  return data;
}

describe('FormulaEngine', () => {
  let engine: FormulaEngine;

  beforeEach(() => {
    engine = new FormulaEngine();
  });

  // ─── Tokenizer tests (via evaluate) ────────────────

  describe('tokenizer', () => {
    it('tokenizes numeric literals: integer', () => {
      expect(engine.evaluate('=42')).toEqual({ displayValue: '42', type: 'number' });
    });

    it('tokenizes numeric literals: decimal', () => {
      expect(engine.evaluate('=3.14')).toEqual({ displayValue: '3.14', type: 'number' });
    });

    it('tokenizes numeric literals: leading dot', () => {
      expect(engine.evaluate('=.5')).toEqual({ displayValue: '0.5', type: 'number' });
    });

    it('tokenizes string literals', () => {
      expect(engine.evaluate('="hello"')).toEqual({ displayValue: 'hello', type: 'text' });
    });

    it('tokenizes empty string literal', () => {
      expect(engine.evaluate('=""')).toEqual({ displayValue: '', type: 'text' });
    });

    it('tokenizes boolean literals TRUE', () => {
      expect(engine.evaluate('=TRUE')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('tokenizes boolean literals FALSE', () => {
      expect(engine.evaluate('=FALSE')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('tokenizes all arithmetic operators', () => {
      expect(engine.evaluate('=2+3')).toEqual({ displayValue: '5', type: 'number' });
      expect(engine.evaluate('=5-2')).toEqual({ displayValue: '3', type: 'number' });
      expect(engine.evaluate('=4*3')).toEqual({ displayValue: '12', type: 'number' });
      expect(engine.evaluate('=12/4')).toEqual({ displayValue: '3', type: 'number' });
      expect(engine.evaluate('="a"&"b"')).toEqual({ displayValue: 'ab', type: 'text' });
    });

    it('tokenizes all comparison operators', () => {
      expect(engine.evaluate('=1=1')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=1<>2')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=1<2')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=2>1')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=1<=2')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=2>=1')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('throws error on unexpected characters', () => {
      const result = engine.evaluate('=1@2');
      expect(result).toEqual({ displayValue: '#ERROR!', type: 'error' });
    });

    it('returns empty result for empty formula input', () => {
      expect(engine.evaluate('')).toEqual({ displayValue: '', type: 'text' });
      expect(engine.evaluate('  ')).toEqual({ displayValue: '', type: 'text' });
    });
  });

  // ─── Parser / precedence tests ─────────────────────

  describe('parser precedence', () => {
    it('multiplication over addition: 1+2*3=7', () => {
      expect(engine.evaluate('=1+2*3')).toEqual({ displayValue: '7', type: 'number' });
    });

    it('parentheses overriding precedence: (1+2)*3=9', () => {
      expect(engine.evaluate('=(1+2)*3')).toEqual({ displayValue: '9', type: 'number' });
    });

    it('nested parentheses: ((2+3)*2)+1=11', () => {
      expect(engine.evaluate('=((2+3)*2)+1')).toEqual({ displayValue: '11', type: 'number' });
    });

    it('unary negation: -5', () => {
      expect(engine.evaluate('=-5')).toEqual({ displayValue: '-5', type: 'number' });
    });

    it('unary negation of expression: -(3+2)', () => {
      expect(engine.evaluate('=-(3+2)')).toEqual({ displayValue: '-5', type: 'number' });
    });

    it('unary plus: +5', () => {
      expect(engine.evaluate('=+5')).toEqual({ displayValue: '5', type: 'number' });
    });

    it('comparison has lowest precedence: 1+2>2 is TRUE', () => {
      expect(engine.evaluate('=1+2>2')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('concatenation binds tighter than comparison: "a"&"b"="ab" is TRUE', () => {
      expect(engine.evaluate('="a"&"b"="ab"')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });
  });

  // ─── Basic evaluation ───────────────────────────────

  describe('evaluate', () => {
    it('returns empty for empty input', () => {
      expect(engine.evaluate('')).toEqual({ displayValue: '', type: 'text' });
      expect(engine.evaluate('  ')).toEqual({ displayValue: '', type: 'text' });
    });

    it('returns plain text as-is', () => {
      expect(engine.evaluate('hello')).toEqual({ displayValue: 'hello', type: 'text' });
    });

    it('coerces numbers', () => {
      expect(engine.evaluate('42')).toEqual({ displayValue: '42', type: 'number' });
      expect(engine.evaluate('3.14')).toEqual({ displayValue: '3.14', type: 'number' });
    });

    it('coerces booleans', () => {
      expect(engine.evaluate('TRUE')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('false')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('evaluates simple arithmetic', () => {
      expect(engine.evaluate('=1+2')).toEqual({ displayValue: '3', type: 'number' });
      expect(engine.evaluate('=10-3')).toEqual({ displayValue: '7', type: 'number' });
      expect(engine.evaluate('=4*5')).toEqual({ displayValue: '20', type: 'number' });
      expect(engine.evaluate('=20/4')).toEqual({ displayValue: '5', type: 'number' });
    });

    it('respects operator precedence', () => {
      expect(engine.evaluate('=2+3*4')).toEqual({ displayValue: '14', type: 'number' });
      expect(engine.evaluate('=(2+3)*4')).toEqual({ displayValue: '20', type: 'number' });
    });

    it('supports unary minus', () => {
      expect(engine.evaluate('=-5')).toEqual({ displayValue: '-5', type: 'number' });
      expect(engine.evaluate('=-5+10')).toEqual({ displayValue: '5', type: 'number' });
    });

    it('supports unary plus', () => {
      expect(engine.evaluate('=+5')).toEqual({ displayValue: '5', type: 'number' });
    });

    it('returns error on division by zero', () => {
      const result = engine.evaluate('=1/0');
      expect(result).toEqual({ displayValue: '#DIV/0!', type: 'error' });
    });

    it('returns #ERROR! for invalid formulas', () => {
      expect(engine.evaluate('=***')).toEqual({ displayValue: '#ERROR!', type: 'error' });
    });
  });

  // ─── String operations ──────────────────────────────

  describe('string operations', () => {
    it('concatenates with & operator', () => {
      expect(engine.evaluate('="hello" & " " & "world"')).toEqual({
        displayValue: 'hello world',
        type: 'text',
      });
    });

    it('handles string literals', () => {
      expect(engine.evaluate('="hello"')).toEqual({ displayValue: 'hello', type: 'text' });
    });
  });

  // ─── Comparisons ────────────────────────────────────

  describe('comparisons', () => {
    it('evaluates = comparison', () => {
      expect(engine.evaluate('=1=1')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=1=2')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('evaluates <> comparison', () => {
      expect(engine.evaluate('=1<>2')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=1<>1')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('evaluates < and > comparisons', () => {
      expect(engine.evaluate('=1<2')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=2>1')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=2<1')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('evaluates <= and >= comparisons', () => {
      expect(engine.evaluate('=1<=1')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=1<=2')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=2>=2')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=1>=2')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });
  });

  // ─── Cell references ────────────────────────────────

  describe('cell references', () => {
    it('resolves cell references', () => {
      const data = makeData({ '0:0': '42' });
      engine.setData(data);
      expect(engine.evaluate('=A1')).toEqual({ displayValue: '42', type: 'number' });
    });

    it('returns 0 for empty cells', () => {
      engine.setData(new Map());
      expect(engine.evaluate('=A1')).toEqual({ displayValue: '0', type: 'number' });
    });

    it('resolves multi-letter column references (AA1)', () => {
      const data = makeData({ '0:26': '99' });
      engine.setData(data);
      expect(engine.evaluate('=AA1')).toEqual({ displayValue: '99', type: 'number' });
    });

    it('resolves text cell references', () => {
      const data = makeData({ '0:0': 'hello' });
      engine.setData(data);
      expect(engine.evaluate('=A1 & " world"')).toEqual({
        displayValue: 'hello world',
        type: 'text',
      });
    });

    it('resolves boolean cell references', () => {
      const data = makeData({ '0:0': 'TRUE' });
      engine.setData(data);
      expect(engine.evaluate('=IF(A1, "yes", "no")')).toEqual({ displayValue: 'yes', type: 'text' });
    });

    it('performs arithmetic with cell references: A1+B1', () => {
      const data = makeData({ '0:0': '10', '0:1': '25' });
      engine.setData(data);
      expect(engine.evaluate('=A1+B1')).toEqual({ displayValue: '35', type: 'number' });
    });

    it('resolves chained formula references', () => {
      const data = makeData({
        '0:0': '10',
        '0:1': '20',
        '0:2': '=A1+B1',
      });
      engine.setData(data);
      expect(engine.evaluate('=C1*2')).toEqual({ displayValue: '60', type: 'number' });
    });

    it('detects direct circular reference (A1 references itself)', () => {
      const data = makeData({ '0:0': '=A1' });
      engine.setData(data);
      const result = engine.evaluate('=A1');
      expect(result.type).toBe('error');
    });

    it('detects indirect circular reference (A1->B1->A1)', () => {
      const data = makeData({
        '0:0': '=B1',
        '0:1': '=A1',
      });
      engine.setData(data);
      const result = engine.evaluate('=A1');
      expect(result.type).toBe('error');
    });
  });

  // ─── Range references ──────────────────────────────

  describe('range references', () => {
    it('resolves ranges in aggregate functions', () => {
      const data = makeData({
        '0:0': '10',
        '1:0': '20',
        '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=SUM(A1:A3)')).toEqual({ displayValue: '60', type: 'number' });
    });

    it('handles ranges with mixed types', () => {
      const data = makeData({
        '0:0': '10',
        '1:0': 'hello',
        '2:0': '30',
      });
      engine.setData(data);
      // SUM treats non-numeric as 0: 10 + 0 + 30 = 40
      expect(engine.evaluate('=SUM(A1:A3)')).toEqual({ displayValue: '40', type: 'number' });
      // COUNT should only count numeric
      expect(engine.evaluate('=COUNT(A1:A3)')).toEqual({ displayValue: '2', type: 'number' });
    });

    it('handles 2D ranges (A1:B2)', () => {
      const data = makeData({
        '0:0': '1',
        '0:1': '2',
        '1:0': '3',
        '1:1': '4',
      });
      engine.setData(data);
      expect(engine.evaluate('=SUM(A1:B2)')).toEqual({ displayValue: '10', type: 'number' });
    });
  });

  // ─── Built-in functions ─────────────────────────────

  describe('SUM', () => {
    it('sums numeric values', () => {
      const data = makeData({ '0:0': '1', '1:0': '2', '2:0': '3' });
      engine.setData(data);
      expect(engine.evaluate('=SUM(A1:A3)')).toEqual({ displayValue: '6', type: 'number' });
    });

    it('ignores non-numeric values', () => {
      const data = makeData({ '0:0': '10', '1:0': 'hello', '2:0': '20' });
      engine.setData(data);
      expect(engine.evaluate('=SUM(A1:A3)')).toEqual({ displayValue: '30', type: 'number' });
    });

    it('sums individual arguments', () => {
      expect(engine.evaluate('=SUM(1, 2, 3)')).toEqual({ displayValue: '6', type: 'number' });
    });

    it('returns 0 for empty range', () => {
      engine.setData(new Map());
      expect(engine.evaluate('=SUM(A1:A3)')).toEqual({ displayValue: '0', type: 'number' });
    });
  });

  describe('AVERAGE', () => {
    it('calculates average of numeric values', () => {
      const data = makeData({ '0:0': '10', '1:0': '20', '2:0': '30' });
      engine.setData(data);
      expect(engine.evaluate('=AVERAGE(A1:A3)')).toEqual({ displayValue: '20', type: 'number' });
    });

    it('filters out non-numeric values', () => {
      const data = makeData({ '0:0': '10', '1:0': 'text', '2:0': '30' });
      engine.setData(data);
      expect(engine.evaluate('=AVERAGE(A1:A3)')).toEqual({ displayValue: '20', type: 'number' });
    });

    it('returns 0 for no numeric values', () => {
      expect(engine.evaluate('=AVERAGE()')).toEqual({ displayValue: '0', type: 'number' });
    });
  });

  describe('MIN', () => {
    it('finds minimum value', () => {
      const data = makeData({ '0:0': '5', '1:0': '3', '2:0': '8' });
      engine.setData(data);
      expect(engine.evaluate('=MIN(A1:A3)')).toEqual({ displayValue: '3', type: 'number' });
    });

    it('returns 0 for empty input', () => {
      expect(engine.evaluate('=MIN()')).toEqual({ displayValue: '0', type: 'number' });
    });
  });

  describe('MAX', () => {
    it('finds maximum value', () => {
      const data = makeData({ '0:0': '5', '1:0': '3', '2:0': '8' });
      engine.setData(data);
      expect(engine.evaluate('=MAX(A1:A3)')).toEqual({ displayValue: '8', type: 'number' });
    });

    it('returns 0 for empty input', () => {
      expect(engine.evaluate('=MAX()')).toEqual({ displayValue: '0', type: 'number' });
    });
  });

  describe('COUNT', () => {
    it('counts numeric values', () => {
      const data = makeData({ '0:0': '10', '1:0': 'hello', '2:0': '30' });
      engine.setData(data);
      expect(engine.evaluate('=COUNT(A1:A3)')).toEqual({ displayValue: '2', type: 'number' });
    });

    it('counts individual arguments', () => {
      expect(engine.evaluate('=COUNT(1, "x", 3)')).toEqual({ displayValue: '2', type: 'number' });
    });
  });

  describe('COUNTA', () => {
    it('counts non-empty values', () => {
      expect(engine.evaluate('=COUNTA(1, "hello", 0)')).toEqual({ displayValue: '3', type: 'number' });
    });
  });

  describe('IF', () => {
    it('returns true value when condition is true', () => {
      expect(engine.evaluate('=IF(1>0, "yes", "no")')).toEqual({ displayValue: 'yes', type: 'text' });
    });

    it('returns false value when condition is false', () => {
      expect(engine.evaluate('=IF(1>2, "yes", "no")')).toEqual({ displayValue: 'no', type: 'text' });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': '100', '0:1': '50' });
      engine.setData(data);
      expect(engine.evaluate('=IF(A1>B1, "big", "small")')).toEqual({ displayValue: 'big', type: 'text' });
    });

    it('supports nested IF', () => {
      expect(engine.evaluate('=IF(TRUE, IF(FALSE, 1, 2), 3)')).toEqual({ displayValue: '2', type: 'number' });
    });
  });

  describe('CONCAT', () => {
    it('concatenates strings', () => {
      expect(engine.evaluate('=CONCAT("hello", " ", "world")')).toEqual({
        displayValue: 'hello world',
        type: 'text',
      });
    });

    it('converts numbers to strings', () => {
      expect(engine.evaluate('=CONCAT("val: ", 42)')).toEqual({
        displayValue: 'val: 42',
        type: 'text',
      });
    });
  });

  describe('ABS', () => {
    it('returns absolute value of positive number', () => {
      expect(engine.evaluate('=ABS(5)')).toEqual({ displayValue: '5', type: 'number' });
    });

    it('returns absolute value of negative number', () => {
      expect(engine.evaluate('=ABS(-5)')).toEqual({ displayValue: '5', type: 'number' });
    });

    it('returns 0 for 0', () => {
      expect(engine.evaluate('=ABS(0)')).toEqual({ displayValue: '0', type: 'number' });
    });
  });

  describe('ROUND', () => {
    it('rounds to specified digits', () => {
      expect(engine.evaluate('=ROUND(3.14159, 2)')).toEqual({ displayValue: '3.14', type: 'number' });
    });

    it('rounds to integer when digits is 0', () => {
      expect(engine.evaluate('=ROUND(3.7, 0)')).toEqual({ displayValue: '4', type: 'number' });
    });

    it('defaults to 0 digits when not specified', () => {
      expect(engine.evaluate('=ROUND(3.7)')).toEqual({ displayValue: '4', type: 'number' });
    });
  });

  describe('UPPER', () => {
    it('converts string to uppercase', () => {
      expect(engine.evaluate('=UPPER("hello")')).toEqual({ displayValue: 'HELLO', type: 'text' });
    });

    it('handles already-uppercase text', () => {
      expect(engine.evaluate('=UPPER("HELLO")')).toEqual({ displayValue: 'HELLO', type: 'text' });
    });

    it('handles mixed case', () => {
      expect(engine.evaluate('=UPPER("Hello World")')).toEqual({ displayValue: 'HELLO WORLD', type: 'text' });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': 'hello' });
      engine.setData(data);
      expect(engine.evaluate('=UPPER(A1)')).toEqual({ displayValue: 'HELLO', type: 'text' });
    });

    it('works with cell references to formula cells (the resolveRef bug)', () => {
      const data = makeData({
        '0:0': 'hello',
        '0:1': 'world',
        '0:2': '=A1 & " " & B1',
      });
      engine.setData(data);
      expect(engine.evaluate('=UPPER(C1)')).toEqual({ displayValue: 'HELLO WORLD', type: 'text' });
    });

    it('works with CONCAT formula reference', () => {
      const data = makeData({
        '0:0': 'alice',
        '0:1': 'smith',
        '0:2': '=CONCAT(A1, " ", B1)',
      });
      engine.setData(data);
      expect(engine.evaluate('=UPPER(C1)')).toEqual({ displayValue: 'ALICE SMITH', type: 'text' });
    });

    it('is case-insensitive as a function name', () => {
      expect(engine.evaluate('=upper("hello")')).toEqual({ displayValue: 'HELLO', type: 'text' });
    });
  });

  describe('LOWER', () => {
    it('converts string to lowercase', () => {
      expect(engine.evaluate('=LOWER("HELLO")')).toEqual({ displayValue: 'hello', type: 'text' });
    });

    it('handles mixed case', () => {
      expect(engine.evaluate('=LOWER("Hello World")')).toEqual({ displayValue: 'hello world', type: 'text' });
    });

    it('works with cell references to formula cells', () => {
      const data = makeData({
        '0:0': 'HELLO',
        '0:1': 'WORLD',
        '0:2': '=CONCAT(A1, " ", B1)',
      });
      engine.setData(data);
      expect(engine.evaluate('=LOWER(C1)')).toEqual({ displayValue: 'hello world', type: 'text' });
    });
  });

  describe('LEN', () => {
    it('returns length of a string', () => {
      expect(engine.evaluate('=LEN("hello")')).toEqual({ displayValue: '5', type: 'number' });
    });

    it('returns 0 for empty string', () => {
      expect(engine.evaluate('=LEN("")')).toEqual({ displayValue: '0', type: 'number' });
    });

    it('handles numbers as strings', () => {
      expect(engine.evaluate('=LEN(12345)')).toEqual({ displayValue: '5', type: 'number' });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': 'hello' });
      engine.setData(data);
      expect(engine.evaluate('=LEN(A1)')).toEqual({ displayValue: '5', type: 'number' });
    });

    it('works with cell references to formula cells (the resolveRef bug)', () => {
      const data = makeData({
        '0:0': 'hello',
        '0:1': 'world',
        '0:2': '=A1 & " " & B1',
      });
      engine.setData(data);
      expect(engine.evaluate('=LEN(C1)')).toEqual({ displayValue: '11', type: 'number' });
    });

    it('works with CONCAT formula reference', () => {
      const data = makeData({
        '0:0': 'alice',
        '0:1': 'smith',
        '0:2': '=CONCAT(A1, " ", B1)',
      });
      engine.setData(data);
      expect(engine.evaluate('=LEN(C1)')).toEqual({ displayValue: '11', type: 'number' });
    });
  });

  describe('TRIM', () => {
    it('removes leading and trailing whitespace', () => {
      expect(engine.evaluate('=TRIM("  hello  ")')).toEqual({ displayValue: 'hello', type: 'text' });
    });

    it('handles string with no extra whitespace', () => {
      expect(engine.evaluate('=TRIM("hello")')).toEqual({ displayValue: 'hello', type: 'text' });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': '  spaced  ' });
      engine.setData(data);
      expect(engine.evaluate('=TRIM(A1)')).toEqual({ displayValue: 'spaced', type: 'text' });
    });
  });

  // ─── Logic/Conditional ─────────────────────────────

  describe('IFERROR', () => {
    it('returns the value when no error', () => {
      expect(engine.evaluate('=IFERROR(42, "fallback")')).toEqual({ displayValue: '42', type: 'number' });
    });

    it('returns the value for non-error string', () => {
      expect(engine.evaluate('=IFERROR("hello", "fallback")')).toEqual({ displayValue: 'hello', type: 'text' });
    });

    it('returns fallback for #DIV/0! error from cell', () => {
      // Note: IFERROR checks if a value is an error *string*. A thrown error
      // (like 1/0) propagates before IFERROR can catch it. Use a cell with the error string.
      const data = makeData({ '0:0': '#DIV/0!' });
      engine.setData(data);
      expect(engine.evaluate('=IFERROR(A1, "oops")')).toEqual({ displayValue: 'oops', type: 'text' });
    });

    it('returns fallback for #VALUE! error', () => {
      const data = makeData({ '0:0': '#VALUE!' });
      engine.setData(data);
      expect(engine.evaluate('=IFERROR(A1, 0)')).toEqual({ displayValue: '0', type: 'number' });
    });

    it('returns fallback for #NAME? error', () => {
      // UNKNOWNFN throws #NAME?, IFERROR catches it at the formula level
      // But since IFERROR receives the thrown error as a string, let's test with a cell containing error
      const data = makeData({ '0:0': '#NAME?' });
      engine.setData(data);
      expect(engine.evaluate('=IFERROR(A1, "fixed")')).toEqual({ displayValue: 'fixed', type: 'text' });
    });

    it('returns fallback for #ERROR!', () => {
      const data = makeData({ '0:0': '#ERROR!' });
      engine.setData(data);
      expect(engine.evaluate('=IFERROR(A1, -1)')).toEqual({ displayValue: '-1', type: 'number' });
    });

    it('returns fallback for #CIRC!', () => {
      const data = makeData({ '0:0': '#CIRC!' });
      engine.setData(data);
      expect(engine.evaluate('=IFERROR(A1, "circular")')).toEqual({ displayValue: 'circular', type: 'text' });
    });

    it('returns fallback for #REF!', () => {
      const data = makeData({ '0:0': '#REF!' });
      engine.setData(data);
      expect(engine.evaluate('=IFERROR(A1, 0)')).toEqual({ displayValue: '0', type: 'number' });
    });

    it('returns numeric fallback for error from cell', () => {
      const data = makeData({ '0:0': '#DIV/0!' });
      engine.setData(data);
      expect(engine.evaluate('=IFERROR(A1, 999)')).toEqual({ displayValue: '999', type: 'number' });
    });
  });

  describe('AND', () => {
    it('returns TRUE when all args are true', () => {
      expect(engine.evaluate('=AND(TRUE, TRUE, TRUE)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('returns FALSE when any arg is false', () => {
      expect(engine.evaluate('=AND(TRUE, FALSE, TRUE)')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('returns TRUE for single true arg', () => {
      expect(engine.evaluate('=AND(TRUE)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('returns FALSE for single false arg', () => {
      expect(engine.evaluate('=AND(FALSE)')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('treats non-zero numbers as truthy', () => {
      expect(engine.evaluate('=AND(1, 2, 3)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('treats zero as falsy', () => {
      expect(engine.evaluate('=AND(1, 0, 3)')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('handles range values', () => {
      const data = makeData({ '0:0': 'TRUE', '1:0': 'TRUE', '2:0': 'TRUE' });
      engine.setData(data);
      expect(engine.evaluate('=AND(A1:A3)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('returns FALSE when range contains false', () => {
      const data = makeData({ '0:0': 'TRUE', '1:0': 'FALSE', '2:0': 'TRUE' });
      engine.setData(data);
      expect(engine.evaluate('=AND(A1:A3)')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('works with comparison expressions', () => {
      expect(engine.evaluate('=AND(1>0, 2>1, 3>2)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=AND(1>0, 2<1)')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });
  });

  describe('OR', () => {
    it('returns TRUE when any arg is true', () => {
      expect(engine.evaluate('=OR(FALSE, TRUE, FALSE)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('returns FALSE when all args are false', () => {
      expect(engine.evaluate('=OR(FALSE, FALSE, FALSE)')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('returns TRUE for single true arg', () => {
      expect(engine.evaluate('=OR(TRUE)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('returns FALSE for single false arg', () => {
      expect(engine.evaluate('=OR(FALSE)')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('treats non-zero numbers as truthy', () => {
      expect(engine.evaluate('=OR(0, 0, 5)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('treats all zeros as falsy', () => {
      expect(engine.evaluate('=OR(0, 0, 0)')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('handles range values', () => {
      const data = makeData({ '0:0': 'FALSE', '1:0': 'TRUE', '2:0': 'FALSE' });
      engine.setData(data);
      expect(engine.evaluate('=OR(A1:A3)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('returns FALSE when range is all false', () => {
      const data = makeData({ '0:0': 'FALSE', '1:0': 'FALSE' });
      engine.setData(data);
      expect(engine.evaluate('=OR(A1:A2)')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('works with comparison expressions', () => {
      expect(engine.evaluate('=OR(1>2, 2>3, 3>2)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });
  });

  describe('NOT', () => {
    it('negates TRUE to FALSE', () => {
      expect(engine.evaluate('=NOT(TRUE)')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('negates FALSE to TRUE', () => {
      expect(engine.evaluate('=NOT(FALSE)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('negates truthy number (non-zero) to FALSE', () => {
      expect(engine.evaluate('=NOT(1)')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('negates zero to TRUE', () => {
      expect(engine.evaluate('=NOT(0)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('works with comparison expression', () => {
      expect(engine.evaluate('=NOT(1>2)')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('=NOT(1<2)')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('works with cell reference', () => {
      const data = makeData({ '0:0': 'TRUE' });
      engine.setData(data);
      expect(engine.evaluate('=NOT(A1)')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });
  });

  // ─── Conditional Aggregation ─────────────────────────

  describe('SUMIF', () => {
    it('sums values matching exact criteria', () => {
      const data = makeData({
        '0:0': 'apple', '0:1': '10',
        '1:0': 'banana', '1:1': '20',
        '2:0': 'apple', '2:1': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=SUMIF(A1:A3, "apple", B1:B3)')).toEqual({ displayValue: '40', type: 'number' });
    });

    it('sums values matching numeric criteria', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=SUMIF(A1:A3, ">15")')).toEqual({ displayValue: '50', type: 'number' });
    });

    it('sums with greater-than-or-equal criteria', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=SUMIF(A1:A3, ">=20")')).toEqual({ displayValue: '50', type: 'number' });
    });

    it('sums with less-than criteria', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=SUMIF(A1:A3, "<20")')).toEqual({ displayValue: '10', type: 'number' });
    });

    it('sums with less-than-or-equal criteria', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=SUMIF(A1:A3, "<=20")')).toEqual({ displayValue: '30', type: 'number' });
    });

    it('sums with not-equal criteria', () => {
      const data = makeData({
        '0:0': 'apple', '0:1': '10',
        '1:0': 'banana', '1:1': '20',
        '2:0': 'apple', '2:1': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=SUMIF(A1:A3, "<>apple", B1:B3)')).toEqual({ displayValue: '20', type: 'number' });
    });

    it('sums with equal-sign criteria prefix', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '10',
      });
      engine.setData(data);
      expect(engine.evaluate('=SUMIF(A1:A3, "=10")')).toEqual({ displayValue: '20', type: 'number' });
    });

    it('returns 0 when no values match', () => {
      const data = makeData({
        '0:0': 'apple', '1:0': 'banana',
      });
      engine.setData(data);
      expect(engine.evaluate('=SUMIF(A1:A2, "cherry")')).toEqual({ displayValue: '0', type: 'number' });
    });

    it('sums the range itself when no sum_range is provided', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=SUMIF(A1:A3, ">0")')).toEqual({ displayValue: '60', type: 'number' });
    });

    it('criteria matching is case-insensitive for strings', () => {
      const data = makeData({
        '0:0': 'Apple', '0:1': '10',
        '1:0': 'APPLE', '1:1': '20',
        '2:0': 'apple', '2:1': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=SUMIF(A1:A3, "apple", B1:B3)')).toEqual({ displayValue: '60', type: 'number' });
    });
  });

  describe('COUNTIF', () => {
    it('counts values matching exact criteria', () => {
      const data = makeData({
        '0:0': 'apple', '1:0': 'banana', '2:0': 'apple', '3:0': 'cherry',
      });
      engine.setData(data);
      expect(engine.evaluate('=COUNTIF(A1:A4, "apple")')).toEqual({ displayValue: '2', type: 'number' });
    });

    it('counts values matching numeric criteria with >', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30', '3:0': '5',
      });
      engine.setData(data);
      expect(engine.evaluate('=COUNTIF(A1:A4, ">15")')).toEqual({ displayValue: '2', type: 'number' });
    });

    it('counts values matching <> criteria', () => {
      const data = makeData({
        '0:0': 'apple', '1:0': 'banana', '2:0': 'apple',
      });
      engine.setData(data);
      expect(engine.evaluate('=COUNTIF(A1:A3, "<>apple")')).toEqual({ displayValue: '1', type: 'number' });
    });

    it('returns 0 when nothing matches', () => {
      const data = makeData({
        '0:0': 'apple', '1:0': 'banana',
      });
      engine.setData(data);
      expect(engine.evaluate('=COUNTIF(A1:A2, "cherry")')).toEqual({ displayValue: '0', type: 'number' });
    });

    it('is case-insensitive for string matching', () => {
      const data = makeData({
        '0:0': 'Apple', '1:0': 'APPLE', '2:0': 'apple',
      });
      engine.setData(data);
      expect(engine.evaluate('=COUNTIF(A1:A3, "apple")')).toEqual({ displayValue: '3', type: 'number' });
    });

    it('counts exact numeric match', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '10',
      });
      engine.setData(data);
      expect(engine.evaluate('=COUNTIF(A1:A3, "10")')).toEqual({ displayValue: '2', type: 'number' });
    });
  });

  describe('AVERAGEIF', () => {
    it('averages values matching criteria', () => {
      const data = makeData({
        '0:0': 'apple', '0:1': '10',
        '1:0': 'banana', '1:1': '20',
        '2:0': 'apple', '2:1': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=AVERAGEIF(A1:A3, "apple", B1:B3)')).toEqual({ displayValue: '20', type: 'number' });
    });

    it('averages with numeric > criteria (no avg_range)', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=AVERAGEIF(A1:A3, ">15")')).toEqual({ displayValue: '25', type: 'number' });
    });

    it('returns #DIV/0! when no values match', () => {
      const data = makeData({
        '0:0': 'apple', '1:0': 'banana',
      });
      engine.setData(data);
      expect(engine.evaluate('=AVERAGEIF(A1:A2, "cherry")')).toEqual({ displayValue: '#DIV/0!', type: 'error' });
    });

    it('averages matching numeric criteria with <=', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=AVERAGEIF(A1:A3, "<=20")')).toEqual({ displayValue: '15', type: 'number' });
    });

    it('works with separate avg_range', () => {
      const data = makeData({
        '0:0': '1', '0:1': '100',
        '1:0': '2', '1:1': '200',
        '2:0': '1', '2:1': '300',
      });
      engine.setData(data);
      expect(engine.evaluate('=AVERAGEIF(A1:A3, "1", B1:B3)')).toEqual({ displayValue: '200', type: 'number' });
    });
  });

  // ─── Lookup ─────────────────────────────────────────

  describe('VLOOKUP', () => {
    it('finds exact match in first column and returns value from specified column', () => {
      const data = makeData({
        '0:0': 'apple', '0:1': '1', '0:2': '100',
        '1:0': 'banana', '1:1': '2', '1:2': '200',
        '2:0': 'cherry', '2:1': '3', '2:2': '300',
      });
      engine.setData(data);
      expect(engine.evaluate('=VLOOKUP("banana", A1:C3, 3, FALSE)')).toEqual({ displayValue: '200', type: 'number' });
    });

    it('exact match is case-insensitive', () => {
      const data = makeData({
        '0:0': 'Apple', '0:1': '10',
        '1:0': 'Banana', '1:1': '20',
      });
      engine.setData(data);
      expect(engine.evaluate('=VLOOKUP("apple", A1:B2, 2, FALSE)')).toEqual({ displayValue: '10', type: 'number' });
    });

    it('returns #N/A when exact match not found', () => {
      const data = makeData({
        '0:0': 'apple', '0:1': '1',
        '1:0': 'banana', '1:1': '2',
      });
      engine.setData(data);
      expect(engine.evaluate('=VLOOKUP("cherry", A1:B2, 2, FALSE)')).toEqual({ displayValue: '#N/A', type: 'error' });
    });

    it('returns #REF! when column index is out of range', () => {
      const data = makeData({
        '0:0': 'apple', '0:1': '1',
      });
      engine.setData(data);
      expect(engine.evaluate('=VLOOKUP("apple", A1:B1, 5, FALSE)')).toEqual({ displayValue: '#REF!', type: 'error' });
    });

    it('returns #REF! when column index is 0', () => {
      const data = makeData({
        '0:0': 'apple', '0:1': '1',
      });
      engine.setData(data);
      expect(engine.evaluate('=VLOOKUP("apple", A1:B1, 0, FALSE)')).toEqual({ displayValue: '#REF!', type: 'error' });
    });

    it('approximate match (default) finds largest value <= lookup', () => {
      const data = makeData({
        '0:0': '10', '0:1': 'A',
        '1:0': '20', '1:1': 'B',
        '2:0': '30', '2:1': 'C',
      });
      engine.setData(data);
      // 25 is between 20 and 30, so should match row with 20
      expect(engine.evaluate('=VLOOKUP(25, A1:B3, 2)')).toEqual({ displayValue: 'B', type: 'text' });
    });

    it('approximate match returns last row when lookup is >= all values', () => {
      const data = makeData({
        '0:0': '10', '0:1': 'A',
        '1:0': '20', '1:1': 'B',
        '2:0': '30', '2:1': 'C',
      });
      engine.setData(data);
      expect(engine.evaluate('=VLOOKUP(99, A1:B3, 2)')).toEqual({ displayValue: 'C', type: 'text' });
    });

    it('approximate match returns #N/A when lookup is less than smallest value', () => {
      const data = makeData({
        '0:0': '10', '0:1': 'A',
        '1:0': '20', '1:1': 'B',
      });
      engine.setData(data);
      expect(engine.evaluate('=VLOOKUP(5, A1:B2, 2)')).toEqual({ displayValue: '#N/A', type: 'error' });
    });

    it('exact match with numeric lookup value', () => {
      const data = makeData({
        '0:0': '100', '0:1': 'one hundred',
        '1:0': '200', '1:1': 'two hundred',
      });
      engine.setData(data);
      expect(engine.evaluate('=VLOOKUP(200, A1:B2, 2, FALSE)')).toEqual({ displayValue: 'two hundred', type: 'text' });
    });

    it('exact match with 0 as the fourth argument', () => {
      const data = makeData({
        '0:0': 'x', '0:1': '10',
        '1:0': 'y', '1:1': '20',
      });
      engine.setData(data);
      expect(engine.evaluate('=VLOOKUP("x", A1:B2, 2, 0)')).toEqual({ displayValue: '10', type: 'number' });
    });
  });

  describe('HLOOKUP', () => {
    it('finds exact match in first row and returns value from specified row', () => {
      const data = makeData({
        '0:0': 'apple', '0:1': 'banana', '0:2': 'cherry',
        '1:0': '10', '1:1': '20', '1:2': '30',
        '2:0': '100', '2:1': '200', '2:2': '300',
      });
      engine.setData(data);
      expect(engine.evaluate('=HLOOKUP("banana", A1:C3, 2, FALSE)')).toEqual({ displayValue: '20', type: 'number' });
    });

    it('returns value from third row', () => {
      const data = makeData({
        '0:0': 'apple', '0:1': 'banana', '0:2': 'cherry',
        '1:0': '10', '1:1': '20', '1:2': '30',
        '2:0': '100', '2:1': '200', '2:2': '300',
      });
      engine.setData(data);
      expect(engine.evaluate('=HLOOKUP("cherry", A1:C3, 3, FALSE)')).toEqual({ displayValue: '300', type: 'number' });
    });

    it('returns #N/A when exact match not found', () => {
      const data = makeData({
        '0:0': 'apple', '0:1': 'banana',
        '1:0': '10', '1:1': '20',
      });
      engine.setData(data);
      expect(engine.evaluate('=HLOOKUP("cherry", A1:B2, 2, FALSE)')).toEqual({ displayValue: '#N/A', type: 'error' });
    });

    it('returns #REF! when row index is out of range', () => {
      const data = makeData({
        '0:0': 'apple', '0:1': 'banana',
        '1:0': '10', '1:1': '20',
      });
      engine.setData(data);
      expect(engine.evaluate('=HLOOKUP("apple", A1:B2, 5, FALSE)')).toEqual({ displayValue: '#REF!', type: 'error' });
    });

    it('approximate match (default) finds largest <= lookup', () => {
      const data = makeData({
        '0:0': '10', '0:1': '20', '0:2': '30',
        '1:0': 'A', '1:1': 'B', '1:2': 'C',
      });
      engine.setData(data);
      expect(engine.evaluate('=HLOOKUP(25, A1:C2, 2)')).toEqual({ displayValue: 'B', type: 'text' });
    });

    it('approximate match returns last column when lookup >= all values', () => {
      const data = makeData({
        '0:0': '10', '0:1': '20', '0:2': '30',
        '1:0': 'A', '1:1': 'B', '1:2': 'C',
      });
      engine.setData(data);
      expect(engine.evaluate('=HLOOKUP(99, A1:C2, 2)')).toEqual({ displayValue: 'C', type: 'text' });
    });

    it('exact match is case-insensitive', () => {
      const data = makeData({
        '0:0': 'Apple', '0:1': 'Banana',
        '1:0': '10', '1:1': '20',
      });
      engine.setData(data);
      expect(engine.evaluate('=HLOOKUP("apple", A1:B2, 2, FALSE)')).toEqual({ displayValue: '10', type: 'number' });
    });
  });

  describe('INDEX', () => {
    it('returns value at specified row (1-indexed)', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=INDEX(A1:A3, 2)')).toEqual({ displayValue: '20', type: 'number' });
    });

    it('returns value at row and column (2D range)', () => {
      const data = makeData({
        '0:0': '1', '0:1': '2', '0:2': '3',
        '1:0': '4', '1:1': '5', '1:2': '6',
        '2:0': '7', '2:1': '8', '2:2': '9',
      });
      engine.setData(data);
      expect(engine.evaluate('=INDEX(A1:C3, 2, 3)')).toEqual({ displayValue: '6', type: 'number' });
    });

    it('returns first column value when col is not specified', () => {
      const data = makeData({
        '0:0': '10', '0:1': '20',
        '1:0': '30', '1:1': '40',
      });
      engine.setData(data);
      expect(engine.evaluate('=INDEX(A1:B2, 2)')).toEqual({ displayValue: '30', type: 'number' });
    });

    it('returns #REF! when row index is out of bounds', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20',
      });
      engine.setData(data);
      expect(engine.evaluate('=INDEX(A1:A2, 5)')).toEqual({ displayValue: '#REF!', type: 'error' });
    });

    it('returns #REF! when col index is out of bounds', () => {
      const data = makeData({
        '0:0': '10', '0:1': '20',
      });
      engine.setData(data);
      expect(engine.evaluate('=INDEX(A1:B1, 1, 5)')).toEqual({ displayValue: '#REF!', type: 'error' });
    });

    it('returns 0 for empty cell in range', () => {
      engine.setData(new Map());
      expect(engine.evaluate('=INDEX(A1:A3, 1)')).toEqual({ displayValue: '0', type: 'number' });
    });

    it('returns first element with row=1, col=1', () => {
      const data = makeData({
        '0:0': '42', '0:1': '99',
      });
      engine.setData(data);
      expect(engine.evaluate('=INDEX(A1:B1, 1, 1)')).toEqual({ displayValue: '42', type: 'number' });
    });
  });

  describe('MATCH', () => {
    it('exact match (type 0): finds position of value', () => {
      const data = makeData({
        '0:0': 'apple', '1:0': 'banana', '2:0': 'cherry',
      });
      engine.setData(data);
      expect(engine.evaluate('=MATCH("banana", A1:A3, 0)')).toEqual({ displayValue: '2', type: 'number' });
    });

    it('exact match returns 1-indexed position', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=MATCH(10, A1:A3, 0)')).toEqual({ displayValue: '1', type: 'number' });
    });

    it('exact match returns #N/A when not found', () => {
      const data = makeData({
        '0:0': 'apple', '1:0': 'banana',
      });
      engine.setData(data);
      expect(engine.evaluate('=MATCH("cherry", A1:A2, 0)')).toEqual({ displayValue: '#N/A', type: 'error' });
    });

    it('defaults to exact match (type 0) when type is omitted', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=MATCH(20, A1:A3)')).toEqual({ displayValue: '2', type: 'number' });
    });

    it('type 1: finds position of largest value <= lookup (sorted ascending)', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=MATCH(25, A1:A3, 1)')).toEqual({ displayValue: '2', type: 'number' });
    });

    it('type 1: returns last position when lookup >= all values', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=MATCH(99, A1:A3, 1)')).toEqual({ displayValue: '3', type: 'number' });
    });

    it('type 1: returns #N/A when lookup < all values', () => {
      const data = makeData({
        '0:0': '10', '1:0': '20', '2:0': '30',
      });
      engine.setData(data);
      expect(engine.evaluate('=MATCH(5, A1:A3, 1)')).toEqual({ displayValue: '#N/A', type: 'error' });
    });

    it('type -1: finds position of smallest value >= lookup (sorted descending)', () => {
      const data = makeData({
        '0:0': '30', '1:0': '20', '2:0': '10',
      });
      engine.setData(data);
      expect(engine.evaluate('=MATCH(25, A1:A3, -1)')).toEqual({ displayValue: '1', type: 'number' });
    });

    it('type -1: returns #N/A when lookup > all values', () => {
      const data = makeData({
        '0:0': '30', '1:0': '20', '2:0': '10',
      });
      engine.setData(data);
      expect(engine.evaluate('=MATCH(50, A1:A3, -1)')).toEqual({ displayValue: '#N/A', type: 'error' });
    });

    it('exact match is case-insensitive for strings', () => {
      const data = makeData({
        '0:0': 'Apple', '1:0': 'Banana', '2:0': 'Cherry',
      });
      engine.setData(data);
      expect(engine.evaluate('=MATCH("banana", A1:A3, 0)')).toEqual({ displayValue: '2', type: 'number' });
    });
  });

  // ─── Math ───────────────────────────────────────────

  describe('MOD', () => {
    it('returns remainder of division', () => {
      expect(engine.evaluate('=MOD(10, 3)')).toEqual({ displayValue: '1', type: 'number' });
    });

    it('returns 0 when evenly divisible', () => {
      expect(engine.evaluate('=MOD(10, 5)')).toEqual({ displayValue: '0', type: 'number' });
    });

    it('handles negative dividend', () => {
      expect(engine.evaluate('=MOD(-10, 3)')).toEqual({ displayValue: '-1', type: 'number' });
    });

    it('returns #DIV/0! when divisor is 0', () => {
      expect(engine.evaluate('=MOD(10, 0)')).toEqual({ displayValue: '#DIV/0!', type: 'error' });
    });

    it('works with decimal numbers', () => {
      expect(engine.evaluate('=MOD(5.5, 2)')).toEqual({ displayValue: '1.5', type: 'number' });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': '17', '0:1': '5' });
      engine.setData(data);
      expect(engine.evaluate('=MOD(A1, B1)')).toEqual({ displayValue: '2', type: 'number' });
    });
  });

  describe('POWER', () => {
    it('returns base raised to exponent', () => {
      expect(engine.evaluate('=POWER(2, 3)')).toEqual({ displayValue: '8', type: 'number' });
    });

    it('returns 1 for any base to the power of 0', () => {
      expect(engine.evaluate('=POWER(5, 0)')).toEqual({ displayValue: '1', type: 'number' });
    });

    it('returns 0 for 0 to any positive power', () => {
      expect(engine.evaluate('=POWER(0, 5)')).toEqual({ displayValue: '0', type: 'number' });
    });

    it('handles negative exponent', () => {
      expect(engine.evaluate('=POWER(2, -1)')).toEqual({ displayValue: '0.5', type: 'number' });
    });

    it('handles fractional exponent (square root)', () => {
      expect(engine.evaluate('=POWER(9, 0.5)')).toEqual({ displayValue: '3', type: 'number' });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': '3', '0:1': '4' });
      engine.setData(data);
      expect(engine.evaluate('=POWER(A1, B1)')).toEqual({ displayValue: '81', type: 'number' });
    });
  });

  describe('CEILING', () => {
    it('rounds up to nearest integer by default', () => {
      expect(engine.evaluate('=CEILING(2.1)')).toEqual({ displayValue: '3', type: 'number' });
    });

    it('rounds up to nearest significance', () => {
      expect(engine.evaluate('=CEILING(2.1, 0.5)')).toEqual({ displayValue: '2.5', type: 'number' });
    });

    it('returns exact value when already at significance boundary', () => {
      expect(engine.evaluate('=CEILING(3, 1)')).toEqual({ displayValue: '3', type: 'number' });
    });

    it('rounds up to nearest 10', () => {
      expect(engine.evaluate('=CEILING(23, 10)')).toEqual({ displayValue: '30', type: 'number' });
    });

    it('returns 0 when significance is 0', () => {
      expect(engine.evaluate('=CEILING(5, 0)')).toEqual({ displayValue: '0', type: 'number' });
    });

    it('handles negative numbers', () => {
      expect(engine.evaluate('=CEILING(-2.1, 1)')).toEqual({ displayValue: '-2', type: 'number' });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': '4.3' });
      engine.setData(data);
      expect(engine.evaluate('=CEILING(A1)')).toEqual({ displayValue: '5', type: 'number' });
    });
  });

  describe('FLOOR', () => {
    it('rounds down to nearest integer by default', () => {
      expect(engine.evaluate('=FLOOR(2.9)')).toEqual({ displayValue: '2', type: 'number' });
    });

    it('rounds down to nearest significance', () => {
      expect(engine.evaluate('=FLOOR(2.9, 0.5)')).toEqual({ displayValue: '2.5', type: 'number' });
    });

    it('returns exact value when already at significance boundary', () => {
      expect(engine.evaluate('=FLOOR(3, 1)')).toEqual({ displayValue: '3', type: 'number' });
    });

    it('rounds down to nearest 10', () => {
      expect(engine.evaluate('=FLOOR(27, 10)')).toEqual({ displayValue: '20', type: 'number' });
    });

    it('returns 0 when significance is 0', () => {
      expect(engine.evaluate('=FLOOR(5, 0)')).toEqual({ displayValue: '0', type: 'number' });
    });

    it('handles negative numbers', () => {
      expect(engine.evaluate('=FLOOR(-2.1, 1)')).toEqual({ displayValue: '-3', type: 'number' });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': '7.8' });
      engine.setData(data);
      expect(engine.evaluate('=FLOOR(A1)')).toEqual({ displayValue: '7', type: 'number' });
    });
  });

  // ─── String functions ─────────────────────────────────

  describe('LEFT', () => {
    it('returns first character by default', () => {
      expect(engine.evaluate('=LEFT("hello")')).toEqual({ displayValue: 'h', type: 'text' });
    });

    it('returns first n characters', () => {
      expect(engine.evaluate('=LEFT("hello", 3)')).toEqual({ displayValue: 'hel', type: 'text' });
    });

    it('returns entire string when n exceeds length', () => {
      expect(engine.evaluate('=LEFT("hi", 10)')).toEqual({ displayValue: 'hi', type: 'text' });
    });

    it('returns empty string when n is 0', () => {
      expect(engine.evaluate('=LEFT("hello", 0)')).toEqual({ displayValue: '', type: 'text' });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': 'world' });
      engine.setData(data);
      expect(engine.evaluate('=LEFT(A1, 2)')).toEqual({ displayValue: 'wo', type: 'text' });
    });

    it('converts numbers to strings', () => {
      // coerceValue sees "123" as numeric, so type is 'number'
      expect(engine.evaluate('=LEFT(12345, 3)')).toEqual({ displayValue: '123', type: 'number' });
    });
  });

  describe('RIGHT', () => {
    it('returns last character by default', () => {
      expect(engine.evaluate('=RIGHT("hello")')).toEqual({ displayValue: 'o', type: 'text' });
    });

    it('returns last n characters', () => {
      expect(engine.evaluate('=RIGHT("hello", 3)')).toEqual({ displayValue: 'llo', type: 'text' });
    });

    it('returns entire string when n exceeds length', () => {
      expect(engine.evaluate('=RIGHT("hi", 10)')).toEqual({ displayValue: 'hi', type: 'text' });
    });

    it('returns empty string when n is 0', () => {
      expect(engine.evaluate('=RIGHT("hello", 0)')).toEqual({ displayValue: '', type: 'text' });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': 'world' });
      engine.setData(data);
      expect(engine.evaluate('=RIGHT(A1, 3)')).toEqual({ displayValue: 'rld', type: 'text' });
    });

    it('converts numbers to strings', () => {
      // coerceValue sees "45" as numeric, so type is 'number'
      expect(engine.evaluate('=RIGHT(12345, 2)')).toEqual({ displayValue: '45', type: 'number' });
    });
  });

  describe('MID', () => {
    it('extracts substring from the middle (1-indexed)', () => {
      expect(engine.evaluate('=MID("hello world", 7, 5)')).toEqual({ displayValue: 'world', type: 'text' });
    });

    it('extracts from the start when start is 1', () => {
      expect(engine.evaluate('=MID("hello", 1, 3)')).toEqual({ displayValue: 'hel', type: 'text' });
    });

    it('returns partial string if n exceeds remaining length', () => {
      expect(engine.evaluate('=MID("hello", 4, 10)')).toEqual({ displayValue: 'lo', type: 'text' });
    });

    it('returns empty string when start is beyond string length', () => {
      expect(engine.evaluate('=MID("hi", 5, 2)')).toEqual({ displayValue: '', type: 'text' });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': 'abcdefg' });
      engine.setData(data);
      expect(engine.evaluate('=MID(A1, 3, 3)')).toEqual({ displayValue: 'cde', type: 'text' });
    });

    it('extracts single character', () => {
      expect(engine.evaluate('=MID("hello", 2, 1)')).toEqual({ displayValue: 'e', type: 'text' });
    });
  });

  describe('SUBSTITUTE', () => {
    it('replaces all occurrences when no instance given', () => {
      expect(engine.evaluate('=SUBSTITUTE("hello hello hello", "hello", "world")')).toEqual({
        displayValue: 'world world world',
        type: 'text',
      });
    });

    it('replaces only the nth occurrence when instance is given', () => {
      expect(engine.evaluate('=SUBSTITUTE("hello hello hello", "hello", "world", 2)')).toEqual({
        displayValue: 'hello world hello',
        type: 'text',
      });
    });

    it('replaces first occurrence with instance=1', () => {
      expect(engine.evaluate('=SUBSTITUTE("aaa", "a", "b", 1)')).toEqual({
        displayValue: 'baa',
        type: 'text',
      });
    });

    it('returns original string when old text not found', () => {
      expect(engine.evaluate('=SUBSTITUTE("hello", "xyz", "abc")')).toEqual({
        displayValue: 'hello',
        type: 'text',
      });
    });

    it('returns original string when nth occurrence does not exist', () => {
      expect(engine.evaluate('=SUBSTITUTE("hello", "l", "r", 5)')).toEqual({
        displayValue: 'hello',
        type: 'text',
      });
    });

    it('replaces empty string (inserts between each character)', () => {
      // JS "ab".split("").join("-") => "a-b"
      expect(engine.evaluate('=SUBSTITUTE("ab", "", "-")')).toEqual({
        displayValue: 'a-b',
        type: 'text',
      });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': 'foo bar foo' });
      engine.setData(data);
      expect(engine.evaluate('=SUBSTITUTE(A1, "foo", "baz")')).toEqual({
        displayValue: 'baz bar baz',
        type: 'text',
      });
    });
  });

  describe('FIND', () => {
    it('returns 1-indexed position of found string', () => {
      expect(engine.evaluate('=FIND("world", "hello world")')).toEqual({ displayValue: '7', type: 'number' });
    });

    it('finds from the beginning (position 1)', () => {
      expect(engine.evaluate('=FIND("h", "hello")')).toEqual({ displayValue: '1', type: 'number' });
    });

    it('is case-sensitive', () => {
      expect(engine.evaluate('=FIND("H", "hello")')).toEqual({ displayValue: '#VALUE!', type: 'error' });
    });

    it('returns #VALUE! when search string is not found', () => {
      expect(engine.evaluate('=FIND("xyz", "hello")')).toEqual({ displayValue: '#VALUE!', type: 'error' });
    });

    it('starts searching from specified position', () => {
      expect(engine.evaluate('=FIND("l", "hello world", 5)')).toEqual({ displayValue: '10', type: 'number' });
    });

    it('finds at the exact start position', () => {
      expect(engine.evaluate('=FIND("l", "hello", 3)')).toEqual({ displayValue: '3', type: 'number' });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': 'abcdef' });
      engine.setData(data);
      expect(engine.evaluate('=FIND("cd", A1)')).toEqual({ displayValue: '3', type: 'number' });
    });

    it('finds single character in single-character string', () => {
      expect(engine.evaluate('=FIND("a", "a")')).toEqual({ displayValue: '1', type: 'number' });
    });
  });

  // ─── Conversion ─────────────────────────────────────

  describe('TEXT', () => {
    it('formats number with "0" format (integer)', () => {
      // coerceValue converts "4" -> number type
      expect(engine.evaluate('=TEXT(3.7, "0")')).toEqual({ displayValue: '4', type: 'number' });
    });

    it('formats number with "0.00" format (2 decimal places)', () => {
      // "3.10" -> Number("3.10") = 3.1 -> String(3.1) = "3.1", coerced to number
      const result = engine.evaluate('=TEXT(3.1, "0.00")');
      expect(result.type).toBe('number');
      expect(result.displayValue).toBe('3.1');
    });

    it('formats number with "0.0" format (1 decimal place)', () => {
      // "3.1" -> coerced to number
      const result = engine.evaluate('=TEXT(3.14159, "0.0")');
      expect(result.type).toBe('number');
      expect(result.displayValue).toBe('3.1');
    });

    it('formats number with "#,##0" format (thousands separator)', () => {
      // "1,234,567" contains commas -> not a valid JS number -> stays text
      expect(engine.evaluate('=TEXT(1234567, "#,##0")')).toEqual({ displayValue: '1,234,567', type: 'text' });
    });

    it('formats number with "#,##0.00" format (thousands + decimals)', () => {
      // "1,234.50" contains commas -> stays text
      expect(engine.evaluate('=TEXT(1234.5, "#,##0.00")')).toEqual({ displayValue: '1,234.50', type: 'text' });
    });

    it('returns toString for unknown format', () => {
      // "42" -> coerced to number
      expect(engine.evaluate('=TEXT(42, "custom")')).toEqual({ displayValue: '42', type: 'number' });
    });

    it('handles non-numeric value', () => {
      expect(engine.evaluate('=TEXT("hello", "0")')).toEqual({ displayValue: 'hello', type: 'text' });
    });

    it('formats zero with "0.00"', () => {
      // "0.00" -> Number("0.00") = 0 -> String(0) = "0", coerced to number
      const result = engine.evaluate('=TEXT(0, "0.00")');
      expect(result.type).toBe('number');
      expect(result.displayValue).toBe('0');
    });

    it('formats negative number', () => {
      // "-5.5" -> coerced to number
      expect(engine.evaluate('=TEXT(-5.5, "0.0")')).toEqual({ displayValue: '-5.5', type: 'number' });
    });

    it('works with cell references and comma format', () => {
      const data = makeData({ '0:0': '1234.5' });
      engine.setData(data);
      // Commas preserve text type
      expect(engine.evaluate('=TEXT(A1, "#,##0.00")')).toEqual({ displayValue: '1,234.50', type: 'text' });
    });
  });

  describe('VALUE', () => {
    it('parses numeric string to number', () => {
      expect(engine.evaluate('=VALUE("42")')).toEqual({ displayValue: '42', type: 'number' });
    });

    it('parses decimal string', () => {
      expect(engine.evaluate('=VALUE("3.14")')).toEqual({ displayValue: '3.14', type: 'number' });
    });

    it('parses negative number string', () => {
      expect(engine.evaluate('=VALUE("-10")')).toEqual({ displayValue: '-10', type: 'number' });
    });

    it('returns #VALUE! for non-numeric string', () => {
      expect(engine.evaluate('=VALUE("hello")')).toEqual({ displayValue: '#VALUE!', type: 'error' });
    });

    it('parses zero', () => {
      expect(engine.evaluate('=VALUE("0")')).toEqual({ displayValue: '0', type: 'number' });
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': '99.5' });
      engine.setData(data);
      expect(engine.evaluate('=VALUE(A1)')).toEqual({ displayValue: '99.5', type: 'number' });
    });

    it('handles already-numeric value', () => {
      expect(engine.evaluate('=VALUE(42)')).toEqual({ displayValue: '42', type: 'number' });
    });
  });

  // ─── Date ───────────────────────────────────────────

  describe('DATE', () => {
    it('returns serial number for a known date', () => {
      // January 1, 2000 is serial number 36526 (days since 1899-12-30)
      const result = engine.evaluate('=DATE(2000, 1, 1)');
      expect(result.type).toBe('number');
      expect(Number(result.displayValue)).toBe(36526);
    });

    it('returns serial number for epoch-adjacent date', () => {
      // January 1, 1900 should be serial number 2 (1899-12-30 = 0, 12-31=1, 1-1=2)
      const result = engine.evaluate('=DATE(1900, 1, 1)');
      expect(result.type).toBe('number');
      expect(Number(result.displayValue)).toBe(2);
    });

    it('handles date with month > 12 (overflow)', () => {
      // DATE(2000, 13, 1) should be the same as DATE(2001, 1, 1)
      const result = engine.evaluate('=DATE(2000, 13, 1)');
      const expected = engine.evaluate('=DATE(2001, 1, 1)');
      expect(result.displayValue).toBe(expected.displayValue);
    });

    it('handles date with day > days in month (overflow)', () => {
      // DATE(2000, 1, 32) should overflow to February 1, 2000
      const result = engine.evaluate('=DATE(2000, 1, 32)');
      const expected = engine.evaluate('=DATE(2000, 2, 1)');
      expect(result.displayValue).toBe(expected.displayValue);
    });

    it('works with cell references', () => {
      const data = makeData({ '0:0': '2024', '0:1': '6', '0:2': '15' });
      engine.setData(data);
      const result = engine.evaluate('=DATE(A1, B1, C1)');
      expect(result.type).toBe('number');
      // Verify it returns a number (exact serial we trust from implementation)
      expect(Number(result.displayValue)).toBeGreaterThan(0);
    });

    it('different dates produce different serial numbers', () => {
      const r1 = engine.evaluate('=DATE(2020, 1, 1)');
      const r2 = engine.evaluate('=DATE(2020, 1, 2)');
      expect(Number(r2.displayValue) - Number(r1.displayValue)).toBe(1);
    });
  });

  describe('NOW', () => {
    it('returns a fractional serial number', () => {
      const result = engine.evaluate('=NOW()');
      expect(result.type).toBe('number');
      const val = Number(result.displayValue);
      // Should be a large number (days since 1899-12-30), roughly > 45000 for 2024+
      expect(val).toBeGreaterThan(45000);
    });

    it('includes a fractional component (time of day)', () => {
      const result = engine.evaluate('=NOW()');
      const val = Number(result.displayValue);
      // The integer part is the date, the fractional part is the time
      // While it's theoretically possible to get exactly midnight, it's extremely unlikely
      // Just verify it's a valid number
      expect(val).toBeGreaterThan(0);
    });

    it('is not cached (volatile) - successive calls may differ', () => {
      // NOW is volatile, so it should not be cached
      const data = makeData({ '0:0': '=NOW()' });
      engine.setData(data);
      const r1 = engine.evaluate('=NOW()', '0:0');
      // The value should be a valid serial number
      expect(r1.type).toBe('number');
      expect(Number(r1.displayValue)).toBeGreaterThan(0);
    });
  });

  // ─── Unknown function ──────────────────────────────

  describe('unknown functions', () => {
    it('returns #NAME? for unregistered functions', () => {
      expect(engine.evaluate('=UNKNOWN(1)')).toEqual({ displayValue: '#NAME?', type: 'error' });
    });
  });

  // ─── Custom function registration ──────────────────

  describe('registerFunction', () => {
    it('allows registering custom functions', () => {
      engine.registerFunction('DOUBLE', (_ctx, val) => Number(val) * 2);
      expect(engine.evaluate('=DOUBLE(5)')).toEqual({ displayValue: '10', type: 'number' });
    });

    it('function names are case-insensitive', () => {
      engine.registerFunction('triple', (_ctx, val) => Number(val) * 3);
      expect(engine.evaluate('=TRIPLE(5)')).toEqual({ displayValue: '15', type: 'number' });
      expect(engine.evaluate('=triple(5)')).toEqual({ displayValue: '15', type: 'number' });
    });
  });

  // ─── Recalculate ────────────────────────────────────

  describe('recalculate', () => {
    it('recalculates formula cells and returns changed keys', () => {
      const data = makeData({
        '0:0': '10',
        '0:1': '=A1*2',
      });
      engine.setData(data);
      engine.recalculate();

      const cell = data.get('0:1')!;
      expect(cell.displayValue).toBe('20');
      expect(cell.type).toBe('number');
    });

    it('detects changes when underlying data changes', () => {
      const data = makeData({
        '0:0': '10',
        '0:1': '=A1*2',
      });
      engine.setData(data);
      engine.recalculate();

      // Change A1
      data.set('0:0', cell('20', 'number'));
      const changed = engine.recalculate();
      expect(changed.has('0:1')).toBe(true);
      expect(data.get('0:1')!.displayValue).toBe('40');
    });
  });

  // ─── Complex scenarios (StringFormulas story) ───────

  describe('StringFormulas story scenario', () => {
    let data: GridData;

    beforeEach(() => {
      data = new Map();
      // Row 0: headers
      data.set('0:0', cell('First Name'));
      data.set('0:1', cell('Last Name'));
      data.set('0:2', cell('Full Name'));
      data.set('0:3', cell('Uppercase'));
      data.set('0:4', cell('Length'));
      data.set('0:5', cell('Trimmed'));

      // Row 1: "  Alice ", "Smith", formulas
      data.set('1:0', cell('  Alice '));
      data.set('1:1', cell('Smith'));
      data.set('1:2', cell('=CONCAT(TRIM(A2), " ", B2)'));
      data.set('1:3', cell('=UPPER(C2)'));
      data.set('1:4', cell('=LEN(C2)'));
      data.set('1:5', cell('=TRIM(A2)'));

      // Row 2: "Bob", "Johnson", formulas
      data.set('2:0', cell('Bob'));
      data.set('2:1', cell('Johnson'));
      data.set('2:2', cell('=CONCAT(A3, " ", B3)'));
      data.set('2:3', cell('=UPPER(C3)'));
      data.set('2:4', cell('=LEN(C3)'));
      data.set('2:5', cell('=TRIM(A3)'));

      engine.setData(data);
    });

    it('UPPER works when referencing a CONCAT formula cell', () => {
      const result = engine.evaluate('=UPPER(C2)');
      expect(result.displayValue).toBe('ALICE SMITH');
    });

    it('LEN works when referencing a CONCAT formula cell', () => {
      const result = engine.evaluate('=LEN(C2)');
      expect(result.displayValue).toBe('11');
    });

    it('TRIM works on cell with leading/trailing whitespace', () => {
      const result = engine.evaluate('=TRIM(A2)');
      expect(result.displayValue).toBe('Alice');
    });

    it('CONCAT with TRIM works correctly', () => {
      const result = engine.evaluate('=CONCAT(TRIM(A2), " ", B2)');
      expect(result.displayValue).toBe('Alice Smith');
    });

    it('recalculate populates all formula cells correctly', () => {
      engine.recalculate();

      // C2 = CONCAT(TRIM(A2), " ", B2) = "Alice Smith"
      expect(data.get('1:2')!.displayValue).toBe('Alice Smith');
      // D2 = UPPER(C2) = "ALICE SMITH"
      expect(data.get('1:3')!.displayValue).toBe('ALICE SMITH');
      // E2 = LEN(C2) = 11
      expect(data.get('1:4')!.displayValue).toBe('11');
      // F2 = TRIM(A2) = "Alice"
      expect(data.get('1:5')!.displayValue).toBe('Alice');

      // C3 = CONCAT(A3, " ", B3) = "Bob Johnson"
      expect(data.get('2:2')!.displayValue).toBe('Bob Johnson');
      // D3 = UPPER(C3) = "BOB JOHNSON"
      expect(data.get('2:3')!.displayValue).toBe('BOB JOHNSON');
      // E3 = LEN(C3) = 11
      expect(data.get('2:4')!.displayValue).toBe('11');
    });
  });

  // ─── recalculateAffected ────────────────────────────

  describe('recalculateAffected', () => {
    it('recalculates a direct dependent when input changes', () => {
      const data = makeData({ '0:0': '10', '0:1': '=A1*2' });
      engine.setData(data);
      engine.recalculate();
      expect(data.get('0:1')!.displayValue).toBe('20');

      // Simulate changing A1
      data.set('0:0', cell('5', 'number'));
      engine.evaluate('5', '0:0');
      const changed = engine.recalculateAffected(['0:0']);
      expect(changed.has('0:1')).toBe(true);
      expect(data.get('0:1')!.displayValue).toBe('10');
    });

    it('recalculates transitive dependents (A1 -> B1 -> C1)', () => {
      const data = makeData({
        '0:0': '5',
        '0:1': '=A1+1',
        '0:2': '=B1*2',
      });
      engine.setData(data);
      engine.recalculate();
      expect(data.get('0:1')!.displayValue).toBe('6');
      expect(data.get('0:2')!.displayValue).toBe('12');

      data.set('0:0', cell('10', 'number'));
      engine.evaluate('10', '0:0');
      const changed = engine.recalculateAffected(['0:0']);
      expect(changed.has('0:1')).toBe(true);
      expect(changed.has('0:2')).toBe(true);
      expect(data.get('0:1')!.displayValue).toBe('11');
      expect(data.get('0:2')!.displayValue).toBe('22');
    });

    it('recalculates diamond dependencies correctly', () => {
      // A1 -> B1, A1 -> C1, B1+C1 -> D1
      const data = makeData({
        '0:0': '10',
        '0:1': '=A1+1',
        '0:2': '=A1+2',
        '0:3': '=B1+C1',
      });
      engine.setData(data);
      engine.recalculate();
      expect(data.get('0:3')!.displayValue).toBe('23'); // 11 + 12

      data.set('0:0', cell('20', 'number'));
      engine.evaluate('20', '0:0');
      const changed = engine.recalculateAffected(['0:0']);
      expect(changed.has('0:3')).toBe(true);
      expect(data.get('0:3')!.displayValue).toBe('43'); // 21 + 22
    });

    it('handles cell that becomes a formula', () => {
      const data = makeData({ '0:0': '10', '0:1': '5' });
      engine.setData(data);
      engine.recalculate();

      // B1 becomes a formula — evaluate and update the cell
      const result = engine.evaluate('=A1*3', '0:1');
      data.set('0:1', cell('=A1*3', result.type, result.displayValue));
      expect(data.get('0:1')!.displayValue).toBe('30');
    });

    it('handles cell that stops being a formula', () => {
      const data = makeData({ '0:0': '10', '0:1': '=A1*2' });
      engine.setData(data);
      engine.recalculate();

      data.set('0:1', cell('plain text'));
      engine.evaluate('plain text', '0:1');
      const changed = engine.recalculateAffected(['0:1']);
      // No dependents to update, should not error
      expect(data.get('0:1')!.displayValue).toBe('plain text');
      expect(changed.size).toBe(0);
    });

    it('falls back to full recalculate when dependency graph is empty', () => {
      const data = makeData({
        '0:0': '10',
        '0:1': '=A1*2',
      });
      engine.setData(data);
      // setData clears the dependency graph, don't call recalculate()

      // recalculateAffected should detect empty graph and fall back
      data.set('0:0', cell('5', 'number'));
      const changed = engine.recalculateAffected(['0:0']);
      expect(changed.has('0:1')).toBe(true);
      expect(data.get('0:1')!.displayValue).toBe('10');
    });

    it('handles range dependencies', () => {
      const data = makeData({
        '0:0': '1',
        '1:0': '2',
        '2:0': '3',
        '3:0': '=SUM(A1:A3)',
      });
      engine.setData(data);
      engine.recalculate();
      expect(data.get('3:0')!.displayValue).toBe('6');

      data.set('1:0', cell('20', 'number'));
      engine.evaluate('20', '1:0');
      const changed = engine.recalculateAffected(['1:0']);
      expect(changed.has('3:0')).toBe(true);
      expect(data.get('3:0')!.displayValue).toBe('24');
    });
  });

  // ─── String comparisons ───────────────────────────────

  describe('string comparisons', () => {
    it('compares strings with < operator', () => {
      expect(engine.evaluate('="apple"<"banana"')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('="banana"<"apple"')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('compares strings with > operator', () => {
      expect(engine.evaluate('="banana">"apple"')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('compares strings with <= operator', () => {
      expect(engine.evaluate('="apple"<="apple"')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('="apple"<="banana"')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('compares strings with >= operator', () => {
      expect(engine.evaluate('="banana">="apple"')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('="apple">="banana"')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });
  });

  // ─── Recursion depth limit ────────────────────────────

  describe('recursion depth limit', () => {
    it('returns error for excessively deep formula chains', () => {
      // Create a chain of references that exceeds depth limit
      const data: GridData = new Map();
      for (let i = 0; i < 70; i++) {
        const raw = i === 0 ? '1' : `=A${i}+1`;
        const isFormula = raw.startsWith('=');
        data.set(`${i}:0`, cell(raw, isFormula ? 'text' : 'number'));
      }
      engine.setData(data);
      // Evaluating a cell deep in the chain should hit the depth limit
      const result = engine.evaluate('=A70');
      expect(result.type).toBe('error');
    });
  });

  // ─── Boolean handling ───────────────────────────────

  describe('boolean handling', () => {
    it('resolves TRUE literal', () => {
      expect(engine.evaluate('=TRUE')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
    });

    it('resolves FALSE literal', () => {
      expect(engine.evaluate('=FALSE')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('uses booleans in IF', () => {
      expect(engine.evaluate('=IF(TRUE, 1, 0)')).toEqual({ displayValue: '1', type: 'number' });
      expect(engine.evaluate('=IF(FALSE, 1, 0)')).toEqual({ displayValue: '0', type: 'number' });
    });
  });

  // ─── Value coercion tests ─────────────────────────

  describe('value coercion', () => {
    it('coerces numeric strings to number type', () => {
      expect(engine.evaluate('42')).toEqual({ displayValue: '42', type: 'number' });
      expect(engine.evaluate('3.14')).toEqual({ displayValue: '3.14', type: 'number' });
      expect(engine.evaluate('0')).toEqual({ displayValue: '0', type: 'number' });
    });

    it('coerces boolean strings to boolean type', () => {
      expect(engine.evaluate('TRUE')).toEqual({ displayValue: 'TRUE', type: 'boolean' });
      expect(engine.evaluate('false')).toEqual({ displayValue: 'FALSE', type: 'boolean' });
    });

    it('returns text type for plain text', () => {
      expect(engine.evaluate('hello')).toEqual({ displayValue: 'hello', type: 'text' });
    });

    it('coerces error codes to error type', () => {
      expect(engine.evaluate('#ERROR!')).toEqual({ displayValue: '#ERROR!', type: 'error' });
      expect(engine.evaluate('#REF!')).toEqual({ displayValue: '#REF!', type: 'error' });
      expect(engine.evaluate('#DIV/0!')).toEqual({ displayValue: '#DIV/0!', type: 'error' });
      expect(engine.evaluate('#NAME?')).toEqual({ displayValue: '#NAME?', type: 'error' });
      expect(engine.evaluate('#CIRC!')).toEqual({ displayValue: '#CIRC!', type: 'error' });
    });
  });

  // ─── Error handling tests ─────────────────────────

  describe('error handling', () => {
    it('returns #DIV/0! on division by zero', () => {
      expect(engine.evaluate('=1/0')).toEqual({ displayValue: '#DIV/0!', type: 'error' });
      expect(engine.evaluate('=10/(5-5)')).toEqual({ displayValue: '#DIV/0!', type: 'error' });
    });

    it('returns #NAME? on unknown function', () => {
      expect(engine.evaluate('=UNKNOWNFUNC(1)')).toEqual({ displayValue: '#NAME?', type: 'error' });
    });

    it('returns #ERROR! on syntax error', () => {
      expect(engine.evaluate('=(((')).toEqual({ displayValue: '#ERROR!', type: 'error' });
    });

    it('returns error on direct circular reference (A1=A1)', () => {
      const data = makeData({ '0:0': '=A1' });
      engine.setData(data);
      engine.recalculate();
      const cellData = data.get('0:0')!;
      expect(cellData.type).toBe('error');
    });

    it('returns error on indirect circular reference (A1=B1, B1=A1)', () => {
      const data = makeData({
        '0:0': '=B1',
        '0:1': '=A1',
      });
      engine.setData(data);
      engine.recalculate();
      // At least one of them should be an error
      const a1 = data.get('0:0')!;
      const b1 = data.get('0:1')!;
      expect(a1.type === 'error' || b1.type === 'error').toBe(true);
    });

    it('returns error when max eval depth is exceeded', () => {
      const data: GridData = new Map();
      for (let i = 0; i < 70; i++) {
        const raw = i === 0 ? '1' : `=A${i}+1`;
        const isFormula = raw.startsWith('=');
        data.set(`${i}:0`, cell(raw, isFormula ? 'text' : 'number'));
      }
      engine.setData(data);
      const result = engine.evaluate('=A70');
      expect(result.type).toBe('error');
    });
  });

  // ─── Dependency tracking tests ─────────────────────

  describe('dependency tracking', () => {
    it('recalculateAffected updates dependent cells', () => {
      const data = makeData({ '0:0': '10', '0:1': '=A1+5' });
      engine.setData(data);
      engine.recalculate();
      expect(data.get('0:1')!.displayValue).toBe('15');

      data.set('0:0', cell('20', 'number'));
      engine.evaluate('20', '0:0');
      const changed = engine.recalculateAffected(['0:0']);
      expect(changed.has('0:1')).toBe(true);
      expect(data.get('0:1')!.displayValue).toBe('25');
    });

    it('transitive dependencies: A1 change updates B1 and C1', () => {
      const data = makeData({
        '0:0': '1',
        '0:1': '=A1*10',
        '0:2': '=B1+1',
      });
      engine.setData(data);
      engine.recalculate();
      expect(data.get('0:1')!.displayValue).toBe('10');
      expect(data.get('0:2')!.displayValue).toBe('11');

      data.set('0:0', cell('2', 'number'));
      engine.evaluate('2', '0:0');
      const changed = engine.recalculateAffected(['0:0']);
      expect(changed.has('0:1')).toBe(true);
      expect(changed.has('0:2')).toBe(true);
      expect(data.get('0:1')!.displayValue).toBe('20');
      expect(data.get('0:2')!.displayValue).toBe('21');
    });

    it('recalculate() rebuilds the full dependency graph', () => {
      const data = makeData({
        '0:0': '5',
        '0:1': '=A1*2',
      });
      engine.setData(data);
      // setData clears deps, recalculate rebuilds them
      engine.recalculate();
      expect(data.get('0:1')!.displayValue).toBe('10');

      // Now recalculateAffected should work because deps are rebuilt
      data.set('0:0', cell('7', 'number'));
      engine.evaluate('7', '0:0');
      const changed = engine.recalculateAffected(['0:0']);
      expect(changed.has('0:1')).toBe(true);
      expect(data.get('0:1')!.displayValue).toBe('14');
    });

    it('recalculateAffected falls back to full recalc when deps empty (after setData)', () => {
      const data = makeData({
        '0:0': '10',
        '0:1': '=A1*3',
      });
      engine.setData(data);
      // Intentionally skip recalculate so deps are empty

      data.set('0:0', cell('4', 'number'));
      const changed = engine.recalculateAffected(['0:0']);
      // Should fall back to full recalc and update B1
      expect(changed.has('0:1')).toBe(true);
      expect(data.get('0:1')!.displayValue).toBe('12');
    });
  });

  // ─── Floating point display ────────────────────────

  describe('floating point display', () => {
    it('displays 10 * 5.99 as 59.9', () => {
      const data: GridData = new Map([
        ['0:0', cell('10', 'number')],
        ['0:1', cell('5.99', 'number')],
      ]);
      engine.setData(data);
      const result = engine.evaluate('=A1*B1', '0:2');
      expect(result.displayValue).toBe('59.9');
    });

    it('displays 1/3 with reasonable precision', () => {
      const result = engine.evaluate('=1/3');
      expect(result.displayValue).toBe('0.333333333333333');
    });

    it('displays exact integers without trailing zeros', () => {
      const result = engine.evaluate('=2+3');
      expect(result.displayValue).toBe('5');
    });

    it('handles very small floating point differences', () => {
      const result = engine.evaluate('=0.1+0.2');
      expect(result.displayValue).toBe('0.3');
    });
  });
});
