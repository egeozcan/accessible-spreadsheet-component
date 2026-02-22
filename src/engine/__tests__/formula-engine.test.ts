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
      expect(result.type).toBe('error');
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

    it('resolves text cell references', () => {
      const data = makeData({ '0:0': 'hello' });
      engine.setData(data);
      expect(engine.evaluate('=A1 & " world"')).toEqual({
        displayValue: 'hello world',
        type: 'text',
      });
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

    it('detects circular references', () => {
      const data = makeData({ '0:0': '=A1' });
      engine.setData(data);
      expect(engine.evaluate('=A1')).toEqual({ displayValue: '#ERROR!', type: 'error' });
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
});
