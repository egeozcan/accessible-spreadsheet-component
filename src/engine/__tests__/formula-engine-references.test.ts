import { describe, it, expect, beforeEach } from 'vitest';
import { FormulaEngine, RangeValue, RefValue } from '../formula-engine.js';
import { type GridData, cellKey, refToCoord } from '../../types.js';

/** Key of a cell given as an A1 reference: key('C2') is "1:2". */
const key = (ref: string): string => {
  const { row, col } = refToCoord(ref);
  return cellKey(row, col);
};

/** Sheet data from raw values keyed by A1 reference. */
function makeData(entries: Record<string, string>): GridData {
  const data: GridData = new Map();
  for (const [ref, raw] of Object.entries(entries)) {
    const isNum = !isNaN(Number(raw)) && raw.trim() !== '';
    data.set(key(ref), { rawValue: raw, displayValue: raw, type: isNum ? 'number' : 'text' });
  }
  return data;
}

/** A value as plain data: arrays as rows of values, blanks as null. */
function plain(value: unknown): unknown {
  if (value instanceof RangeValue) {
    return Array.from({ length: value.rows }, (_, r) => value.getRow(r).map(plain));
  }
  return value === undefined ? null : value;
}

/** Set a cell's raw value in place (creating the cell if needed) and return its key. */
function write(data: GridData, ref: string, raw: string): string {
  const cell = data.get(key(ref));
  if (cell) cell.rawValue = raw;
  else data.set(key(ref), { rawValue: raw, displayValue: raw, type: 'text' });
  return key(ref);
}

const shown = (data: GridData, ref: string) => data.get(key(ref))?.displayValue;

describe('FormulaEngine – whole-column and whole-row references', () => {
  let engine: FormulaEngine;
  const value = (formula: string, sheet?: string) => plain(engine.evaluateValue(formula, sheet));

  beforeEach(() => {
    engine = new FormulaEngine();
    // A1:A3 = 1,2,3   B1:B3 = 10,20,30   C1 = 100
    engine.setData(makeData({ A1: '1', A2: '2', A3: '3', B1: '10', B2: '20', B3: '30', C1: '100' }));
  });

  it('aggregates whole columns', () => {
    expect(value('=SUM(A:A)')).toBe(6);
    expect(value('=SUM(A:B)')).toBe(66);
    expect(value('=SUM($A:$B)')).toBe(66);
    expect(value('=SUM(B:A)')).toBe(66);
    expect(value('=COUNT(A:A)')).toBe(3);
    expect(value('=MAX(A:B)')).toBe(30);
    expect(value('=SUM(Z:Z)')).toBe(0);
  });

  it('aggregates whole rows', () => {
    expect(value('=SUM(1:1)')).toBe(111);
    expect(value('=SUM(1:2)')).toBe(133);
    expect(value('=SUM($1:$2)')).toBe(133);
    expect(value('=COUNTA(1:1)')).toBe(3);
    expect(value('=SUM(9:9)')).toBe(0);
  });

  it('lookups and criteria functions accept whole columns', () => {
    expect(value('=COUNTIF(A:A,">1")')).toBe(2);
    expect(value('=SUMIF(A:A,">1",B:B)')).toBe(50);
    expect(value('=VLOOKUP(2,A:B,2,FALSE)')).toBe(20);
    expect(value('=XLOOKUP(3,A:A,B:B)')).toBe(30);
    expect(value('=MATCH(20,B:B,0)')).toBe(2);
    expect(value('=INDEX(B:B,2)')).toBe(20);
  });

  it('the values of a whole column stop where the sheet content ends', () => {
    expect(value('=A:A')).toEqual([[1], [2], [3]]);
    expect(value('=1:1')).toEqual([[1, 10, 100]]);
    expect(value('=A:A*2')).toEqual([[2], [4], [6]]);
  });

  it('ROWS and COLUMNS of whole references follow the sheet size', () => {
    expect(value('=ROWS(A:A)')).toBe(1048576);
    expect(value('=COLUMNS(1:1)')).toBe(16384);
    engine.setSheetSize(50, 10);
    expect(value('=ROWS(A:A)')).toBe(50);
    expect(value('=COLUMNS(A:A)')).toBe(1);
    expect(value('=COLUMNS(A:C)')).toBe(3);
    expect(value('=COLUMNS(1:1)')).toBe(10);
    expect(value('=ROWS(1:3)')).toBe(3);
  });

  it('each sheet has its own size', () => {
    engine.addSheet('Small', makeData({ A1: '7', A2: '8' }), 20, 4);
    engine.setSheetSize(50, 10);
    expect(value('=ROWS(Small!A:A)')).toBe(20);
    expect(value('=COLUMNS(Small!1:1)')).toBe(4);
    expect(value('=SUM(Small!A:A)')).toBe(15);
    expect(value('=ROWS(A:A)', 'Small')).toBe(20);
  });

  it('whole-column formulas see cells added later', () => {
    const data = makeData({ A1: '1', A2: '2', C1: '=SUM(A:A)', C2: '=COUNT(A:A)' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'C1')).toBe('3');
    expect(shown(data, 'C2')).toBe('2');

    const changed = engine.recalculateAffected([write(data, 'A1000', '100')]);
    expect(shown(data, 'C1')).toBe('103');
    expect(shown(data, 'C2')).toBe('3');
    expect(changed.has(key('C1'))).toBe(true);

    data.delete(key('A1000'));
    engine.recalculateAffected([key('A1000')]);
    expect(shown(data, 'C1')).toBe('3');
    expect(shown(data, 'C2')).toBe('2');
  });

  it('whole-row formulas see cells added far to the right', () => {
    const data = makeData({ A1: '1', B1: '2', A3: '=SUM(1:1)' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'A3')).toBe('3');
    engine.recalculateAffected([write(data, 'AX1', '7')]);
    expect(shown(data, 'A3')).toBe('10');
  });

  it('a formula inside the column it sums is circular', () => {
    const data = makeData({ A1: '1', A4: '=SUM(A:A)', C13: '=SUM(1:13)' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'A4')).toBe('#CIRC!');
    expect(shown(data, 'C13')).toBe('#CIRC!');
  });
});

describe('FormulaEngine – OFFSET', () => {
  let engine: FormulaEngine;
  const value = (formula: string) => plain(engine.evaluateValue(formula));

  beforeEach(() => {
    engine = new FormulaEngine();
    engine.setData(makeData({ A1: '1', A2: '2', A3: '3', B1: '10', B2: '20', B3: '30', D1: '2' }));
  });

  it('moves a reference by rows and columns', () => {
    expect(value('=OFFSET(A1,1,1)')).toBe(20);
    expect(value('=OFFSET(B3,-2,-1)')).toBe(1);
    expect(value('=OFFSET(A1,0,0)')).toBe(1);
    // Offsets are truncated to whole numbers
    expect(value('=OFFSET(A1,1.9,0)')).toBe(2);
  });

  it('keeps the size of the base range unless height and width are given', () => {
    expect(value('=OFFSET(A1:A2,1,1)')).toEqual([[20], [30]]);
    expect(value('=OFFSET(A1,1,0,2,1)')).toEqual([[2], [3]]);
    expect(value('=SUM(OFFSET(A1,0,0,3,2))')).toBe(66);
    expect(value('=SUM(OFFSET(A1,0,0,D1,2))')).toBe(33);
    // A negative height extends upwards
    expect(value('=OFFSET(A3,0,0,-2,1)')).toEqual([[2], [3]]);
  });

  it('returns a reference that reference functions can use', () => {
    expect(value('=ROW(OFFSET(A1,4,2))')).toBe(5);
    expect(value('=COLUMN(OFFSET(A1,4,2))')).toBe(3);
    expect(value('=ROWS(OFFSET(A1,0,0,3,2))')).toBe(3);
    expect(value('=SUM(A1:OFFSET(A1,2,1))')).toBe(66);
  });

  it('is #REF! when it leaves the sheet or has no size', () => {
    expect(value('=OFFSET(A1,-1,0)')).toBe('#REF!');
    expect(value('=OFFSET(A1,0,-1)')).toBe('#REF!');
    expect(value('=OFFSET(A1,2000000,0)')).toBe('#REF!');
    expect(value('=OFFSET(A1,0,20000)')).toBe('#REF!');
    expect(value('=OFFSET(A1,0,0,0,1)')).toBe('#REF!');
    expect(value('=OFFSET(A1,0,0,1,0)')).toBe('#REF!');
  });

  it('needs a reference and at least rows and cols', () => {
    expect(value('=OFFSET(5,1,1)')).toBe('#VALUE!');
    expect(value('=OFFSET(A1,1)')).toBe('#VALUE!');
  });

  it('recalculates when a cell it lands on changes', () => {
    const data = makeData({ A1: '1', A2: '2', C1: '=SUM(OFFSET(A1,0,0,2,1))', C2: '=OFFSET(A1,1,0)*10' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'C1')).toBe('3');
    expect(shown(data, 'C2')).toBe('20');
    engine.recalculateAffected([write(data, 'A2', '7')]);
    expect(shown(data, 'C1')).toBe('8');
    expect(shown(data, 'C2')).toBe('70');
  });
});

describe('FormulaEngine – INDIRECT', () => {
  let engine: FormulaEngine;
  const value = (formula: string) => plain(engine.evaluateValue(formula));

  beforeEach(() => {
    engine = new FormulaEngine();
    engine.setData(
      makeData({ A1: '1', A2: '2', A3: '3', B1: '10', B2: '20', B3: '30', D1: 'A2', D2: 'B1:B3', D3: 'Sheet2!A1', D4: 'hello world' })
    );
    engine.addSheet('Sheet2', makeData({ A1: '77' }));
  });

  it('resolves text to a cell', () => {
    expect(value('=INDIRECT("A2")')).toBe(2);
    expect(value('=INDIRECT("a2")')).toBe(2);
    expect(value('=INDIRECT("$B$2")')).toBe(20);
    expect(value('=INDIRECT(D1)')).toBe(2);
    expect(value('=INDIRECT("A"&3)')).toBe(3);
    expect(value('=INDIRECT(ADDRESS(2,2))')).toBe(20);
  });

  it('resolves text to a range', () => {
    expect(value('=INDIRECT("B1:B3")')).toEqual([[10], [20], [30]]);
    expect(value('=SUM(INDIRECT(D2))')).toBe(60);
    expect(value('=SUM(INDIRECT("A:A"))')).toBe(6);
    expect(value('=ROW(INDIRECT("C7"))')).toBe(7);
  });

  it('resolves sheet-qualified text', () => {
    expect(value('=INDIRECT("Sheet2!A1")')).toBe(77);
    expect(value('=INDIRECT(D3)+1')).toBe(78);
  });

  it('is #REF! for text that is not a reference', () => {
    expect(value('=INDIRECT(D4)')).toBe('#REF!');
    expect(value('=INDIRECT("")')).toBe('#REF!');
    expect(value('=INDIRECT("1+1")')).toBe('#REF!');
    expect(value('=INDIRECT("SUM(A1)")')).toBe('#REF!');
    expect(value('=INDIRECT("nope!")')).toBe('#REF!');
    expect(value('=INDIRECT("Sheet9!A1")')).toBe('#REF!');
    // A blank cell
    expect(value('=INDIRECT(Z99)')).toBe('#REF!');
  });

  it('recalculates when the cell it points to changes', () => {
    const data = makeData({ A1: '1', A2: '2', B1: 'A1', C1: '=INDIRECT(B1)', C2: '=INDIRECT("A2")*2' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'C1')).toBe('1');
    expect(shown(data, 'C2')).toBe('4');

    let changed = engine.recalculateAffected([write(data, 'A1', '5')]);
    expect(shown(data, 'C1')).toBe('5');
    expect([...changed]).toEqual([key('C1')]);

    changed = engine.recalculateAffected([write(data, 'A2', '7')]);
    expect(shown(data, 'C2')).toBe('14');
    expect([...changed]).toEqual([key('C2')]);
  });

  it('follows the text when it is changed to point elsewhere', () => {
    const data = makeData({ A1: '1', A2: '2', B1: 'A1', C1: '=INDIRECT(B1)' });
    engine.setData(data);
    engine.recalculate();
    engine.recalculateAffected([write(data, 'B1', 'A2')]);
    expect(shown(data, 'C1')).toBe('2');
    // Now A2 is the precedent
    engine.recalculateAffected([write(data, 'A2', '9')]);
    expect(shown(data, 'C1')).toBe('9');
  });
});

describe('FormulaEngine – INDEX and the range operator', () => {
  let engine: FormulaEngine;
  const value = (formula: string) => plain(engine.evaluateValue(formula));

  beforeEach(() => {
    engine = new FormulaEngine();
    engine.setData(makeData({ A1: '1', A2: '2', A3: '3', B1: '10', B2: '20', B3: '30' }));
  });

  it('picks one value', () => {
    expect(value('=INDEX(A1:B3,3,2)')).toBe(30);
    expect(value('=INDEX(A1:A3,2)')).toBe(2);
    expect(value('=INDEX({1,2,3},2)')).toBe(2);
    expect(value('=INDEX({1;2;3},3)')).toBe(3);
  });

  it('returns a whole row or column for index 0', () => {
    expect(value('=INDEX(A1:B3,2,0)')).toEqual([[2, 20]]);
    expect(value('=INDEX(A1:B3,0,2)')).toEqual([[10], [20], [30]]);
    expect(value('=SUM(INDEX(A1:B3,2,0))')).toBe(22);
    expect(value('=SUM(INDEX(A1:B3,0,2))')).toBe(60);
    expect(value('=ROWS(INDEX(A1:B3,0,2))')).toBe(3);
    expect(value('=COLUMNS(INDEX(A1:B3,2,0))')).toBe(2);
    expect(value('=INDEX(A1:B3,0,0)')).toEqual([
      [1, 10],
      [2, 20],
      [3, 30],
    ]);
    expect(value('=INDEX({1,2;3,4},0,2)')).toEqual([[2], [4]]);
  });

  it('is #REF! outside the array', () => {
    expect(value('=INDEX(A1:B3,4,1)')).toBe('#REF!');
    expect(value('=INDEX(A1:B3,1,3)')).toBe('#REF!');
    expect(value('=INDEX(A1:B3,-1,1)')).toBe('#REF!');
    expect(value('=INDEX(A1:B3)')).toBe('#VALUE!');
  });

  it('returns a reference, so it can be an end of a range', () => {
    expect(value('=SUM(A1:INDEX(A1:A3,2))')).toBe(3);
    expect(value('=SUM(A1:INDEX(B:B,2))')).toBe(33);
    expect(value('=SUM(INDEX(A1:A3,2):B3)')).toBe(55);
    expect(value('=INDEX(A1:A3,2):B3')).toEqual([
      [2, 20],
      [3, 30],
    ]);
    expect(value('=ROW(INDEX(A1:B3,3,2))')).toBe(3);
  });

  it('the range operator spans the corners of both operands', () => {
    expect(value('=SUM(A1:A2:B3)')).toBe(66);
    expect(value('=SUM(B3:A1)')).toBe(66);
    expect(value('=SUM((A1):(B2))')).toBe(33);
    // Both sides must be references
    expect(value('=A1:5')).toBe('#VALUE!');
  });

  it('a range ending in INDEX recalculates when a cell inside changes', () => {
    const data = makeData({ A1: '1', A2: '2', A3: '3', C1: '=SUM(A1:INDEX(A1:A3,2))' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'C1')).toBe('3');
    engine.recalculateAffected([write(data, 'A2', '10')]);
    expect(shown(data, 'C1')).toBe('11');
  });
});

describe('FormulaEngine – ROW, COLUMN, ROWS, COLUMNS, ADDRESS', () => {
  let engine: FormulaEngine;
  const value = (formula: string) => plain(engine.evaluateValue(formula));

  beforeEach(() => {
    engine = new FormulaEngine();
    engine.setData(makeData({ A1: '1' }));
    engine.addSheet('Sheet2');
  });

  it('ROW and COLUMN of a cell', () => {
    expect(value('=ROW(A5)')).toBe(5);
    expect(value('=ROW($C$9)')).toBe(9);
    expect(value('=COLUMN(C1)')).toBe(3);
    expect(value('=COLUMN(AA1)')).toBe(27);
    expect(value('=ROW(Sheet2!A7)')).toBe(7);
  });

  it('ROW and COLUMN of a range give every row or column number', () => {
    expect(value('=ROW(B2:B4)')).toEqual([[2], [3], [4]]);
    expect(value('=ROW(B2:D4)')).toEqual([[2], [3], [4]]);
    expect(value('=COLUMN(B1:D1)')).toEqual([[2, 3, 4]]);
    expect(value('=COLUMN(B2:D4)')).toEqual([[2, 3, 4]]);
    expect(value('=COLUMN(A:C)')).toEqual([[1, 2, 3]]);
    expect(value('=SUM(ROW(A1:A10))')).toBe(55);
  });

  it('ROW and COLUMN need a reference', () => {
    expect(value('=ROW(5)')).toBe('#VALUE!');
    expect(value('=COLUMN("A1")')).toBe('#VALUE!');
    expect(value('=ROW(A1,A2)')).toBe('#VALUE!');
  });

  it('ROWS and COLUMNS', () => {
    expect(value('=ROWS(A1:B3)')).toBe(3);
    expect(value('=COLUMNS(A1:B3)')).toBe(2);
    expect(value('=ROWS({1,2;3,4;5,6})')).toBe(3);
    expect(value('=COLUMNS({1,2;3,4;5,6})')).toBe(2);
    expect(value('=ROWS(5)')).toBe(1);
    expect(value('=ROWS(A1)')).toBe(1);
  });

  it('ADDRESS writes a reference as text', () => {
    expect(value('=ADDRESS(1,1)')).toBe('$A$1');
    expect(value('=ADDRESS(2,3,1)')).toBe('$C$2');
    expect(value('=ADDRESS(2,3,2)')).toBe('C$2');
    expect(value('=ADDRESS(2,3,3)')).toBe('$C2');
    expect(value('=ADDRESS(2,3,4)')).toBe('C2');
    expect(value('=ADDRESS(1,27)')).toBe('$AA$1');
  });

  it('ADDRESS adds a sheet name, quoted when needed', () => {
    expect(value('=ADDRESS(1,1,1,TRUE,"Sheet2")')).toBe('Sheet2!$A$1');
    expect(value('=ADDRESS(1,1,4,TRUE,"My Sheet")')).toBe("'My Sheet'!A1");
  });

  it('ADDRESS rejects positions off the sheet and unknown modes', () => {
    expect(value('=ADDRESS(0,1)')).toBe('#VALUE!');
    expect(value('=ADDRESS(1,0)')).toBe('#VALUE!');
    expect(value('=ADDRESS(1,1,5)')).toBe('#VALUE!');
    expect(value('=ADDRESS({1;2},1,4)')).toEqual([['A1'], ['A2']]);
  });

  it('references off the sheet are #REF!', () => {
    expect(value('=A0')).toBe('#REF!');
    expect(value('=A1048577')).toBe('#REF!');
    expect(value('=XFE1')).toBe('#REF!');
    // The last cell of the largest sheet is a valid, empty cell (a blank result is 0)
    expect(value('=XFD1048576')).toBe(0);
  });
});

describe('FormulaEngine – dependency tracking', () => {
  let engine: FormulaEngine;

  beforeEach(() => {
    engine = new FormulaEngine();
  });

  it('detects circular references through ranges', () => {
    const data = makeData({
      A1: '=SUM(B1:B3)', B2: '=A1',
      D1: '=D1',
      E1: '=F1', F1: '=E1',
      H7: '=SUM(H7:H9)',
      A10: '=OFFSET(A10,0,0)',
      B11: '=INDIRECT("B11")',
      J1: '=SUM(K1:K200)', K150: '=J1',
    });
    engine.setData(data);
    engine.recalculate();
    for (const ref of ['A1', 'B2', 'D1', 'E1', 'F1', 'H7', 'A10', 'B11', 'J1', 'K150']) {
      expect(data.get(key(ref)), ref).toMatchObject({ displayValue: '#CIRC!', type: 'error' });
    }
  });

  it('IFERROR around a circular cell handles the error', () => {
    const data = makeData({ D1: '=D1', G7: '=IFERROR(D1,"c")' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'G7')).toBe('c');
  });

  it('a cycle goes away when it is broken', () => {
    const data = makeData({ A1: '=SUM(B1:B3)', B2: '=A1', B1: '4' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'A1')).toBe('#CIRC!');
    engine.recalculateAffected([write(data, 'B2', '6')]);
    expect(shown(data, 'A1')).toBe('10');
  });

  it('recalculates dependents of a large range when a cell inside changes', () => {
    const entries: Record<string, string> = {};
    for (let i = 1; i <= 500; i++) entries[`A${i}`] = String(i);
    Object.assign(entries, { C1: '=SUM(A1:A500)', C2: '=C1*2', C3: '=SUM(A:A)', C4: '=SUM(A1:A10)' });
    const data = makeData(entries);
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'C1')).toBe('125250');
    expect(shown(data, 'C2')).toBe('250500');

    const changed = engine.recalculateAffected([write(data, 'A300', '1300')]);
    expect(shown(data, 'C1')).toBe('126250');
    expect(shown(data, 'C2')).toBe('252500');
    expect(shown(data, 'C3')).toBe('126250');
    // Only what reads A300 is touched
    expect([...changed].sort()).toEqual([key('C1'), key('C2'), key('C3')]);
    expect(shown(data, 'C4')).toBe('55');
  });

  it('a change outside a large range leaves its dependents alone', () => {
    const entries: Record<string, string> = {};
    for (let i = 1; i <= 200; i++) entries[`A${i}`] = '1';
    Object.assign(entries, { C1: '=SUM(A1:A200)', B1: '5', D1: '=B1+1' });
    const data = makeData(entries);
    engine.setData(data);
    engine.recalculate();
    const changed = engine.recalculateAffected([write(data, 'B1', '6')]);
    expect([...changed]).toEqual([key('D1')]);
    expect(shown(data, 'C1')).toBe('200');
  });

  it('a formula cell inside a referenced range triggers the range’s dependents', () => {
    const entries: Record<string, string> = {};
    for (let i = 1; i <= 100; i++) entries[`A${i}`] = '1';
    Object.assign(entries, { A50: '=E1*2', E1: '5', C1: '=SUM(A1:A100)', C2: '=SUM(A49:A51)', C3: '=C1+C2' });
    const data = makeData(entries);
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'C1')).toBe('109');
    expect(shown(data, 'C2')).toBe('12');

    const changed = engine.recalculateAffected([write(data, 'E1', '10')]);
    expect(shown(data, 'A50')).toBe('20');
    expect(shown(data, 'C1')).toBe('119');
    expect(shown(data, 'C2')).toBe('22');
    expect(shown(data, 'C3')).toBe('141');
    expect([...changed].sort()).toEqual([key('A50'), key('C1'), key('C2'), key('C3')].sort());
  });

  it('replacing a value inside a range with a formula is picked up', () => {
    const data = makeData({ A1: '1', A2: '2', A3: '3', C1: '=SUM(A1:A3)', E1: '40' });
    engine.setData(data);
    engine.recalculate();
    engine.recalculateAffected([write(data, 'A2', '=E1')]);
    expect(shown(data, 'C1')).toBe('44');
    engine.recalculateAffected([write(data, 'E1', '50')]);
    expect(shown(data, 'C1')).toBe('54');
  });

  it('resolves long reference chains without overflowing the stack', () => {
    const entries: Record<string, string> = { A1: '1' };
    for (let i = 2; i <= 5000; i++) entries[`A${i}`] = `=A${i - 1}+1`;
    const data = makeData(entries);
    engine.setData(data);

    // Standalone, from the far end
    expect(engine.evaluate('=A5000')).toEqual({ displayValue: '5000', type: 'number' });

    engine.recalculate();
    expect(shown(data, 'A5000')).toBe('5000');
    expect(shown(data, 'A2500')).toBe('2500');

    const changed = engine.recalculateAffected([write(data, 'A1', '10')]);
    expect(changed.size).toBe(4999);
    expect(shown(data, 'A5000')).toBe('5009');
  });

  it('reports a long cycle as circular rather than hanging', () => {
    const entries: Record<string, string> = { A1: '=A3000' };
    for (let i = 2; i <= 3000; i++) entries[`A${i}`] = `=A${i - 1}+1`;
    const data = makeData(entries);
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'A1')).toBe('#CIRC!');
    expect(shown(data, 'A3000')).toBe('#CIRC!');
    expect(shown(data, 'A1500')).toBe('#CIRC!');
  });

  it('evaluate() with a cell key records dependencies for that cell', () => {
    const data = makeData({ A1: '1', B1: '=A1+1', C1: '=B1*2' });
    engine.setData(data);
    expect(engine.evaluate('=A1+1', key('B1')).displayValue).toBe('2');
    expect(engine.evaluate('=B1*2', key('C1')).displayValue).toBe('4');

    const changed = engine.recalculateAffected([write(data, 'A1', '5')]);
    expect([...changed].sort()).toEqual([key('B1'), key('C1')]);
    expect(shown(data, 'B1')).toBe('6');
    expect(shown(data, 'C1')).toBe('12');
  });

  it('recalculateAffected after setData falls back to a full recalculation', () => {
    const data = makeData({ A1: '2', B1: '=A1*3', F6: '=1+1' });
    engine.setData(data);
    const changed = engine.recalculateAffected([key('A1')]);
    expect([...changed].sort()).toEqual([key('B1'), key('F6')]);
    expect(shown(data, 'B1')).toBe('6');
    expect(shown(data, 'F6')).toBe('2');
  });

  it('updateData swaps the data map but keeps the graph of unchanged cells', () => {
    const data = makeData({ A1: '1', B1: '=A1+1', C1: '=B1*2' });
    engine.setData(data);
    engine.recalculate();

    const next: GridData = new Map(data);
    next.set(key('A1'), { rawValue: '7', displayValue: '7', type: 'number' });
    next.set(key('B1'), { rawValue: '=A1+100', displayValue: '', type: 'text' });
    engine.updateData(next, [key('A1'), key('B1')]);
    expect(engine.getSheetData()).toBe(next);

    const changed = engine.recalculateAffected([key('A1'), key('B1')]);
    expect([...changed].sort()).toEqual([key('B1'), key('C1')]);
    expect(shown(next, 'B1')).toBe('107');
    expect(shown(next, 'C1')).toBe('214');
  });

  it('resultFor returns the cached result, or evaluates when the text differs', () => {
    const data = makeData({ A1: '2', B1: '=A1*3' });
    engine.setData(data);
    engine.recalculate();
    expect(engine.resultFor(key('B1'), '=A1*3')).toEqual({ displayValue: '6', type: 'number' });
    expect(engine.resultFor(key('B1'), '=A1*4')).toEqual({ displayValue: '8', type: 'number' });
    expect(engine.resultFor(key('J10'), '=ROW()')).toEqual({ displayValue: '10', type: 'number' });
    expect(engine.resultFor(key('J10'), '5')).toEqual({ displayValue: '5', type: 'number' });
  });
});

describe('FormulaEngine – custom function argument kinds', () => {
  let engine: FormulaEngine;
  let seen: unknown[];
  const value = (formula: string) => plain(engine.evaluateValue(formula));

  beforeEach(() => {
    engine = new FormulaEngine();
    seen = [];
    // A3 is blank, A4 is an error
    engine.setData(makeData({ A1: '1', A2: '2', A4: '=1/0', B1: '10', B2: '20' }));
  });

  it("a 'number' argument given a range is called once per element", () => {
    engine.registerFunction(
      'DBL',
      (_ctx, x) => {
        seen.push(x);
        return (x as number) * 2;
      },
      { args: ['number'] }
    );
    expect(value('=DBL(A1:A2)')).toEqual([[2], [4]]);
    expect(seen).toEqual([1, 2]);
    expect(value('=DBL({1,2;3,4})')).toEqual([
      [2, 4],
      [6, 8],
    ]);
    expect(value('=DBL(5)')).toBe(10);
    expect(value('=dbl(A2)')).toBe(4);
    expect(value('=SUM(DBL(A1:A2))')).toBe(6);
    // A blank cell arrives as 0
    expect(value('=DBL(A3)')).toBe(0);
  });

  it("a 'number' argument keeps an error in its element without calling the function", () => {
    engine.registerFunction(
      'DBL',
      (_ctx, x) => {
        seen.push(x);
        return (x as number) * 2;
      },
      { args: ['number'] }
    );
    expect(value('=DBL(A1:A4)')).toEqual([[2], [4], [0], ['#DIV/0!']]);
    expect(seen).toEqual([1, 2, 0]);
    expect(value('=DBL(A4)')).toBe('#DIV/0!');
  });

  it('the per-element result spills when recalculated', () => {
    engine.registerFunction('DBL', (_ctx, x) => (x as number) * 2, { args: ['number'] });
    const data = makeData({ A1: '1', A2: '2', C1: '=DBL(A1:A2)' });
    engine.setData(data);
    engine.recalculate();
    expect(data.get(key('C1'))).toMatchObject({ displayValue: '2', spill: { rows: 2, cols: 1 } });
    expect(data.get(key('C2'))).toMatchObject({ displayValue: '4', spilledFrom: key('C1') });

    engine.recalculateAffected([write(data, 'A2', '21')]);
    expect(shown(data, 'C2')).toBe('42');
  });

  it('without options a range arrives as a RangeValue', () => {
    engine.registerFunction('SHAPE', (_ctx, x) => {
      seen.push(x);
      return x instanceof RangeValue ? `${x.rows}x${x.cols}` : typeof x;
    });
    expect(value('=SHAPE(A1:B2)')).toBe('2x2');
    expect(seen[0]).toBeInstanceOf(RangeValue);
    expect((seen[0] as RangeValue).values).toEqual([1, 10, 2, 20]);
    expect(value('=SHAPE({1,2,3})')).toBe('1x3');
    expect(value('=SHAPE(A1)')).toBe('number');
    expect(value('=SHAPE("x")')).toBe('string');
    // The array is passed whole even when it contains an error
    expect(value('=SHAPE(A1:A4)')).toBe('4x1');
  });

  it("a 'reference' argument arrives as a RefValue", () => {
    engine.registerFunction(
      'WHERE',
      (_ctx, x) => {
        seen.push(x);
        return x instanceof RefValue ? `${x.sheet.name}:${x.top},${x.left}-${x.bottom},${x.right}` : 'not a reference';
      },
      { args: ['reference'] }
    );
    expect(value('=WHERE(B2:C5)')).toBe('Sheet1:1,1-4,2');
    expect(seen[0]).toBeInstanceOf(RefValue);
    expect(seen[0]).toMatchObject({ rows: 4, cols: 2, isCell: false });
    expect(value('=WHERE(A1)')).toBe('Sheet1:0,0-0,0');
    expect((seen[1] as RefValue).isCell).toBe(true);
    // A reference to an error cell is still just a reference
    expect(value('=WHERE(A4)')).toBe('Sheet1:3,0-3,0');
    expect(value('=WHERE(5)')).toBe('not a reference');
  });

  it("a 'value' argument receives error codes and undefined for blanks", () => {
    engine.registerFunction(
      'KIND',
      (_ctx, x) => {
        seen.push(x);
        return x === undefined ? 'blank' : typeof x === 'string' && x.startsWith('#') ? 'error' : typeof x;
      },
      { args: ['value'] }
    );
    expect(value('=KIND(A1:A4)')).toEqual([['number'], ['number'], ['blank'], ['error']]);
    expect(seen).toEqual([1, 2, undefined, '#DIV/0!']);
    expect(value('=KIND(A4)')).toBe('error');
    expect(value('=KIND(1/0)')).toBe('error');
    expect(value('=KIND(A3)')).toBe('blank');
    expect(value('=KIND("x")')).toBe('string');
  });

  it("a 'text' argument receives '' for blanks", () => {
    engine.registerFunction('WRAP', (_ctx, x) => `[${x}]`, { args: ['text'] });
    expect(value('=WRAP(A1:A3)')).toEqual([['[1]'], ['[2]'], ['[]']]);
    expect(value('=WRAP(A3)')).toBe('[]');
  });

  it('kinds apply per argument and the last one repeats', () => {
    engine.registerFunction('MIX', (_ctx, n, arr) => (n as number) + (arr as RangeValue).values.length, {
      args: ['number', 'array'],
    });
    expect(value('=MIX(A1:A2,B1:B2)')).toEqual([[3], [4]]);

    engine.registerFunction('JOIN', (_ctx, ...xs) => xs.join('+'), { args: ['number'] });
    expect(value('=JOIN(1,2,3)')).toBe('1+2+3');
    expect(value('=JOIN(A1:A2,10)')).toEqual([['1+10'], ['2+10']]);
  });

  it('a volatile function is recomputed by every recalculateAffected', () => {
    let ticks = 0;
    let calls = 0;
    engine.registerFunction('TICK', () => ++ticks, { volatile: true });
    engine.registerFunction('ONCE', () => ++calls);
    const data = makeData({ A1: '1', B1: '=A1+1', D1: '=TICK()', D2: '=D1*10', D3: '=ONCE()' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'D1')).toBe('1');
    expect(shown(data, 'D2')).toBe('10');
    expect(shown(data, 'D3')).toBe('1');

    // An unrelated change still recomputes the volatile cell and its dependents
    let changed = engine.recalculateAffected([write(data, 'A1', '5')]);
    expect(shown(data, 'D1')).toBe('2');
    expect(shown(data, 'D2')).toBe('20');
    expect(shown(data, 'D3')).toBe('1');
    expect([...changed].sort()).toEqual([key('B1'), key('D1'), key('D2')]);

    changed = engine.recalculateAffected([]);
    expect(shown(data, 'D1')).toBe('3');
    expect(shown(data, 'D2')).toBe('30');
    expect(calls).toBe(1);
    expect([...changed].sort()).toEqual([key('D1'), key('D2')]);
  });

  it('custom functions are listed and removed by setFunctions', () => {
    engine.registerFunction('dbl', (_ctx, x) => (x as number) * 2, { args: ['number'] });
    expect(engine.functionNames()).toContain('DBL');
    expect(engine.functionNames()).toEqual(expect.arrayContaining(['SEQUENCE', 'LET', 'LAMBDA', 'FLATTEN', 'INDIRECT']));
    engine.setFunctions({});
    expect(value('=DBL(1)')).toBe('#NAME?');
    expect(value('=SUM(1,2)')).toBe(3);
  });
});
