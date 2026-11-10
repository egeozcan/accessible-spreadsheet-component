import { describe, it, expect, beforeEach } from 'vitest';
import { FormulaEngine, RangeValue, isValidName, renameSheetInFormula } from '../formula-engine.js';
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
const raw = (data: GridData, ref: string) => data.get(key(ref))?.rawValue;
const sorted = (keys: Iterable<string> | undefined) => [...(keys ?? [])].sort();

describe('FormulaEngine – LET', () => {
  let engine: FormulaEngine;
  const value = (formula: string) => plain(engine.evaluateValue(formula));

  beforeEach(() => {
    engine = new FormulaEngine();
    engine.setData(makeData({ A1: '1', A2: '2', A3: '3' }));
  });

  it('binds names for the final expression', () => {
    expect(value('=LET(x,2,x*3)')).toBe(6);
    expect(value('=LET(x,A1+A2,x*x)')).toBe(9);
    expect(value('=LET(greeting,"hi",greeting&"!")')).toBe('hi!');
  });

  it('later bindings can use earlier ones', () => {
    expect(value('=LET(x,2,y,x+1,x*y)')).toBe(6);
    expect(value('=LET(a,1,b,a+1,c,b+1,c)')).toBe(3);
    // A name can be bound again
    expect(value('=LET(x,1,x,x+1,x)')).toBe(2);
  });

  it('names are case-insensitive and may look almost like cells', () => {
    expect(value('=LET(Total,5,TOTAL+total)')).toBe(10);
    expect(value('=LET(a1x,2,a1x)')).toBe(2);
  });

  it('can bind ranges, which stay references', () => {
    expect(value('=LET(r,A1:A3,SUM(r))')).toBe(6);
    expect(value('=LET(r,A1:A3,ROWS(r))')).toBe(3);
    expect(value('=LET(r,A2:A3,ROW(r))')).toEqual([[2], [3]]);
    expect(value('=LET(r,A1:A3,r*2)')).toEqual([[2], [4], [6]]);
  });

  it('can be nested, with inner names shadowing outer ones', () => {
    expect(value('=LET(x,1,LET(y,x+1,x+y))')).toBe(3);
    expect(value('=LET(x,1,LET(x,10,x)+x)')).toBe(11);
  });

  it('rejects malformed calls', () => {
    expect(value('=LET(x,1)')).toBe('#VALUE!');
    expect(value('=LET(x)')).toBe('#VALUE!');
    expect(value('=LET(1,2,3)')).toBe('#VALUE!');
    expect(value('=LET(x,1,y,2)')).toBe('#VALUE!');
    expect(value('=LET(x,1,y)')).toBe('#NAME?');
  });

  it('a name bound to an error can be handled by IFERROR', () => {
    expect(value('=LET(x,1/0,IFERROR(x,5))')).toBe(5);
  });

  it('works as a cell formula and tracks its precedents', () => {
    const data = makeData({ A1: '1', A2: '2', C1: '=LET(s,SUM(A1:A2),s*10)' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'C1')).toBe('30');
    engine.recalculateAffected([write(data, 'A2', '5')]);
    expect(shown(data, 'C1')).toBe('60');
  });
});

describe('FormulaEngine – LAMBDA', () => {
  let engine: FormulaEngine;
  const value = (formula: string) => plain(engine.evaluateValue(formula));

  beforeEach(() => {
    engine = new FormulaEngine();
    engine.setData(makeData({ A1: '1', A2: '2', A3: '3' }));
  });

  it('can be called right away', () => {
    expect(value('=LAMBDA(x,x*2)(4)')).toBe(8);
    expect(value('=LAMBDA(a,b,a&b)("x","y")')).toBe('xy');
    expect(value('=(LAMBDA(x,x*x))(5)')).toBe(25);
    expect(value('=LAMBDA(5)()')).toBe(5);
    expect(value('=LAMBDA(x,x+A1)(A2)')).toBe(3);
  });

  it('bound by LET, it is called like a function', () => {
    expect(value('=LET(f,LAMBDA(x,x+1),f(2))')).toBe(3);
    expect(value('=LET(f,LAMBDA(a,b,a&b),f("x","y"))')).toBe('xy');
    expect(value('=LET(double,LAMBDA(x,x*2),double(double(3)))')).toBe(12);
    // A LET name wins over a built-in of the same name
    expect(value('=LET(SUM,LAMBDA(a,a*100),SUM(2))')).toBe(200);
  });

  it('closes over the names in scope where it was made', () => {
    expect(value('=LET(x,5,LAMBDA(y,x+y))(3)')).toBe(8);
    expect(value('=LAMBDA(x,LAMBDA(y,x+y))(1)(2)')).toBe(3);
    expect(value('=LET(n,10,f,LAMBDA(x,x+n),n,1000,f(1))')).toBe(11);
  });

  it('takes ranges as references', () => {
    expect(value('=LAMBDA(r,ROWS(r))(A1:A3)')).toBe(3);
    expect(value('=LAMBDA(r,ROW(r))(A2)')).toBe(2);
    expect(value('=SUM(LAMBDA(r,r)(A1:A3))')).toBe(6);
  });

  it('treats missing arguments as blank and rejects extra ones', () => {
    expect(value('=LAMBDA(x,y,ISBLANK(y))(5)')).toBe(true);
    expect(value('=LAMBDA(x,y,x-y)(5,)')).toBe(5);
    expect(value('=LAMBDA(x,x)(1,2)')).toBe('#VALUE!');
  });

  it('rejects malformed definitions and calling what is not a lambda', () => {
    expect(value('=LAMBDA()')).toBe('#VALUE!');
    expect(value('=LAMBDA(1,2)(3)')).toBe('#VALUE!');
    expect(value('=(1)(2)')).toBe('#VALUE!');
    expect(value('=LET(x,5,x(1))')).toBe('#VALUE!');
  });

  it('a lambda that is not called is #CALC!', () => {
    expect(value('=LAMBDA(x,x)')).toBe('#CALC!');
    expect(engine.evaluate('=LAMBDA(x,x)')).toEqual({ displayValue: '#CALC!', type: 'error' });
    expect(value('=LET(f,LAMBDA(x,x),f)')).toBe('#CALC!');
  });

  it('can be passed to MAP, REDUCE, BYROW and BYCOL by name', () => {
    expect(value('=LET(f,LAMBDA(x,x*2),MAP(A1:A3,f))')).toEqual([[2], [4], [6]]);
    expect(value('=LET(add,LAMBDA(a,b,a+b),REDUCE(0,A1:A3,add))')).toBe(6);
    expect(value('=LET(total,LAMBDA(r,SUM(r)),BYROW({1,2;3,4},total))')).toEqual([[3], [7]]);
  });
});

describe('FormulaEngine – defined names', () => {
  let engine: FormulaEngine;
  let sheet1: GridData;
  let sheet2: GridData;
  const value = (formula: string, sheet?: string) => plain(engine.evaluateValue(formula, sheet));

  beforeEach(() => {
    engine = new FormulaEngine();
    sheet1 = makeData({ A1: '1', A2: '2', A3: '3' });
    sheet2 = makeData({ A1: '100' });
    engine.setData(sheet1);
    engine.addSheet('Sheet2', sheet2);
  });

  it('names a constant', () => {
    engine.defineName('Rate', '0.5');
    engine.defineName('Label', '="total"');
    expect(value('=Rate*100')).toBe(50);
    expect(value('=Label&": "&Rate')).toBe('total: 0.5');
  });

  it('names a range', () => {
    engine.defineName('Nums', '=A1:A3');
    expect(value('=SUM(Nums)')).toBe(6);
    expect(value('=ROWS(Nums)')).toBe(3);
    expect(value('=Nums')).toEqual([[1], [2], [3]]);
    expect(value('=INDEX(Nums,2)')).toBe(2);
    expect(value('=ROW(Nums)')).toEqual([[1], [2], [3]]);
  });

  it('names a formula, which may use other names', () => {
    engine.defineName('Rate', '0.5');
    engine.defineName('Total', '=SUM(A1:A3)+Rate');
    expect(value('=Total+1')).toBe(7.5);
  });

  it('names a LAMBDA that is called like a function', () => {
    engine.defineName('Twice', '=LAMBDA(x,x*2)');
    expect(value('=Twice(4)')).toBe(8);
    expect(value('=twice(A2)')).toBe(4);
    expect(value('=Twice(Twice(1))')).toBe(4);
    expect(value('=MAP(A1:A3,Twice)')).toEqual([[2], [4], [6]]);
    // Calling a name that is not a lambda
    engine.defineName('Rate', '0.5');
    expect(value('=Rate(1)')).toBe('#VALUE!');
  });

  it('a LAMBDA name can call itself', () => {
    engine.defineName('Fact', '=LAMBDA(n,IF(n<=1,1,n*Fact(n-1)))');
    expect(value('=Fact(5)')).toBe(120);
    expect(value('=Fact(10)')).toBe(3628800);
  });

  it('runaway recursion ends as an error', () => {
    engine.defineName('Forever', '=LAMBDA(n,Forever(n+1))');
    engine.defineName('Loop', '=Loop+1');
    expect(value('=Forever(1)')).toBe('#NUM!');
    expect(value('=Loop')).toBe('#CIRC!');
  });

  it('names are case-insensitive', () => {
    engine.defineName('Rate', '0.5');
    expect(value('=rate*2')).toBe(1);
    expect(value('=RATE*2')).toBe(1);
    // Defining it again in another case replaces it
    engine.defineName('RATE', '2');
    expect(engine.getNames()).toEqual([{ name: 'RATE', formula: '=2', sheet: 'Sheet1' }]);
    expect(value('=Rate*2')).toBe(4);
  });

  it('an undefined name is #NAME?', () => {
    expect(value('=Missing+1')).toBe('#NAME?');
    expect(value('=Missing(1)')).toBe('#NAME?');
    engine.defineName('Rate', '0.5');
    expect(engine.removeName('RATE')).toBe(true);
    expect(engine.removeName('RATE')).toBe(false);
    expect(value('=Rate')).toBe('#NAME?');
  });

  it('LET and LAMBDA names shadow defined names', () => {
    engine.defineName('K', '5');
    expect(value('=K*2')).toBe(10);
    expect(value('=LET(K,1,K)')).toBe(1);
    expect(value('=LAMBDA(K,K+1)(1)')).toBe(2);
  });

  it('unqualified references belong to the sheet given at definition', () => {
    engine.defineName('First', '=A1');
    engine.defineName('FirstOnTwo', '=A1', 'Sheet2');
    engine.defineName('Explicit', '=Sheet2!A1*2');
    // Same result whichever sheet the formula using the name is on
    expect(value('=First')).toBe(1);
    expect(value('=First', 'Sheet2')).toBe(1);
    expect(value('=FirstOnTwo')).toBe(100);
    expect(value('=FirstOnTwo', 'Sheet2')).toBe(100);
    expect(value('=Explicit')).toBe(200);
    expect(engine.getNames()).toEqual([
      { name: 'First', formula: '=A1', sheet: 'Sheet1' },
      { name: 'FirstOnTwo', formula: '=A1', sheet: 'Sheet2' },
      { name: 'Explicit', formula: '=Sheet2!A1*2', sheet: 'Sheet1' },
    ]);
  });

  it('the default sheet of a name is the active sheet at definition', () => {
    engine.setActiveSheet('Sheet2');
    engine.defineName('Here', '=A1');
    engine.setActiveSheet('Sheet1');
    expect(value('=Here')).toBe(100);
  });

  it('setNames replaces every name', () => {
    engine.defineName('Old', '1');
    engine.setNames([
      { name: 'K', formula: '5' },
      { name: 'OnTwo', formula: '=A1', sheet: 'Sheet2' },
      // A sheet that doesn't exist is not replaced by another one: the name has none
      { name: 'Lost', formula: '=A1', sheet: 'Gone' },
    ]);
    expect(engine.getNames()).toEqual([
      { name: 'K', formula: '=5', sheet: 'Sheet1' },
      { name: 'OnTwo', formula: '=A1', sheet: 'Sheet2' },
      { name: 'Lost', formula: '=A1', sheet: '' },
    ]);
    expect(value('=Old')).toBe('#NAME?');
    expect(value('=K+OnTwo')).toBe(105);
    expect(value('=Lost')).toBe('#REF!');
  });

  it('cells using a name recalculate when the cells behind it change', () => {
    engine.defineName('Nums', '=A1:A3');
    engine.defineName('Rate', '0.5');
    write(sheet1, 'C1', '=SUM(Nums)*Rate');
    write(sheet1, 'C2', '=Nums');
    write(sheet2, 'B1', '=SUM(Nums)');
    engine.recalculate();
    expect(shown(sheet1, 'C1')).toBe('3');
    expect(sheet1.get(key('C2'))?.spill).toEqual({ rows: 3, cols: 1 });
    expect(shown(sheet2, 'B1')).toBe('6');

    const changed = engine.recalculateAffected([write(sheet1, 'A1', '11')]);
    expect(shown(sheet1, 'C1')).toBe('8');
    expect(shown(sheet1, 'C2')).toBe('11');
    expect(shown(sheet2, 'B1')).toBe('16');
    expect(sorted(changed)).toEqual([key('C1'), key('C2')]);
    expect(sorted(engine.changes.get('Sheet2'))).toEqual([key('B1')]);
  });

  it('redefining a name takes effect on the next recalculate()', () => {
    engine.defineName('Rate', '0.5');
    write(sheet1, 'C1', '=SUM(A1:A3)*Rate');
    engine.recalculate();
    expect(shown(sheet1, 'C1')).toBe('3');

    engine.defineName('Rate', '2');
    expect([...engine.recalculate()]).toEqual([key('C1')]);
    expect(shown(sheet1, 'C1')).toBe('12');

    engine.removeName('Rate');
    engine.recalculate();
    expect(sheet1.get(key('C1'))).toMatchObject({ displayValue: '#NAME?', type: 'error' });
  });

  it('defineName validates the name', () => {
    expect(() => engine.defineName('A1', '=1')).toThrow('Invalid name: A1');
    expect(() => engine.defineName('my name', '=1')).toThrow('Invalid name');
    expect(() => engine.defineName('TRUE', '=1')).toThrow('Invalid name');
    expect(() => engine.defineName('Fine', '=1', 'Nowhere')).toThrow('Unknown sheet');
    expect(engine.getNames()).toEqual([]);
  });
});

describe('isValidName', () => {
  it('accepts identifiers', () => {
    for (const name of ['Rate', 'rate2x', '_x', 'x.y', 'a', '\\x', 'Tax_Rate', 'ABCD1', 'SUM']) {
      expect(isValidName(name), name).toBe(true);
    }
  });

  it('rejects cell references', () => {
    for (const name of ['A1', 'a1', 'ABC1', 'XFD1', 'ZZZ99', 'R1C1', 'RC', 'r2c3']) {
      expect(isValidName(name), name).toBe(false);
    }
  });

  it('rejects booleans', () => {
    for (const name of ['TRUE', 'FALSE', 'true', 'False']) {
      expect(isValidName(name), name).toBe(false);
    }
  });

  it('rejects names that are not identifiers', () => {
    for (const name of ['1abc', 'my name', 'x-y', '', ' x', 'a!b', 'a,b', '$x', '.x']) {
      expect(isValidName(name), JSON.stringify(name)).toBe(false);
    }
  });
});

describe('FormulaEngine – sheets', () => {
  let engine: FormulaEngine;
  let sheet1: GridData;
  let sheet2: GridData;
  let mySheet: GridData;
  const value = (formula: string, sheet?: string) => plain(engine.evaluateValue(formula, sheet));

  beforeEach(() => {
    engine = new FormulaEngine();
    sheet1 = makeData({
      A1: '=Sheet2!A1+1',
      A2: "=SUM('My Sheet'!A1:B2)",
      A3: '=SUM(Sheet2!A:A)',
      A4: "='My Sheet'!A1&Sheet2!B3",
      A5: '=A1*2',
      A6: '=Nope!A1',
    });
    sheet2 = makeData({ A1: '10', A2: '20', B1: '=Sheet1!A5', B3: '=A1+A2' });
    mySheet = makeData({ A1: '1', B1: '2', A2: '3', B2: '4' });
    engine.setData(sheet1);
    engine.addSheet('Sheet2', sheet2);
    engine.addSheet('My Sheet', mySheet);
  });

  it('starts with one sheet and lists added ones in order', () => {
    const fresh = new FormulaEngine();
    expect(fresh.sheetNames()).toEqual(['Sheet1']);
    expect(fresh.activeSheet).toBe('Sheet1');
    expect(engine.sheetNames()).toEqual(['Sheet1', 'Sheet2', 'My Sheet']);
    expect(engine.hasSheet('sheet2')).toBe(true);
    expect(engine.hasSheet('MY SHEET')).toBe(true);
    expect(engine.hasSheet('Nope')).toBe(false);
    expect(engine.getSheetData('Sheet2')).toBe(sheet2);
    expect(engine.getSheetData()).toBe(sheet1);
  });

  it('rejects duplicate and unknown sheet names', () => {
    expect(() => engine.addSheet('SHEET2')).toThrow('Sheet already exists');
    expect(() => engine.setActiveSheet('Nope')).toThrow('Unknown sheet');
    expect(() => engine.evaluate('=1', undefined, 'Nope')).toThrow('Unknown sheet');
    expect(() => engine.renameSheet('Sheet2', 'sheet1')).toThrow('Sheet already exists');
    expect(() => engine.removeSheet('Nope')).toThrow('Unknown sheet');
  });

  it('evaluates sheet-qualified references', () => {
    expect(value('=Sheet2!A1')).toBe(10);
    expect(value('=sheet2!a2')).toBe(20);
    expect(value("='My Sheet'!B2")).toBe(4);
    expect(value("=SUM('My Sheet'!A1:B2)")).toBe(10);
    expect(value('=SUM(Sheet2!A1:A2)*2')).toBe(60);
    expect(value('=Sheet2!A1:A2')).toEqual([[10], [20]]);
    expect(value('=Sheet1!A1')).toBe(11);
  });

  it('an unknown sheet is #REF!', () => {
    expect(value('=Nope!A1')).toBe('#REF!');
    expect(value('=SUM(Nope!A1:A3)')).toBe('#REF!');
    expect(value("='No such'!A:A")).toBe('#REF!');
  });

  it('unqualified references are on the sheet the formula is on', () => {
    expect(value('=A1', 'Sheet2')).toBe(10);
    expect(value('=A1', 'My Sheet')).toBe(1);
    expect(value('=A1+Sheet1!A1', 'Sheet2')).toBe(21);
    expect(engine.evaluate('=A1+A2', undefined, 'Sheet2')).toEqual({ displayValue: '30', type: 'number' });
    expect(engine.evaluate('=ROW()', key('D4'), 'Sheet2').displayValue).toBe('4');
    // A range cannot span two sheets
    expect(value('=SUM(Sheet2!A1:INDEX(A1:A3,2))')).toBe('#VALUE!');
  });

  it('the active sheet is the default for formulas and keys', () => {
    engine.setActiveSheet('Sheet2');
    expect(engine.activeSheet).toBe('Sheet2');
    expect(value('=A1')).toBe(10);
    expect(engine.getSheetData()).toBe(sheet2);
  });

  it('recalculate() computes every sheet and reports changes per sheet', () => {
    const changed = engine.recalculate();
    expect(shown(sheet1, 'A1')).toBe('11');
    expect(shown(sheet1, 'A2')).toBe('10');
    expect(shown(sheet1, 'A3')).toBe('30');
    expect(shown(sheet1, 'A4')).toBe('130');
    expect(shown(sheet1, 'A5')).toBe('22');
    expect(sheet1.get(key('A6'))).toMatchObject({ displayValue: '#REF!', type: 'error' });
    expect(shown(sheet2, 'B1')).toBe('22');
    expect(shown(sheet2, 'B3')).toBe('30');

    // The return value is the active sheet's part of `changes`
    expect(sorted(changed)).toEqual(['A1', 'A2', 'A3', 'A4', 'A5', 'A6'].map(key));
    expect(sorted(engine.changes.get('Sheet1'))).toEqual(sorted(changed));
    expect(sorted(engine.changes.get('Sheet2'))).toEqual([key('B1'), key('B3')]);
    expect(engine.changes.has('My Sheet')).toBe(false);
  });

  it('dependents on other sheets recalculate when a sheet changes', () => {
    engine.recalculate();
    const changed = engine.recalculateAffected([write(sheet2, 'A1', '100')], 'Sheet2');

    // Sheet1 reads Sheet2!A1 directly, through a whole column and through Sheet2!B3
    expect(shown(sheet1, 'A1')).toBe('101');
    expect(shown(sheet1, 'A3')).toBe('120');
    expect(shown(sheet1, 'A4')).toBe('1120');
    expect(shown(sheet1, 'A5')).toBe('202');
    // ...and Sheet2!B1 reads Sheet1!A5 back
    expect(shown(sheet2, 'B1')).toBe('202');
    expect(shown(sheet2, 'B3')).toBe('120');

    // The return value is for the sheet that was named
    expect(sorted(changed)).toEqual([key('B1'), key('B3')]);
    expect(sorted(engine.changes.get('Sheet2'))).toEqual([key('B1'), key('B3')]);
    expect(sorted(engine.changes.get('Sheet1'))).toEqual(['A1', 'A3', 'A4', 'A5'].map(key));
    // Untouched
    expect(shown(sheet1, 'A2')).toBe('10');
  });

  it('a change on a quoted-name sheet reaches its dependents', () => {
    engine.recalculate();
    const changed = engine.recalculateAffected([write(mySheet, 'B2', '40')], 'My Sheet');
    expect(shown(sheet1, 'A2')).toBe('46');
    expect(changed.size).toBe(0);
    expect(sorted(engine.changes.get('Sheet1'))).toEqual([key('A2')]);
    expect(engine.changes.has('My Sheet')).toBe(false);
  });

  it('a formula added on another sheet is picked up by recalculateAffected', () => {
    engine.recalculate();
    engine.recalculateAffected([write(mySheet, 'D1', '=Sheet2!A2*3')], 'My Sheet');
    expect(shown(mySheet, 'D1')).toBe('60');
    engine.recalculateAffected([write(sheet2, 'A2', '5')], 'Sheet2');
    expect(shown(mySheet, 'D1')).toBe('15');
    expect(sorted(engine.changes.get('My Sheet'))).toEqual([key('D1')]);
  });

  describe('renameSheet', () => {
    it('rewrites formulas in every sheet, quoting the new name when needed', () => {
      engine.recalculate();
      const changed = engine.renameSheet('Sheet2', 'Data 2');

      expect(raw(sheet1, 'A1')).toBe("='Data 2'!A1+1");
      expect(raw(sheet1, 'A3')).toBe("=SUM('Data 2'!A:A)");
      expect(raw(sheet1, 'A4')).toBe("='My Sheet'!A1&'Data 2'!B3");
      // References to other sheets and unqualified ones are untouched
      expect(raw(sheet1, 'A2')).toBe("=SUM('My Sheet'!A1:B2)");
      expect(raw(sheet2, 'B1')).toBe('=Sheet1!A5');
      expect(raw(sheet2, 'B3')).toBe('=A1+A2');

      expect(changed).toEqual(new Map([['Sheet1', [key('A1'), key('A3'), key('A4')]]]));
      expect(engine.sheetNames()).toEqual(['Sheet1', 'Data 2', 'My Sheet']);
      expect(engine.hasSheet('Sheet2')).toBe(false);

      engine.recalculate();
      expect(shown(sheet1, 'A1')).toBe('11');
      expect(shown(sheet1, 'A4')).toBe('130');
    });

    it('drops the quotes when the new name does not need them', () => {
      const changed = engine.renameSheet('My Sheet', 'Mine');
      expect(raw(sheet1, 'A2')).toBe('=SUM(Mine!A1:B2)');
      expect(raw(sheet1, 'A4')).toBe('=Mine!A1&Sheet2!B3');
      expect(changed).toEqual(new Map([['Sheet1', [key('A2'), key('A4')]]]));
    });

    it('escapes apostrophes in the new name', () => {
      engine.renameSheet('My Sheet', "O'Brien");
      expect(raw(sheet1, 'A2')).toBe("=SUM('O''Brien'!A1:B2)");
      engine.recalculate();
      expect(shown(sheet1, 'A2')).toBe('10');
    });

    it('reports changed cells under the new name of the renamed sheet', () => {
      write(sheet2, 'C1', '=Sheet2!A1*2');
      const changed = engine.renameSheet('Sheet2', 'Two');
      expect(raw(sheet2, 'C1')).toBe('=Two!A1*2');
      expect(changed.get('Two')).toEqual([key('C1')]);
      expect(changed.has('Sheet2')).toBe(false);
    });

    it('matches the old name whatever its case in the formula', () => {
      write(sheet1, 'B1', '=sheet2!A1+SHEET2!A2');
      engine.renameSheet('Sheet2', 'Two');
      expect(raw(sheet1, 'B1')).toBe('=Two!A1+Two!A2');
    });

    it('leaves text inside strings alone', () => {
      write(sheet1, 'B1', '="Sheet2!A1"&Sheet2!A1');
      engine.renameSheet('Sheet2', 'Two');
      expect(raw(sheet1, 'B1')).toBe('="Sheet2!A1"&Two!A1');
    });

    it('rewrites defined names', () => {
      engine.defineName('Doubled', '=Sheet2!A1*2');
      engine.defineName('OnTwo', '=A1', 'Sheet2');
      engine.renameSheet('Sheet2', 'Two Words');
      expect(engine.getNames()).toEqual([
        { name: 'Doubled', formula: "='Two Words'!A1*2", sheet: 'Sheet1' },
        { name: 'OnTwo', formula: '=A1', sheet: 'Two Words' },
      ]);
      expect(value('=Doubled+OnTwo')).toBe(30);
    });

    it('can change only the case of a name', () => {
      engine.renameSheet('Sheet2', 'SHEET2');
      expect(engine.sheetNames()).toEqual(['Sheet1', 'SHEET2', 'My Sheet']);
      expect(raw(sheet1, 'A1')).toBe('=SHEET2!A1+1');
    });

    it('renameSheetInFormula rewrites one formula', () => {
      expect(renameSheetInFormula('=Sheet2!A1+A1', 'Sheet2', 'New')).toBe('=New!A1+A1');
      expect(renameSheetInFormula("='My Sheet'!A:A", 'my sheet', 'S')).toBe('=S!A:A');
      expect(renameSheetInFormula('=Sheet2!A1', 'Sheet2', 'A B')).toBe("='A B'!A1");
      expect(renameSheetInFormula('=Sheet22!A1', 'Sheet2', 'New')).toBe('=Sheet22!A1');
      expect(renameSheetInFormula('Sheet2!A1', 'Sheet2', 'New')).toBe('Sheet2!A1');
    });
  });

  describe('removeSheet', () => {
    it('turns references to the sheet into #REF! after recalculate()', () => {
      engine.recalculate();
      engine.removeSheet('Sheet2');
      engine.recalculate();

      expect(engine.sheetNames()).toEqual(['Sheet1', 'My Sheet']);
      for (const ref of ['A1', 'A3', 'A4']) {
        expect(sheet1.get(key(ref)), ref).toMatchObject({ displayValue: '#REF!', type: 'error' });
      }
      // The formula text is kept, and the error propagates to dependents
      expect(raw(sheet1, 'A1')).toBe('=Sheet2!A1+1');
      expect(shown(sheet1, 'A5')).toBe('#REF!');
      // Other sheets are unaffected
      expect(shown(sheet1, 'A2')).toBe('10');
    });

    it('formulas resolve again when a sheet of that name is added back', () => {
      engine.removeSheet('Sheet2');
      engine.recalculate();
      expect(shown(sheet1, 'A1')).toBe('#REF!');
      engine.addSheet('Sheet2', makeData({ A1: '5' }));
      engine.recalculate();
      expect(shown(sheet1, 'A1')).toBe('6');
      expect(shown(sheet1, 'A5')).toBe('12');
    });

    it('moves the active sheet off a removed sheet and keeps the last sheet', () => {
      engine.setActiveSheet('Sheet2');
      engine.removeSheet('Sheet2');
      expect(engine.activeSheet).toBe('Sheet1');
      engine.removeSheet('My Sheet');
      expect(() => engine.removeSheet('Sheet1')).toThrow('at least one sheet');
      expect(engine.sheetNames()).toEqual(['Sheet1']);
    });

    it('a name defined on the removed sheet is #REF!', () => {
      engine.defineName('OnTwo', '=A1', 'Sheet2');
      expect(value('=OnTwo')).toBe(10);
      engine.removeSheet('Sheet2');
      expect(value('=OnTwo')).toBe('#REF!');
    });
  });

  it('moveSheet changes the order only', () => {
    engine.recalculate();
    engine.moveSheet('My Sheet', 0);
    expect(engine.sheetNames()).toEqual(['My Sheet', 'Sheet1', 'Sheet2']);
    expect(engine.activeSheet).toBe('Sheet1');
    engine.moveSheet('My Sheet', 99);
    expect(engine.sheetNames()).toEqual(['Sheet1', 'Sheet2', 'My Sheet']);
    expect(value('=Sheet2!A1')).toBe(10);
  });

  it('setData replaces one sheet’s data', () => {
    engine.setData(makeData({ A1: '1000' }), 'Sheet2');
    engine.recalculate();
    expect(shown(sheet1, 'A1')).toBe('1001');
    expect(shown(sheet1, 'A3')).toBe('1000');
  });
});
