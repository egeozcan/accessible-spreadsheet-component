import { describe, it, expect, beforeEach } from 'vitest';
import { FormulaEngine, RangeValue } from '../formula-engine.js';
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

/** A value as plain data: arrays as rows of values (nested arrays likewise), blanks as null. */
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
const column = (data: GridData, col: string, from: number, to: number) =>
  Array.from({ length: to - from + 1 }, (_, i) => shown(data, `${col}${from + i}`));

describe('FormulaEngine – broadcasting', () => {
  let engine: FormulaEngine;
  const value = (formula: string) => plain(engine.evaluateValue(formula));

  beforeEach(() => {
    engine = new FormulaEngine();
    // A1:A3 = 1,2,3   B1:B3 = a,b,c   C1:E1 = 10,20,30   A5 = error, B5 = 4, C5 = #N/A
    engine.setData(
      makeData({
        A1: '1', A2: '2', A3: '3',
        B1: 'a', B2: 'b', B3: 'c',
        C1: '10', D1: '20', E1: '30',
        A5: '=1/0', B5: '4', C5: '=NA()',
      })
    );
  });

  it('applies operators to every element of a range', () => {
    expect(value('=A1:A3*2')).toEqual([[2], [4], [6]]);
    expect(value('=A1:A3&"x"')).toEqual([['1x'], ['2x'], ['3x']]);
    expect(value('=A1:A3^2')).toEqual([[1], [4], [9]]);
    expect(value('=A1:A3>1')).toEqual([[false], [true], [true]]);
    expect(value('=-A1:A3')).toEqual([[-1], [-2], [-3]]);
    expect(value('=A1:A3%')).toEqual([[0.01], [0.02], [0.03]]);
  });

  it('applies single-value functions to every element', () => {
    expect(value('=ABS(-A1:A3)')).toEqual([[1], [2], [3]]);
    expect(value('=ABS({-1,2;-3,4})')).toEqual([[1, 2], [3, 4]]);
    expect(value('=UPPER(B1:B3)')).toEqual([['A'], ['B'], ['C']]);
    expect(value('=LEN(B1:B3&B1:B3)')).toEqual([[2], [2], [2]]);
    expect(value('=NOT({1,0})')).toEqual([[false, true]]);
  });

  it('broadcasts over an array in any single-value argument', () => {
    expect(value('=LEFT("hello",{1,2,3})')).toEqual([['h', 'he', 'hel']]);
    expect(value('=ROUND({1.234,2.345},{1;2})')).toEqual([
      [1.2, 2.3],
      [1.23, 2.35],
    ]);
  });

  it('combines a column with a row into a table', () => {
    expect(value('=A1:A3+C1:E1')).toEqual([
      [11, 21, 31],
      [12, 22, 32],
      [13, 23, 33],
    ]);
    expect(value('={1,2}&{"a";"b"}')).toEqual([
      ['1a', '2a'],
      ['1b', '2b'],
    ]);
  });

  it('pads with #N/A where sizes do not match', () => {
    expect(value('={1,2,3}+{1,2}')).toEqual([[2, 4, '#N/A']]);
    expect(value('={1;2;3}+{1;2}')).toEqual([[2], [4], ['#N/A']]);
  });

  it('keeps an error in the element it happened in', () => {
    expect(value('={1,2}/{1,0}')).toEqual([[1, '#DIV/0!']]);
    expect(value('=B5:C5+1')).toEqual([[5, '#N/A']]);
    expect(value('=ISERROR(A5:C5)')).toEqual([[true, false, true]]);
  });

  it('evaluateValue returns the whole array when its first element is an error', () => {
    expect(value('={#N/A,1}')).toEqual([['#N/A', 1]]);
  });

  it('IF with an array condition picks element by element', () => {
    expect(value('=IF(A1:A3>1,1,0)')).toEqual([[0], [1], [1]]);
    expect(value('=IF(A1:A3>1,"big")')).toEqual([[false], ['big'], ['big']]);
    expect(value('=IF(A1:A3>1,A1:A3*10,B1:B3)')).toEqual([['a'], [20], [30]]);
    expect(value('=IF({1,0},{1;2},{3;4})')).toEqual([
      [1, 3],
      [2, 4],
    ]);
    // An error in a branch only shows where that branch is picked
    expect(value('=IF({TRUE,FALSE},1,1/0)')).toEqual([[1, '#DIV/0!']]);
  });

  it('IF with a single condition still only evaluates the branch it returns', () => {
    expect(value('=IF(TRUE,1,1/0)')).toBe(1);
    expect(value('=IF(FALSE,1/0,2)')).toBe(2);
  });

  it('aggregates see the broadcast result', () => {
    expect(value('=SUM(A1:A3*2)')).toBe(12);
    expect(value('=SUM(IF(A1:A3>1,A1:A3))')).toBe(5);
    expect(value('=MAX(A1:A3*C1:E1)')).toBe(90);
    expect(value('=AVERAGE(A1:A3*2)')).toBe(4);
    expect(value('=MIN(-A1:A3)')).toBe(-3);
    expect(value('=AND(A1:A3>0)')).toBe(true);
    expect(value('=OR(A1:A3>3)')).toBe(false);
    expect(value('=CONCAT(A1:A3&"-")')).toBe('1-2-3-');
    expect(value('=TEXTJOIN(",",TRUE,B1:B3&A1:A3)')).toBe('a1,b2,c3');
  });

  it('SUMPRODUCT accepts array expressions', () => {
    expect(value('=SUMPRODUCT(A1:A3,A1:A3)')).toBe(14);
    expect(value('=SUMPRODUCT((A1:A3>1)*A1:A3)')).toBe(5);
    expect(value('=SUMPRODUCT(A1:A3*2,{1;1;1})')).toBe(12);
    expect(value('=SUMPRODUCT(--(A1:A3>=2))')).toBe(2);
  });

  it('IFERROR replaces only the failing elements', () => {
    expect(value('=IFERROR(A5:C5,0)')).toEqual([[0, 4, 0]]);
    expect(value('=IFERROR(1/{1,0,2},"x")')).toEqual([[1, 'x', 0.5]]);
    expect(value('=IFERROR({1,2}/{0,1},{9,8})')).toEqual([[9, 2]]);
    expect(value('=IFNA(B5:C5,"na")')).toEqual([[4, 'na']]);
  });

  it('array literals hold numbers, text, booleans, errors and expressions', () => {
    expect(value('={1,2;3,4}')).toEqual([
      [1, 2],
      [3, 4],
    ]);
    expect(value('={1,"a",TRUE}')).toEqual([[1, 'a', true]]);
    expect(value('={1,#N/A}')).toEqual([[1, '#N/A']]);
    expect(value('={-1,2}')).toEqual([[-1, 2]]);
    expect(value('={A1,A2+1}')).toEqual([[1, 3]]);
    expect(value('=SUM({1,2;3,4})')).toBe(10);
    expect(value('={1,2}={1,3}')).toEqual([[true, false]]);
  });

  it('evaluate() shows the top-left value of an array result', () => {
    expect(engine.evaluate('={1,2;3,4}')).toEqual({ displayValue: '1', type: 'number' });
    expect(engine.evaluate('=B1:B3&"!"')).toEqual({ displayValue: 'a!', type: 'text' });
    expect(engine.evaluate('={#N/A,1}')).toEqual({ displayValue: '#N/A', type: 'error' });
  });
});

describe('FormulaEngine – array functions', () => {
  let engine: FormulaEngine;
  const value = (formula: string) => plain(engine.evaluateValue(formula));

  beforeEach(() => {
    engine = new FormulaEngine();
    // A1:C4 is a table of fruit, amount, tag; E1:E3 = 1,2,3; G1:I2 has a blank row and errors
    engine.setData(
      makeData({
        A1: 'pear', B1: '3', C1: 'x',
        A2: 'apple', B2: '1', C2: 'y',
        A3: 'fig', B3: '2', C3: 'x',
        A4: 'apple', B4: '5', C4: 'z',
        E1: '1', E2: '2', E3: '3',
        G2: '=1/0', H2: '4', I2: '=NA()',
      })
    );
  });

  it('SEQUENCE', () => {
    expect(value('=SEQUENCE(3)')).toEqual([[1], [2], [3]]);
    expect(value('=SEQUENCE(2,3)')).toEqual([
      [1, 2, 3],
      [4, 5, 6],
    ]);
    expect(value('=SEQUENCE(2,2,10,5)')).toEqual([
      [10, 15],
      [20, 25],
    ]);
    expect(value('=SEQUENCE(3,1,3,-1)')).toEqual([[3], [2], [1]]);
    expect(value('=SEQUENCE(1)')).toBe(1);
    expect(value('=SEQUENCE(0)')).toBe('#CALC!');
    expect(value('=SUM(SEQUENCE(100))')).toBe(5050);
  });

  it('FILTER keeps the rows (or columns) whose condition is true', () => {
    expect(value('=FILTER(A1:C4,B1:B4>1)')).toEqual([
      ['pear', 3, 'x'],
      ['fig', 2, 'x'],
      ['apple', 5, 'z'],
    ]);
    expect(value('=FILTER(A1:A4,C1:C4="x")')).toEqual([['pear'], ['fig']]);
    expect(value('=FILTER(A1:C4,(B1:B4>1)*(C1:C4="x"))')).toEqual([
      ['pear', 3, 'x'],
      ['fig', 2, 'x'],
    ]);
    expect(value('=FILTER(A1:C2,{1,0,1})')).toEqual([
      ['pear', 'x'],
      ['apple', 'y'],
    ]);
  });

  it('FILTER with no match is #CALC!, or if_empty when given', () => {
    expect(value('=FILTER(A1:C4,B1:B4>10)')).toBe('#CALC!');
    expect(value('=FILTER(A1:C4,B1:B4>10,"none")')).toBe('none');
    // An error in if_empty only matters when it is used
    expect(value('=FILTER(A1:A4,B1:B4>10,1/0)')).toBe('#DIV/0!');
    expect(value('=FILTER(A1:A4,B1:B4>4,1/0)')).toBe('apple');
    // The condition must match the array's height or width
    expect(value('=FILTER(A1:C4,{1,0})')).toBe('#VALUE!');
  });

  it('SORT', () => {
    expect(value('=SORT(B1:B4)')).toEqual([[1], [2], [3], [5]]);
    expect(value('=SORT(B1:B4,1,-1)')).toEqual([[5], [3], [2], [1]]);
    expect(value('=SORT(A1:C4,2)')).toEqual([
      ['apple', 1, 'y'],
      ['fig', 2, 'x'],
      ['pear', 3, 'x'],
      ['apple', 5, 'z'],
    ]);
    expect(value('=SORT(A1:C4,2,-1)')).toEqual([
      ['apple', 5, 'z'],
      ['pear', 3, 'x'],
      ['fig', 2, 'x'],
      ['apple', 1, 'y'],
    ]);
    // Equal keys keep their original order
    expect(value('=SORT(A1:B4)')).toEqual([
      ['apple', 1],
      ['apple', 5],
      ['fig', 2],
      ['pear', 3],
    ]);
    expect(value('=SORT({3,1,2},1,1,TRUE)')).toEqual([[1, 2, 3]]);
    // Numbers before text (ignoring case) before booleans
    expect(value('=SORT({"b";"A";3;TRUE;1})')).toEqual([[1], [3], ['A'], ['b'], [true]]);
    expect(value('=SORT(A1:C4,4)')).toBe('#VALUE!');
    expect(value('=SORT(B1:B4,1,2)')).toBe('#VALUE!');
  });

  it('SORTBY sorts by one or more other arrays', () => {
    expect(value('=SORTBY(A1:A4,B1:B4)')).toEqual([['apple'], ['fig'], ['pear'], ['apple']]);
    expect(value('=SORTBY(A1:A4,B1:B4,-1)')).toEqual([['apple'], ['pear'], ['fig'], ['apple']]);
    expect(value('=SORTBY(A1:B4,C1:C4,1,B1:B4,-1)')).toEqual([
      ['pear', 3],
      ['fig', 2],
      ['apple', 1],
      ['apple', 5],
    ]);
    expect(value('=SORTBY(A1:A4,B1:B3)')).toBe('#VALUE!');
  });

  it('SORTBY sorts a single row by a row of keys', () => {
    expect(value('=SORTBY({1,2,3},{3,1,2})')).toEqual([[2, 3, 1]]);
  });

  it('UNIQUE', () => {
    expect(value('=UNIQUE(A1:A4)')).toEqual([['pear'], ['apple'], ['fig']]);
    expect(value('=UNIQUE(C1:C4)')).toEqual([['x'], ['y'], ['z']]);
    expect(value('=UNIQUE({"a";"A";"b"})')).toEqual([['a'], ['b']]);
    // Whole rows are compared
    expect(value('=UNIQUE({1,2;1,2;3,4})')).toEqual([
      [1, 2],
      [3, 4],
    ]);
    expect(value('=UNIQUE({1,1,2},TRUE)')).toEqual([[1, 2]]);
  });

  it('UNIQUE with exactly_once keeps only what occurs once', () => {
    expect(value('=UNIQUE(A1:A4,FALSE,TRUE)')).toEqual([['pear'], ['fig']]);
    expect(value('=UNIQUE({1;1},,TRUE)')).toBe('#CALC!');
  });

  it('TRANSPOSE', () => {
    expect(value('=TRANSPOSE(E1:E3)')).toEqual([[1, 2, 3]]);
    expect(value('=TRANSPOSE({1,2;3,4})')).toEqual([
      [1, 3],
      [2, 4],
    ]);
    expect(value('=TRANSPOSE(5)')).toBe(5);
  });

  it('XMATCH', () => {
    expect(value('=XMATCH(2,E1:E3)')).toBe(2);
    expect(value('=XMATCH("FIG",A1:A4)')).toBe(3);
    // Search from the end
    expect(value('=XMATCH("apple",A1:A4,0,-1)')).toBe(4);
    expect(value('=XMATCH("apple",A1:A4)')).toBe(2);
    // Next smaller / next larger
    expect(value('=XMATCH(2.5,E1:E3,-1)')).toBe(2);
    expect(value('=XMATCH(2.5,E1:E3,1)')).toBe(3);
    expect(value('=XMATCH(9,E1:E3)')).toBe('#N/A');
    expect(value('=XMATCH({1,3},E1:E3)')).toEqual([[1, 3]]);
  });

  it('XLOOKUP returns a whole row when the return range has several columns', () => {
    expect(value('=XLOOKUP("fig",A1:A4,B1:C4)')).toEqual([[2, 'x']]);
    expect(value('=XLOOKUP("fig",A1:A4,B1:B4)')).toBe(2);
    expect(value('=XLOOKUP(2,{1,2,3},{"a","b","c";"d","e","f"})')).toEqual([['b'], ['e']]);
    expect(value('=XLOOKUP("zz",A1:A4,B1:C4,"none")')).toBe('none');
    expect(value('=XLOOKUP("zz",A1:A4,B1:C4)')).toBe('#N/A');
    expect(value('=XLOOKUP({"fig";"pear"},A1:A4,B1:B4)')).toEqual([[2], [3]]);
  });

  it('VSTACK and HSTACK pad the shorter arrays with #N/A', () => {
    expect(value('=VSTACK({1,2},{3,4})')).toEqual([
      [1, 2],
      [3, 4],
    ]);
    expect(value('=VSTACK(1,2)')).toEqual([[1], [2]]);
    expect(value('=VSTACK(E1:E2,{10,20})')).toEqual([
      [1, '#N/A'],
      [2, '#N/A'],
      [10, 20],
    ]);
    expect(value('=HSTACK(E1:E3,A1:A3)')).toEqual([
      [1, 'pear'],
      [2, 'apple'],
      [3, 'fig'],
    ]);
    expect(value('=HSTACK({1;2},3)')).toEqual([
      [1, 3],
      [2, '#N/A'],
    ]);
  });

  it('TAKE', () => {
    expect(value('=TAKE(A1:C4,2)')).toEqual([
      ['pear', 3, 'x'],
      ['apple', 1, 'y'],
    ]);
    expect(value('=TAKE(A1:C4,-1)')).toEqual([['apple', 5, 'z']]);
    expect(value('=TAKE(A1:C4,,1)')).toEqual([['pear'], ['apple'], ['fig'], ['apple']]);
    expect(value('=TAKE(A1:C4,2,-2)')).toEqual([
      [3, 'x'],
      [1, 'y'],
    ]);
    // More than there is gives everything; nothing is #CALC!
    expect(value('=ROWS(TAKE(A1:C4,10))')).toBe(4);
    expect(value('=TAKE(A1:C4,0)')).toBe('#CALC!');
  });

  it('DROP', () => {
    expect(value('=DROP(A1:C4,1)')).toEqual([
      ['apple', 1, 'y'],
      ['fig', 2, 'x'],
      ['apple', 5, 'z'],
    ]);
    expect(value('=DROP(A1:C4,-3)')).toEqual([['pear', 3, 'x']]);
    expect(value('=DROP(A1:C4,,2)')).toEqual([['x'], ['y'], ['x'], ['z']]);
    expect(value('=DROP(A1:C4,1,-1)')).toEqual([
      ['apple', 1],
      ['fig', 2],
      ['apple', 5],
    ]);
    expect(value('=DROP(A1:C4,4)')).toBe('#CALC!');
  });

  it('CHOOSECOLS and CHOOSEROWS pick by position, negative from the end', () => {
    expect(value('=CHOOSECOLS(A1:C2,1,3)')).toEqual([
      ['pear', 'x'],
      ['apple', 'y'],
    ]);
    expect(value('=CHOOSECOLS(A1:C2,-1)')).toEqual([['x'], ['y']]);
    expect(value('=CHOOSECOLS(A1:C2,{3,1})')).toEqual([
      ['x', 'pear'],
      ['y', 'apple'],
    ]);
    expect(value('=CHOOSEROWS(A1:C4,1,-1)')).toEqual([
      ['pear', 3, 'x'],
      ['apple', 5, 'z'],
    ]);
    expect(value('=CHOOSEROWS(A1:B4,2,2)')).toEqual([
      ['apple', 1],
      ['apple', 1],
    ]);
    expect(value('=CHOOSECOLS(A1:C4,4)')).toBe('#VALUE!');
    expect(value('=CHOOSECOLS(A1:C4,0)')).toBe('#VALUE!');
    expect(value('=CHOOSEROWS(A1:C4,5)')).toBe('#VALUE!');
  });

  it('TOCOL and TOROW', () => {
    expect(value('=TOCOL({1,2;3,4})')).toEqual([[1], [2], [3], [4]]);
    expect(value('=TOCOL({1,2;3,4},,TRUE)')).toEqual([[1], [3], [2], [4]]);
    expect(value('=TOROW({1,2;3,4})')).toEqual([[1, 2, 3, 4]]);
    expect(value('=TOROW(A1:B2)')).toEqual([['pear', 3, 'apple', 1]]);
  });

  it('TOCOL can ignore blanks and errors', () => {
    // G1:I1 is blank, G2 = #DIV/0!, H2 = 4, I2 = #N/A
    expect(value('=TOCOL(H1:I2,1)')).toEqual([[4], ['#N/A']]);
    expect(value('=TOCOL(H2:I2,2)')).toBe(4);
    expect(value('=TOCOL(G1:I2,3)')).toBe(4);
    expect(value('=TOCOL(G1:I1,1)')).toBe('#CALC!');
    expect(value('=TOCOL({1,#N/A;2,3},2)')).toEqual([[1], [2], [3]]);
  });

  it('TOCOL keeps blank cells of a range when not told to ignore them', () => {
    expect(value('=ROWS(TOCOL(H1:I2))')).toBe(4);
  });

  it('TEXTSPLIT', () => {
    expect(value('=TEXTSPLIT("a,b,c",",")')).toEqual([['a', 'b', 'c']]);
    expect(value('=TEXTSPLIT("a,b;c,d",",",";")')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
    expect(value('=TEXTSPLIT("a;b",,";")')).toEqual([['a'], ['b']]);
    expect(value('=TEXTSPLIT("a-b_c",{"-","_"})')).toEqual([['a', 'b', 'c']]);
    expect(value('=TEXTSPLIT("abc",)')).toBe('#VALUE!');
  });

  it('TEXTSPLIT options: padding, empty items and case', () => {
    expect(value('=TEXTSPLIT("a,b;c",",",";")')).toEqual([
      ['a', 'b'],
      ['c', '#N/A'],
    ]);
    expect(value('=TEXTSPLIT("a,b;c",",",";",,,"-")')).toEqual([
      ['a', 'b'],
      ['c', '-'],
    ]);
    expect(value('=TEXTSPLIT("a,,b",",")')).toEqual([['a', '', 'b']]);
    expect(value('=TEXTSPLIT("a,,b",",",,TRUE)')).toEqual([['a', 'b']]);
    expect(value('=TEXTSPLIT("aXbxc","x")')).toEqual([['aXb', 'c']]);
    expect(value('=TEXTSPLIT("aXbxc","x",,,1)')).toEqual([['a', 'b', 'c']]);
  });

  it('MAP, REDUCE, BYROW and BYCOL call a LAMBDA over an array', () => {
    expect(value('=MAP(E1:E3,LAMBDA(x,x*x))')).toEqual([[1], [4], [9]]);
    expect(value('=MAP({1,2;3,4},LAMBDA(x,x+1))')).toEqual([
      [2, 3],
      [4, 5],
    ]);
    expect(value('=MAP(E1:E2,{10,20},LAMBDA(a,b,a+b))')).toEqual([
      [11, 21],
      [12, 22],
    ]);
    expect(value('=REDUCE(0,E1:E3,LAMBDA(acc,v,acc+v*v))')).toBe(14);
    expect(value('=REDUCE("",A1:A3,LAMBDA(acc,v,acc&LEFT(v,1)))')).toBe('paf');
    expect(value('=BYROW({1,2;3,4},LAMBDA(r,SUM(r)))')).toEqual([[3], [7]]);
    expect(value('=BYCOL({1,2;3,4},LAMBDA(c,MAX(c)))')).toEqual([[3, 4]]);
    expect(value('=MAP(E1:E3,5)')).toBe('#VALUE!');
  });

  it('array functions compose', () => {
    expect(value('=SORT(UNIQUE(A1:A4))')).toEqual([['apple'], ['fig'], ['pear']]);
    expect(value('=SUM(TAKE(SORT(B1:B4,1,-1),2))')).toBe(8);
    expect(value('=ROWS(FILTER(A1:A4,B1:B4>1))')).toBe(3);
    expect(value('=COLUMNS(SEQUENCE(2,5))')).toBe(5);
    expect(value('=TEXTJOIN("|",TRUE,SORT(FILTER(A1:A4,C1:C4="x")))')).toBe('fig|pear');
    expect(value('=INDEX(SORT(A1:C4,2,-1),1,1)')).toBe('apple');
  });
});

describe('FormulaEngine – spilling', () => {
  let engine: FormulaEngine;

  beforeEach(() => {
    engine = new FormulaEngine();
  });

  it('writes the spilled values into the cells below the formula', () => {
    const data = makeData({ A1: '1', A2: '2', A3: '3', C1: '=A1:A3*2' });
    engine.setData(data);
    const changed = engine.recalculate();

    expect(data.get(key('C1'))).toEqual({
      rawValue: '=A1:A3*2',
      displayValue: '2',
      type: 'number',
      spill: { rows: 3, cols: 1 },
    });
    expect(data.get(key('C2'))).toEqual({ rawValue: '', displayValue: '4', type: 'number', spilledFrom: key('C1') });
    expect(data.get(key('C3'))).toEqual({ rawValue: '', displayValue: '6', type: 'number', spilledFrom: key('C1') });
    expect([...changed].sort()).toEqual([key('C1'), key('C2'), key('C3')]);
    expect(engine.spillRange(key('C1'))).toEqual({ rows: 3, cols: 1 });
    expect(engine.spillRange(key('A1'))).toBeNull();
  });

  it('spills in two dimensions with each value keeping its type', () => {
    const data = makeData({ A1: '={1,#N/A;"a",TRUE}' });
    engine.setData(data);
    engine.recalculate();
    expect(data.get(key('A1'))).toMatchObject({ displayValue: '1', type: 'number', spill: { rows: 2, cols: 2 } });
    expect(data.get(key('B1'))).toMatchObject({ displayValue: '#N/A', type: 'error', spilledFrom: key('A1') });
    expect(data.get(key('A2'))).toMatchObject({ displayValue: 'a', type: 'text', spilledFrom: key('A1') });
    expect(data.get(key('B2'))).toMatchObject({ displayValue: 'TRUE', type: 'boolean', spilledFrom: key('A1') });
  });

  it('spills an array whose first value is an error', () => {
    const data = makeData({ A1: '={1,2}/{0,1}' });
    engine.setData(data);
    engine.recalculate();
    expect(data.get(key('A1'))).toMatchObject({ displayValue: '#DIV/0!', type: 'error', spill: { rows: 1, cols: 2 } });
    expect(shown(data, 'B1')).toBe('2');
  });

  it('resultFor returns the unformatted value of a spilled cell', () => {
    const data = makeData({ A1: '=SEQUENCE(3)*1.5' });
    engine.setData(data);
    engine.recalculate();
    expect(engine.resultFor(key('A2'), '')).toEqual({ displayValue: '3', type: 'number' });
    expect(engine.resultFor(key('A3'), '')).toEqual({ displayValue: '4.5', type: 'number' });
    expect(engine.resultFor(key('A1'), '=SEQUENCE(3)*1.5')).toEqual({ displayValue: '1.5', type: 'number' });
    // An ordinary empty cell
    expect(engine.resultFor(key('F9'), '')).toEqual({ displayValue: '', type: 'text' });
  });

  it('formulas can read spilled cells, alone and through ranges', () => {
    const data = makeData({ A1: '=SEQUENCE(2,2)', D1: '=B2&"!"', D2: '=SUM(A1:B2)', D3: '=A2*10' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'D1')).toBe('4!');
    expect(shown(data, 'D2')).toBe('10');
    expect(shown(data, 'D3')).toBe('30');
  });

  it('updates dependents of spilled cells when an input changes', () => {
    const data = makeData({ A1: '1', A2: '2', A3: '3', C1: '=A1:A3*2', E1: '=C2+1', E2: '=SUM(C1:C10)' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'E1')).toBe('5');
    expect(shown(data, 'E2')).toBe('12');

    const changed = engine.recalculateAffected([write(data, 'A2', '20')]);
    expect(shown(data, 'C2')).toBe('40');
    expect(shown(data, 'E1')).toBe('41');
    expect(shown(data, 'E2')).toBe('48');
    expect([...changed].sort()).toEqual([key('C2'), key('E1'), key('E2')].sort());
  });

  it('a value in the way turns the formula into #SPILL! and removes the spilled cells', () => {
    const data = makeData({ A1: '1', A2: '2', A3: '3', C1: '=A1:A3*2', E1: '=C2+1', E2: '=SUM(C1:C10)' });
    engine.setData(data);
    engine.recalculate();

    engine.recalculateAffected([write(data, 'C3', 'x')]);
    expect(data.get(key('C1'))).toEqual({ rawValue: '=A1:A3*2', displayValue: '#SPILL!', type: 'error' });
    expect(data.has(key('C2'))).toBe(false);
    expect(data.get(key('C3'))?.rawValue).toBe('x');
    expect(engine.spillRange(key('C1'))).toBeNull();
    // Readers of the vanished cells follow
    expect(shown(data, 'E1')).toBe('1');
    expect(shown(data, 'E2')).toBe('#SPILL!');
  });

  it('removing the obstruction restores the spill', () => {
    const data = makeData({ A1: '1', A2: '2', A3: '3', C1: '=A1:A3*2', C3: 'x', E1: '=C2+1' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'C1')).toBe('#SPILL!');
    expect(shown(data, 'E1')).toBe('1');

    data.delete(key('C3'));
    engine.recalculateAffected([key('C3')]);
    expect(column(data, 'C', 1, 3)).toEqual(['2', '4', '6']);
    expect(data.get(key('C3'))?.spilledFrom).toBe(key('C1'));
    expect(engine.spillRange(key('C1'))).toEqual({ rows: 3, cols: 1 });
    expect(shown(data, 'E1')).toBe('5');
  });

  it('a formula written into the spill range blocks it too', () => {
    const data = makeData({ A1: '=SEQUENCE(3)' });
    engine.setData(data);
    engine.recalculate();
    engine.recalculateAffected([write(data, 'A2', '=1+1')]);
    expect(shown(data, 'A1')).toBe('#SPILL!');
    expect(shown(data, 'A2')).toBe('2');
    expect(data.has(key('A3'))).toBe(false);
  });

  it('growing a spill adds cells and updates formulas that read them', () => {
    // E1 comes before the spilling formula, so it reads C3 before the spill reaches it
    const data = makeData({
      E1: '=C3',
      E2: '=SUM(C1:C10)',
      E3: '=COUNT(C:C)',
      E4: '=C4+1',
      A1: '2',
      C1: '=SEQUENCE(A1)',
    });
    engine.setData(data);
    engine.recalculate();
    expect(column(data, 'C', 1, 4)).toEqual(['1', '2', undefined, undefined]);
    expect(column(data, 'E', 1, 4)).toEqual(['0', '3', '2', '1']);

    const changed = engine.recalculateAffected([write(data, 'A1', '4')]);
    expect(column(data, 'C', 1, 4)).toEqual(['1', '2', '3', '4']);
    expect(data.get(key('C1'))?.spill).toEqual({ rows: 4, cols: 1 });
    expect(column(data, 'E', 1, 4)).toEqual(['3', '10', '4', '5']);
    expect(changed.has(key('C3'))).toBe(true);
    expect(changed.has(key('C4'))).toBe(true);
    expect(changed.has(key('C2'))).toBe(false);
  });

  it('a full recalculation gets readers right whatever the order of the cells', () => {
    const data = makeData({ E1: '=C3', E2: '=SUM(C1:C10)', E3: '=C2+C3', C1: '=SEQUENCE(3)' });
    engine.setData(data);
    engine.recalculate();
    expect(column(data, 'E', 1, 3)).toEqual(['3', '6', '5']);
  });

  it('shrinking a spill removes cells and updates formulas that read them', () => {
    const data = makeData({ A1: '4', C1: '=SEQUENCE(A1)', E1: '=C3', E2: '=SUM(C1:C10)', E3: '=COUNT(C:C)' });
    engine.setData(data);
    engine.recalculate();
    expect(column(data, 'E', 1, 3)).toEqual(['3', '10', '4']);

    engine.recalculateAffected([write(data, 'A1', '2')]);
    expect(column(data, 'C', 1, 4)).toEqual(['1', '2', undefined, undefined]);
    expect(data.has(key('C3'))).toBe(false);
    expect(data.get(key('C1'))?.spill).toEqual({ rows: 2, cols: 1 });
    expect(column(data, 'E', 1, 3)).toEqual(['0', '3', '2']);
  });

  it('a spill that shrinks to one value stops being a spill', () => {
    const data = makeData({ A1: '3', C1: '=SEQUENCE(A1)' });
    engine.setData(data);
    engine.recalculate();
    engine.recalculateAffected([write(data, 'A1', '1')]);
    expect(data.get(key('C1'))).toEqual({ rawValue: '=SEQUENCE(A1)', displayValue: '1', type: 'number' });
    expect(data.has(key('C2'))).toBe(false);
    expect(engine.spillRange(key('C1'))).toBeNull();
  });

  it('removes the spilled cells when the formula is deleted', () => {
    const data = makeData({ A1: '=SEQUENCE(3)', C1: '=SUM(A1:A3)' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'C1')).toBe('6');

    data.delete(key('A1'));
    engine.recalculateAffected([key('A1')]);
    expect(data.has(key('A2'))).toBe(false);
    expect(data.has(key('A3'))).toBe(false);
    expect(shown(data, 'C1')).toBe('0');
  });

  it('of two spills that would overlap, the later one is #SPILL!', () => {
    const down = makeData({ B1: '=SEQUENCE(3)', A2: '=SEQUENCE(1,3)' });
    engine.setData(down);
    engine.recalculate();
    expect(column(down, 'B', 1, 3)).toEqual(['1', '2', '3']);
    expect(shown(down, 'A2')).toBe('#SPILL!');
    expect(down.has(key('C2'))).toBe(false);

    const across = makeData({ A2: '=SEQUENCE(1,3)', B1: '=SEQUENCE(3)' });
    engine.setData(across);
    engine.recalculate();
    expect(shown(across, 'B1')).toBe('#SPILL!');
    expect([shown(across, 'A2'), shown(across, 'B2'), shown(across, 'C2')]).toEqual(['1', '2', '3']);
    expect(across.has(key('B3'))).toBe(false);
  });

  it('a spill beyond the sheet bounds is #SPILL!', () => {
    const data = makeData({ A1: '=SEQUENCE(5)', B1: '=SEQUENCE(1,3)', B3: '=SEQUENCE(3)', C5: '=SEQUENCE(2)' });
    engine.setData(data);
    engine.setSheetSize(5, 3);
    engine.recalculate();
    // Fits exactly: rows 1-5
    expect(column(data, 'A', 1, 5)).toEqual(['1', '2', '3', '4', '5']);
    expect(column(data, 'B', 3, 5)).toEqual(['1', '2', '3']);
    // Would need column D / row 6
    expect(shown(data, 'B1')).toBe('#SPILL!');
    expect(shown(data, 'C5')).toBe('#SPILL!');
    expect(data.has(key('C1'))).toBe(false);
  });

  it('the spill operator refers to the whole spilled range', () => {
    const data = makeData({
      A1: '2',
      C1: '=SEQUENCE(A1)',
      E1: '=SUM(C1#)',
      E2: '=ROWS(C1#)',
      G1: '=C1#*10',
      // Not a spill: the cell itself
      E3: '=SUM(A1#)',
      // No formula there at all
      E4: '=Z9#',
    });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'E1')).toBe('3');
    expect(shown(data, 'E2')).toBe('2');
    expect(column(data, 'G', 1, 3)).toEqual(['10', '20', undefined]);
    expect(shown(data, 'E3')).toBe('#REF!');
    expect(shown(data, 'E4')).toBe('#REF!');

    engine.recalculateAffected([write(data, 'A1', '3')]);
    expect(shown(data, 'E1')).toBe('6');
    expect(shown(data, 'E2')).toBe('3');
    expect(column(data, 'G', 1, 3)).toEqual(['10', '20', '30']);
  });

  it('the spill operator on a blocked spill gives #SPILL!', () => {
    const data = makeData({ C1: '=SEQUENCE(3)', C2: 'x', E1: '=SUM(C1#)' });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'E1')).toBe('#SPILL!');
  });

  it('@ picks the value on the formula’s own row or column', () => {
    const data = makeData({
      A1: '10', A2: '20', A3: '30', B1: '11', C1: '12',
      E2: '=@A1:A3',
      E3: '=@A:A',
      B5: '=@A1:C1',
      // Neither on a row nor a column of the range: its first cell
      F7: '=@A1:A3',
      G1: '=@SEQUENCE(3)',
      G2: '=@A1',
    });
    engine.setData(data);
    engine.recalculate();
    expect(shown(data, 'E2')).toBe('20');
    expect(shown(data, 'E3')).toBe('30');
    expect(shown(data, 'B5')).toBe('11');
    expect(shown(data, 'F7')).toBe('10');
    expect(data.get(key('G1'))).toEqual({ rawValue: '=@SEQUENCE(3)', displayValue: '1', type: 'number' });
    expect(shown(data, 'G2')).toBe('10');
  });

  it('a formula that reads its own spill range ends as an error', () => {
    const data = makeData({ A1: '=SEQUENCE(3)+A2', C1: '=C1:C3*2', E1: '=E2:E3', G1: '=SUM(G1#)' });
    engine.setData(data);
    engine.recalculate();
    for (const ref of ['A1', 'C1', 'E1', 'G1']) {
      expect(data.get(key(ref))?.type, ref).toBe('error');
    }
    // Still an error on later passes
    engine.recalculateAffected([key('A1'), key('C1'), key('E1')]);
    expect(data.get(key('A1'))?.type).toBe('error');
    expect(data.get(key('C1'))?.type).toBe('error');
  });

  it('a spilled-into cell with a format survives the spill going away', () => {
    const data = makeData({ A1: '3', C1: '=SEQUENCE(A1)' });
    data.set(key('C2'), { rawValue: '', displayValue: '', type: 'text', format: { bold: true } });
    engine.setData(data);
    engine.recalculate();
    expect(data.get(key('C2'))).toEqual({
      rawValue: '',
      displayValue: '2',
      type: 'number',
      format: { bold: true },
      spilledFrom: key('C1'),
    });
    const formatted = data.get(key('C2'));

    engine.recalculateAffected([write(data, 'A1', '1')]);
    // The formatted cell is emptied but kept (the same object); the plain one is deleted
    expect(data.get(key('C2'))).toBe(formatted);
    expect(formatted).toEqual({ rawValue: '', displayValue: '', type: 'text', format: { bold: true } });
    expect(data.has(key('C3'))).toBe(false);
  });

  it('recalculate() cleans up stale spilled cells', () => {
    const data = makeData({ A1: '=SEQUENCE(2)' });
    // Left over from an earlier state: nothing spills into these any more
    data.set(key('G8'), { rawValue: '', displayValue: '99', type: 'number', spilledFrom: key('G7') });
    data.set(key('G9'), { rawValue: '', displayValue: '99', type: 'number', spilledFrom: key('G7'), format: { bold: true } });
    data.set(key('G10'), { rawValue: '5', displayValue: '5', type: 'number', spilledFrom: key('G7') });
    engine.setData(data);
    const changed = engine.recalculate();

    expect(data.has(key('G8'))).toBe(false);
    expect(data.get(key('G9'))).toEqual({ rawValue: '', displayValue: '', type: 'text', format: { bold: true } });
    expect(data.get(key('G10'))).toEqual({ rawValue: '5', displayValue: '5', type: 'number' });
    expect(changed.has(key('G8'))).toBe(true);
    expect(changed.has(key('G9'))).toBe(true);
    // The real spill is rebuilt
    expect(data.get(key('A2'))).toMatchObject({ displayValue: '2', spilledFrom: key('A1') });
  });

  it('recalculate() twice gives the same sheet', () => {
    const data = makeData({ A1: '=SEQUENCE(2,2)', D1: '=SUM(A1#)', A4: '=TRANSPOSE(A1#)' });
    engine.setData(data);
    engine.recalculate();
    const snapshot = JSON.stringify([...data]);
    engine.recalculate();
    expect(JSON.stringify([...data])).toBe(snapshot);
    expect(shown(data, 'D1')).toBe('10');
    expect([shown(data, 'A4'), shown(data, 'B4'), shown(data, 'A5'), shown(data, 'B5')]).toEqual(['1', '3', '2', '4']);
  });

  it('a spill can feed another spill', () => {
    const data = makeData({ A1: '2', C1: '=SEQUENCE(A1)', E1: '=C1#*10', G1: '=SUM(E1#)' });
    engine.setData(data);
    engine.recalculate();
    expect(column(data, 'E', 1, 3)).toEqual(['10', '20', undefined]);
    expect(shown(data, 'G1')).toBe('30');

    engine.recalculateAffected([write(data, 'A1', '3')]);
    expect(column(data, 'E', 1, 3)).toEqual(['10', '20', '30']);
    expect(data.get(key('E1'))?.spill).toEqual({ rows: 3, cols: 1 });
    expect(shown(data, 'G1')).toBe('60');
  });

  it('reports spilled cells of other sheets in engine.changes', () => {
    const sheet1 = makeData({ A1: '=Sheet2!A1#', C1: '=SUM(Sheet2!A1#)', C2: '=Sheet2!A2*10' });
    const sheet2 = makeData({ A1: '=SEQUENCE(B1)', B1: '3' });
    engine.setData(sheet1);
    engine.addSheet('Sheet2', sheet2);
    engine.recalculate();
    expect(column(sheet2, 'A', 1, 3)).toEqual(['1', '2', '3']);
    expect(column(sheet1, 'A', 1, 3)).toEqual(['1', '2', '3']);
    expect(shown(sheet1, 'C1')).toBe('6');
    expect(shown(sheet1, 'C2')).toBe('20');
    expect(engine.spillRange(key('A1'), 'Sheet2')).toEqual({ rows: 3, cols: 1 });
    expect(engine.resultFor(key('A2'), '', 'Sheet2')).toEqual({ displayValue: '2', type: 'number' });

    const changed = engine.recalculateAffected([write(sheet2, 'B1', '2')], 'Sheet2');
    // The return value is for the sheet that was named; `changes` has every sheet
    expect([...changed].sort()).toEqual([key('A1'), key('A3')]);
    expect([...(engine.changes.get('Sheet1') ?? [])].sort()).toEqual([key('A1'), key('C1'), key('A3')]);
    expect(sheet2.has(key('A3'))).toBe(false);
    expect(sheet1.has(key('A3'))).toBe(false);
    expect(shown(sheet1, 'C1')).toBe('3');
  });
});

describe('second review regressions', () => {
  it('IFS picks per element for array conditions', () => {
    const engine = new FormulaEngine();
    const picked = engine.evaluateValue('=IFS({FALSE,TRUE},1,TRUE,2)') as RangeValue;
    expect(picked.values).toEqual([2, 1]);
    expect((engine.evaluateValue('=IFS({FALSE,TRUE},1)') as RangeValue).values).toEqual(['#N/A', 1]);
    expect(engine.evaluateValue('=IFS(FALSE,1,TRUE,2)')).toBe(2);
  });

  it('a failing IFERROR fallback only fails the elements that needed it', () => {
    const engine = new FormulaEngine();
    expect((engine.evaluateValue('=IFERROR({1,#N/A},1/0)') as RangeValue).values).toEqual([1, '#DIV/0!']);
  });

  it('A1# gives all spilled values before the sheet has been recalculated', () => {
    const engine = new FormulaEngine();
    engine.setData(new Map([['0:0', { rawValue: '=SEQUENCE(3)', displayValue: '', type: 'text' }]]));
    expect(engine.evaluate('=SUM(A1#)').displayValue).toBe('6');
    engine.recalculate();
    expect(engine.evaluate('=SUM(A1#)').displayValue).toBe('6');
    expect(engine.evaluate('=ROWS(A1#)').displayValue).toBe('3');
  });

  it('INDEX of a cell with an error is still a reference', () => {
    const engine = new FormulaEngine();
    engine.setData(new Map([['0:0', { rawValue: '=1/0', displayValue: '', type: 'text' }]]));
    expect(engine.evaluate('=ROW(INDEX(A1,1))').displayValue).toBe('1');
    expect(engine.evaluate('=INDEX(A1,1)').displayValue).toBe('#DIV/0!');
  });
});

describe('third review regressions', () => {
  type Cell = GridData extends Map<string, infer C> ? C : never;
  const cellOf = (rawValue: string): Cell => ({ rawValue, displayValue: '', type: 'text' });
  const numberCell = (rawValue: string): Cell => ({ rawValue, displayValue: rawValue, type: 'number' });

  it('A1# is a reference before and after the sheet shows the spill', () => {
    const engine = new FormulaEngine();
    engine.setData(new Map([['0:0', cellOf('=SEQUENCE(3)')]]));
    expect((engine.evaluateValue('=ROW(A1#)') as RangeValue).values).toEqual([1, 2, 3]);
    expect(engine.evaluate('=SUM(A1:A3)').displayValue).toBe('6');
    engine.recalculate();
    expect((engine.evaluateValue('=ROW(A1#)') as RangeValue).values).toEqual([1, 2, 3]);
  });

  it('A1# keeps the other values when the first one is an error', () => {
    const engine = new FormulaEngine();
    const data = new Map([['0:0', cellOf('={#N/A;1;2}')]]);
    engine.setData(data);
    engine.recalculate();
    expect(engine.evaluate('=SUM(IFERROR(A1#,0))').displayValue).toBe('3');
  });

  it('a long chain of spills that depend on each other settles', () => {
    const engine = new FormulaEngine();
    const data: GridData = new Map([
      ['4:0', numberCell('1')],
      ['0:0', cellOf('=SEQUENCE($A$5)')],
    ]);
    const col = (c: number) => {
      let s = '';
      for (let n = c; n >= 0; n = Math.floor(n / 26) - 1) s = String.fromCharCode(65 + (n % 26)) + s;
      return s;
    };
    for (let c = 1; c <= 40; c++) data.set(`0:${c}`, cellOf(`=SEQUENCE(IF(${col(c - 1)}2>0,2,1))`));
    engine.setData(data);
    engine.recalculate();
    expect(data.has('1:40')).toBe(false);
    data.get('4:0')!.rawValue = '2';
    engine.recalculateAffected(['4:0']);
    expect(data.get('1:40')?.displayValue).toBe('2');
  });

  it('a volatile function that fails is still recomputed, and so is one that became volatile', () => {
    const engine = new FormulaEngine();
    const data: GridData = new Map([
      ['0:0', cellOf('=TICK()')],
      ['0:1', numberCell('1')],
    ]);
    let calls = 0;
    engine.registerFunction(
      'TICK',
      () => {
        if (++calls === 1) throw new Error('#N/A');
        return calls;
      },
      { volatile: true }
    );
    engine.setData(data);
    engine.recalculate();
    expect(data.get('0:0')!.displayValue).toBe('#N/A');
    engine.recalculateAffected(['0:1']);
    expect(data.get('0:0')!.displayValue).toBe('2');

    let plain = 0;
    engine.registerFunction('TICK', () => ++plain);
    engine.recalculate();
    engine.registerFunction('TICK', () => ++plain, { volatile: true });
    engine.recalculate();
    const before = data.get('0:0')!.displayValue;
    engine.recalculateAffected(['0:1']);
    expect(data.get('0:0')!.displayValue).not.toBe(before);
  });
});

describe('fourth review regressions', () => {
  type Cell = GridData extends Map<string, infer C> ? C : never;
  const formula = (rawValue: string): Cell => ({ rawValue, displayValue: '', type: 'text' });

  it('standalone evaluation sees a blocked spill as #SPILL! and a free one in whole columns', () => {
    const engine = new FormulaEngine();
    engine.setData(new Map([['0:0', formula('=SEQUENCE(3)')]]));
    expect(engine.evaluateValue('=SUM(A:A)')).toBe(6);
    expect(engine.evaluateValue('=COUNTA(A:A)')).toBe(3);

    const blocked: GridData = new Map([
      ['0:0', formula('=SEQUENCE(3)')],
      ['1:0', { rawValue: 'x', displayValue: 'x', type: 'text' }],
    ]);
    engine.setData(blocked);
    expect(engine.evaluateValue('=ISERROR(A1)')).toBe(true);
    expect(engine.evaluate('=A1').displayValue).toBe('#SPILL!');
    engine.recalculate();
    expect(engine.evaluateValue('=ISERROR(A1)')).toBe(true);
  });

  it('a volatile formula registered with evaluate(raw, key) is recomputed by later recalculations', () => {
    const engine = new FormulaEngine();
    let calls = 0;
    engine.registerFunction('TICK', () => ++calls, { volatile: true });
    const data: GridData = new Map([
      ['0:0', formula('=TICK()')],
      ['0:2', { rawValue: '1', displayValue: '1', type: 'number' }],
    ]);
    engine.setData(data);
    engine.evaluate('=TICK()', '0:0');
    const before = calls;
    data.get('0:2')!.rawValue = '2';
    engine.updateData(data, ['0:2']);
    engine.recalculateAffected(['0:2']);
    expect(calls).toBeGreaterThan(before);
  });
});
