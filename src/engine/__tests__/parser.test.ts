import { describe, it, expect, beforeEach } from 'vitest';
import { FormulaEngine } from '../formula-engine.js';
import type { GridData } from '../../types.js';
import {
  type Node,
  type Token,
  callsAny,
  parseFormula,
  sheetNameNeedsQuotes,
  sheetPrefix,
  tokenizeFormula,
} from '../parser.js';

/** Token stream as `TYPE:value` words, without the trailing EOF. */
const tokens = (input: string): string =>
  tokenizeFormula(input)
    .filter((t) => t.type !== 'EOF')
    .map((t) => `${t.type}:${t.value}`)
    .join(' ');

const refTokens = (input: string): Token[] => tokenizeFormula(input).filter((t) => t.type === 'REF');

const num = (v: number): Node => ({ t: 'num', v });
const end = (row: number, col: number, rowAbs = false, colAbs = false) => ({ row, col, rowAbs, colAbs });

describe('tokenizeFormula', () => {
  it('splits operands and operators and always ends with EOF', () => {
    expect(tokens('A1+B2')).toBe('REF:A1 OPERATOR:+ REF:B2');
    expect(tokens('1.5e3')).toBe('NUMBER:1.5e3');
    expect(tokens('.5')).toBe('NUMBER:.5');
    expect(tokens('50%')).toBe('NUMBER:50 OPERATOR:%');
    expect(tokens('@A1')).toBe('OPERATOR:@ REF:A1');
    const all = tokenizeFormula('1');
    expect(all[all.length - 1]).toMatchObject({ type: 'EOF', start: 1, end: 1 });
  });

  it('reads two-character comparison operators as one token', () => {
    expect(tokens('a<>b')).toBe('NAME:a OPERATOR:<> NAME:b');
    expect(tokens('a<=b')).toBe('NAME:a OPERATOR:<= NAME:b');
    expect(tokens('a>=b')).toBe('NAME:a OPERATOR:>= NAME:b');
    expect(tokens('a<b')).toBe('NAME:a OPERATOR:< NAME:b');
  });

  it('unescapes doubled quotes in string literals', () => {
    expect(tokens('"a""b"')).toBe('STRING:a"b');
    expect(tokens('""')).toBe('STRING:');
    expect(() => tokenizeFormula('"open')).toThrow('Unterminated string literal');
  });

  it('distinguishes booleans, names, functions and references', () => {
    expect(tokens('TRUE')).toBe('BOOLEAN:TRUE');
    expect(tokens('true')).toBe('BOOLEAN:TRUE');
    expect(tokens('foo')).toBe('NAME:foo');
    expect(tokens('x.y')).toBe('NAME:x.y');
    expect(tokens('_a1')).toBe('NAME:_a1');
    expect(tokens('sum (1)')).toBe('FUNC:SUM LPAREN:( NUMBER:1 RPAREN:)');
    // Four letters are not a column, so this is a name
    expect(tokens('AAAA1')).toBe('NAME:AAAA1');
    expect(tokens('a1')).toBe('REF:A1');
  });

  it('treats something that looks like a cell but is called as a function name', () => {
    expect(tokens('LOG10(5)')).toBe('FUNC:LOG10 LPAREN:( NUMBER:5 RPAREN:)');
    expect(refTokens('LOG10(A1)').map((t) => t.value)).toEqual(['A1']);
  });

  it('reads a literal range as one REF token', () => {
    expect(tokens('SUM(A1:B2)')).toBe('FUNC:SUM LPAREN:( REF:A1:B2 RPAREN:)');
    const [ref] = refTokens('$A1:B$2');
    expect(ref.ref).toEqual({ kind: 'range', sheet: undefined, start: end(0, 0, false, true), end: end(1, 1, true, false) });
    // With spaces around the colon it is the range operator between two cells
    expect(tokens('A1 :B2')).toBe('REF:A1 COLON:: REF:B2');
  });

  it('reads whole-column and whole-row references', () => {
    expect(tokens('A:A')).toBe('REF:A:A');
    expect(tokens('$A:$C')).toBe('REF:$A:$C');
    expect(tokens('1:3')).toBe('REF:1:3');
    expect(tokens('$1:$3')).toBe('REF:$1:$3');
    expect(refTokens('B:D')[0].ref).toEqual({ kind: 'cols', sheet: undefined, start: end(-1, 1), end: end(-1, 3) });
    expect(refTokens('2:$5')[0].ref).toEqual({ kind: 'rows', sheet: undefined, start: end(1, -1), end: end(4, -1, true) });
  });

  it('keeps the sheet name on the reference, unquoted', () => {
    const [plain] = refTokens('Sheet2!A1');
    expect(plain.value).toBe('A1');
    expect(plain.ref).toMatchObject({ kind: 'cell', sheet: 'Sheet2' });
    const [quoted] = refTokens("'My sheet'!A1:B2");
    expect(quoted.ref).toMatchObject({ kind: 'range', sheet: 'My sheet' });
    expect(refTokens("'It''s'!A1")[0].ref?.sheet).toBe("It's");
    expect(refTokens('Sheet2!C:D')[0].ref).toMatchObject({ kind: 'cols', sheet: 'Sheet2' });
    expect(refTokens('Sheet2!1:2')[0].ref).toMatchObject({ kind: 'rows', sheet: 'Sheet2' });
  });

  it('records where the token and the part after the sheet name start', () => {
    const input = "'My sheet'!$A1:B$2 + Sheet2!C:D + 2:$5";
    const refs = refTokens(input);
    expect(refs.map((t) => input.slice(t.start, t.end))).toEqual(["'My sheet'!$A1:B$2", 'Sheet2!C:D', '2:$5']);
    expect(refs.map((t) => input.slice(t.refStart, t.end))).toEqual(['$A1:B$2', 'C:D', '2:$5']);
  });

  it('reads error literals, and # after an operand as the spill operator', () => {
    expect(tokens('#N/A')).toBe('ERROR:#N/A');
    expect(tokens('#DIV/0!+1')).toBe('ERROR:#DIV/0! OPERATOR:+ NUMBER:1');
    expect(tokens('A1#')).toBe('REF:A1 HASH:#');
    expect(tokens('Sheet2!#REF!')).toBe('ERROR:#REF!');
    expect(() => tokenizeFormula('1#')).toThrow();
    expect(() => tokenizeFormula('#NULL!')).toThrow();
  });

  it('reads array braces and separators', () => {
    expect(tokens('{1,2;3}')).toBe('LBRACE:{ NUMBER:1 COMMA:, NUMBER:2 SEMICOLON:; NUMBER:3 RBRACE:}');
  });

  it('rejects text that is not part of the formula language', () => {
    expect(() => tokenizeFormula('~')).toThrow('Unexpected character: ~');
    expect(() => tokenizeFormula('Sheet2!x')).toThrow('Expected a reference after the sheet name');
    expect(() => tokenizeFormula('Sheet2!5')).toThrow('Expected a reference after the sheet name');
    expect(() => tokenizeFormula('$x')).toThrow('Unexpected identifier');
    expect(() => tokenizeFormula('A1:$x')).toThrow('Invalid range');
    expect(() => tokenizeFormula("'unclosed")).toThrow();
    expect(() => tokenizeFormula('Sheet 2!A1')).toThrow();
  });
});

describe('parseFormula', () => {
  it('gives * and / precedence over + and -', () => {
    expect(parseFormula('1+2*3')).toEqual({
      t: 'bin',
      op: '+',
      l: num(1),
      r: { t: 'bin', op: '*', l: num(2), r: num(3) },
    });
  });

  it('binds unary minus tighter than ^ (Excel precedence)', () => {
    expect(parseFormula('-2^2')).toEqual({
      t: 'bin',
      op: '^',
      l: { t: 'unary', op: '-', x: num(2) },
      r: num(2),
    });
    expect(parseFormula('2^-1')).toEqual({ t: 'bin', op: '^', l: num(2), r: { t: 'unary', op: '-', x: num(1) } });
  });

  it('makes ^ left-associative', () => {
    expect(parseFormula('2^3^2')).toEqual({
      t: 'bin',
      op: '^',
      l: { t: 'bin', op: '^', l: num(2), r: num(3) },
      r: num(2),
    });
  });

  it('parses % as a postfix operator that can repeat', () => {
    expect(parseFormula('50%')).toEqual({ t: 'unary', op: '%', x: num(50) });
    expect(parseFormula('1%%')).toEqual({ t: 'unary', op: '%', x: { t: 'unary', op: '%', x: num(1) } });
    // Postfix binds tighter than the unary minus in front
    expect(parseFormula('-5%')).toEqual({ t: 'unary', op: '-', x: { t: 'unary', op: '%', x: num(5) } });
  });

  it('parses literals', () => {
    expect(parseFormula('"a"')).toEqual({ t: 'str', v: 'a' });
    expect(parseFormula('TRUE')).toEqual({ t: 'bool', v: true });
    expect(parseFormula('#N/A')).toEqual({ t: 'err', v: '#N/A' });
    expect(parseFormula('1e3')).toEqual(num(1000));
    expect(parseFormula('x')).toEqual({ t: 'name', name: 'x' });
  });

  it('parses comparison operators at the lowest precedence', () => {
    expect(parseFormula('1+1<=2&"a"')).toMatchObject({
      t: 'bin',
      op: '<=',
      l: { t: 'bin', op: '+' },
      r: { t: 'bin', op: '&' },
    });
    expect(parseFormula('1<>2')).toMatchObject({ t: 'bin', op: '<>' });
  });

  it('parses calls, keeping omitted arguments as null', () => {
    expect(parseFormula('SUM(1,,2)')).toEqual({ t: 'call', name: 'SUM', args: [num(1), null, num(2)] });
    expect(parseFormula('sum()')).toEqual({ t: 'call', name: 'SUM', args: [] });
    expect(parseFormula('F(1,)')).toEqual({ t: 'call', name: 'F', args: [num(1), null] });
  });

  it('parses calling the result of a call or a parenthesised expression', () => {
    expect(parseFormula('LAMBDA(x,x)(1)')).toMatchObject({
      t: 'invoke',
      callee: { t: 'call', name: 'LAMBDA' },
      args: [num(1)],
    });
    expect(parseFormula('f(1)(2)(3)')).toMatchObject({ t: 'invoke', callee: { t: 'invoke' }, args: [num(3)] });
    expect(parseFormula('(f)(2)')).toMatchObject({ t: 'invoke', callee: { t: 'name', name: 'f' } });
    // A number can't be called
    expect(() => parseFormula('1(2)')).toThrow();
  });

  it('parses array literals, including nested ones', () => {
    expect(parseFormula('{1,2;3,4}')).toEqual({ t: 'array', rows: [[num(1), num(2)], [num(3), num(4)]] });
    expect(parseFormula('{1,{2;3}}')).toEqual({
      t: 'array',
      rows: [[num(1), { t: 'array', rows: [[num(2)], [num(3)]] }]],
    });
    expect(parseFormula('{A1}')).toMatchObject({ t: 'array', rows: [[{ t: 'ref' }]] });
    expect(() => parseFormula('{1,2;3}')).toThrow('Array rows differ in length');
    expect(() => parseFormula('{}')).toThrow('Empty array');
  });

  it('parses references, the spill operator and implicit intersection', () => {
    expect(parseFormula('A1:B2')).toMatchObject({ t: 'ref', ref: { kind: 'range' } });
    expect(parseFormula('A1#')).toMatchObject({ t: 'unary', op: '#', x: { t: 'ref', ref: { kind: 'cell' } } });
    expect(parseFormula('@A1:A3')).toMatchObject({ t: 'unary', op: '@', x: { t: 'ref' } });
  });

  it('parses : between expressions as the range operator', () => {
    expect(parseFormula('A1:INDEX(B:B,2)')).toMatchObject({
      t: 'bin',
      op: ':',
      l: { t: 'ref', ref: { kind: 'cell' } },
      r: { t: 'call', name: 'INDEX' },
    });
    expect(parseFormula('A1:B2:C3')).toMatchObject({ t: 'bin', op: ':', l: { t: 'ref', ref: { kind: 'range' } } });
  });

  it('throws on incomplete or malformed formulas', () => {
    for (const bad of ['', ' ', ')', 'SUM(', '1 +', '(1', 'A1 B1', '1,2', 'x y']) {
      expect(() => parseFormula(bad), bad).toThrow();
    }
  });

  it('fails cleanly on pathological nesting instead of overflowing the stack', () => {
    const parens = '('.repeat(5000) + '1' + ')'.repeat(5000);
    expect(() => parseFormula(parens)).toThrow('Formula is nested too deeply');
    expect(() => parseFormula('-'.repeat(5000) + '1')).toThrow('Formula is nested too deeply');
    expect(() => parseFormula('SUM('.repeat(1000) + '1' + ')'.repeat(1000))).toThrow('Formula is nested too deeply');
    // Reasonable nesting is fine
    expect(parseFormula('('.repeat(100) + '1' + ')'.repeat(100))).toEqual(num(1));
  });
});

describe('callsAny', () => {
  const now = new Set(['NOW']);

  it('finds a call anywhere in the tree', () => {
    expect(callsAny(parseFormula('1+SUM(NOW())'), now)).toBe(true);
    expect(callsAny(parseFormula('now()'), now)).toBe(true);
    expect(callsAny(parseFormula('{1,NOW()}'), now)).toBe(true);
    expect(callsAny(parseFormula('-NOW()%'), now)).toBe(true);
    expect(callsAny(parseFormula('A1:OFFSET(A1,NOW(),1)'), now)).toBe(true);
    expect(callsAny(parseFormula('LAMBDA(x,x)(NOW())'), now)).toBe(true);
    expect(callsAny(parseFormula('LAMBDA(x,NOW())(1)'), now)).toBe(true);
  });

  it('ignores other calls, names and string contents', () => {
    expect(callsAny(parseFormula('1+SUM(2)'), now)).toBe(false);
    expect(callsAny(parseFormula('"NOW()"'), now)).toBe(false);
    expect(callsAny(parseFormula('NOW'), now)).toBe(false);
  });
});

describe('sheet names in references', () => {
  it('quotes names that are not plain identifiers or look like a cell', () => {
    expect(sheetNameNeedsQuotes('Sheet2')).toBe(false);
    expect(sheetNameNeedsQuotes('sheet_2')).toBe(false);
    expect(sheetNameNeedsQuotes('x.y')).toBe(false);
    expect(sheetNameNeedsQuotes('ABCD1')).toBe(false);
    expect(sheetNameNeedsQuotes('My sheet')).toBe(true);
    expect(sheetNameNeedsQuotes('1st')).toBe(true);
    expect(sheetNameNeedsQuotes('Data-1')).toBe(true);
    expect(sheetNameNeedsQuotes('A1')).toBe(true);
    expect(sheetNameNeedsQuotes('ABC123')).toBe(true);
  });

  it('sheetPrefix writes the name as it appears before a reference', () => {
    expect(sheetPrefix('Sheet2')).toBe('Sheet2!');
    expect(sheetPrefix('My sheet')).toBe("'My sheet'!");
    expect(sheetPrefix("It's")).toBe("'It''s'!");
    expect(sheetPrefix('A1')).toBe("'A1'!");
  });

  it('a prefix written by sheetPrefix tokenizes back to the same sheet name', () => {
    for (const name of ['Sheet2', 'My sheet', "It's", 'A1', 'Data-1', '1st']) {
      const [ref] = refTokens(`${sheetPrefix(name)}B2`);
      expect(ref.ref?.sheet, name).toBe(name);
    }
  });
});

describe('FormulaEngine – operators and literals', () => {
  let engine: FormulaEngine;
  const value = (formula: string) => engine.evaluateValue(formula);

  beforeEach(() => {
    engine = new FormulaEngine();
    const data: GridData = new Map([
      ['0:0', { rawValue: '4', displayValue: '4', type: 'number' }],
      ['0:1', { rawValue: 'text', displayValue: 'text', type: 'text' }],
    ]);
    engine.setData(data);
  });

  it('evaluates exponentiation with Excel precedence and associativity', () => {
    expect(value('=2^3')).toBe(8);
    expect(value('=-2^2')).toBe(4);
    expect(value('=2^3^2')).toBe(64);
    expect(value('=2^-1')).toBe(0.5);
    expect(value('=2*3^2')).toBe(18);
    expect(value('=A1^0.5')).toBe(2);
  });

  it('reports impossible powers as errors', () => {
    expect(value('=0^-1')).toBe('#DIV/0!');
    expect(value('=(-8)^0.5')).toBe('#NUM!');
  });

  it('reports an overflowing power as #NUM!', () => {
    expect(value('=2^1024')).toBe('#NUM!');
  });

  it('evaluates percent and scientific notation', () => {
    expect(value('=50%')).toBe(0.5);
    expect(value('=200%%')).toBe(0.02);
    expect(value('=A1%')).toBe(0.04);
    expect(value('=1e3')).toBe(1000);
    expect(value('=1.5E-2')).toBe(0.015);
    expect(value('=.5+.5')).toBe(1);
  });

  it('evaluates unary operators', () => {
    expect(value('=-A1')).toBe(-4);
    expect(value('=+A1')).toBe(4);
    expect(value('=--"3"')).toBe(3);
    expect(value('=2*-3')).toBe(-6);
    expect(value('=-B1')).toBe('#VALUE!');
  });

  it('evaluates string literals with escaped quotes', () => {
    expect(value('="a""b"')).toBe('a"b');
    expect(value('="say ""hi"""')).toBe('say "hi"');
    expect(value('=""')).toBe('');
    expect(value('="a"&"b"&1')).toBe('ab1');
    expect(value('="=1+1"')).toBe('=1+1');
  });

  it('evaluates error literals and propagates them', () => {
    expect(value('=#N/A')).toBe('#N/A');
    expect(value('=#DIV/0!')).toBe('#DIV/0!');
    expect(value('=#REF!+1')).toBe('#REF!');
    expect(value('=1+#N/A')).toBe('#N/A');
    expect(engine.evaluate('=#NAME?')).toEqual({ displayValue: '#NAME?', type: 'error' });
  });

  it('evaluates booleans and comparisons', () => {
    expect(value('=TRUE')).toBe(true);
    expect(value('=false')).toBe(false);
    expect(value('=TRUE()')).toBe(true);
    expect(value('=FALSE()')).toBe(false);
    expect(value('=NA()')).toBe('#N/A');
    expect(value('=1=1')).toBe(true);
    expect(value('=1<>1')).toBe(false);
    expect(value('="a"="A"')).toBe(true);
    expect(value('=1=1=TRUE')).toBe(true);
    // TRUE() takes no arguments
    expect(value('=TRUE(1)')).toBe('#VALUE!');
  });

  it('turns syntax errors into #ERROR!', () => {
    for (const bad of ['=1+', '=(1', '=1 2', '=1,2', '={1,2;3}', '={}', "='x", '=#NULL!', '=1#', '=~']) {
      expect(engine.evaluate(bad), bad).toEqual({ displayValue: '#ERROR!', type: 'error' });
    }
  });

  it('turns unknown names and functions into #NAME?', () => {
    expect(value('=nope')).toBe('#NAME?');
    expect(value('=NOPE(1)')).toBe('#NAME?');
    expect(value('=A1+nope')).toBe('#NAME?');
  });

  it('does not crash on deeply nested formulas', () => {
    const ok = '=' + '('.repeat(100) + '1' + ')'.repeat(100);
    expect(engine.evaluate(ok)).toEqual({ displayValue: '1', type: 'number' });
    const tooDeep = '=' + '('.repeat(5000) + '1' + ')'.repeat(5000);
    expect(engine.evaluate(tooDeep)).toEqual({ displayValue: '#ERROR!', type: 'error' });
    expect(engine.evaluate('=' + '-'.repeat(5000) + '1')).toEqual({ displayValue: '#ERROR!', type: 'error' });
    expect(engine.evaluate('=' + 'SUM('.repeat(1000) + '1' + ')'.repeat(1000)).type).toBe('error');
  });

  it('ROW() and COLUMN() without arguments use the cell being evaluated', () => {
    expect(value('=ROW()')).toBe(1);
    expect(engine.evaluateValue('=ROW()', undefined, { row: 4, col: 2 })).toBe(5);
    expect(engine.evaluateValue('=COLUMN()', undefined, { row: 4, col: 2 })).toBe(3);
    expect(engine.evaluate('=ROW()*10+COLUMN()', undefined, undefined, { row: 1, col: 1 }).displayValue).toBe('22');
    // A cell key takes the place of `at`
    expect(engine.evaluate('=ROW()', '6:3').displayValue).toBe('7');
    expect(engine.evaluate('=COLUMN()', '6:3').displayValue).toBe('4');
  });
});
