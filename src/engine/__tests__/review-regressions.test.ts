import { describe, it, expect } from 'vitest';
import { FormulaEngine, RangeValue } from '../formula-engine.js';
import { compileWildcard, wildcardMatches } from '../functions.js';

const ev = (f: string, e = new FormulaEngine()) => e.evaluateValue(f);
const timed = (f: string) => {
  const t = Date.now();
  const v = ev(f);
  return { v, ms: Date.now() - t };
};

describe('review regressions – no hangs or huge allocations', () => {
  it('SEARCH with many wildcards does not backtrack exponentially', () => {
    const r = timed('=SEARCH(REPT("*a",25)&"b",REPT("a",50))');
    expect(r.v).toBe('#VALUE!');
    expect(r.ms).toBeLessThan(1000);
  });
  it('factorial and combinatorial functions stop at overflow', () => {
    expect(ev('=FACTDOUBLE(1E20)')).toBe('#NUM!');
    expect(ev('=FACT(171)')).toBe('#NUM!');
    expect(ev('=PERMUT(1E20,1E19)')).toBe('#NUM!');
    expect(ev('=COMBIN(1E20,1E19)')).toBe('#NUM!');
  });
  it('array producers refuse oversized results before allocating', () => {
    expect(ev('=EXPAND(1,1000000000,1)')).toBe('#NUM!');
    expect(ev('=MUNIT(100000)')).toBe('#NUM!');
    expect(ev('=RANDARRAY(100000,100000)')).toBe('#NUM!');
    expect(ev('=WRAPROWS(SEQUENCE(10),1000000000)')).toBe('#NUM!');
  });
  it('BINOM.INV and the depreciation functions are not linear in the periods', () => {
    expect(timed('=BINOM.INV(1000000,0.5,0.5)').ms).toBeLessThan(1000);
    expect(ev('=BINOM.INV(1000000,0.5,0.5)')).toBe(500000);
    expect(ev('=BINOM.INV(100,0.5,0.5)')).toBe(50);
    expect(ev('=BINOM.INV(10,0.3,0.3)')).toBe(2);
    expect(timed('=DDB(1,1,1000000000000,1000000000000)').ms).toBeLessThan(1000);
    expect(ev('=DDB(1,1,1000000000000,1000000000000)')).toBe(0);
    expect(timed('=DB(1000,100,1000000000000,5)').ms).toBeLessThan(1000);
  });
});

describe('review regressions – results', () => {
  it('DDB and DB closed forms match the period loop (Microsoft examples)', () => {
    expect(ev('=ROUND(DDB(2400,300,10*365,1,2),2)')).toBe(1.32);
    expect(ev('=ROUND(DDB(2400,300,10,2,2),2)')).toBe(384);
    expect(ev('=ROUND(DB(1000000,100000,6,1,7),2)')).toBe(186083.33);
    expect(ev('=ROUND(DB(1000000,100000,6,2,7),2)')).toBe(259639.42);
    expect(ev('=ROUND(DB(1000000,100000,6,7,7),2)')).toBe(15845.1);
  });
  it('QUOTIENT and EVEN keep digits beyond the tenth', () => {
    expect(ev('=QUOTIENT(12345678901,1)')).toBe(12345678901);
    expect(ev('=QUOTIENT(0.99999999999,1)')).toBe(0);
    expect(ev('=CEILING(0.07,0.01)')).toBe(0.07);
  });
  it('MAXA of negatives', () => expect(ev('=MAXA(-5,-2)')).toBe(-2));
  it('IPMT / PPMT for an annuity due', () => {
    expect(ev('=ROUND(IPMT(0.1,2,2,1000,0,1),10)')).toBeCloseTo(-47.6190476190, 8);
    expect(ev('=ROUND(PPMT(0.1,2,2,1000,0,1),10)')).toBeCloseTo(-476.190476190, 8);
    expect(ev('=IPMT(0.1,1,2,1000,0,1)')).toBe(0);
  });
  it('a blank lookup value looks for 0', () => {
    const e = new FormulaEngine();
    expect(e.evaluateValue('=VLOOKUP(Z9,{0,"zero";1,"one"},2,FALSE)')).toBe('zero');
    expect(e.evaluateValue('=MATCH(Z9,{0,1},0)')).toBe(1);
  });
  it('VALUE reads currency, dates and times', () => {
    expect(ev('=VALUE("$1,000")')).toBe(1000);
    expect(ev('=VALUE("-$5")')).toBe(-5);
    expect(ev('=VALUE("12:00")')).toBe(0.5);
    expect(ev('=VALUE("2020-01-01")')).toBe(43831);
    expect(ev('=VALUE("1,00")')).toBe('#VALUE!');
  });
  it('SWITCH matches an array expression element by element', () => {
    const v = ev('=SWITCH({1,2},1,"one",2,"two")') as RangeValue;
    expect(v.values).toEqual(['one', 'two']);
    expect(ev('=SWITCH(3,1,"one",2,"two")')).toBe('#N/A');
  });
});

describe('review regressions – round two', () => {
  it('SEARCH with a leading wildcard over long text is linear-ish', () => {
    const r = timed('=SEARCH("*"&REPT("a",3000)&"b",REPT("a",9000))');
    expect(r.v).toBe('#VALUE!');
    expect(r.ms).toBeLessThan(2000);
    expect(ev('=SEARCH("b*d","abcd")')).toBe(2);
    expect(ev('=SEARCH("c?e","abcde")')).toBe(3);
    expect(ev('=SEARCH("*c","abcabc")')).toBe(1);
    expect(ev('=SEARCH("~*","a*b")')).toBe(2);
    expect(ev('=SEARCH("a*z","xxabcz",3)')).toBe(3);
    expect(ev('=SEARCH("a*z","xxabcz",4)')).toBe('#VALUE!');
  });
  it('loops bounded by user counts terminate', () => {
    expect(timed('=CUMIPMT(0.1,1E20,1000,1E20,1E20,0)').ms).toBeLessThan(1000);
    expect(timed('=BINOM.INV(1E20,1,0.5)').ms).toBeLessThan(1000);
    expect(timed('=HYPGEOM.DIST(1E20,1E20,1E20,1E20,TRUE)').ms).toBeLessThan(1000);
    expect(timed('=BINOM.DIST(5E19,1E20,0.5,FALSE)').ms).toBeLessThan(1000);
    expect(ev('=TEXTSPLIT(REPT(",",17000)&REPT(";",10000),",",";")')).toBe('#NUM!');
    expect(ev('=TEXTJOIN(",",TRUE,REPT("a",32767),"b")')).toBe('#VALUE!');
  });
  it('BINOM.INV keeps lower and upper tails', () => {
    expect(ev('=BINOM.INV(100,0.5,1E-13)')).toBe(15);
    expect(ev('=BINOM.INV(100,0.5,1)')).toBe('#NUM!'); // Excel: alpha of 1 is #NUM!
  });
  it('CHOOSE and SWITCH broadcast array alternatives', () => {
    const rows = (f: string) => {
      const v = ev(f) as RangeValue;
      return Array.from({ length: v.rows }, (_, r) => v.getRow(r));
    };
    expect(rows('=CHOOSE({1,2},{1;2},{3;4})')).toEqual([[1, 3], [2, 4]]);
    expect(rows('=SWITCH({1,2},1,{10,20},2,{30,40})')).toEqual([[10, 40]]);
    expect(rows('=SWITCH({1/0,2},1,"one",2,"two","other")')).toEqual([['#DIV/0!', 'two']]);
  });
});

describe('review regressions – round three', () => {
  it('cumulative HYPGEOM.DIST works on large counts quickly', () => {
    const r = timed('=HYPGEOM.DIST(100000,200000,200000,400000,TRUE)');
    expect(r.ms).toBeLessThan(1000);
    expect(r.v as number).toBeCloseTo(0.5012615638955756, 6);
    expect(ev('=HYPGEOM.DIST(1,4,8,20,TRUE)') as number).toBeCloseTo(0.4654281, 6);
    expect(ev('=HYPGEOM.DIST(3,4,8,20,TRUE)') as number).toBeCloseTo(4775 / 4845, 9);
  });
  it('SWITCH broadcasts match values too', () => {
    const v = ev('=SWITCH(TRUE,{TRUE,FALSE},"yes","no")') as RangeValue;
    expect(v.values).toEqual(['yes', 'no']);
  });
  it('a non-numeric CHOOSE selector only fails its own element', () => {
    const v = ev('=CHOOSE({1,"x"},10,20)') as RangeValue;
    expect(v.values).toEqual([10, '#VALUE!']);
  });
});

describe('review regressions – round four', () => {
  it('a match error only affects the elements that reach it', () => {
    const v = ev('=SWITCH({1,2},1,"one",1/0,"two")') as RangeValue;
    expect(v.values).toEqual(['one', '#DIV/0!']);
    expect(ev('=SWITCH(1,1,"one",1/0,"two")')).toBe('one');
  });
  it('DAYS360 US method leaves other month ends alone', () => {
    expect(ev('=DAYS360(DATE(2021,1,1),DATE(2021,4,30))')).toBe(119);
    expect(ev('=DAYS360(DATE(2021,1,30),DATE(2021,2,28))')).toBe(28);
    expect(ev('=DAYS360(DATE(2021,2,28),DATE(2021,3,31))')).toBe(30);
    expect(ev('=DAYS360(DATE(2021,1,31),DATE(2021,3,31))')).toBe(60);
    expect(ev('=YEARFRAC(DATE(2021,1,1),DATE(2021,2,28),0)') as number).toBeCloseTo(0.158333333333333, 12);
  });
  it('PERCENTRANK interpolates from the last of tied values', () => {
    expect(ev('=PERCENTRANK.INC({1,2,2,4},3)')).toBe(0.833);
    expect(ev('=PERCENTRANK.EXC({1,2,2,4},3)')).toBe(0.7);
    expect(ev('=PERCENTRANK.INC({1,2,2,4},2)')).toBe(0.333);
  });
});

describe('review regressions – round five', () => {
  it('DAYS360 / YEARFRAC February rules match Excel', () => {
    expect(ev('=DAYS360(DATE(2021,2,28),DATE(2022,2,28))')).toBe(358);
    expect(ev('=DAYS360(DATE(2021,2,28),DATE(2021,2,28))')).toBe(-2);
    expect(ev('=YEARFRAC(DATE(2021,2,28),DATE(2022,2,28),0)')).toBe(1);
    expect(ev('=YEARFRAC(DATE(2021,2,28),DATE(2021,3,31),0)') as number).toBeCloseTo(31 / 360, 12);
  });
  it('SEARCH on maximum-length text answers promptly', () => {
    const r = timed('=SEARCH("*"&REPT("a",30000)&"b",REPT("a",32767))');
    expect(r.v).toBe('#VALUE!');
    expect(r.ms).toBeLessThan(3000);
  });
  it('XIRR takes blank amounts as 0', () => {
    const e = new FormulaEngine();
    e.setData(new Map([
      ['0:0', { rawValue: '-100', displayValue: '-100', type: 'number' }],
      ['2:0', { rawValue: '121', displayValue: '121', type: 'number' }],
      ['0:1', { rawValue: '43831', displayValue: '43831', type: 'number' }],
      ['1:1', { rawValue: '44013', displayValue: '44013', type: 'number' }],
      ['2:1', { rawValue: '44561', displayValue: '44561', type: 'number' }],
    ]) as never);
    expect(e.evaluateValue('=XIRR(A1:A3,B1:B3)') as number).toBeCloseTo(0.1, 6);
    expect(e.evaluateValue('=MDETERM(A1:B2)')).toBe('#VALUE!');
  });
  it('TEXTBEFORE / TEXTAFTER fallbacks apply per element', () => {
    expect((ev('=TEXTBEFORE({"a-b","x"},"-",1,0,0,{"no","none"})') as RangeValue).values).toEqual(['a', 'none']);
    expect((ev('=TEXTAFTER({"a-b","x"},"-",1,0,0,{"no","none"})') as RangeValue).values).toEqual(['b', 'none']);
  });
  it('year-first slashed dates and hour-only AM/PM times', () => {
    expect(ev('=DATEVALUE("2024/01/05")')).toBe(45296);
    expect(ev('=VALUE("2024/01/05")')).toBe(45296);
    expect(ev('=TIMEVALUE("2 PM")') as number).toBeCloseTo(0.5833333333, 9);
  });
});

describe('review regressions – round six', () => {
  it('FREQUENCY is not quadratic', () => {
    const r = timed('=SUM(FREQUENCY(SEQUENCE(300000),SEQUENCE(300000)))');
    expect(r.v).toBe(300000);
    expect(r.ms).toBeLessThan(3000);
    expect((ev('=FREQUENCY({1,2,3,4,5},{2,4})') as RangeValue).values).toEqual([2, 2, 1]);
  });
  it('VALUE / DATEVALUE / TIMEVALUE read more date and time text', () => {
    expect(ev('=DATEVALUE("01-05-2024")')).toBe(45296);
    expect(ev('=VALUE("01-05-2024")')).toBe(45296);
    expect(ev('=VALUE("2 PM")') as number).toBeCloseTo(0.5833333333, 9);
    expect(ev('=TIMEVALUE("12:30:00.5")') as number).toBeCloseTo(0.5208391203703704, 12);
    expect(ev('=DATEVALUE("2024-01-05 12:30:00.5")')).toBe(45296);
    expect(ev('=VALUE("1,00")')).toBe('#VALUE!');
  });
});

describe('review regressions – round seven', () => {
  it('BINOM.INV terminates near 2^53', () => {
    expect(timed('=BINOM.INV(9007199254740988,1,0.5)').ms).toBeLessThan(1000);
  });
  it('number parsers have no quadratic backtracking', () => {
    expect(timed('=VALUE(REPT("1",32766)&"x")').ms).toBeLessThan(500);
    expect(timed('=NUMBERVALUE(REPT("1",32766)&"x")').ms).toBeLessThan(500);
    expect(ev('=VALUE("(1,234.56)")')).toBe(-1234.56);
  });
  it('wildcard SEARCH / XMATCH are linear for plain segments', () => {
    expect(timed('=SEARCH("*"&REPT("a",16000)&"b",REPT("a",32767))').ms).toBeLessThan(1000);
    expect(timed('=XMATCH("*"&REPT("a",16000)&"b",REPT("a",32767),2)').ms).toBeLessThan(1000);
    expect(ev('=XMATCH("a*b*c",{"abc","axxbyyc","ab"},2)')).toBe(1);
    expect(ev('=XMATCH("a?c*",{"ab","abcd"},2)')).toBe(2);
  });
  it('ARRAYTOTEXT stops before building gigabytes', () => {
    expect(ev('=ARRAYTOTEXT(EXPAND(REPT(CHAR(34),32767),1000,100,REPT(CHAR(34),32767)),1)')).toBe('#VALUE!');
  });
  it('matrix work is bounded; big maxima do not spread arguments', () => {
    expect(timed('=SUM(MMULT(MUNIT(2000),MUNIT(2000)))').v).toBe('#NUM!');
    expect(ev('=SUM(MINVERSE(MUNIT(400)))')).toBe(400);
    expect(ev('=SUMXMY2({1,2,3},{1;2;3})')).toBe(0);
  });
  it('fractional seconds up to 59.9', () => {
    expect(ev('=VALUE("2024-01-05 12:30:59.5")') as number).toBeCloseTo(45296.52152199074, 9);
  });
});

describe('wildcard matcher against a regex reference', () => {
  it('agrees on 20,000 random small patterns', () => {
    let seed = 12345;
    const rnd = (n: number) => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) % n);
    const pick = (alphabet: string, len: number) => Array.from({ length: len }, () => alphabet[rnd(alphabet.length)]).join('');
    for (let i = 0; i < 20000; i++) {
      const pattern = pick('ab*?~', 1 + rnd(6));
      const text = pick('abAB~*?', rnd(8));
      const tokens = compileWildcard(pattern);
      if (!tokens) continue;
      let re = '';
      for (let k = 0; k < pattern.length; k++) {
        const ch = pattern[k];
        if (ch === '~' && k + 1 < pattern.length && '*?~'.includes(pattern[k + 1])) re += '\\' + pattern[++k];
        else if (ch === '*') re += '[\\s\\S]*';
        else if (ch === '?') re += '[\\s\\S]';
        else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      }
      const expected = new RegExp('^' + re + '$', 'i').test(text);
      expect(wildcardMatches(tokens, text), `${pattern} vs ${text}`).toBe(expected);
    }
  });
});

describe('review regressions – round eight', () => {
  it('array-valued padding is rejected before it multiplies', () => {
    expect(ev('=SUM(EXPAND(1,100,100,SEQUENCE(10000)))')).toBe('#VALUE!');

  });
  it('NETWORKDAYS and WORKDAY count whole weeks arithmetically', () => {
    expect(timed('=SUM(NETWORKDAYS(SEQUENCE(100000),2958465))').ms).toBeLessThan(2000);
    expect(ev('=NETWORKDAYS(DATE(2024,1,1),DATE(2024,12,31))')).toBe(262);
    expect(ev('=WORKDAY(DATE(2024,1,1),1000)')).toBe(46692); // values below checked in Excel
    expect(ev('=WORKDAY(DATE(2024,1,1),-1000)')).toBe(ev('=DATE(2020,3,2)'));
    expect(ev('=WORKDAY(DATE(2024,1,6),3,DATE(2024,1,9))')).toBe(45302);
    expect(ev('=WORKDAY.INTL(DATE(2024,1,1),37,"0000011",{45296,45297,45298,45299})')).toBe(45345); // 39 working days on: two holidays fall on workdays
    expect(ev('=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,3,31),11)')).toBe(78);
    expect(ev('=WORKDAY(DATE(2024,1,1),10,DATE(2024,1,3))')).toBe(ev('=WORKDAY(DATE(2024,1,1),11)'));
  });
  it('wildcard lookups stay fast with long patterns and many candidates', () => {
    expect(timed('=XMATCH(REPT("?",32767),EXPAND("x",100000,1,"x"),2)').ms).toBeLessThan(2000);
  });
  it('TEXTJOIN checks its length before building the text', () => {
    expect(ev('=TEXTJOIN(REPT("a",32767),FALSE,SEQUENCE(10000))')).toBe('#VALUE!');
  });
});

describe('review regressions – round nine', () => {
  it('lifted rank and order queries sort the reference once', () => {
    expect(timed('=SUM(LARGE(SEQUENCE(100000),SEQUENCE(100000)))').ms).toBeLessThan(3000);
    expect(timed('=SUM(RANK.EQ(SEQUENCE(100000),SEQUENCE(100000)))').ms).toBeLessThan(3000);
    expect(timed('=SUM(PERCENTRANK.INC(SEQUENCE(20000),SEQUENCE(20000)))').ms).toBeLessThan(3000);
    expect(ev('=SUM(LARGE(SEQUENCE(1000),SEQUENCE(1000)))')).toBe(500500);
  });
  it('RANK and PERCENTRANK take array queries (as Excel does)', () => {
    const e = new FormulaEngine();
    e.setData(new Map([['0:0', { rawValue: '1', displayValue: '1', type: 'number' }], ['1:0', { rawValue: '2', displayValue: '2', type: 'number' }], ['2:0', { rawValue: '3', displayValue: '3', type: 'number' }]]) as never);
    expect(e.evaluateValue('=SUM(RANK.EQ({1,2},A1:A3))')).toBe(5);
    expect(e.evaluateValue('=SUM(PERCENTRANK.INC({1,2,3},{1,2}))')).toBe(0.5);
    expect(e.evaluateValue('=SUM(PERCENTRANK.EXC({1,2,3},{1,2}))')).toBe(0.75);
  });
  it('WORKDAY with a million consecutive holidays is not quadratic', () => {
    expect(timed('=WORKDAY(45000,1,SEQUENCE(1000000,1,45001))').ms).toBeLessThan(3000);
  });
  it('a call per element cannot hold a billion cells together', () => {
    expect(ev('=SUM(MUNIT(SEQUENCE(1000,1,1000,0)))')).toBe('#NUM!');
  });
});

describe('review regressions – round ten', () => {
  it('AGGREGATE 14-19 take an array of k', () => {
    expect(ev('=SUM(AGGREGATE(14,6,{10,20,30,40},{1,2}))')).toBe(70);
    expect(ev('=SUM(AGGREGATE(15,6,{10,20,30,40},{1,2}))')).toBe(30);
    expect(ev('=AGGREGATE(14,6,{10,20,30,40},2)')).toBe(30);
  });
  it('ISREF is FALSE for an invalid reference and TRUE for a real one', () => {
    expect(ev('=ISREF(INDIRECT("MissingSheet!A1"))')).toBe(false);
    expect(ev('=ISREF(1/0)')).toBe(false);
    expect(ev('=ISREF(A1:B2)')).toBe(true);
    expect(ev('=ISREF(5)')).toBe(false);
  });
});

describe('review regressions – round eleven', () => {
  it('LOOKUP skips errors in the lookup vector', () => {
    expect(ev('=LOOKUP("zzz",{"a",#N/A})')).toBe('a');
    expect(ev('=LOOKUP("zzz",{"a",#N/A},{10,20})')).toBe(10);
  });
});

describe('review regressions – round twelve', () => {
  const grid = (cells: Record<string, string>) => {
    const e = new FormulaEngine();
    const m = new Map();
    for (const [ref, raw] of Object.entries(cells)) {
      const col = ref.charCodeAt(0) - 65;
      const row = Number(ref.slice(1)) - 1;
      m.set(`${row}:${col}`, { rawValue: raw, displayValue: raw, type: raw !== '' && !isNaN(Number(raw)) ? 'number' : 'text' });
    }
    e.setData(m as never);
    return e;
  };
  it('COUNTBLANK counts a whole column at the size of the grid', () => {
    const e = grid({ A1: '1' });
    expect(e.evaluateValue('=COUNTBLANK(A:A)')).toBeGreaterThan(1000);
    expect(e.evaluateValue('=COUNTBLANK(A1:A3)')).toBe(2);
    expect(e.evaluateValue('=COUNTBLANK(Z9)')).toBe(1);
    expect(grid({ A1: '=1/0' }).evaluateValue('=COUNTBLANK(A1)')).toBe(0);
  });
  it('CHOOSE and SWITCH keep the chosen reference', () => {
    const e = grid({ A1: '1', A2: '2', B1: '3', B2: '4' });
    expect(e.evaluateValue('=ROWS(OFFSET(CHOOSE(1,A1:A2,B1:B2),0,0))')).toBe(2);
    expect(e.evaluateValue('=ROWS(OFFSET(SWITCH(1,1,A1:A2,B1:B2),0,0))')).toBe(2);
    expect(e.evaluateValue('=ISREF(CHOOSE(1,A1,A2))')).toBe(true);
    expect(e.evaluateValue('=SUM(CHOOSE(2,A1:A2,B1:B2))')).toBe(7);
    expect(e.evaluateValue('=CHOOSE(2,1/0,"ok")')).toBe('ok');
    expect(e.evaluateValue('=SWITCH(3,1,1/0,3,"three")')).toBe('three');
  });
  it('array broadcasting reads blanks as 0', () => {
    const e = grid({ A2: '1' });
    expect(e.evaluateValue('=SUM(SWITCH(A1:A2,0,10,20))')).toBe(30);
    expect(e.evaluateValue('=TEXTJOIN("|",FALSE,CHOOSE({1,1},A1:A2))')).toBe('0|0|1|1');
  });
  it('MAXIFS / MINIFS criteria do not match logical cells', () => {
    const e = new FormulaEngine();
    e.setData(new Map([
      ['0:0', { rawValue: '100', displayValue: '100', type: 'number' }],
      ['1:0', { rawValue: '5', displayValue: '5', type: 'number' }],
      ['0:1', { rawValue: 'TRUE', displayValue: 'TRUE', type: 'boolean' }],
      ['1:1', { rawValue: '1', displayValue: '1', type: 'number' }],
    ]) as never);
    expect(e.evaluateValue('=MAXIFS(A1:A2,B1:B2,1)')).toBe(5);
  });
});

describe('review regressions – round thirteen', () => {
  const cells = (entries: Array<[string, unknown, string]>) => {
    const e = new FormulaEngine();
    e.setData(new Map(entries.map(([k, v, type]) => [k, { rawValue: String(v), displayValue: String(v), type }])) as never);
    return e;
  };
  it('blank scalars read as 0 in CHOOSE / SWITCH over arrays', () => {
    const e = new FormulaEngine();
    expect(e.evaluateValue('=AVERAGE(CHOOSE({1,2},Z9,20))')).toBe(10);
    expect((e.evaluateValue('=SWITCH({0,1},Z9,10,20)') as RangeValue).values).toEqual([10, 20]);
    expect(e.evaluateValue('=SWITCH(0,,"yes","no")')).toBe('yes');
  });
  it('criteria: logical, relational and blank cases match Excel', () => {
    const e = cells([
      ['0:0', 10, 'number'], ['1:0', 100, 'number'],
      ['0:1', 'TRUE', 'boolean'], ['1:1', '="TRUE"', 'text'],
    ]);
    expect(e.evaluateValue('=MAXIFS(A1:A2,B1:B2,TRUE)')).toBe(10);
    expect(e.evaluateValue('=COUNTIF(B1:B2,TRUE)')).toBe(1);
    const n = cells([
      ['0:0', 10, 'number'], ['1:0', 100, 'number'],
      ['0:1', 1, 'number'], ['1:1', '="1"', 'text'],
    ]);
    expect(n.evaluateValue('=COUNTIF(B1:B2,">0")')).toBe(1);
    expect(n.evaluateValue('=COUNTIF(B1:B2,1)')).toBe(2);
    const b = cells([['0:0', 10, 'number'], ['1:0', 100, 'number'], ['1:1', '=""', 'text']]);
    expect(b.evaluateValue('=COUNTIF(B1:B2,"<>")')).toBe(1);
    expect(b.evaluateValue('=COUNTIF(B1:B2,"=")')).toBe(1);
    expect(b.evaluateValue('=COUNTIF(B1:B2,"")')).toBe(2);
  });
});

describe('review regressions – round fourteen', () => {
  const cells = (entries: Array<[string, unknown, string]>) => {
    const e = new FormulaEngine();
    e.setData(new Map(entries.map(([k, v, type]) => [k, { rawValue: String(v), displayValue: String(v), type }])) as never);
    return e;
  };
  it('AVERAGEIF "=" and "<>x" see empty cells; relational text skips logicals', () => {
    const e = cells([['1:0', '=""', 'text'], ['0:1', 10, 'number'], ['1:1', 100, 'number']]);
    expect(e.evaluateValue('=AVERAGEIF(A1:A2,"=",B1:B2)')).toBe(10);
    const x = cells([['1:0', 'x', 'text'], ['0:1', 100, 'number'], ['1:1', 10, 'number']]);
    expect(x.evaluateValue('=MAXIFS(B1:B2,A1:A2,"<>x")')).toBe(100);
    const b = cells([['0:0', '=TRUE()', 'boolean'], ['1:0', '=FALSE()', 'boolean']]);
    expect(b.evaluateValue('=COUNTIF(A1:A2,">a")')).toBe(0);
  });
  it('<> counts numeric text as different; relational text compares text cells', () => {
    const e = cells([['0:0', '="1"', 'text'], ['1:0', '="2"', 'text'], ['0:1', 100, 'number'], ['1:1', 10, 'number']]);
    expect(e.evaluateValue('=MAXIFS(B1:B2,A1:A2,"<>1")')).toBe(100);
    expect(e.evaluateValue('=COUNTIF(A1:A2,"<a")')).toBe(2);
  });
  it('SUMIF / AVERAGEIF sum numbers only and return a summed error', () => {
    const e = cells([['0:0', 1, 'number'], ['1:0', 1, 'number'], ['0:1', '="10"', 'text'], ['1:1', 20, 'number']]);
    expect(e.evaluateValue('=SUMIF(A1:A2,1,B1:B2)')).toBe(20);
    expect(e.evaluateValue('=AVERAGEIF(A1:A2,1,B1:B2)')).toBe(20);
    const err = cells([['0:0', 1, 'number'], ['1:0', 1, 'number'], ['0:1', '=1/0', 'text'], ['1:1', 20, 'number']]);
    expect(err.evaluateValue('=SUMIF(A1:A2,1,B1:B2)')).toBe('#DIV/0!');
  });
});

describe('review regressions – round fifteen', () => {
  it('criteria operands may be dates, percentages and currency', () => {
    const e = new FormulaEngine();
    expect(e.evaluateValue('=COUNTIF({45292;45658},">=1/1/2025")')).toBe(1);
    expect(e.evaluateValue('=COUNTIF({0.1;0.2},"10%")')).toBe(1);
    expect(e.evaluateValue('=COUNTIF({1000;2000},">$1,000")')).toBe(1);
    expect(e.evaluateValue('=SUMIFS({10;20},{45292;45658},">=1/1/2025")')).toBe(20);
  });
  it('SUMIF / AVERAGEIF align a wider result range by coordinates', () => {
    const e = new FormulaEngine();
    expect(e.evaluateValue('=SUMIF({1,1;1,1},">0",{10,20,30;40,50,60;70,80,90})')).toBe(120);
    expect(e.evaluateValue('=AVERAGEIF({1,1;1,1},">0",{10,20,30;40,50,60;70,80,90})')).toBe(30);
  });
});

describe('review regressions – round sixteen', () => {
  it('DOLLARDE keeps a fractional numerator', () => {
    expect(ev('=DOLLARDE(99.015,32)')).toBe(99.046875);
    expect(ev('=DOLLARDE(DOLLARFR(1.15625,16),16)') as number).toBeCloseTo(1.15625, 12);
    expect(ev('=DOLLARDE(1.02,16)')).toBe(1.125);
  });
});

describe('review regressions – round seventeen', () => {
  it('RATE of a loan repaid exactly is 0', () => {
    expect(ev('=RATE(12,-1000,10000,2000)')).toBe(0);
  });
  it('T and N read a reference range as its first cell', () => {
    const e = new FormulaEngine();
    e.setData(new Map([
      ['0:0', { rawValue: '1', displayValue: '1', type: 'number' }], ['1:0', { rawValue: '2', displayValue: '2', type: 'number' }],
      ['0:1', { rawValue: 'a', displayValue: 'a', type: 'text' }], ['1:1', { rawValue: 'b', displayValue: 'b', type: 'text' }],
    ]) as never);
    expect(e.evaluateValue('=SUM(N(A1:A2))')).toBe(1);
    expect(e.evaluateValue('=T(B1:B2)')).toBe('a');
    expect(e.evaluateValue('=N(TRUE)')).toBe(1);
    expect(e.evaluateValue('=T(5)')).toBe('');
  });
  it('FIXED and DOLLAR reject empty text', () => {
    expect(ev('=FIXED("",2)')).toBe('#VALUE!');
    expect(ev('=DOLLAR("",2)')).toBe('#VALUE!');
    expect(ev('=FIXED(1234.567,1)')).toBe('1,234.6');
  });
  it('reverse TEXTBEFORE / TEXTAFTER with match_end use the start as the boundary', () => {
    expect(ev('=TEXTAFTER("a/b/c","/",-1,0,1)')).toBe('c');
    expect(ev('=TEXTBEFORE("a/b/c","/",-1,0,1)')).toBe('a/b');
    expect(ev('=TEXTAFTER("abc","/",-1,0,1)')).toBe('abc');
  });
});

describe('review regressions – round eighteen', () => {
  const grid = (cells: Record<string, string>) => {
    const e = new FormulaEngine();
    const m = new Map();
    for (const [ref, raw] of Object.entries(cells)) {
      m.set(`${Number(ref.slice(1)) - 1}:${ref.charCodeAt(0) - 65}`, { rawValue: raw, displayValue: raw, type: raw !== '' && !isNaN(Number(raw)) ? 'number' : 'text' });
    }
    e.setData(m as never);
    return e;
  };
  it('N / T read only the first cell, so the rest of the range may depend on the formula', () => {
    const e = grid({ A1: '1', B1: '=N(A1:A2)', A2: '=B1+1' });
    e.recalculate();
    expect(e.evaluateValue('=A2')).not.toBe('#CIRC!');
    const data = new Map(
      [['0:0', '1'], ['0:1', '=N(A1:A2)'], ['1:0', '=B1+1']].map(([k, raw]) => [k, { rawValue: raw, displayValue: raw, type: 'text' }])
    );
    const e2 = new FormulaEngine();
    e2.setData(data as never);
    e2.recalculate();
    expect(data.get('1:0')?.displayValue).toBe('2');
  });
  it('RATE with a guess finds the root near the guess', () => {
    expect(ev('=RATE(2,-300,100,500,0,1)') as number).toBeCloseTo(1, 8);
    expect(ev('=RATE(2,-150,100,200,0,-0.4)') as number).toBeCloseTo(-0.5, 8);
  });
  it('LCM treats a blank as 0; MROUND uses the plain quotient', () => {
    expect(grid({ A1: '6' }).evaluateValue('=LCM(A1:A2)')).toBe(0);
    expect(ev('=MROUND(6.05,0.1)')).toBe(6);
    expect(ev('=MROUND(1.5,1)')).toBe(2);
  });
  it('MAXIFS with a single blank criteria cell', () => {
    expect(grid({ A1: '100' }).evaluateValue('=MAXIFS(A1,B1,"")')).toBe(100);
  });
});

describe('review regressions – round nineteen', () => {
  it('a blank lookup value matches a blank entry before 0', () => {
    const e = new FormulaEngine();
    e.setData(new Map([
      ['1:6', { rawValue: '0', displayValue: '0', type: 'number' }],
      ['0:2', { rawValue: '10', displayValue: '10', type: 'number' }],
      ['1:2', { rawValue: '20', displayValue: '20', type: 'number' }],
    ]) as never);
    expect(e.evaluateValue('=XLOOKUP(G5,G1:G2,C1:C2,"missing")')).toBe(10);
    expect(e.evaluateValue('=XMATCH(G5,G1:G2)')).toBe(1);
    expect(e.evaluateValue('=MATCH(G5,{0,1},0)')).toBe(1);
  });
  it('XIRR accepts numeric text; XNPV rejects a blank amount', () => {
    expect(ev('=XIRR({-1000;"300";400;500},{45292;45658;46023;46388})') as number).toBeCloseTo(0.08884315192699434, 8);
    const e = new FormulaEngine();
    e.setData(new Map([
      ['0:0', { rawValue: '-1000', displayValue: '-1000', type: 'number' }],
      ['2:0', { rawValue: '400', displayValue: '400', type: 'number' }],
      ['3:0', { rawValue: '500', displayValue: '500', type: 'number' }],
      ...[45292, 45658, 46023, 46388].map((d, i) => [`${i}:1`, { rawValue: String(d), displayValue: String(d), type: 'number' }] as const),
    ]) as never);
    expect(e.evaluateValue('=XNPV(0.1,A1:A4,B1:B4)')).toBe('#NUM!');
  });
});

describe('review regressions – round twenty', () => {
  const cells = (entries: Array<[string, string]>) => {
    const e = new FormulaEngine();
    e.setData(new Map(entries.map(([k, v]) => [k, { rawValue: v, displayValue: v, type: 'number' }])) as never);
    return e;
  };
  it('classic lookups read a blank lookup value as 0, XLOOKUP/XMATCH keep it blank', () => {
    // A1 blank, A2 = 0, B1 = 10, B2 = 20, Z1 blank
    const e = cells([['1:0', '0'], ['0:1', '10'], ['1:1', '20']]);
    expect(e.evaluateValue('=MATCH(Z1,A1:A2,0)')).toBe(2);
    expect(e.evaluateValue('=VLOOKUP(Z1,A1:B2,2,FALSE)')).toBe(20);
    expect(e.evaluateValue('=XMATCH(Z1,A1:A2)')).toBe(1);
    // no blank entry: XLOOKUP / XMATCH do not fall back to 0
    const z = cells([['0:0', '0'], ['1:0', '1'], ['0:1', '10'], ['1:1', '20']]);
    expect(z.evaluateValue('=XLOOKUP(Z1,A1:A2,B1:B2,"missing")')).toBe('missing');
    expect(z.evaluateValue('=XMATCH(Z1,A1:A2)')).toBe('#N/A');
  });
  it('XNPV is strict about text; XIRR reads numeric text and date text', () => {
    expect(ev('=XNPV(0.1,{-1000;"300";400;500},{45292;45658;46023;46388})')).toBe('#NUM!');
    expect(ev('=XNPV(0.1,{-1000;1100},{45292;"45658"})')).toBe('#NUM!');
    expect(ev('=XIRR({-1000;1100},{"2024-01-01";"2025-01-01"})') as number).toBeCloseTo(0.0997135818004608, 8);
    expect(ev('=XIRR({-1000;1100},{"1/1/2024";"1/1/2025"})') as number).toBeCloseTo(0.0997135818004608, 8);
  });
});

describe('review regressions – round twenty-one', () => {
  it('a blank lookup value still takes the approximate fallback', () => {
    const e = new FormulaEngine();
    e.setData(new Map([0, 1, 2, 3].flatMap((v, r) => [
      [`${r}:0`, { rawValue: String(v), displayValue: String(v), type: 'number' }],
      [`${r}:1`, { rawValue: String((r + 1) * 10), displayValue: String((r + 1) * 10), type: 'number' }],
    ])) as never);
    expect(e.evaluateValue('=XLOOKUP(Z1,A1:A4,B1:B4,"missing",-1,1)')).toBe(40);
    expect(e.evaluateValue('=XMATCH(Z1,A1:A4,-1,1)')).toBe(4);
    expect(e.evaluateValue('=XMATCH(Z1,A1:A4)')).toBe('#N/A');
  });
  it('XIRR / FIXED / DOLLAR read formatted number text', () => {
    expect(ev('=XIRR({-1000;"$300";400;500},{45292;45658;46023;46388})') as number).toBeCloseTo(0.08884315192699434, 8);
    expect(ev('=XIRR({-1000;"1,100"},{45292;45658})') as number).toBeCloseTo(ev('=XIRR({-1000;1100},{45292;45658})') as number, 12);
    expect(ev('=FIXED("12%",2)')).toBe('0.12');
    expect(ev('=FIXED("1,200.50",2)')).toBe('1,200.50');
    expect(ev('=DOLLAR("$1,200.50",2)')).toBe('$1,200.50');
  });
});

describe('review regressions – round twenty-two', () => {
  it('N / T inside BYROW / BYCOL read the first cell of each slice', () => {
    const e = new FormulaEngine();
    e.setData(new Map([[0, 0, '1'], [0, 1, '2'], [1, 0, '3'], [1, 1, '4']].map(([r, c, v]) => [`${r}:${c}`, { rawValue: String(v), displayValue: String(v), type: 'number' }])) as never);
    expect(e.evaluateValue('=SUM(BYROW(A1:B2,LAMBDA(rr,N(rr))))')).toBe(4);
    expect(e.evaluateValue('=SUM(BYCOL(A1:B2,LAMBDA(rr,N(rr))))')).toBe(3);
  });
  it('QUOTIENT uses the plain quotient', () => {
    expect(ev('=QUOTIENT(0.3,0.1)')).toBe(2);
    expect(ev('=QUOTIENT(7,2)')).toBe(3);
  });
});

describe('review regressions – round twenty-three', () => {
  it('N / T go element by element over computed arrays, first cell over references', () => {
    const e = new FormulaEngine();
    e.setData(new Map([[0, 0, '1'], [0, 1, '2'], [1, 0, '3'], [1, 1, '4'], [2, 0, '5'], [3, 0, '6'], [4, 0, '7']].map(([r, c, v]) => [`${r}:${c}`, { rawValue: String(v), displayValue: String(v), type: 'number' }])) as never);
    expect(e.evaluateValue('=SUMPRODUCT(N(A1:A5>2))')).toBe(4);
    expect(e.evaluateValue('=LET(vals,SEQUENCE(5),SUM(N(vals>2)))')).toBe(3);
    expect(e.evaluateValue('=SUM(BYROW({1,2;3,4},LAMBDA(rr,SUM(N(rr)))))')).toBe(10);
    expect(e.evaluateValue('=TEXTJOIN(",",TRUE,T({"a","b","c"}))')).toBe('a,b,c');
    expect(e.evaluateValue('=SUM(N(IF({TRUE,FALSE},1,1/0)))')).toBe('#DIV/0!');
    // over a reference range, each row / column slice is a reference
    expect(e.evaluateValue('=SUM(BYROW(A1:B2,LAMBDA(rr,N(rr))))')).toBe(4);
    expect(e.evaluateValue('=SUM(BYCOL(A1:B2,LAMBDA(rr,N(rr))))')).toBe(3);
    expect(e.evaluateValue('=SUM(BYROW(A1:B2,LAMBDA(rr,SUM(rr))))')).toBe(10);
    expect(e.evaluateValue('=SUM(BYROW(A1:B2,LAMBDA(rr,COLUMNS(rr))))')).toBe(4);
  });
});

describe('review regressions – round twenty-four', () => {
  it('BYROW / BYCOL over a reference do not read cells the lambda does not use', () => {
    const data = new Map(
      [['0:0', '1'], ['0:1', '=C1+1'], ['1:0', '3'], ['1:1', '4'], ['0:2', '=SUM(BYROW(A1:B2,LAMBDA(rr,N(rr))))']].map(([k, raw]) => [k, { rawValue: raw, displayValue: raw, type: 'text' }])
    );
    const e = new FormulaEngine();
    e.setData(data as never);
    e.recalculate();
    expect(data.get('0:2')?.displayValue).toBe('4');
    expect(data.get('0:1')?.displayValue).toBe('5');
    data.get('0:0')!.rawValue = '10';
    e.recalculateAffected(['0:0']);
    expect(data.get('0:2')?.displayValue).toBe('13');
    expect(data.get('0:1')?.displayValue).toBe('14');
  });
  it('a single-cell source stays a reference', () => {
    const e = new FormulaEngine();
    e.setData(new Map([['0:0', { rawValue: '=1/0', displayValue: '', type: 'text' }]]) as never);
    expect(e.evaluateValue('=BYROW(A1,LAMBDA(rr,IFERROR(SUM(rr),0)))')).toBe(0);
    expect(e.evaluateValue('=BYROW(A1,LAMBDA(rr,ISREF(rr)))')).toBe(true);
  });
  it('a computed source is evaluated once', () => {
    const e = new FormulaEngine();
    let calls = 0;
    e.registerFunction('SOURCE', () => ++calls);
    e.evaluateValue('=BYROW(SOURCE(),LAMBDA(rr,SUM(rr)))');
    expect(calls).toBe(1);
  });
});
