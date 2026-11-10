import { describe, expect, it } from 'vitest';
import {
  adjustFormulaForStructureChange as adjust,
  formatReference,
  mapFormulaReferences,
  shiftFormulaReferences as shift,
  type StructureChange,
} from '../references.js';
import type { RefInfo } from '../../engine/parser.js';
import { FormulaEngine } from '../../engine/formula-engine.js';

const insertRows = (index: number, count = 1): StructureChange => ({ axis: 'row', kind: 'insert', index, count });
const deleteRows = (index: number, count = 1): StructureChange => ({ axis: 'row', kind: 'delete', index, count });
const insertCols = (index: number, count = 1): StructureChange => ({ axis: 'col', kind: 'insert', index, count });
const deleteCols = (index: number, count = 1): StructureChange => ({ axis: 'col', kind: 'delete', index, count });

describe('mapFormulaReferences', () => {
  it('visits every reference with its sheet prefix and text', () => {
    const seen: Array<{ kind: string; sheet: string | undefined; prefix: string; text: string }> = [];
    const formula = "=A1+Sheet2!B2:C3+'x y'!A:A+1:2+\"A1\"";
    const result = mapFormulaReferences(formula, (ref, text) => {
      seen.push({ kind: ref.kind, sheet: ref.sheet, prefix: text.sheet, text: text.ref });
      return null;
    });
    expect(result).toBe(formula);
    expect(seen).toEqual([
      { kind: 'cell', sheet: undefined, prefix: '', text: 'A1' },
      { kind: 'range', sheet: 'Sheet2', prefix: 'Sheet2!', text: 'B2:C3' },
      { kind: 'cols', sheet: 'x y', prefix: "'x y'!", text: 'A:A' },
      { kind: 'rows', sheet: undefined, prefix: '', text: '1:2' },
    ]);
  });

  it('replaces only the references the callback returns text for', () => {
    expect(mapFormulaReferences('=A1+B2', (ref) => (ref.start.col === 0 ? 'Z9' : null))).toBe('=Z9+B2');
    expect(mapFormulaReferences('= A1 * ( B2 ) ', () => 'X')).toBe('= X * ( X ) ');
    // The replacement stands for the whole reference, prefix included
    expect(mapFormulaReferences('=Sheet2!A1', (_ref, text) => text.sheet + 'B5')).toBe('=Sheet2!B5');
    expect(mapFormulaReferences('=Sheet2!A1', () => 'B5')).toBe('=B5');
  });

  it('returns non-formulas and formulas that do not tokenize unchanged', () => {
    expect(mapFormulaReferences('A1', () => 'x')).toBe('A1');
    expect(mapFormulaReferences('=A1+"', () => 'x')).toBe('=A1+"');
    expect(mapFormulaReferences('=A1 ~ B1', () => 'x')).toBe('=A1 ~ B1');
  });

  it('does not treat names, functions or booleans as references', () => {
    const texts: string[] = [];
    mapFormulaReferences('=LOG10(A2)+Total+TRUE+SUM(B1)+x1y', (_ref, text) => {
      texts.push(text.ref);
      return null;
    });
    expect(texts).toEqual(['A2', 'B1']);
  });
});

describe('formatReference', () => {
  const end = (row: number, col: number, rowAbs = false, colAbs = false) => ({ row, col, rowAbs, colAbs });

  it('writes cells, ranges, whole columns and whole rows', () => {
    const cell: RefInfo = { kind: 'cell', start: end(4, 1), end: end(4, 1) };
    const range: RefInfo = { kind: 'range', start: end(0, 0, true, true), end: end(2, 27, true, false) };
    const cols: RefInfo = { kind: 'cols', start: end(-1, 0, false, true), end: end(-1, 2) };
    const rows: RefInfo = { kind: 'rows', start: end(0, -1), end: end(9, -1, true) };
    expect(formatReference(cell)).toBe('B5');
    expect(formatReference(range)).toBe('$A$1:AB$3');
    expect(formatReference(cols)).toBe('$A:C');
    expect(formatReference(rows)).toBe('1:$10');
  });

  it('never writes the sheet name', () => {
    expect(formatReference({ kind: 'cell', sheet: 'Sheet2', start: end(0, 0), end: end(0, 0) })).toBe('A1');
  });
});

describe('shiftFormulaReferences – whole columns and rows', () => {
  it('shifts whole-column references sideways only', () => {
    expect(shift('=SUM(A:A)', 0, 1)).toBe('=SUM(B:B)');
    expect(shift('=SUM(A:A)', 5, 0)).toBe('=SUM(A:A)');
    expect(shift('=SUM(A:A)', 3, 2)).toBe('=SUM(C:C)');
    expect(shift('=SUM(B:D)', 0, -1)).toBe('=SUM(A:C)');
  });

  it('shifts whole-row references up and down only', () => {
    expect(shift('=SUM(1:3)', 2, 0)).toBe('=SUM(3:5)');
    expect(shift('=SUM(1:3)', 0, 7)).toBe('=SUM(1:3)');
    expect(shift('=SUM(2:3)', -1, 4)).toBe('=SUM(1:2)');
  });

  it('keeps $-anchored ends in place', () => {
    expect(shift('=SUM($A:B)', 0, 2)).toBe('=SUM($A:D)');
    expect(shift('=SUM($A:$B)', 0, 2)).toBe('=SUM($A:$B)');
    expect(shift('=SUM($1:3)', 2, 0)).toBe('=SUM($1:5)');
    expect(shift('=SUM($1:$3)', 2, 0)).toBe('=SUM($1:$3)');
  });

  it('turns references pushed off the sheet into #REF!', () => {
    expect(shift('=SUM(A:A)', 0, -1)).toBe('=SUM(#REF!)');
    expect(shift('=SUM(1:1)', -1, 0)).toBe('=SUM(#REF!)');
    expect(shift('=XFD1', 0, 1)).toBe('=#REF!');
    expect(shift('=A1048576', 1, 0)).toBe('=#REF!');
    expect(shift('=XFD:XFD', 0, 1)).toBe('=#REF!');
  });
});

describe('shiftFormulaReferences – sheet-qualified references', () => {
  it('shifts the reference and keeps the sheet prefix as written', () => {
    expect(shift('=Sheet2!A1', 1, 1)).toBe('=Sheet2!B2');
    expect(shift("='My sheet'!A1:B2", 1, 1)).toBe("='My sheet'!B2:C3");
    expect(shift("='It''s'!A1", 1, 0)).toBe("='It''s'!A2");
    expect(shift('=sheet2!a1', 1, 1)).toBe('=sheet2!B2');
    expect(shift('=Sheet2!A:A', 0, 1)).toBe('=Sheet2!B:B');
    expect(shift('=Sheet2!1:1', 1, 0)).toBe('=Sheet2!2:2');
  });

  it('keeps absolute sheet-qualified references unchanged', () => {
    expect(shift("='My sheet'!$A$1", 5, 5)).toBe("='My sheet'!$A$1");
    expect(shift('=Sheet2!$A1+Sheet2!A$1', 1, 1)).toBe('=Sheet2!$A2+Sheet2!B$1');
  });

  it('handles several sheets in one formula', () => {
    expect(shift("=Sheet2!A1+'My sheet'!B2+C3", 1, 1)).toBe("=Sheet2!B2+'My sheet'!C3+D4");
  });

  it('a sheet-qualified reference pushed off the sheet becomes a plain #REF!', () => {
    expect(shift('=Sheet2!A1', -1, 0)).toBe('=#REF!');
    expect(shift('=Sheet2!A1+Sheet2!B2', -1, 0)).toBe('=#REF!+Sheet2!B1');
  });

  it('leaves an error after a sheet name alone', () => {
    expect(shift('=Sheet2!#REF!+A1', 1, 0)).toBe('=Sheet2!#REF!+A2');
  });
});

describe('shiftFormulaReferences – what is not a reference', () => {
  it('leaves references inside strings untouched', () => {
    expect(shift('="A1"&A1', 1, 0)).toBe('="A1"&A2');
    expect(shift('=IF(A1,"B2",C3)', 1, 1)).toBe('=IF(B2,"B2",D4)');
    expect(shift('="Sheet2!A1 and ""B2"""', 3, 3)).toBe('="Sheet2!A1 and ""B2"""');
  });

  it('leaves function names that look like cells untouched', () => {
    expect(shift('=LOG10(A1)', 1, 0)).toBe('=LOG10(A2)');
    expect(shift('=LOG10(5)+A10', 1, 0)).toBe('=LOG10(5)+A11');
    expect(shift('=ATAN2(B1,C1)', 0, 1)).toBe('=ATAN2(C1,D1)');
  });

  it('leaves names, booleans and LET variables untouched', () => {
    expect(shift('=TRUE+A1', 1, 0)).toBe('=TRUE+A2');
    expect(shift('=Total+A1', 1, 0)).toBe('=Total+A2');
    expect(shift('=LET(x,A1,x+B1)', 1, 0)).toBe('=LET(x,A2,x+B2)');
  });

  it('shifts the cell of a spill reference and keeps the #', () => {
    expect(shift('=A1#', 1, 1)).toBe('=B2#');
    expect(shift('=SUM(A1#)+B2', 2, 0)).toBe('=SUM(A3#)+B4');
    expect(shift('=$A$1#', 2, 2)).toBe('=$A$1#');
  });

  it('shifts both sides of a range made with an expression', () => {
    expect(shift('=A1:INDEX(B:B,2)', 1, 1)).toBe('=B2:INDEX(C:C,2)');
  });

  it('preserves spacing and still rewrites formulas that do not parse', () => {
    expect(shift('= A1 + B2 ', 1, 1)).toBe('= B2 + C3 ');
    expect(shift('=A1+', 1, 1)).toBe('=B2+');
  });
});

describe('adjustFormulaForStructureChange – whole columns and rows', () => {
  it('whole-column references follow column changes', () => {
    expect(adjust('=SUM(A:A)', insertCols(0))).toBe('=SUM(B:B)');
    expect(adjust('=SUM(B:D)', deleteCols(0))).toBe('=SUM(A:C)');
    expect(adjust('=SUM($A:$B)', insertCols(0))).toBe('=SUM($B:$C)');
  });

  it('whole-column references grow and shrink with the columns inside them', () => {
    expect(adjust('=SUM(B:C)', insertCols(2))).toBe('=SUM(B:D)');
    expect(adjust('=SUM(B:C)', deleteCols(1))).toBe('=SUM(B:B)');
    expect(adjust('=SUM(B:C)', insertCols(3))).toBe('=SUM(B:C)');
  });

  it('a whole-column reference whose columns are all deleted becomes #REF!', () => {
    expect(adjust('=SUM(B:C)', deleteCols(1, 2))).toBe('=SUM(#REF!)');
    expect(adjust('=SUM(B:B)+1', deleteCols(1))).toBe('=SUM(#REF!)+1');
  });

  it('whole-column references ignore row changes', () => {
    expect(adjust('=SUM(A:A)', insertRows(0))).toBe('=SUM(A:A)');
    expect(adjust('=SUM(A:C)', deleteRows(0, 100))).toBe('=SUM(A:C)');
  });

  it('whole-row references follow row changes', () => {
    expect(adjust('=SUM(1:3)', insertRows(1))).toBe('=SUM(1:4)');
    expect(adjust('=SUM(2:4)', deleteRows(0))).toBe('=SUM(1:3)');
    expect(adjust('=SUM($2:$3)', deleteRows(2))).toBe('=SUM($2:$2)');
    expect(adjust('=SUM(1:3)', deleteRows(0, 3))).toBe('=SUM(#REF!)');
  });

  it('whole-row references ignore column changes', () => {
    expect(adjust('=SUM(1:3)', insertCols(0))).toBe('=SUM(1:3)');
    expect(adjust('=SUM(1:3)', deleteCols(0, 5))).toBe('=SUM(1:3)');
  });

  it('adjusts whole references next to ordinary ones', () => {
    expect(adjust('=SUM(A:A)+B5+SUM(3:3)', insertRows(2))).toBe('=SUM(A:A)+B6+SUM(4:4)');
    expect(adjust('=SUM(A1:INDEX(A:A,3))', insertRows(0))).toBe('=SUM(A2:INDEX(A:A,3))');
  });
});

describe('adjustFormulaForStructureChange – sheets and scope', () => {
  const onSheet1 = { formulaSheet: 'Sheet1' };

  it('without a scope, sheet-qualified references are left alone', () => {
    expect(adjust('=Sheet2!A1', insertRows(0))).toBe('=Sheet2!A1');
    expect(adjust('=Sheet2!A1+A1', insertRows(0))).toBe('=Sheet2!A1+A2');
  });

  it('by default the formula’s own sheet is the one that changed', () => {
    expect(adjust('=Sheet2!A1+A1', insertRows(0), onSheet1)).toBe('=Sheet2!A1+A2');
    // A reference that names the formula's own sheet moves too
    expect(adjust('=Sheet1!A1+A1', insertRows(0), onSheet1)).toBe('=Sheet1!A2+A2');
  });

  it('only references to the changed sheet move', () => {
    const scope = { formulaSheet: 'Sheet1', changedSheet: 'Sheet2' };
    expect(adjust('=Sheet2!A1+A1', insertRows(0), scope)).toBe('=Sheet2!A2+A1');
    expect(adjust('=A1+B2', insertRows(0), scope)).toBe('=A1+B2');
    expect(adjust("=Sheet2!A1+'Other sheet'!A1", insertRows(0), scope)).toBe("=Sheet2!A2+'Other sheet'!A1");
  });

  it('unqualified references belong to the formula’s sheet', () => {
    const onChanged = { formulaSheet: 'Sheet2', changedSheet: 'Sheet2' };
    expect(adjust('=Sheet2!A1+A1', insertRows(0), onChanged)).toBe('=Sheet2!A2+A2');
    const elsewhere = { formulaSheet: 'Sheet2', changedSheet: 'Sheet1' };
    expect(adjust('=Sheet2!A1+A1', insertRows(0), elsewhere)).toBe('=Sheet2!A1+A1');
    expect(adjust('=Sheet1!A1+A1', insertRows(0), elsewhere)).toBe('=Sheet1!A2+A1');
  });

  it('with only changedSheet given, unqualified references do not move', () => {
    expect(adjust('=A1', insertRows(0), { changedSheet: 'Sheet2' })).toBe('=A1');
    expect(adjust('=Sheet2!A1+A1', insertRows(0), { changedSheet: 'Sheet2' })).toBe('=Sheet2!A2+A1');
  });

  it('compares sheet names ignoring case', () => {
    expect(adjust('=sheet2!A1+A1', insertRows(0), { formulaSheet: 'Sheet1', changedSheet: 'SHEET2' })).toBe(
      '=sheet2!A2+A1'
    );
    expect(adjust('=SHEET1!A1', insertRows(0), { formulaSheet: 'sheet1' })).toBe('=SHEET1!A2');
  });

  it('keeps the sheet prefix exactly as written', () => {
    const scope = { formulaSheet: 'Sheet1', changedSheet: 'My sheet' };
    expect(adjust("='My sheet'!A1:B2+A1", insertRows(1), scope)).toBe("='My sheet'!A1:B3+A1");
    expect(adjust("= 'My sheet'!$A$1", insertCols(0), scope)).toBe("= 'My sheet'!$B$1");
    expect(adjust("='my SHEET'!C3", deleteCols(0), scope)).toBe("='my SHEET'!B3");
    expect(adjust("='It''s'!A1", insertRows(0), { formulaSheet: 'Sheet1', changedSheet: "It's" })).toBe("='It''s'!A2");
  });

  it('adjusts sheet-qualified whole columns and rows', () => {
    const scope = { formulaSheet: 'Sheet1', changedSheet: 'Sheet2' };
    expect(adjust('=SUM(Sheet2!A:A)', insertCols(0), scope)).toBe('=SUM(Sheet2!B:B)');
    expect(adjust('=SUM(Sheet2!2:4)', deleteRows(0), scope)).toBe('=SUM(Sheet2!1:3)');
    expect(adjust('=SUM(Sheet2!A:A)+SUM(A:A)', insertCols(0), scope)).toBe('=SUM(Sheet2!B:B)+SUM(A:A)');
  });

  it('a deleted sheet-qualified reference becomes a plain #REF!', () => {
    const scope = { formulaSheet: 'Sheet1', changedSheet: 'Sheet2' };
    expect(adjust('=Sheet2!A1+1', deleteRows(0), scope)).toBe('=#REF!+1');
    expect(adjust("='My sheet'!A1:B2", deleteRows(0, 2), { formulaSheet: 'Sheet1', changedSheet: 'My sheet' })).toBe(
      '=#REF!'
    );
    expect(adjust('=SUM(Sheet2!2:2)', deleteRows(1), scope)).toBe('=SUM(#REF!)');
  });
});

describe('adjustFormulaForStructureChange – what is not a reference', () => {
  it('leaves references inside strings untouched', () => {
    expect(adjust('="A2"&A2', insertRows(0))).toBe('="A2"&A3');
    expect(adjust('=IF(A2="A2","Sheet2!A2",B2)', insertRows(0))).toBe('=IF(A3="A2","Sheet2!A2",B3)');
  });

  it('leaves function names that look like cells untouched', () => {
    expect(adjust('=LOG10(A2)', insertRows(0))).toBe('=LOG10(A3)');
    expect(adjust('=LOG10(5)', insertRows(0))).toBe('=LOG10(5)');
    expect(adjust('=LOG10(5)', deleteRows(9))).toBe('=LOG10(5)');
    expect(adjust('=LOG10(B1)', deleteCols(0))).toBe('=LOG10(A1)');
  });

  it('moves the cell of a spill reference and keeps the #', () => {
    expect(adjust('=A2#', insertRows(0))).toBe('=A3#');
    expect(adjust('=SUM(C5#)', deleteCols(0))).toBe('=SUM(B5#)');
    expect(adjust('=SUM(A2#)', insertRows(5))).toBe('=SUM(A2#)');
  });

  it('rewrites formulas that do not parse, as long as they tokenize', () => {
    expect(adjust('=A1+', insertRows(0))).toBe('=A2+');
    expect(adjust('=A1+"', insertRows(0))).toBe('=A1+"');
  });
});

describe('rewritten formulas still evaluate', () => {
  const engine = new FormulaEngine();

  it('a whole-column reference deleted entirely evaluates to #REF!', () => {
    expect(engine.evaluate(adjust('=SUM(B:C)', deleteCols(1, 2)))).toEqual({ displayValue: '#REF!', type: 'error' });
    expect(engine.evaluate(shift('=Sheet2!A1+1', -1, 0))).toEqual({ displayValue: '#REF!', type: 'error' });
  });

  it('a deleted spill reference evaluates to #REF!', () => {
    const rewritten = adjust('=SUM(A2#)', deleteRows(1));
    expect(engine.evaluate(rewritten)).toEqual({ displayValue: '#REF!', type: 'error' });
  });
});
