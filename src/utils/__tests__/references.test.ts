import { describe, expect, it } from 'vitest';
import {
  adjustFormulaForStructureChange as adjust,
  shiftFormulaReferences as shift,
  type StructureChange,
} from '../references.js';

const insertRows = (index: number, count = 1): StructureChange => ({ axis: 'row', kind: 'insert', index, count });
const deleteRows = (index: number, count = 1): StructureChange => ({ axis: 'row', kind: 'delete', index, count });
const insertCols = (index: number, count = 1): StructureChange => ({ axis: 'col', kind: 'insert', index, count });
const deleteCols = (index: number, count = 1): StructureChange => ({ axis: 'col', kind: 'delete', index, count });

describe('adjustFormulaForStructureChange', () => {
  it('leaves non-formulas alone', () => {
    expect(adjust('A1', insertRows(0))).toBe('A1');
    expect(adjust('hello A5', deleteRows(0))).toBe('hello A5');
  });

  describe('inserting rows', () => {
    it('moves references at or below the insertion point', () => {
      // index 1 = before row 2
      expect(adjust('=A1+A2+A3', insertRows(1))).toBe('=A1+A3+A4');
      expect(adjust('=B5', insertRows(0, 3))).toBe('=B8');
    });

    it('moves absolute and mixed references too', () => {
      expect(adjust('=$A$2+A$2+$A2', insertRows(0))).toBe('=$A$3+A$3+$A3');
    });

    it('grows a range that spans the insertion point', () => {
      expect(adjust('=SUM(A1:A3)', insertRows(1))).toBe('=SUM(A1:A4)');
      expect(adjust('=SUM(A2:B3)', insertRows(1, 2))).toBe('=SUM(A4:B5)');
      // Inserting right after a range does not grow it
      expect(adjust('=SUM(A1:A3)', insertRows(3))).toBe('=SUM(A1:A3)');
    });

    it('handles reversed ranges', () => {
      expect(adjust('=SUM(A3:A1)', insertRows(1))).toBe('=SUM(A4:A1)');
    });

    it('ignores columns', () => {
      expect(adjust('=Z1', insertRows(5))).toBe('=Z1');
    });
  });

  describe('deleting rows', () => {
    it('moves references below the deletion up and breaks deleted ones', () => {
      // delete row 2
      expect(adjust('=A1+A2+A3', deleteRows(1))).toBe('=A1+#REF!+A2');
    });

    it('shrinks ranges that lose rows', () => {
      expect(adjust('=SUM(A1:A10)', deleteRows(2, 2))).toBe('=SUM(A1:A8)');
      expect(adjust('=SUM(A1:A10)', deleteRows(0, 3))).toBe('=SUM(A1:A7)');
      expect(adjust('=SUM(A1:A10)', deleteRows(8, 5))).toBe('=SUM(A1:A8)');
      expect(adjust('=SUM(A5:A10)', deleteRows(0, 2))).toBe('=SUM(A3:A8)');
    });

    it('breaks a range that is deleted entirely', () => {
      expect(adjust('=SUM(A2:B3)', deleteRows(1, 2))).toBe('=SUM(#REF!)');
      expect(adjust('=SUM(A2:B3)', deleteRows(0, 5))).toBe('=SUM(#REF!)');
    });
  });

  describe('columns', () => {
    it('inserts and deletes columns', () => {
      expect(adjust('=A1+B1+C1', insertCols(1))).toBe('=A1+C1+D1');
      expect(adjust('=SUM(A1:C1)', insertCols(1, 2))).toBe('=SUM(A1:E1)');
      expect(adjust('=A1+B1+C1', deleteCols(1))).toBe('=A1+#REF!+B1');
      expect(adjust('=SUM(A1:D2)', deleteCols(0, 2))).toBe('=SUM(A1:B2)');
      expect(adjust('=Z1', insertCols(0))).toBe('=AA1');
    });
  });

  it('does not touch string literals, function names or non-reference identifiers', () => {
    expect(adjust('=CONCAT("A2", A2)', insertRows(0))).toBe('=CONCAT("A2", A3)');
    expect(adjust('="say ""B2"" "&B2', insertRows(0))).toBe('="say ""B2"" "&B3');
    expect(adjust('=LOG10(A2)', insertRows(0))).toBe('=LOG10(A3)');
    expect(adjust('=MY_A2(1)', insertRows(0))).toBe('=MY_A2(1)');
  });

  it('accepts lowercase references', () => {
    expect(adjust('=sum(a1:a3)', insertRows(0))).toBe('=sum(A2:A4)');
  });
});

describe('shiftFormulaReferences', () => {
  it('shifts relative parts and keeps absolute ones', () => {
    expect(shift('=A1+$B$2+C$3+$D4', 1, 1)).toBe('=B2+$B$2+D$3+$D5');
    expect(shift('=SUM(A1:B2)', 2, 0)).toBe('=SUM(A3:B4)');
    expect(shift('plain', 5, 5)).toBe('plain');
  });

  it('turns a reference, or a whole range with an invalid end, into a parseable #REF!', () => {
    expect(shift('=A1+1', -1, 0)).toBe('=#REF!+1');
    expect(shift('=IFERROR(SUM(A1:A2),0)', -1, 0)).toBe('=IFERROR(SUM(#REF!),0)');
    expect(shift('=SUM(A1:B1)', 0, -1)).toBe('=SUM(#REF!)');
    expect(shift('=SUM($A$1:B1)', 0, -1)).toBe('=SUM($A$1:A1)');
  });

  it('leaves string literals alone', () => {
    expect(shift('="A1"&A1', 1, 0)).toBe('="A1"&A2');
  });
});

