import { describe, it, expect } from 'vitest';
import {
  cellKey,
  parseKey,
  colToLetter,
  letterToCol,
  refToCoord,
  coordToRef,
} from '../types.js';

describe('types utility functions', () => {
  describe('cellKey', () => {
    it('creates a key from row and col', () => {
      expect(cellKey(0, 0)).toBe('0:0');
      expect(cellKey(5, 3)).toBe('5:3');
      expect(cellKey(100, 25)).toBe('100:25');
    });
  });

  describe('parseKey', () => {
    it('parses a key back to coordinates', () => {
      expect(parseKey('0:0')).toEqual({ row: 0, col: 0 });
      expect(parseKey('5:3')).toEqual({ row: 5, col: 3 });
      expect(parseKey('100:25')).toEqual({ row: 100, col: 25 });
    });
  });

  describe('colToLetter', () => {
    it('converts single-letter columns', () => {
      expect(colToLetter(0)).toBe('A');
      expect(colToLetter(1)).toBe('B');
      expect(colToLetter(25)).toBe('Z');
    });

    it('converts multi-letter columns', () => {
      expect(colToLetter(26)).toBe('AA');
      expect(colToLetter(27)).toBe('AB');
      expect(colToLetter(51)).toBe('AZ');
      expect(colToLetter(52)).toBe('BA');
    });
  });

  describe('letterToCol', () => {
    it('converts single-letter columns', () => {
      expect(letterToCol('A')).toBe(0);
      expect(letterToCol('B')).toBe(1);
      expect(letterToCol('Z')).toBe(25);
    });

    it('converts multi-letter columns', () => {
      expect(letterToCol('AA')).toBe(26);
      expect(letterToCol('AB')).toBe(27);
      expect(letterToCol('AZ')).toBe(51);
      expect(letterToCol('BA')).toBe(52);
    });
  });

  describe('colToLetter / letterToCol roundtrip', () => {
    it('converts back and forth correctly', () => {
      for (let i = 0; i < 100; i++) {
        expect(letterToCol(colToLetter(i))).toBe(i);
      }
    });
  });

  describe('refToCoord', () => {
    it('converts cell references to coordinates', () => {
      expect(refToCoord('A1')).toEqual({ row: 0, col: 0 });
      expect(refToCoord('B2')).toEqual({ row: 1, col: 1 });
      expect(refToCoord('Z10')).toEqual({ row: 9, col: 25 });
      expect(refToCoord('AA1')).toEqual({ row: 0, col: 26 });
    });

    it('is case-insensitive', () => {
      expect(refToCoord('a1')).toEqual({ row: 0, col: 0 });
      expect(refToCoord('b2')).toEqual({ row: 1, col: 1 });
    });

    it('throws on invalid reference', () => {
      expect(() => refToCoord('123')).toThrow('Invalid cell reference');
      expect(() => refToCoord('')).toThrow('Invalid cell reference');
    });
  });

  describe('coordToRef', () => {
    it('converts coordinates to cell references', () => {
      expect(coordToRef({ row: 0, col: 0 })).toBe('A1');
      expect(coordToRef({ row: 1, col: 1 })).toBe('B2');
      expect(coordToRef({ row: 9, col: 25 })).toBe('Z10');
      expect(coordToRef({ row: 0, col: 26 })).toBe('AA1');
    });
  });

  describe('refToCoord / coordToRef roundtrip', () => {
    it('converts back and forth correctly', () => {
      const refs = ['A1', 'B2', 'Z10', 'AA1', 'AB27', 'AZ100'];
      for (const ref of refs) {
        expect(coordToRef(refToCoord(ref))).toBe(ref);
      }
    });
  });
});
