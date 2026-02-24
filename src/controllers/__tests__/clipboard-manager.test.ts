import { describe, it, expect, beforeEach } from 'vitest';
import { ClipboardManager } from '../clipboard-manager.js';
import type { GridData, CellData, SelectionRange } from '../../types.js';

function cell(rawValue: string, displayValue?: string): CellData {
  return {
    rawValue,
    displayValue: displayValue ?? rawValue,
    type: 'text',
  };
}

describe('ClipboardManager', () => {
  let manager: ClipboardManager;

  beforeEach(() => {
    manager = new ClipboardManager();
  });

  describe('parseTSV', () => {
    it('parses single-cell TSV', () => {
      const result = manager.parseTSV('hello', 0, 0, 100, 26);
      expect(result).toEqual([{ id: '0:0', value: 'hello' }]);
    });

    it('parses a single row with tabs', () => {
      const result = manager.parseTSV('a\tb\tc', 0, 0, 100, 26);
      expect(result).toEqual([
        { id: '0:0', value: 'a' },
        { id: '0:1', value: 'b' },
        { id: '0:2', value: 'c' },
      ]);
    });

    it('parses multiple rows', () => {
      const result = manager.parseTSV('a\tb\nc\td', 0, 0, 100, 26);
      expect(result).toEqual([
        { id: '0:0', value: 'a' },
        { id: '0:1', value: 'b' },
        { id: '1:0', value: 'c' },
        { id: '1:1', value: 'd' },
      ]);
    });

    it('respects target offset', () => {
      const result = manager.parseTSV('x', 3, 2, 100, 26);
      expect(result).toEqual([{ id: '3:2', value: 'x' }]);
    });

    it('clips to grid bounds', () => {
      const result = manager.parseTSV('a\tb', 0, 0, 100, 1);
      expect(result).toEqual([{ id: '0:0', value: 'a' }]);
    });

    it('handles trailing newline', () => {
      const result = manager.parseTSV('a\nb\n', 0, 0, 100, 26);
      expect(result).toEqual([
        { id: '0:0', value: 'a' },
        { id: '1:0', value: 'b' },
      ]);
    });

    it('handles Windows-style line endings', () => {
      const result = manager.parseTSV('a\r\nb\r\n', 0, 0, 100, 26);
      expect(result).toEqual([
        { id: '0:0', value: 'a' },
        { id: '1:0', value: 'b' },
      ]);
    });

    it('clips rows to grid bounds', () => {
      const result = manager.parseTSV('a\nb\nc', 0, 0, 2, 26);
      expect(result).toEqual([
        { id: '0:0', value: 'a' },
        { id: '1:0', value: 'b' },
      ]);
    });

    it('handles empty input', () => {
      const result = manager.parseTSV('', 0, 0, 100, 26);
      expect(result).toEqual([]);
    });

    it('handles paste at offset near grid edge', () => {
      const result = manager.parseTSV('a\tb\nc\td', 99, 24, 100, 26);
      expect(result).toEqual([
        { id: '99:24', value: 'a' },
        { id: '99:25', value: 'b' },
      ]);
    });

    it('handles bare \\r line endings (old Mac style)', () => {
      // The robust _parseTSVRows parser treats bare \r as a row separator (RFC 4180)
      const result = manager.parseTSV('a\rb\rc', 0, 0, 100, 26);
      expect(result).toEqual([
        { id: '0:0', value: 'a' },
        { id: '1:0', value: 'b' },
        { id: '2:0', value: 'c' },
      ]);
    });

    it('handles multi-row multi-col grid', () => {
      const result = manager.parseTSV('a\tb\tc\n1\t2\t3\nx\ty\tz', 0, 0, 100, 26);
      expect(result).toEqual([
        { id: '0:0', value: 'a' },
        { id: '0:1', value: 'b' },
        { id: '0:2', value: 'c' },
        { id: '1:0', value: '1' },
        { id: '1:1', value: '2' },
        { id: '1:2', value: '3' },
        { id: '2:0', value: 'x' },
        { id: '2:1', value: 'y' },
        { id: '2:2', value: 'z' },
      ]);
    });

    it('handles target offset with multi-row grid', () => {
      const result = manager.parseTSV('a\tb\n1\t2', 2, 3, 100, 26);
      expect(result).toEqual([
        { id: '2:3', value: 'a' },
        { id: '2:4', value: 'b' },
        { id: '3:3', value: '1' },
        { id: '3:4', value: '2' },
      ]);
    });

    it('clips both rows and columns to grid bounds', () => {
      const result = manager.parseTSV('a\tb\tc\n1\t2\t3\nx\ty\tz', 0, 0, 2, 2);
      expect(result).toEqual([
        { id: '0:0', value: 'a' },
        { id: '0:1', value: 'b' },
        { id: '1:0', value: '1' },
        { id: '1:1', value: '2' },
      ]);
    });

    it('handles tab-only input (empty cells)', () => {
      const result = manager.parseTSV('\t\t', 0, 0, 100, 26);
      expect(result).toEqual([
        { id: '0:0', value: '' },
        { id: '0:1', value: '' },
        { id: '0:2', value: '' },
      ]);
    });
  });

  // parseHTMLTable uses DOMParser which is a browser API not available in
  // Node. These code paths are covered by E2E tests (Playwright).

  describe('cut', () => {
    it('returns keys for all cells in range', async () => {
      const data: GridData = new Map();
      data.set('0:0', cell('a', 'a'));
      data.set('0:1', cell('b', 'b'));
      data.set('1:0', cell('c', 'c'));
      data.set('1:1', cell('d', 'd'));

      const range: SelectionRange = {
        start: { row: 0, col: 0 },
        end: { row: 1, col: 1 },
      };

      // Mock clipboard
      const original = globalThis.navigator;
      Object.defineProperty(globalThis, 'navigator', {
        value: {
          clipboard: {
            writeText: async () => {},
            readText: async () => '',
            write: async () => {},
          },
        },
        writable: true,
        configurable: true,
      });

      try {
        const keys = await manager.cut(data, range);
        expect(keys).toEqual(['0:0', '0:1', '1:0', '1:1']);
      } finally {
        Object.defineProperty(globalThis, 'navigator', {
          value: original,
          writable: true,
          configurable: true,
        });
      }
    });
  });

  describe('adjustFormulaReferences', () => {
    it('returns non-formula strings unchanged', () => {
      expect(manager.adjustFormulaReferences('hello', 2, 3)).toBe('hello');
      expect(manager.adjustFormulaReferences('123', 1, 1)).toBe('123');
      expect(manager.adjustFormulaReferences('', 1, 1)).toBe('');
    });

    it('adjusts relative references by offset', () => {
      // A1 (col=0, row=0) + rowOffset=2, colOffset=1 → B3 (col=1, row=2)
      expect(manager.adjustFormulaReferences('=A1', 2, 1)).toBe('=B3');
    });

    it('does not adjust absolute column reference ($A1)', () => {
      // $A1 with colOffset=3 → col stays A, row 1+2=3
      expect(manager.adjustFormulaReferences('=$A1', 2, 3)).toBe('=$A3');
    });

    it('does not adjust absolute row reference (A$1)', () => {
      // A$1 with rowOffset=5 → row stays 1, col A+2=C
      expect(manager.adjustFormulaReferences('=A$1', 5, 2)).toBe('=C$1');
    });

    it('does not adjust fully absolute reference ($A$1)', () => {
      expect(manager.adjustFormulaReferences('=$A$1', 3, 4)).toBe('=$A$1');
    });

    it('adjusts mixed reference $A1 - only row adjusts', () => {
      // $A1: col locked at A, row 1+3=4
      expect(manager.adjustFormulaReferences('=$A1', 3, 5)).toBe('=$A4');
    });

    it('adjusts mixed reference A$1 - only col adjusts', () => {
      // A$1: row locked at 1, col A+2=C
      expect(manager.adjustFormulaReferences('=A$1', 3, 2)).toBe('=C$1');
    });

    it('adjusts references in ranges', () => {
      // SUM(A1:B2) with rowOffset=1, colOffset=1
      // A1 → B2, B2 → C3
      expect(manager.adjustFormulaReferences('=SUM(A1:B2)', 1, 1)).toBe('=SUM(B2:C3)');
    });

    it('produces #REF! when adjustment would go negative', () => {
      // A1 (col=0, row=0) with colOffset=-1 → col would be -1
      expect(manager.adjustFormulaReferences('=A1', 0, -1)).toBe('=#REF!');
      // A1 with rowOffset=-1 → row would be -1
      expect(manager.adjustFormulaReferences('=A1', -1, 0)).toBe('=#REF!');
    });

    it('handles multiple references in one formula', () => {
      // A1+B2 with rowOffset=1, colOffset=1 → B2+C3
      expect(manager.adjustFormulaReferences('=A1+B2', 1, 1)).toBe('=B2+C3');
    });

    it('handles zero offset (no change)', () => {
      expect(manager.adjustFormulaReferences('=A1', 0, 0)).toBe('=A1');
      expect(manager.adjustFormulaReferences('=SUM(A1:B2)', 0, 0)).toBe('=SUM(A1:B2)');
    });
  });
});
