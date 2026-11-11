// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { CellData, GridData, HistoryChangeDetail } from '../types.js';
import type { Y11nSpreadsheet } from '../y11n-spreadsheet.js';
import '../y11n-spreadsheet.js';

/** The methods that back toolbar buttons: the selection, its format, history, rows and columns. */

function data(entries: Record<string, string | Partial<CellData>>): GridData {
  const map: GridData = new Map();
  for (const [key, value] of Object.entries(entries)) {
    const cell = typeof value === 'string' ? { rawValue: value } : value;
    const rawValue = cell.rawValue ?? '';
    const type = rawValue !== '' && !isNaN(Number(rawValue)) ? 'number' : 'text';
    map.set(key, { displayValue: rawValue, type, ...cell, rawValue } as CellData);
  }
  return map;
}

async function createSpreadsheet(rows = 20, cols = 6): Promise<Y11nSpreadsheet> {
  const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
  el.rows = rows;
  el.cols = cols;
  document.body.appendChild(el);
  await el.updateComplete;
  return el;
}

const announcement = (el: Y11nSpreadsheet) => el.shadowRoot!.querySelector('[role="log"]')!.textContent!.trim();
const raw = (el: Y11nSpreadsheet, key: string) => el.getData().get(key)?.rawValue ?? '';
const range = (r1: number, c1: number, r2: number, c2: number) => ({
  start: { row: r1, col: c1 },
  end: { row: r2, col: c2 },
});

describe('selection and command methods', () => {
  let sheet: Y11nSpreadsheet;

  afterEach(() => sheet?.remove());

  describe('getSelection', () => {
    it('returns the range and the active cell, as copies', async () => {
      sheet = await createSpreadsheet();
      sheet.goTo('B2:C4');
      const selection = sheet.getSelection();
      expect(selection).toEqual({ range: range(1, 1, 3, 2), activeCell: { row: 1, col: 1 } });
      selection.range.start.row = 9;
      expect(sheet.getSelection().range.start.row).toBe(1);
    });
  });

  describe('getSelectionFormat', () => {
    it('is the format every selected cell shares', async () => {
      sheet = await createSpreadsheet();
      const currency = { type: 'currency' as const, decimals: 2 };
      sheet.setData(
        data({
          '0:0': { rawValue: '1', format: { bold: true, textColor: '#ff0000', numberFormat: currency } },
          '0:1': { rawValue: '2', format: { bold: true, italic: true, numberFormat: { ...currency } } },
          '1:0': { rawValue: '3', format: { bold: false, italic: true } },
        })
      );
      sheet.goTo('A1');
      expect(sheet.getSelectionFormat()).toEqual({ bold: true, textColor: '#ff0000', numberFormat: currency });
      sheet.goTo('A1:B1');
      expect(sheet.getSelectionFormat()).toEqual({ bold: true, numberFormat: currency });
      sheet.goTo('A1:A2');
      expect(sheet.getSelectionFormat()).toEqual({});
      // A cell without a format: nothing is shared
      sheet.goTo('A1:C1');
      expect(sheet.getSelectionFormat()).toEqual({});
      sheet.goTo('D4');
      expect(sheet.getSelectionFormat()).toEqual({});
    });

    it('handles a selection larger than the data', async () => {
      sheet = await createSpreadsheet(200, 6);
      sheet.setData(data({ '0:0': { rawValue: 'x', format: { bold: true } } }));
      sheet.goTo('A1:F200');
      expect(sheet.getSelectionFormat()).toEqual({});

      sheet = await createSpreadsheet(2, 1);
      sheet.setData(
        data({
          '0:0': { rawValue: 'x', format: { bold: true } },
          '1:0': { rawValue: '', format: { bold: true } },
        })
      );
      sheet.goTo('A1:A2');
      expect(sheet.getSelectionFormat()).toEqual({ bold: true });
    });

    it('leaves out rows hidden by the filter', async () => {
      sheet = await createSpreadsheet();
      sheet.setData(data({ '0:0': 'Head', '1:0': 'a', '2:0': 'b', '3:0': 'a' }));
      sheet.setFilter(range(0, 0, 3, 0), [{ col: 0, values: ['a'] }]);
      sheet.goTo('A2:A4');
      sheet.formatSelection({ italic: true });
      expect(sheet.getCellFormat('2:0')).toBeUndefined();
      expect(sheet.getSelectionFormat()).toEqual({ italic: true });
    });
  });

  describe('formatSelection and clearSelectionFormat', () => {
    it('format the selection as one undo step and fire format-change as the user', async () => {
      sheet = await createSpreadsheet();
      sheet.setData(data({ '0:0': 'a', '0:1': 'b' }));
      const handler = vi.fn();
      sheet.addEventListener('format-change', handler);
      sheet.goTo('A1:B1');
      sheet.formatSelection({ textAlign: 'center', wrap: true });
      expect(sheet.getCellFormat('0:1')).toEqual({ textAlign: 'center', wrap: true });
      expect(handler.mock.calls[0][0].detail.source).toBe('user');

      sheet.formatSelection({ wrap: undefined });
      expect(sheet.getCellFormat('0:0')).toEqual({ textAlign: 'center' });

      sheet.clearSelectionFormat();
      expect(sheet.getCellFormat('0:0')).toBeUndefined();
      sheet.undo();
      expect(sheet.getCellFormat('0:0')).toEqual({ textAlign: 'center' });
    });

    it('skip read-only cells and are refused in read-only mode', async () => {
      sheet = await createSpreadsheet();
      sheet.setData(data({ '0:0': { rawValue: 'a', readOnly: true }, '0:1': 'b' }));
      sheet.goTo('A1:B1');
      sheet.formatSelection({ bold: true });
      await sheet.updateComplete;
      expect(sheet.getCellFormat('0:0')).toBeUndefined();
      expect(sheet.getCellFormat('0:1')).toEqual({ bold: true });
      expect(announcement(sheet)).toContain('1 read-only cell was not changed');

      sheet.readOnly = true;
      sheet.clearSelectionFormat();
      sheet.formatSelection({ italic: true });
      expect(sheet.getCellFormat('0:1')).toEqual({ bold: true });
    });
  });

  describe('changeDecimals', () => {
    it('starts from what a number without a format shows, cell by cell', async () => {
      sheet = await createSpreadsheet();
      sheet.setData(data({ '0:0': '10.5', '1:0': '7', '2:0': 'text' }));
      sheet.goTo('A1:A4');
      sheet.changeDecimals(1);
      expect(sheet.getCellFormat('0:0')?.numberFormat).toEqual({ type: 'number', decimals: 2 });
      expect(sheet.getCellFormat('1:0')?.numberFormat).toEqual({ type: 'number', decimals: 1 });
      expect(sheet.getData().get('1:0')?.displayValue).toBe('7.0');
      // Text and empty cells are left alone
      expect(sheet.getCellFormat('2:0')).toBeUndefined();
      expect(sheet.getCellFormat('3:0')).toBeUndefined();
      await sheet.updateComplete;
      expect(announcement(sheet)).toBe('2 decimal places');

      // One undo step for the whole selection
      sheet.undo();
      expect(sheet.getCellFormat('0:0')).toBeUndefined();
      expect(sheet.getCellFormat('1:0')).toBeUndefined();
    });

    it('keeps each cell\'s kind of format, and stops at zero', async () => {
      sheet = await createSpreadsheet();
      const date = { type: 'date' as const, datePattern: 'YYYY-MM-DD' };
      sheet.setData(
        data({
          '0:0': { rawValue: '0.5', format: { bold: true, numberFormat: { type: 'percent', decimals: 1 } } },
          '0:1': { rawValue: '12', format: { numberFormat: { type: 'currency', decimals: 2 } } },
          '0:2': { rawValue: '46000', format: { numberFormat: date } },
        })
      );
      sheet.goTo('A1:C1');
      sheet.changeDecimals(-1);
      sheet.changeDecimals(-1);
      expect(sheet.getCellFormat('0:0')).toEqual({ bold: true, numberFormat: { type: 'percent', decimals: 0 } });
      expect(sheet.getCellFormat('0:1')?.numberFormat).toEqual({ type: 'currency', decimals: 0 });
      expect(sheet.getCellFormat('0:2')?.numberFormat).toEqual(date);
    });

    it('says so when the selection has no numbers', async () => {
      sheet = await createSpreadsheet();
      sheet.setData(data({ '0:0': 'text' }));
      sheet.goTo('A1');
      sheet.changeDecimals(1);
      await sheet.updateComplete;
      expect(announcement(sheet)).toBe('No numbers in the selection');
      expect(sheet.canUndo).toBe(false);
    });

    it('leaves dates alone', async () => {
      sheet = await createSpreadsheet();
      const date = { type: 'date' as const, datePattern: 'YYYY-MM-DD' };
      sheet.setData(data({ '0:0': { rawValue: '46000', format: { numberFormat: date } } }));
      sheet.goTo('A1');
      sheet.changeDecimals(1);
      expect(sheet.getCellFormat('0:0')?.numberFormat).toEqual(date);
    });
  });

  describe('undo, redo and history-change', () => {
    it('undo and redo report whether they did something', async () => {
      sheet = await createSpreadsheet();
      sheet.setData(data({ '0:0': 'a' }));
      expect(sheet.canUndo).toBe(false);
      expect(sheet.undo()).toBe(false);

      sheet.setCellFormat('0:0', { bold: true });
      expect(sheet.canUndo).toBe(true);
      expect(sheet.canRedo).toBe(false);
      expect(sheet.undo()).toBe(true);
      expect(sheet.getCellFormat('0:0')).toBeUndefined();
      expect(sheet.canRedo).toBe(true);
      expect(sheet.redo()).toBe(true);
      expect(sheet.getCellFormat('0:0')).toEqual({ bold: true });
      expect(sheet.redo()).toBe(false);
    });

    it('are refused in read-only mode', async () => {
      sheet = await createSpreadsheet();
      sheet.setData(data({ '0:0': 'a' }));
      sheet.setCellFormat('0:0', { bold: true });
      sheet.readOnly = true;
      expect(sheet.undo()).toBe(false);
      expect(sheet.canUndo).toBe(true);
    });

    it('history-change fires when undo or redo become available or not', async () => {
      sheet = await createSpreadsheet();
      const states: HistoryChangeDetail[] = [];
      sheet.addEventListener('history-change', (e) => states.push((e as CustomEvent<HistoryChangeDetail>).detail));

      sheet.setCellFormat('0:0', { bold: true });
      await sheet.updateComplete;
      sheet.setCellFormat('0:0', { italic: true }); // no change in what is available
      await sheet.updateComplete;
      sheet.undo();
      await sheet.updateComplete;
      sheet.undo();
      await sheet.updateComplete;
      sheet.redo();
      await sheet.updateComplete;
      sheet.setData(new Map()); // a fresh history
      await sheet.updateComplete;

      expect(states).toEqual([
        { canUndo: true, canRedo: false },
        { canUndo: true, canRedo: true },
        { canUndo: false, canRedo: true },
        { canUndo: true, canRedo: true },
        { canUndo: false, canRedo: false },
      ]);
    });

    it('history-change also fires when new data starts a fresh history', async () => {
      sheet = await createSpreadsheet();
      const handler = vi.fn();
      sheet.addEventListener('history-change', handler);
      sheet.setData(data({ '0:0': { rawValue: 'a', format: { bold: true } } }));
      await sheet.updateComplete;
      expect(handler).toHaveBeenCalledOnce();
      expect(handler.mock.calls[0][0].detail).toEqual({ canUndo: false, canRedo: false });
    });
  });

  describe('rows and columns at the selection', () => {
    const table = () => data({ '0:0': 'a', '1:0': 'b', '2:0': 'c', '0:1': 'x', '0:2': 'y' });

    it('insert as many rows as the selection spans, above or below it', async () => {
      sheet = await createSpreadsheet(10, 4);
      sheet.setData(table());
      sheet.goTo('A2:A3');
      expect(sheet.insertRowsAtSelection()).toBe(true);
      expect(sheet.rows).toBe(12);
      expect([raw(sheet, '0:0'), raw(sheet, '1:0'), raw(sheet, '3:0'), raw(sheet, '4:0')]).toEqual(['a', '', 'b', 'c']);

      sheet.goTo('A1');
      sheet.insertRowsAtSelection('below');
      expect([raw(sheet, '0:0'), raw(sheet, '1:0')]).toEqual(['a', '']);
      expect(sheet.rows).toBe(13);

      sheet.undo();
      sheet.undo();
      expect(sheet.rows).toBe(10);
      expect(raw(sheet, '1:0')).toBe('b');
    });

    it('insert columns left or right of the selection', async () => {
      sheet = await createSpreadsheet(10, 4);
      sheet.setData(table());
      sheet.goTo('B1');
      sheet.insertColumnsAtSelection('right');
      expect([raw(sheet, '0:1'), raw(sheet, '0:2'), raw(sheet, '0:3')]).toEqual(['x', '', 'y']);
      sheet.insertColumnsAtSelection();
      expect([raw(sheet, '0:1'), raw(sheet, '0:2')]).toEqual(['', 'x']);
      expect(sheet.cols).toBe(6);
    });

    it('delete the rows or columns the selection spans, cancelably', async () => {
      sheet = await createSpreadsheet(10, 4);
      sheet.setData(table());
      sheet.goTo('A1:B2');
      expect(sheet.deleteSelectedRows()).toBe(true);
      expect(raw(sheet, '0:0')).toBe('c');
      expect(sheet.rows).toBe(8);

      sheet.undo();
      sheet.goTo('B1');
      expect(sheet.deleteSelectedColumns()).toBe(true);
      expect(raw(sheet, '0:1')).toBe('y');

      sheet.addEventListener('before-structure-change', (e) => e.preventDefault());
      expect(sheet.deleteSelectedColumns()).toBe(false);
      expect(sheet.cols).toBe(3);
    });

    it('refuse to delete everything, and anything in read-only mode', async () => {
      sheet = await createSpreadsheet(3, 3);
      sheet.setData(table());
      sheet.goTo('A1:A3');
      expect(sheet.deleteSelectedRows()).toBe(false);
      await sheet.updateComplete;
      expect(announcement(sheet)).toBe('Cannot delete every row');

      sheet.readOnly = true;
      sheet.goTo('A1');
      expect(sheet.insertRowsAtSelection()).toBe(false);
      expect(sheet.deleteSelectedColumns()).toBe(false);
      expect(sheet.rows).toBe(3);
    });
  });

  describe('other commands', () => {
    it('autoSum writes a SUM under the selected numbers', async () => {
      sheet = await createSpreadsheet();
      sheet.setData(data({ '0:0': '1', '1:0': '2' }));
      sheet.goTo('A1:A2');
      sheet.autoSum();
      expect(raw(sheet, '2:0')).toBe('=SUM(A1:A2)');
    });

    it('toggleLists turns text into lists and back', async () => {
      sheet = await createSpreadsheet();
      sheet.setData(data({ '0:0': 'Ann, Bob' }));
      sheet.goTo('A1');
      sheet.toggleLists();
      expect(sheet.getCellValues('0:0')?.values).toEqual(['Ann', 'Bob']);
      sheet.toggleLists();
      expect(sheet.getCellValues('0:0')).toBeNull();
      expect(raw(sheet, '0:0')).toBe('Ann, Bob');
    });

    it('freezePanesAsUser freezes as the user and announces it', async () => {
      sheet = await createSpreadsheet();
      const handler = vi.fn();
      sheet.addEventListener('freeze-change', handler);
      expect(sheet.freezePanesAsUser(1, 0)).toBe(true);
      await sheet.updateComplete;
      expect([sheet.frozenRows, sheet.frozenCols]).toEqual([1, 0]);
      expect(handler.mock.calls[0][0].detail).toEqual({ rows: 1, cols: 0, source: 'user' });
      expect(announcement(sheet)).toBe('Froze row 1');

      sheet.goTo('C3');
      sheet.freezePanesAsUser();
      expect([sheet.frozenRows, sheet.frozenCols]).toEqual([2, 2]);

      expect(sheet.freezePanesAsUser(0, 0)).toBe(true);
      await sheet.updateComplete;
      expect(announcement(sheet)).toBe('Panes unfrozen');
      expect(sheet.freezePanesAsUser(0, 0)).toBe(false);
      expect(() => sheet.freezePanesAsUser(-1)).toThrow(RangeError);
    });

    it('reapplyFilterAsUser takes in rows added under the table and can be cancelled', async () => {
      sheet = await createSpreadsheet();
      sheet.setData(data({ '0:0': 'Head', '1:0': 'a', '2:0': 'b', '3:0': 'a' }));
      sheet.setFilter(range(0, 0, 3, 0), [{ col: 0, values: ['a'] }]);
      const next = sheet.getData();
      next.set('4:0', { rawValue: 'b', displayValue: 'b', type: 'text' });
      sheet.setData(next);
      expect(sheet.isRowHidden(4)).toBe(false);

      const veto = (e: Event) => e.preventDefault();
      sheet.addEventListener('before-filter', veto);
      sheet.reapplyFilterAsUser();
      expect(sheet.isRowHidden(4)).toBe(false);
      sheet.removeEventListener('before-filter', veto);

      sheet.reapplyFilterAsUser();
      expect(sheet.isRowHidden(4)).toBe(true);
      expect(sheet.getFilter()?.range.end.row).toBe(4);
    });

    it('changes made while the focus is outside the grid leave it there', async () => {
      sheet = await createSpreadsheet();
      sheet.setData(data({ '0:0': 'a' }));
      const button = document.createElement('button');
      document.body.append(button);
      button.focus();
      sheet.goTo('A1');
      sheet.formatSelection({ bold: true });
      sheet.undo();
      sheet.insertRowsAtSelection();
      await sheet.updateComplete;
      await Promise.resolve();
      expect(document.activeElement).toBe(button);
      button.remove();
    });

    it('announce speaks through the live region', async () => {
      sheet = await createSpreadsheet();
      sheet.announce('Saved');
      await sheet.updateComplete;
      expect(announcement(sheet)).toBe('Saved');
    });
  });
});
