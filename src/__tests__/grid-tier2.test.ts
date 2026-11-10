// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import type {
  BeforeChangeDetail,
  CellData,
  CellEditStartDetail,
  DataChangeDetail,
  GridData,
  SelectionRange,
  StructureChangeDetail,
  ValidationErrorDetail,
} from '../types.js';
import type { Y11nSpreadsheet } from '../y11n-spreadsheet.js';
import '../y11n-spreadsheet.js';

function data(entries: Record<string, string | Partial<CellData>>): GridData {
  const map: GridData = new Map();
  for (const [key, value] of Object.entries(entries)) {
    const cell = typeof value === 'string' ? { rawValue: value } : value;
    const rawValue = cell.rawValue ?? '';
    map.set(key, { displayValue: rawValue, type: 'text', ...cell, rawValue } as CellData);
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

const shadow = (el: Y11nSpreadsheet) => el.shadowRoot!;
const grid = (el: Y11nSpreadsheet) => shadow(el).querySelector('.ls-grid') as HTMLElement;
const editor = (el: Y11nSpreadsheet) => shadow(el).querySelector('#editor') as HTMLInputElement;
const isEditing = (el: Y11nSpreadsheet) => editor(el).style.display !== 'none';
const announcement = (el: Y11nSpreadsheet) =>
  shadow(el).querySelector('[role="log"]')!.textContent!.trim();

async function press(el: Y11nSpreadsheet, key: string, mods: KeyboardEventInit = {}) {
  grid(el).dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, composed: true, ...mods }));
  await el.updateComplete;
}

async function pressInEditor(el: Y11nSpreadsheet, key: string, mods: KeyboardEventInit = {}) {
  editor(el).dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, composed: true, ...mods }));
  await el.updateComplete;
}

async function typeInEditor(el: Y11nSpreadsheet, value: string) {
  editor(el).value = value;
  editor(el).dispatchEvent(new Event('input', { bubbles: true }));
  await el.updateComplete;
}

async function enterValue(el: Y11nSpreadsheet, value: string) {
  await press(el, 'F2');
  await typeInEditor(el, value);
  await pressInEditor(el, 'Enter');
}

const raw = (el: Y11nSpreadsheet, key: string) => el.getData().get(key)?.rawValue ?? '';
const display = (el: Y11nSpreadsheet, key: string) => el.getData().get(key)?.displayValue ?? '';

type Internals = {
  _selection: {
    activeCell: { row: number; col: number };
    range: SelectionRange;
    moveTo(row: number, col: number, extend?: boolean): void;
  };
  _autoFill(source: SelectionRange, target: SelectionRange): void;
  _undo(): void;
  _redo(): void;
};
const internals = (el: Y11nSpreadsheet) => el as unknown as Internals;

async function select(el: Y11nSpreadsheet, r1: number, c1: number, r2 = r1, c2 = c1) {
  internals(el)._selection.moveTo(r1, c1);
  internals(el)._selection.moveTo(r2, c2, true);
  el.requestUpdate();
  await el.updateComplete;
}

const last = <T>(items: T[]): T => items[items.length - 1];

const range = (r1: number, c1: number, r2: number, c2: number): SelectionRange => ({
  start: { row: r1, col: c1 },
  end: { row: r2, col: c2 },
});

describe('tier 2 grid features', () => {
  beforeEach(() => {
    document.body.innerHTML = '';
  });

  describe('read-only cells', () => {
    it('cannot be edited, cleared or formatted by the user, and are exposed as read-only', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': { rawValue: 'locked', readOnly: true }, '0:1': 'free' }));
      await el.updateComplete;

      const cell = shadow(el).querySelector('[data-key="0:0"]')!;
      expect(cell.getAttribute('aria-readonly')).toBe('true');
      expect(cell.classList.contains('read-only-cell')).toBe(true);
      expect(shadow(el).querySelector('[data-key="0:1"]')!.getAttribute('aria-readonly')).toBe('false');

      await press(el, 'x');
      expect(isEditing(el)).toBe(false);
      expect(announcement(el)).toBe('A1 is read-only');

      await select(el, 0, 0, 0, 1);
      await press(el, 'Delete');
      expect(raw(el, '0:0')).toBe('locked');
      expect(raw(el, '0:1')).toBe('');
      expect(announcement(el)).toBe('1 read-only cell was not changed');

      await press(el, 'b', { ctrlKey: true });
      expect(el.getCellFormat('0:0')).toBeUndefined();
      expect(el.getCellFormat('0:1')).toEqual({ bold: true });
    });

    it('are skipped by fill and keep the flag when cleared programmatically', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'src', '1:0': { rawValue: '', readOnly: true } }));
      await el.updateComplete;
      await select(el, 0, 0, 2, 0);
      await press(el, 'd', { ctrlKey: true });
      expect(raw(el, '1:0')).toBe('');
      expect(raw(el, '2:0')).toBe('src');
      expect(el.getData().get('1:0')?.readOnly).toBe(true);

      // A format change on an empty locked cell keeps the entry (and the flag)
      el.setCellFormat('1:0', { bold: true });
      el.clearCellFormat('1:0');
      expect(el.getData().get('1:0')?.readOnly).toBe(true);
    });
  });

  describe('validateCell', () => {
    it('keeps the editor open with a message when Enter commits an invalid value', async () => {
      const el = await createSpreadsheet();
      const errors: ValidationErrorDetail[] = [];
      el.addEventListener('validation-error', (e) => errors.push((e as CustomEvent).detail));
      el.validateCell = (_id, value) => (/^\d+$/.test(value) ? null : 'Enter a whole number');

      await enterValue(el, 'abc');
      expect(isEditing(el)).toBe(true);
      expect(raw(el, '0:0')).toBe('');
      const message = shadow(el).querySelector('#editor-error')!;
      expect(message.textContent).toBe('Enter a whole number');
      expect(editor(el).getAttribute('aria-invalid')).toBe('true');
      expect(editor(el).getAttribute('aria-describedby')).toBe('editor-error');
      expect(announcement(el)).toBe('Invalid value: Enter a whole number');
      expect(errors).toEqual([
        { errors: [{ cellId: '0:0', value: 'abc', message: 'Enter a whole number' }], operation: 'edit' },
      ]);
      expect(internals(el)._selection.activeCell).toEqual({ row: 0, col: 0 });

      // Typing clears the message; a valid value commits and moves down
      await typeInEditor(el, '42');
      expect(shadow(el).querySelector('#editor-error')).toBeNull();
      await pressInEditor(el, 'Enter');
      expect(isEditing(el)).toBe(false);
      expect(raw(el, '0:0')).toBe('42');
      expect(internals(el)._selection.activeCell).toEqual({ row: 1, col: 0 });
    });

    it('cancels an invalid edit when focus moves away, and never validates clearing', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': '5' }));
      el.validateCell = (_id, value) => (Number(value) > 0 ? null : 'Must be positive');
      await el.updateComplete;

      await press(el, 'F2');
      await typeInEditor(el, '-1');
      editor(el).dispatchEvent(new FocusEvent('blur'));
      await el.updateComplete;
      expect(isEditing(el)).toBe(false);
      expect(raw(el, '0:0')).toBe('5');
      expect(announcement(el)).toBe('Invalid value: Must be positive. Edit cancelled');

      await press(el, 'Delete');
      expect(raw(el, '0:0')).toBe('');
    });

    it('drops rejected cells from bulk operations', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': '1', '0:1': 'x' }));
      el.validateCell = (id, value) => (id === '2:0' && value === '1' ? 'No ones in A3' : null);
      await el.updateComplete;
      await select(el, 0, 0, 3, 0);
      await press(el, 'd', { ctrlKey: true });
      expect(raw(el, '1:0')).toBe('1');
      expect(raw(el, '2:0')).toBe('');
      expect(raw(el, '3:0')).toBe('1');
      expect(announcement(el)).toBe('Invalid value: No ones in A3');
    });

    it('validates Ctrl+Enter against the edited cell first', async () => {
      const el = await createSpreadsheet();
      el.validateCell = (_id, value) => (value === 'bad' ? 'Nope' : null);
      await select(el, 0, 0, 1, 0);
      await press(el, 'F2');
      await typeInEditor(el, 'bad');
      await pressInEditor(el, 'Enter', { ctrlKey: true });
      expect(isEditing(el)).toBe(true);
      expect(raw(el, '1:0')).toBe('');
    });
  });

  describe('cell-edit-start', () => {
    it('fires before editing and can be cancelled', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'hi' }));
      await el.updateComplete;
      const seen: CellEditStartDetail[] = [];
      const handler = (e: Event) => {
        seen.push((e as CustomEvent).detail);
        e.preventDefault();
      };
      el.addEventListener('cell-edit-start', handler);
      await press(el, 'Enter');
      expect(isEditing(el)).toBe(false);
      expect(seen).toEqual([{ cellId: '0:0', value: 'hi' }]);

      el.removeEventListener('cell-edit-start', handler);
      await press(el, 'Enter');
      expect(isEditing(el)).toBe(true);
    });
  });

  describe('before-change', () => {
    it('can cancel, rewrite or drop changes, and is not fired for undo', async () => {
      const el = await createSpreadsheet();
      const events: BeforeChangeDetail[] = [];
      let mode: 'cancel' | 'rewrite' | 'drop' | 'none' = 'cancel';
      el.addEventListener('before-change', (e) => {
        const detail = (e as CustomEvent<BeforeChangeDetail>).detail;
        events.push(JSON.parse(JSON.stringify(detail)));
        if (mode === 'cancel') e.preventDefault();
        if (mode === 'rewrite') detail.changes[0].value = detail.changes[0].value.toUpperCase();
        if (mode === 'drop') detail.changes.splice(1, 1);
      });

      await enterValue(el, 'one');
      expect(raw(el, '0:0')).toBe('');
      expect(events[0]).toEqual({
        changes: [{ cellId: '0:0', value: 'one', oldValue: '' }],
        operation: 'edit',
      });
      expect(announcement(el)).toBe('Change cancelled');

      mode = 'rewrite';
      const cellChanges: string[] = [];
      el.addEventListener('cell-change', (e) => cellChanges.push((e as CustomEvent).detail.value));
      await select(el, 0, 0);
      await enterValue(el, 'two');
      expect(raw(el, '0:0')).toBe('TWO');
      expect(cellChanges).toEqual(['TWO']);

      mode = 'drop';
      await select(el, 0, 0, 2, 0);
      await press(el, 'd', { ctrlKey: true });
      expect(last(events).operation).toBe('fill');
      expect(raw(el, '1:0')).toBe('TWO');
      expect(raw(el, '2:0')).toBe('');

      const count = events.length;
      internals(el)._undo();
      await el.updateComplete;
      expect(raw(el, '1:0')).toBe('');
      expect(events.length).toBe(count);
    });
  });

  describe('inserting and deleting rows and columns', () => {
    it('inserts rows, moving cells, formats and references, and undoes it', async () => {
      const el = await createSpreadsheet(10, 4);
      el.setData(
        data({
          '0:0': '1',
          '1:0': { rawValue: '2', format: { bold: true } },
          '2:0': { rawValue: '3', readOnly: true },
          '3:0': '=SUM(A1:A3)',
          '3:1': '=A2*10',
        })
      );
      await el.updateComplete;
      const dataChanges: DataChangeDetail[] = [];
      const structureChanges: StructureChangeDetail[] = [];
      el.addEventListener('data-change', (e) => dataChanges.push((e as CustomEvent).detail));
      el.addEventListener('structure-change', (e) => structureChanges.push((e as CustomEvent).detail));

      el.insertRows(1, 2);
      await el.updateComplete;
      expect(el.rows).toBe(12);
      expect(raw(el, '1:0')).toBe('');
      expect(raw(el, '3:0')).toBe('2');
      expect(el.getCellFormat('3:0')).toEqual({ bold: true });
      expect(el.getData().get('4:0')?.readOnly).toBe(true);
      expect(raw(el, '5:0')).toBe('=SUM(A1:A5)');
      expect(display(el, '5:0')).toBe('6');
      expect(raw(el, '5:1')).toBe('=A4*10');
      expect(display(el, '5:1')).toBe('20');
      expect(structureChanges).toEqual([{ axis: 'row', kind: 'insert', index: 1, count: 2, source: 'programmatic' }]);
      expect(last(dataChanges).operation).toBe('insert-rows');
      expect(last(dataChanges).updates).toContainEqual({ id: '1:0', value: '' });

      internals(el)._undo();
      await el.updateComplete;
      expect(el.rows).toBe(10);
      expect(raw(el, '1:0')).toBe('2');
      expect(raw(el, '3:0')).toBe('=SUM(A1:A3)');
      expect(display(el, '3:0')).toBe('6');
      expect(last(structureChanges)).toEqual({ axis: 'row', kind: 'delete', index: 1, count: 2, source: 'undo' });

      internals(el)._redo();
      await el.updateComplete;
      expect(el.rows).toBe(12);
      expect(raw(el, '5:0')).toBe('=SUM(A1:A5)');
    });

    it('deletes columns, turning references to deleted cells into #REF!', async () => {
      const el = await createSpreadsheet(5, 5);
      el.setData(data({ '0:0': '1', '0:1': '2', '0:2': '3', '1:0': '=A1+B1+C1', '1:1': '=SUM(A1:C1)' }));
      await el.updateComplete;
      el.deleteColumns(1);
      await el.updateComplete;
      expect(el.cols).toBe(4);
      expect(raw(el, '0:1')).toBe('3');
      expect(raw(el, '1:0')).toBe('=A1+#REF!+B1');
      expect(display(el, '1:0')).toBe('#REF!');
      expect(raw(el, '1:1')).toBe('');
      expect(raw(el, '1:0')).not.toBe('');
      // The SUM in B2 was in the deleted column; C2's neighbour moved into B
      expect(el.getData().has('1:1')).toBe(false);

      internals(el)._undo();
      await el.updateComplete;
      expect(el.cols).toBe(5);
      expect(raw(el, '1:0')).toBe('=A1+B1+C1');
      expect(display(el, '1:0')).toBe('6');
      expect(raw(el, '1:1')).toBe('=SUM(A1:C1)');
    });

    it('rejects out-of-range arguments', async () => {
      const el = await createSpreadsheet(5, 5);
      expect(() => el.insertRows(6)).toThrow(RangeError);
      expect(() => el.insertRows(-1)).toThrow(RangeError);
      expect(() => el.insertRows(0, 0)).toThrow(RangeError);
      expect(() => el.deleteRows(4, 2)).toThrow(RangeError);
      expect(() => el.deleteColumns(0, 5)).toThrow(RangeError);
      expect(() => el.insertColumns(1.5)).toThrow(RangeError);
      el.insertRows(5);
      expect(el.rows).toBe(6);
    });

    it('inserts and deletes whole selected rows from the keyboard', async () => {
      const el = await createSpreadsheet(10, 3);
      el.setData(data({ '2:0': 'a', '3:0': 'b' }));
      await el.updateComplete;
      const before = vi.fn();
      el.addEventListener('before-structure-change', before);

      // Not a whole-row selection: explain instead of acting
      await select(el, 2, 0);
      await press(el, '+', { ctrlKey: true, shiftKey: true, code: 'Equal' });
      expect(el.rows).toBe(10);
      expect(announcement(el)).toBe('Select entire rows with Shift+Space or columns with Ctrl+Space first');

      await press(el, ' ', { shiftKey: true });
      await press(el, '+', { ctrlKey: true, shiftKey: true, code: 'Equal' });
      expect(el.rows).toBe(11);
      expect(raw(el, '3:0')).toBe('a');
      expect(announcement(el)).toBe('Inserted row 3');
      expect(before).toHaveBeenCalledTimes(1);
      expect(before.mock.calls[0][0].detail).toEqual({ axis: 'row', kind: 'insert', index: 2, count: 1, source: 'user' });
      // The selection is still the whole (now empty) row
      expect(internals(el)._selection.range).toEqual(range(2, 0, 2, 2));

      await select(el, 3, 0, 4, 2);
      await press(el, '-', { ctrlKey: true, code: 'Minus' });
      expect(el.rows).toBe(9);
      expect(raw(el, '3:0')).toBe('');
      expect(announcement(el)).toBe('Deleted rows 4 to 5');
    });

    it('lets the host veto, and refuses to delete read-only cells', async () => {
      const el = await createSpreadsheet(5, 3);
      el.setData(data({ '1:0': { rawValue: 'x', readOnly: true } }));
      await el.updateComplete;
      await select(el, 1, 0, 1, 2);
      await press(el, '-', { ctrlKey: true, code: 'Minus' });
      expect(el.rows).toBe(5);
      expect(announcement(el)).toBe('Cannot delete: the rows contain read-only cells');

      el.addEventListener('before-structure-change', (e) => e.preventDefault());
      await select(el, 0, 0, 4, 0);
      await press(el, ' ', { ctrlKey: true });
      await press(el, '+', { ctrlKey: true, altKey: true, code: 'Equal' });
      expect(el.cols).toBe(3);
    });

    it('keeps a whole-column selection whole when columns are inserted', async () => {
      const el = await createSpreadsheet(5, 3);
      await select(el, 0, 1);
      await press(el, ' ', { shiftKey: true });
      el.insertColumns(0);
      await el.updateComplete;
      expect(internals(el)._selection.range).toEqual(range(0, 0, 0, 3));
    });
  });

  describe('review regressions', () => {
    it('resets a rejected formula-bar draft so it is not committed into the next cell', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': '1', '0:1': '1' }));
      el.validateCell = (id, value) => (id === '0:0' && isNaN(Number(value)) ? 'Number please' : null);
      await el.updateComplete;
      const bar = shadow(el).querySelector('y11n-formula-bar')!;
      await (bar as unknown as { updateComplete: Promise<unknown> }).updateComplete;
      const input = bar.shadowRoot!.querySelector<HTMLInputElement>('input.formula-input')!;
      input.value = 'bad';
      input.dispatchEvent(new Event('input'));
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
      await el.updateComplete;
      await (bar as unknown as { updateComplete: Promise<unknown> }).updateComplete;
      expect(raw(el, '0:0')).toBe('1');
      expect(input.value).toBe('1');

      await select(el, 0, 1);
      await (bar as unknown as { updateComplete: Promise<unknown> }).updateComplete;
      input.dispatchEvent(new FocusEvent('blur'));
      await el.updateComplete;
      expect(raw(el, '0:1')).toBe('1');
    });

    it('does not let a cut clear moved cells when rows change while the clipboard is written', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'cut me', '1:0': 'keep me' }));
      await el.updateComplete;
      const clipboard = (el as unknown as { _clipboardManager: { cut: (...a: unknown[]) => Promise<string[]> } })
        ._clipboardManager;
      let release!: () => void;
      clipboard.cut = () => new Promise((resolve) => (release = () => resolve(['0:0'])));
      await press(el, 'x', { ctrlKey: true });
      el.deleteRows(0);
      release();
      await new Promise((r) => setTimeout(r, 0));
      await el.updateComplete;
      expect(raw(el, '0:0')).toBe('keep me');
      expect(announcement(el)).toBe('Cut not applied because the data changed');
    });

    it('fills a range formula upward into a parseable #REF! that IFERROR can catch', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '1:1': '=IFERROR(SUM(A1:A2),0)' }));
      await el.updateComplete;
      await select(el, 1, 1);
      internals(el)._autoFill(range(1, 1, 1, 1), range(0, 1, 1, 1));
      await el.updateComplete;
      expect(raw(el, '0:1')).toBe('=IFERROR(SUM(#REF!),0)');
      expect(display(el, '0:1')).toBe('0');
    });

    it('keeps IFERROR results after a row its fallback referenced is deleted', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': '=IFERROR(1,B2)' }));
      await el.updateComplete;
      el.deleteRows(1);
      await el.updateComplete;
      expect(raw(el, '0:0')).toBe('=IFERROR(1,#REF!)');
      expect(display(el, '0:0')).toBe('1');
    });

    it('Ctrl+Alt+D continues a series from the keyboard', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': '1', '1:0': '2', '0:1': 'x' }));
      await el.updateComplete;
      await select(el, 0, 0, 4, 1);
      await press(el, 'd', { ctrlKey: true, altKey: true, code: 'KeyD' });
      expect([2, 3, 4].map((r) => raw(el, `${r}:0`))).toEqual(['3', '4', '5']);
      // Column B's seed is "x", empty: the pattern repeats
      expect([2, 3, 4].map((r) => raw(el, `${r}:1`))).toEqual(['x', '', 'x']);
      expect(announcement(el)).toBe('Filled 6 cells');

      // Nothing to extend: explain instead
      await select(el, 6, 0, 8, 0);
      await press(el, 'r', { ctrlKey: true, altKey: true, code: 'KeyR' });
      expect(announcement(el)).toBe('Select the cells to continue and the empty cells to the right of them');
    });

    it('does not apply a paste that resolves while a progressive fill is running', async () => {
      const el = await createSpreadsheet(1200, 2);
      el.setData(data({ '0:0': '1', '1:0': '2' }));
      await el.updateComplete;
      const clipboard = (el as unknown as {
        _clipboardManager: { paste: (...a: unknown[]) => Promise<Array<{ id: string; value: string }>> };
      })._clipboardManager;
      let release!: () => void;
      clipboard.paste = () =>
        new Promise((resolve) => (release = () => resolve([{ id: '599:0', value: 'pasted' }])));
      await select(el, 599, 0);
      await press(el, 'v', { ctrlKey: true });

      await select(el, 0, 0, 999, 0);
      await press(el, 'd', { ctrlKey: true, altKey: true, code: 'KeyD' });
      release(); // the fill is now applying in animation-frame chunks
      await new Promise((r) => setTimeout(r, 0));
      for (let i = 0; i < 10 && raw(el, '999:0') !== '1000'; i++) {
        await new Promise((r) => requestAnimationFrame(() => r(null)));
      }
      expect(raw(el, '599:0')).toBe('600');
      expect(raw(el, '999:0')).toBe('1000');
      expect(announcement(el)).toBe('Paste not applied because another change is in progress');
      internals(el)._undo();
      await el.updateComplete;
      expect(raw(el, '599:0')).toBe('');
    });

    it('stops user writes when the grid becomes read-only, including an open edit', async () => {
      const el = await createSpreadsheet();
      await press(el, 'F2');
      await typeInEditor(el, 'late');
      el.readOnly = true;
      await el.updateComplete;
      expect(isEditing(el)).toBe(false);
      await pressInEditor(el, 'Enter');
      expect(raw(el, '0:0')).toBe('');
    });

    it('structure and CSV APIs see a data assignment Lit has not processed yet', async () => {
      const el = await createSpreadsheet();
      el.data = data({ '0:0': 'keep' });
      el.insertRows(0);
      await el.updateComplete;
      expect(raw(el, '1:0')).toBe('keep');

      el.data = data({ '0:0': 'a' });
      el.fromCSV('b', { start: { row: 0, col: 1 } });
      await el.updateComplete;
      expect(raw(el, '0:0')).toBe('a');
      expect(raw(el, '0:1')).toBe('b');
    });

    it('imports pending data before committing an open edit for a structure change', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'old' }));
      await el.updateComplete;
      await press(el, 'F2');
      await typeInEditor(el, 'draft');
      el.data = data({ '0:0': 'new', '1:0': 'keep' });
      el.insertRows(0);
      await el.updateComplete;
      expect(raw(el, '1:0')).toBe('draft');
      expect(raw(el, '2:0')).toBe('keep');
    });

    it('stops a progressive fill when the data is replaced between frames', async () => {
      const el = await createSpreadsheet(1200, 2);
      el.setData(data({ '0:0': '1', '1:0': '2' }));
      await el.updateComplete;
      await select(el, 0, 0, 999, 0);
      await press(el, 'd', { ctrlKey: true, altKey: true, code: 'KeyD' });
      // First chunk applied synchronously; replace the data before the next frame
      el.setData(data({ '599:0': 'server' }));
      for (let i = 0; i < 5; i++) await new Promise((r) => requestAnimationFrame(() => r(null)));
      await el.updateComplete;
      expect(raw(el, '599:0')).toBe('server');
      expect(raw(el, '999:0')).toBe('');
      internals(el)._undo();
      await el.updateComplete;
      expect(raw(el, '599:0')).toBe('server');
    });

    it('does not resurrect history cleared by setData() in a structure-change listener', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'old' }));
      await el.updateComplete;
      el.addEventListener('structure-change', () => el.setData(data({ '0:0': 'replacement' })), { once: true });
      el.insertRows(0);
      internals(el)._undo();
      await el.updateComplete;
      expect(raw(el, '0:0')).toBe('replacement');
    });

    it('abandons an operation when a cancelable-event listener changes the grid', async () => {
      const el = await createSpreadsheet(3, 2);
      el.setData(data({ '0:0': 'a', '1:0': 'b', '2:0': 'c' }));
      await el.updateComplete;
      el.addEventListener('before-structure-change', () => el.deleteRows(1, 2), { once: true });
      await select(el, 0, 0, 0, 1);
      await press(el, '-', { ctrlKey: true, code: 'Minus' });
      expect(el.rows).toBe(1);
      expect(raw(el, '0:0')).toBe('a');

      el.addEventListener('before-change', () => el.insertRows(0), { once: true });
      await select(el, 0, 1);
      await enterValue(el, 'x');
      expect(raw(el, '1:0')).toBe('a');
      expect(raw(el, '0:1')).toBe('');
      expect(raw(el, '1:1')).toBe('');
      expect(announcement(el)).toBe('Change not applied because the data changed');
    });

    it('a pending cut or paste does not touch data replaced while the clipboard was busy', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'old document' }));
      await el.updateComplete;
      const clipboard = (el as unknown as {
        _clipboardManager: {
          cut: (...a: unknown[]) => Promise<string[]>;
          paste: (...a: unknown[]) => Promise<Array<{ id: string; value: string }>>;
        };
      })._clipboardManager;
      let release!: () => void;
      clipboard.cut = () => new Promise((resolve) => (release = () => resolve(['0:0'])));
      await press(el, 'x', { ctrlKey: true });
      el.setData(data({ '0:0': 'new document' }));
      release();
      await new Promise((r) => setTimeout(r, 0));
      expect(raw(el, '0:0')).toBe('new document');

      clipboard.paste = () => new Promise((resolve) => (release = () => resolve([{ id: '0:0', value: 'pasted' }])));
      await press(el, 'v', { ctrlKey: true });
      el.data = data({ '0:0': 'newer document' });
      release();
      await new Promise((r) => setTimeout(r, 0));
      await el.updateComplete;
      expect(raw(el, '0:0')).toBe('newer document');
    });

    it('drops a batch when a validation-error listener replaces the data', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': '1', '1:0': 'old' }));
      el.validateCell = (id) => (id === '2:0' ? 'No' : null);
      await el.updateComplete;
      el.addEventListener('validation-error', () => el.setData(data({ '1:0': 'server' })), { once: true });
      await select(el, 0, 0, 2, 0);
      await press(el, 'd', { ctrlKey: true });
      expect(raw(el, '1:0')).toBe('server');
      expect(announcement(el)).toBe('Change not applied because the data changed');
    });

    it('abandons a keyboard structure change when a listener assigns new data', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'old' }));
      await el.updateComplete;
      el.addEventListener('before-structure-change', () => (el.data = data({ '0:0': 'fresh' })), { once: true });
      await press(el, ' ', { shiftKey: true });
      await press(el, '+', { ctrlKey: true, shiftKey: true, code: 'Equal' });
      await el.updateComplete;
      expect(raw(el, '0:0')).toBe('fresh');
      expect(el.rows).toBe(20);
    });

    it('Ctrl+F in the inline editor opens the find bar', async () => {
      const el = await createSpreadsheet();
      await press(el, 'F2');
      await typeInEditor(el, 'abc');
      const event = new KeyboardEvent('keydown', { key: 'f', ctrlKey: true, bubbles: true, composed: true, cancelable: true });
      editor(el).dispatchEvent(event);
      await el.updateComplete;
      await new Promise((r) => setTimeout(r, 0));
      expect(event.defaultPrevented).toBe(true);
      expect(raw(el, '0:0')).toBe('abc');
      expect(shadow(el).activeElement).toBe(shadow(el).querySelector('#find-input'));
    });

    it('openFind() while editing commits the edit and keeps focus in the find bar', async () => {
      const el = await createSpreadsheet();
      await press(el, 'F2');
      await typeInEditor(el, 'typed');
      await el.openFind();
      await el.updateComplete;
      await new Promise((r) => setTimeout(r, 0));
      expect(raw(el, '0:0')).toBe('typed');
      expect(shadow(el).activeElement).toBe(shadow(el).querySelector('#find-input'));
    });
  });

  describe('find', () => {
    it('findCells searches display values case-insensitively in reading order', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '2:0': 'Apple', '0:3': 'pineapple', '1:1': '=UPPER("apple")', '4:4': 'pear' }));
      await el.updateComplete;
      expect(el.findCells('apple')).toEqual(['0:3', '1:1', '2:0']);
      expect(el.findCells('Apple', { matchCase: true })).toEqual(['2:0']);
      expect(el.findCells('upper', { formulas: true })).toEqual(['1:1']);
      expect(el.findCells('')).toEqual([]);
    });

    it('Ctrl+F opens a find bar that steps through matches', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '1:0': 'cat', '3:2': 'catalog', '5:1': 'dog' }));
      await el.updateComplete;

      await press(el, 'f', { ctrlKey: true });
      await el.updateComplete;
      const input = shadow(el).querySelector('#find-input') as HTMLInputElement;
      expect(input).not.toBeNull();
      expect(shadow(el).activeElement).toBe(input);

      input.value = 'cat';
      input.dispatchEvent(new Event('input'));
      await el.updateComplete;
      expect(internals(el)._selection.activeCell).toEqual({ row: 1, col: 0 });
      const status = () => shadow(el).querySelector('#find-status')!.textContent;
      expect(status()).toBe('1 of 2');

      const key = async (k: string, mods: KeyboardEventInit = {}) => {
        input.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, ...mods }));
        await el.updateComplete;
      };
      await key('Enter');
      expect(internals(el)._selection.activeCell).toEqual({ row: 3, col: 2 });
      expect(status()).toBe('2 of 2');
      await key('Enter');
      expect(internals(el)._selection.activeCell).toEqual({ row: 1, col: 0 });
      await key('Enter', { shiftKey: true });
      expect(internals(el)._selection.activeCell).toEqual({ row: 3, col: 2 });

      input.value = 'zebra';
      input.dispatchEvent(new Event('input'));
      await el.updateComplete;
      expect(status()).toBe('No matches');

      await key('Escape');
      await el.updateComplete;
      expect(shadow(el).querySelector('#find-input')).toBeNull();
      expect(shadow(el).activeElement).toBe(shadow(el).querySelector('[data-key="3:2"]'));
    });
  });

  describe('fill handle', () => {
    it('is rendered at the bottom-right of the selection, hidden from assistive tech', async () => {
      const el = await createSpreadsheet();
      await select(el, 1, 1, 2, 3);
      const handles = shadow(el).querySelectorAll('.ls-fill-handle');
      expect(handles.length).toBe(1);
      expect(handles[0].closest('[data-key]')!.getAttribute('data-key')).toBe('2:3');
      expect(handles[0].getAttribute('aria-hidden')).toBe('true');

      el.readOnly = true;
      await el.updateComplete;
      expect(shadow(el).querySelector('.ls-fill-handle')).toBeNull();
    });

    it('continues numeric series and repeats everything else', async () => {
      const el = await createSpreadsheet();
      el.setData(
        data({
          '0:0': '1', '1:0': '2',
          '0:1': '0.1', '1:1': '0.2',
          '0:2': 'a', '1:2': 'b',
          '0:3': '5', '1:3': { rawValue: '=A1*2', format: { italic: true } },
        })
      );
      await el.updateComplete;
      await select(el, 0, 0, 1, 3);
      internals(el)._autoFill(range(0, 0, 1, 3), range(0, 0, 4, 3));
      await el.updateComplete;

      expect([2, 3, 4].map((r) => raw(el, `${r}:0`))).toEqual(['3', '4', '5']);
      expect([2, 3, 4].map((r) => raw(el, `${r}:1`))).toEqual(['0.3', '0.4', '0.5']);
      expect([2, 3, 4].map((r) => raw(el, `${r}:2`))).toEqual(['a', 'b', 'a']);
      // Mixed column: the number and formula repeat, the formula shifted
      expect([2, 3, 4].map((r) => raw(el, `${r}:3`))).toEqual(['5', '=A3*2', '5']);
      expect(el.getCellFormat('3:3')).toEqual({ italic: true });
      expect(internals(el)._selection.range).toEqual(range(0, 0, 4, 3));

      internals(el)._undo();
      await el.updateComplete;
      expect(raw(el, '2:0')).toBe('');
    });

    it('copies a single number, and fills upward and leftward', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '3:3': '7', '4:3': '8', '0:5': 'x' }));
      await el.updateComplete;
      await select(el, 3, 3, 4, 3);
      internals(el)._autoFill(range(3, 3, 4, 3), range(1, 3, 4, 3));
      await el.updateComplete;
      expect(raw(el, '2:3')).toBe('6');
      expect(raw(el, '1:3')).toBe('5');

      await select(el, 0, 5);
      internals(el)._autoFill(range(0, 5, 0, 5), range(0, 2, 0, 5));
      await el.updateComplete;
      expect([2, 3, 4].map((c) => raw(el, `0:${c}`))).toEqual(['x', 'x', 'x']);

      el.setData(data({ '0:0': '10' }));
      await el.updateComplete;
      await select(el, 0, 0);
      internals(el)._autoFill(range(0, 0, 0, 0), range(0, 0, 2, 0));
      await el.updateComplete;
      expect(raw(el, '2:0')).toBe('10');
    });
  });
});
