// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';
import type { CellData, DataChangeDetail, GridData, SelectionRange, SortDetail } from '../types.js';
import type { Y11nSpreadsheet } from '../y11n-spreadsheet.js';
import '../y11n-spreadsheet.js';

function data(entries: Record<string, string | Partial<CellData>>): GridData {
  const map: GridData = new Map();
  for (const [key, value] of Object.entries(entries)) {
    const cell = typeof value === 'string' ? { rawValue: value } : value;
    const rawValue = cell.rawValue ?? '';
    // Type literals the way the engine would, so sorting sees real numbers and booleans
    const type: CellData['type'] =
      rawValue !== '' && !rawValue.startsWith('=') && !Number.isNaN(Number(rawValue))
        ? 'number'
        : rawValue === 'TRUE' || rawValue === 'FALSE'
          ? 'boolean'
          : 'text';
    map.set(key, { displayValue: rawValue, type, ...cell, rawValue } as CellData);
  }
  return map;
}

async function createSpreadsheet(entries: Record<string, string | Partial<CellData>> = {}, rows = 20, cols = 6) {
  const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
  el.rows = rows;
  el.cols = cols;
  document.body.appendChild(el);
  await el.updateComplete;
  el.setData(data(entries));
  await el.updateComplete;
  return el;
}

const shadow = (el: Y11nSpreadsheet) => el.shadowRoot!;
const grid = (el: Y11nSpreadsheet) => shadow(el).querySelector('.ls-grid') as HTMLElement;
const announcement = (el: Y11nSpreadsheet) => shadow(el).querySelector('[role="log"]')!.textContent!.trim();

async function press(el: Y11nSpreadsheet, key: string, mods: KeyboardEventInit = {}) {
  grid(el).dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, composed: true, ...mods }));
  await el.updateComplete;
}

const raw = (el: Y11nSpreadsheet, key: string) => el.getData().get(key)?.rawValue ?? '';
const display = (el: Y11nSpreadsheet, key: string) => el.getData().get(key)?.displayValue ?? '';
/** Raw values of column `col`, rows `r0..r1` */
const column = (el: Y11nSpreadsheet, col: number, r0: number, r1: number) =>
  Array.from({ length: r1 - r0 + 1 }, (_, i) => raw(el, `${r0 + i}:${col}`));

type Internals = {
  _selection: { moveTo(row: number, col: number, extend?: boolean): void };
  _undo(): void;
  _redo(): void;
  _undoStack: unknown[];
};
const internals = (el: Y11nSpreadsheet) => el as unknown as Internals;

async function select(el: Y11nSpreadsheet, r1: number, c1: number, r2 = r1, c2 = c1) {
  internals(el)._selection.moveTo(r1, c1);
  internals(el)._selection.moveTo(r2, c2, true);
  el.requestUpdate();
  await el.updateComplete;
}

const range = (r1: number, c1: number, r2: number, c2: number): SelectionRange => ({
  start: { row: r1, col: c1 },
  end: { row: r2, col: c2 },
});

function listen<T>(el: Y11nSpreadsheet, type: string): T[] {
  const events: T[] = [];
  el.addEventListener(type, (e) => events.push((e as CustomEvent<T>).detail));
  return events;
}

/** A small table: header row, then name / qty / price rows, with totals in D */
const TABLE = {
  '0:0': 'Item',
  '0:1': 'Qty',
  '0:2': 'Price',
  '0:3': 'Total',
  '1:0': 'Widget C',
  '1:1': '7',
  '1:2': '12',
  '1:3': '=B2*C2',
  '2:0': 'Widget A',
  '2:1': '10',
  '2:2': '5',
  '2:3': '=B3*C3',
  '3:0': 'Widget B',
  '3:1': '25',
  '3:2': '3',
  '3:3': '=B4*C4',
};

describe('sorting', () => {
  beforeEach(() => {
    document.body.innerHTML = '';
  });

  describe('sort()', () => {
    it('reorders whole rows of the range by the key column', async () => {
      const el = await createSpreadsheet(TABLE);
      expect(el.sort(range(1, 0, 3, 3), { col: 0 })).toBe(true);
      expect(column(el, 0, 1, 3)).toEqual(['Widget A', 'Widget B', 'Widget C']);
      expect(column(el, 1, 1, 3)).toEqual(['10', '25', '7']);
      expect(raw(el, '0:0')).toBe('Item');
    });

    it('moves only cells inside the range', async () => {
      const el = await createSpreadsheet({ '0:0': 'b', '0:1': 'x', '1:0': 'a', '1:1': 'y' });
      el.sort(range(0, 0, 1, 0), { col: 0 });
      expect(column(el, 0, 0, 1)).toEqual(['a', 'b']);
      expect(column(el, 1, 0, 1)).toEqual(['x', 'y']);
    });

    it('orders mixed values numbers, text, FALSE, TRUE, errors, blanks', async () => {
      const el = await createSpreadsheet({
        '0:0': 'TRUE',
        '1:0': '=1/0',
        '2:0': 'banana',
        '3:0': '42',
        '4:0': 'FALSE',
        '5:0': 'Apple',
        '7:0': '-3',
      });
      el.sort(range(0, 0, 7, 0), { col: 0 });
      expect(column(el, 0, 0, 7)).toEqual(['-3', '42', 'Apple', 'banana', 'FALSE', 'TRUE', '=1/0', '']);
      el.sort(range(0, 0, 7, 0), { col: 0, direction: 'descending' });
      expect(column(el, 0, 0, 7)).toEqual(['=1/0', 'TRUE', 'FALSE', 'banana', 'Apple', '42', '-3', '']);
    });

    it('keeps the header row in place with hasHeader', async () => {
      const el = await createSpreadsheet(TABLE);
      el.sort(range(0, 0, 3, 3), { col: 1, direction: 'descending' }, { hasHeader: true });
      expect(column(el, 1, 0, 3)).toEqual(['Qty', '25', '10', '7']);
    });

    it('breaks ties with later keys and is otherwise stable', async () => {
      const el = await createSpreadsheet({
        '0:0': 'b', '0:1': '2', '0:2': 'first b2',
        '1:0': 'a', '1:1': '9', '1:2': 'a9',
        '2:0': 'b', '2:1': '1', '2:2': 'b1',
        '3:0': 'B', '3:1': '2', '3:2': 'second b2',
      });
      el.sort(range(0, 0, 3, 2), [{ col: 0 }, { col: 1 }]);
      expect(column(el, 2, 0, 3)).toEqual(['a9', 'b1', 'first b2', 'second b2']);
    });

    it('sorts number cells by their value, not their formatted display', async () => {
      const currency = { numberFormat: { type: 'currency' as const, decimals: 2, thousandsSep: true } };
      const el = await createSpreadsheet({
        '0:0': { rawValue: '1000', type: 'number', format: currency },
        '1:0': { rawValue: '200', type: 'number', format: currency },
        '2:0': { rawValue: '=5*10', type: 'number', format: currency },
      });
      el.sort(range(0, 0, 2, 0), { col: 0 });
      expect(column(el, 0, 0, 2)).toEqual(['=5*10', '200', '1000']);
      expect(display(el, '2:0')).toBe('$1,000.00');
    });

    it('shifts relative references in moved formulas and recalculates them', async () => {
      const el = await createSpreadsheet(TABLE);
      el.sort(range(1, 0, 3, 3), { col: 0 });
      // Widget A moved from row 3 to row 2: its total follows it
      expect(raw(el, '1:3')).toBe('=B2*C2');
      expect(display(el, '1:3')).toBe('50');
      expect(display(el, '2:3')).toBe('75');
      expect(display(el, '3:3')).toBe('84');
    });

    it('keeps absolute references and leaves formulas outside the range alone', async () => {
      const el = await createSpreadsheet({
        '0:0': '3',
        '1:0': '1',
        '2:0': '2',
        '0:1': '=A1*$E$1',
        '0:4': '10',
        '5:0': '=A1',
        '6:0': '=SUM(A1:A3)',
      });
      el.sort(range(0, 0, 2, 1), { col: 0 });
      expect(raw(el, '2:1')).toBe('=A3*$E$1');
      expect(display(el, '2:1')).toBe('30');
      // Outside references keep pointing at the same cells, which now hold other values
      expect(raw(el, '5:0')).toBe('=A1');
      expect(display(el, '5:0')).toBe('1');
      expect(display(el, '6:0')).toBe('6');
    });

    it('moves formats and read-only flags with their cells', async () => {
      const el = await createSpreadsheet({
        '0:0': { rawValue: 'b', format: { bold: true }, readOnly: true },
        '1:0': 'a',
      });
      el.sort(range(0, 0, 1, 0), { col: 0 });
      expect(el.getData().get('1:0')).toMatchObject({ rawValue: 'b', format: { bold: true }, readOnly: true });
      expect(el.getData().get('0:0')?.format).toBeUndefined();
      expect(el.getData().get('0:0')?.readOnly).toBeUndefined();
    });

    it('is one undoable step and fires data-change and sort-change', async () => {
      const el = await createSpreadsheet(TABLE);
      const sorts = listen<SortDetail>(el, 'sort-change');
      const changes = listen<DataChangeDetail>(el, 'data-change');
      el.sort(range(1, 0, 3, 3), { col: 0 });

      expect(sorts).toEqual([
        { range: range(1, 0, 3, 3), keys: [{ col: 0, direction: 'ascending' }], hasHeader: false, source: 'programmatic' },
      ]);
      expect(changes[0].operation).toBe('sort');
      expect(changes[0].updates).toContainEqual({ id: '1:0', value: 'Widget A' });

      internals(el)._undo();
      await el.updateComplete;
      expect(column(el, 0, 1, 3)).toEqual(['Widget C', 'Widget A', 'Widget B']);
      expect(display(el, '1:3')).toBe('84');
      expect(sorts[1].source).toBe('undo');

      internals(el)._redo();
      await el.updateComplete;
      expect(column(el, 0, 1, 3)).toEqual(['Widget A', 'Widget B', 'Widget C']);
      expect(display(el, '1:3')).toBe('50');
      expect(sorts[2].source).toBe('redo');
    });

    it('keeps undo in order when a data-change listener re-sorts after each change', async () => {
      const el = await createSpreadsheet({ '0:0': '1', '0:1': 'one', '1:0': '2', '1:1': 'two', '2:0': '3', '2:1': 'three' });
      el.addEventListener('data-change', (e) => {
        const { operation, source } = (e as CustomEvent<DataChangeDetail>).detail;
        if (source === 'user' && operation !== 'sort') el.sort(range(0, 0, 2, 1), { col: 0 });
      });
      await select(el, 0, 0);
      await press(el, 'Delete');
      expect(column(el, 0, 0, 2)).toEqual(['2', '3', '']);
      expect(column(el, 1, 0, 2)).toEqual(['two', 'three', 'one']);

      // The sort was recorded after the clear, so it is undone first
      internals(el)._undo();
      expect(column(el, 0, 0, 2)).toEqual(['', '2', '3']);
      expect(column(el, 1, 0, 2)).toEqual(['one', 'two', 'three']);
      internals(el)._undo();
      expect(column(el, 0, 0, 2)).toEqual(['1', '2', '3']);

      internals(el)._redo();
      internals(el)._redo();
      expect(column(el, 0, 0, 2)).toEqual(['2', '3', '']);
      expect(column(el, 1, 0, 2)).toEqual(['two', 'three', 'one']);
    });

    it('clears the redo stack when a listener makes a change during undo', async () => {
      const el = await createSpreadsheet({ '0:0': 'b', '1:0': 'a' });
      el.sort(range(0, 0, 1, 0), { col: 0 });
      el.addEventListener('data-change', (e) => {
        if ((e as CustomEvent<DataChangeDetail>).detail.source === 'undo') el.fromCSV('z', { start: { row: 5, col: 0 } });
      }, { once: true });
      internals(el)._undo();
      expect(column(el, 0, 0, 1)).toEqual(['b', 'a']);
      // Redoing the sort on top of the listener's change would be a new history branch
      internals(el)._redo();
      expect(column(el, 0, 0, 1)).toEqual(['b', 'a']);
      expect(raw(el, '5:0')).toBe('z');
    });

    it('does nothing, and records no history, when the range is already sorted', async () => {
      const el = await createSpreadsheet({ '0:0': '1', '1:0': '2' });
      const sorts = listen<SortDetail>(el, 'sort-change');
      expect(el.sort(range(0, 0, 1, 0), { col: 0 })).toBe(false);
      expect(sorts).toHaveLength(0);
      expect(internals(el)._undoStack).toHaveLength(0);
    });

    it('rejects ranges outside the grid and keys outside the range', async () => {
      const el = await createSpreadsheet({}, 10, 4);
      expect(() => el.sort(range(0, 0, 10, 0), { col: 0 })).toThrow(RangeError);
      expect(() => el.sort(range(0, 0, 3, 1), { col: 2 })).toThrow(RangeError);
      expect(() => el.sort(range(0, 0, 3, 1), [])).toThrow(RangeError);
      expect(() => el.sort(range(-1, 0, 3, 1), { col: 0 })).toThrow(RangeError);
      expect(() => el.sort(range(0, 0, 3, 1), { col: 0, direction: 'up' as never })).toThrow(RangeError);
    });

    it('accepts a range given bottom-right to top-left', async () => {
      const el = await createSpreadsheet({ '0:0': 'b', '1:0': 'a' });
      el.sort(range(1, 0, 0, 0), { col: 0 });
      expect(column(el, 0, 0, 1)).toEqual(['a', 'b']);
    });
  });

  describe('sortSelection() and the keyboard', () => {
    it('sorts the data block around a single active cell and detects its header', async () => {
      const el = await createSpreadsheet({ ...TABLE, '5:0': 'not part of the block' });
      await select(el, 2, 1);
      expect(el.sortSelection('descending')).toBe(true);
      await el.updateComplete;
      expect(column(el, 1, 0, 3)).toEqual(['Qty', '25', '10', '7']);
      expect(raw(el, '5:0')).toBe('not part of the block');
      expect(announcement(el)).toBe('Sorted A2:D4 by column B, descending. Row 1 kept as the header');
    });

    it('sorts the header too when the first row does not look like one', async () => {
      const el = await createSpreadsheet({ '0:0': 'pear', '1:0': 'apple', '2:0': 'fig' });
      await select(el, 0, 0);
      el.sortSelection();
      expect(column(el, 0, 0, 2)).toEqual(['apple', 'fig', 'pear']);
    });

    it('treats a bold first row over plain text as a header', async () => {
      const el = await createSpreadsheet({
        '0:0': { rawValue: 'Name', format: { bold: true } },
        '1:0': 'pear',
        '2:0': 'apple',
      });
      await select(el, 1, 0);
      el.sortSelection();
      expect(column(el, 0, 0, 2)).toEqual(['Name', 'apple', 'pear']);
    });

    it('sorts an explicit selection, trimmed to its data, by the active column', async () => {
      const el = await createSpreadsheet({ '0:0': 'x', '0:1': '3', '1:0': 'y', '1:1': '1', '2:0': 'z', '2:1': '2' });
      await select(el, 0, 0, 19, 1);
      el.sortSelection();
      await el.updateComplete;
      expect(column(el, 0, 0, 2)).toEqual(['y', 'z', 'x']);
      expect(announcement(el)).toBe('Sorted A1:B3 by column B, ascending');
    });

    it('sorts the whole table when only one of its columns is selected', async () => {
      const el = await createSpreadsheet({ '0:0': 'x', '0:1': '3', '1:0': 'y', '1:1': '1', '2:0': 'z', '2:1': '2' });
      await select(el, 0, 1, 19, 1);
      el.sortSelection();
      await el.updateComplete;
      expect(column(el, 0, 0, 2)).toEqual(['y', 'z', 'x']);
      expect(column(el, 1, 0, 2)).toEqual(['1', '2', '3']);
      expect(announcement(el)).toBe('Sorted A1:B3 by column B, ascending');
    });

    it('sorts a lone column of data on its own', async () => {
      const el = await createSpreadsheet({ '0:1': 'c', '1:1': 'a', '2:1': 'b', '0:3': 'other' });
      await select(el, 0, 1, 2, 1);
      el.sortSelection();
      expect(column(el, 1, 0, 2)).toEqual(['a', 'b', 'c']);
      expect(raw(el, '0:3')).toBe('other');
    });

    it('sorts with Alt+Shift+Up and Alt+Shift+Down', async () => {
      const el = await createSpreadsheet({ '0:0': 'b', '1:0': 'c', '2:0': 'a' });
      await select(el, 1, 0);
      await press(el, 'ArrowUp', { altKey: true, shiftKey: true });
      expect(column(el, 0, 0, 2)).toEqual(['a', 'b', 'c']);
      await press(el, 'ArrowDown', { altKey: true, shiftKey: true });
      expect(column(el, 0, 0, 2)).toEqual(['c', 'b', 'a']);
    });

    it('refuses to move read-only cells and says why', async () => {
      const el = await createSpreadsheet({ '0:0': 'b', '0:1': { rawValue: 'locked', readOnly: true }, '1:0': 'a' });
      await select(el, 0, 0);
      expect(el.sortSelection()).toBe(false);
      await el.updateComplete;
      expect(column(el, 0, 0, 1)).toEqual(['b', 'a']);
      expect(announcement(el)).toBe('Cannot sort: the rows to move contain read-only cells');
    });

    it('allows the sort when the read-only cells stay where they are', async () => {
      const el = await createSpreadsheet({ '0:0': 'a', '0:1': { rawValue: 'locked', readOnly: true }, '1:0': 'c', '2:0': 'b' });
      await select(el, 0, 0);
      expect(el.sortSelection()).toBe(true);
      expect(column(el, 0, 0, 2)).toEqual(['a', 'b', 'c']);
    });

    it('fires a cancelable before-sort', async () => {
      const el = await createSpreadsheet({ '0:0': 'b', '1:0': 'a' });
      const seen: SortDetail[] = [];
      el.addEventListener('before-sort', (e) => {
        seen.push((e as CustomEvent<SortDetail>).detail);
        e.preventDefault();
      });
      await select(el, 0, 0);
      expect(el.sortSelection()).toBe(false);
      await el.updateComplete;
      expect(seen).toEqual([
        { range: range(0, 0, 1, 0), keys: [{ col: 0, direction: 'ascending' }], hasHeader: false, source: 'user' },
      ]);
      expect(column(el, 0, 0, 1)).toEqual(['b', 'a']);
      expect(announcement(el)).toBe('Sort cancelled');
    });

    it('ignores a before-sort listener that mutates the event detail', async () => {
      const el = await createSpreadsheet({
        '0:0': 'b', '0:1': 'x', '0:2': 'keep 1',
        '1:0': 'a', '1:1': 'y', '1:2': 'keep 2',
      });
      el.addEventListener('before-sort', (e) => {
        (e as CustomEvent<SortDetail>).detail.range.end.col = 2;
      });
      await select(el, 1, 1, 0, 0);
      el.sortSelection();
      expect(column(el, 0, 0, 1)).toEqual(['a', 'b']);
      expect(column(el, 2, 0, 1)).toEqual(['keep 1', 'keep 2']);
    });

    it('does not sort when a before-sort listener changes the data', async () => {
      const el = await createSpreadsheet({ '0:0': 'b', '1:0': 'a' });
      el.addEventListener('before-sort', () => el.setData(data({ '0:0': 'z', '1:0': 'y' })));
      await select(el, 0, 0);
      expect(el.sortSelection()).toBe(false);
      expect(column(el, 0, 0, 1)).toEqual(['z', 'y']);
    });

    it('does nothing in a read-only grid, on an empty cell, or when already sorted', async () => {
      const el = await createSpreadsheet({ '0:0': 'a', '1:0': 'b' });
      const sorts = listen<SortDetail>(el, 'sort-change');
      await select(el, 5, 5);
      expect(el.sortSelection()).toBe(false);
      await el.updateComplete;
      expect(announcement(el)).toBe('Nothing to sort');

      await select(el, 0, 0);
      expect(el.sortSelection()).toBe(false);
      await el.updateComplete;
      expect(announcement(el)).toBe('Already sorted');

      el.readOnly = true;
      await el.updateComplete;
      expect(el.sortSelection('descending')).toBe(false);
      await press(el, 'ArrowDown', { altKey: true, shiftKey: true });
      expect(column(el, 0, 0, 1)).toEqual(['a', 'b']);
      expect(sorts).toHaveLength(0);
    });

    it('commits an open edit before sorting', async () => {
      const el = await createSpreadsheet({ '0:0': 'b', '1:0': 'c' });
      await select(el, 1, 0);
      await press(el, 'F2');
      const editor = shadow(el).querySelector('#editor') as HTMLInputElement;
      editor.value = 'a';
      editor.dispatchEvent(new Event('input', { bubbles: true }));
      await el.updateComplete;
      el.sortSelection();
      expect(column(el, 0, 0, 1)).toEqual(['a', 'b']);
    });
  });
});
