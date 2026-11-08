// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { CellData, FilterChangeDetail, GridData } from '../types.js';
import type { Y11nSpreadsheet } from '../y11n-spreadsheet.js';
import '../y11n-spreadsheet.js';

function data(entries: Record<string, string>): GridData {
  const map: GridData = new Map();
  for (const [key, rawValue] of Object.entries(entries)) {
    const isNumber = rawValue !== '' && !Number.isNaN(Number(rawValue));
    map.set(key, { rawValue, displayValue: rawValue, type: isNumber ? 'number' : 'text' } as CellData);
  }
  return map;
}

/**
 *      A      B       C
 *  1   Name   Fruit   Qty
 *  2   Ann    Apple   3
 *  3   Bob    Pear    5
 *  4   Cid    Apple   2
 *  5   Dee            7
 *  6   Eve    Plum    1
 */
const TABLE = {
  '0:0': 'Name', '0:1': 'Fruit', '0:2': 'Qty',
  '1:0': 'Ann', '1:1': 'Apple', '1:2': '3',
  '2:0': 'Bob', '2:1': 'Pear', '2:2': '5',
  '3:0': 'Cid', '3:1': 'Apple', '3:2': '2',
  '4:0': 'Dee', '4:2': '7',
  '5:0': 'Eve', '5:1': 'Plum', '5:2': '1',
};
const TABLE_RANGE = { start: { row: 0, col: 0 }, end: { row: 5, col: 2 } };

async function createSpreadsheet(entries: Record<string, string> = TABLE, rows = 30, cols = 6) {
  const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
  el.rows = rows;
  el.cols = cols;
  el.data = data(entries);
  document.body.appendChild(el);
  await el.updateComplete;
  return el;
}

const shadow = (el: Y11nSpreadsheet) => el.shadowRoot!;
const grid = (el: Y11nSpreadsheet) => shadow(el).querySelector('.ls-grid') as HTMLElement;
const cell = (el: Y11nSpreadsheet, row: number, col: number) =>
  shadow(el).querySelector(`[data-row="${row}"][data-col="${col}"]`) as HTMLElement | null;
const announcement = (el: Y11nSpreadsheet) => shadow(el).querySelector('[role="log"]')!.textContent!.trim();
const status = (el: Y11nSpreadsheet) => shadow(el).querySelector('[role="status"][aria-live]')!.textContent!.trim();
const menu = (el: Y11nSpreadsheet) => shadow(el).querySelector('.ls-filter-menu') as HTMLElement | null;
const raw = (el: Y11nSpreadsheet, row: number, col: number) => el.getData().get(`${row}:${col}`)?.rawValue ?? '';

async function press(el: Y11nSpreadsheet, key: string, mods: KeyboardEventInit = {}) {
  grid(el).dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, composed: true, ...mods }));
  await el.updateComplete;
}

type Internals = {
  _selection: {
    moveTo(row: number, col: number, extend?: boolean): void;
    activeCell: { row: number; col: number };
    range: { start: { row: number; col: number }; end: { row: number; col: number } };
  };
  _undo(): void;
  _redo(): void;
  _handleCopy(): Promise<void>;
  _handlePaste(): Promise<void>;
};
const internals = (el: Y11nSpreadsheet) => el as unknown as Internals;
const active = (el: Y11nSpreadsheet) => internals(el)._selection.activeCell;

async function select(el: Y11nSpreadsheet, row: number, col: number, toRow?: number, toCol?: number) {
  internals(el)._selection.moveTo(row, col);
  if (toRow !== undefined) internals(el)._selection.moveTo(toRow, toCol ?? col, true);
  el.requestUpdate();
  await el.updateComplete;
}

function listen<T>(el: Y11nSpreadsheet, type: string): T[] {
  const events: T[] = [];
  el.addEventListener(type, (e) => events.push((e as CustomEvent<T>).detail));
  return events;
}

/** Labels of the values listed in the open filter menu (after Select all). */
function menuValues(el: Y11nSpreadsheet): string[] {
  return [...menu(el)!.querySelectorAll('.ls-filter-option')].slice(1).map((o) => o.textContent!.trim());
}

function menuCheckbox(el: Y11nSpreadsheet, label: string): HTMLInputElement {
  const option = [...menu(el)!.querySelectorAll('.ls-filter-option')].find((o) => o.textContent!.trim() === label);
  return option!.querySelector('input')!;
}

function menuButton(el: Y11nSpreadsheet, text: string): HTMLButtonElement {
  return [...menu(el)!.querySelectorAll('button')].find((b) => b.textContent!.trim().startsWith(text))!;
}

async function click(el: Y11nSpreadsheet, target: HTMLElement) {
  target.click();
  await el.updateComplete;
}

async function toggle(el: Y11nSpreadsheet, checkbox: HTMLInputElement) {
  checkbox.checked = !checkbox.checked;
  checkbox.dispatchEvent(new Event('change', { bubbles: true }));
  await el.updateComplete;
}

describe('filter', () => {
  beforeEach(() => {
    document.body.innerHTML = '';
  });

  describe('API', () => {
    it('hides the rows that do not match and reports them', async () => {
      const el = await createSpreadsheet();
      const events = listen<FilterChangeDetail>(el, 'filter-change');
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      expect(el.getFilter()).toEqual({
        range: TABLE_RANGE,
        columns: [{ col: 1, values: ['Apple'] }],
        hiddenRows: [2, 4, 5],
      });
      expect([0, 1, 2, 3, 4, 5].map((r) => el.isRowHidden(r))).toEqual([false, false, true, false, true, true]);
      expect(events).toEqual([
        { range: TABLE_RANGE, columns: [{ col: 1, values: ['Apple'] }], hiddenRows: [2, 4, 5], source: 'programmatic' },
      ]);
    });

    it('never hides the header row, and keeps blank cells with an empty value', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Nope', ''] }]);
      expect(el.getFilter()!.hiddenRows).toEqual([1, 2, 3, 5]);
      expect(el.isRowHidden(0)).toBe(false);
    });

    it('combines columns: a row must match every filtered column', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [
        { col: 1, values: ['Apple', 'Pear'] },
        { col: 2, values: ['5', '2'] },
      ]);
      expect(el.getFilter()!.hiddenRows).toEqual([1, 4, 5]);
    });

    it('compares displayed values, number formats included', async () => {
      const el = await createSpreadsheet({ '0:0': 'Price', '1:0': '1', '2:0': '2' });
      el.setRangeFormat({ start: { row: 1, col: 0 }, end: { row: 2, col: 0 } }, { numberFormat: { type: 'currency' } });
      el.setFilter({ start: { row: 0, col: 0 }, end: { row: 2, col: 0 } }, [{ col: 0, values: ['$2.00'] }]);
      expect(el.getFilter()!.hiddenRows).toEqual([1]);
    });

    it('setColumnFilter changes or clears one column', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      el.setColumnFilter(2, ['3']);
      expect(el.getFilter()!.hiddenRows).toEqual([2, 3, 4, 5]);
      el.setColumnFilter(1, null);
      expect(el.getFilter()!.columns).toEqual([{ col: 2, values: ['3'] }]);
      expect(el.getFilter()!.hiddenRows).toEqual([2, 3, 4, 5]);
    });

    it('does not re-filter on edits until reapplyFilter()', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      el.fromCSV('Pear', { start: { row: 1, col: 1 } });
      expect(el.isRowHidden(1)).toBe(false);
      el.reapplyFilter();
      expect(el.isRowHidden(1)).toBe(true);
    });

    it('removeFilter shows every row', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      const events = listen<FilterChangeDetail>(el, 'filter-change');
      el.removeFilter();
      expect(el.getFilter()).toBeNull();
      expect(el.isRowHidden(2)).toBe(false);
      expect(events).toEqual([{ range: null, columns: [], hiddenRows: [], source: 'programmatic' }]);
    });

    it('throws for a range outside the grid or a column outside the range', async () => {
      const el = await createSpreadsheet();
      expect(() => el.setFilter({ start: { row: 0, col: 0 }, end: { row: 99, col: 0 } })).toThrow(RangeError);
      expect(() => el.setFilter(TABLE_RANGE, [{ col: 4, values: [] }])).toThrow(RangeError);
      expect(() => el.setColumnFilter(0, ['x'])).toThrow(RangeError);
      el.setFilter(TABLE_RANGE);
      expect(() => el.setColumnFilter(3, ['x'])).toThrow(RangeError);
    });

    it('fires no event when nothing changes', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      const events = listen(el, 'filter-change');
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      el.reapplyFilter();
      expect(events).toHaveLength(0);
    });
  });

  describe('rendering', () => {
    it('leaves hidden rows out of the DOM and keeps real row indexes', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await el.updateComplete;
      expect(cell(el, 2, 0)).toBeNull();
      expect(cell(el, 3, 0)).not.toBeNull();
      const rowIndexes = [...shadow(el).querySelectorAll('.ls-row')].slice(0, 4).map((r) => r.getAttribute('aria-rowindex'));
      expect(rowIndexes).toEqual(['2', '3', '5', '8']);
    });

    it('lays rows out without gaps for hidden rows', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await el.updateComplete;
      // No spacer is rendered between visible rows 1 and 3 (row 2 is hidden)
      const row1 = cell(el, 1, 0)!.closest('.ls-row')!;
      expect(row1.nextElementSibling).toBe(cell(el, 3, 0)!.closest('.ls-row'));
    });

    it('puts a filter button in each header cell and marks filtered columns', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await el.updateComplete;
      for (const c of [0, 1, 2]) {
        expect(cell(el, 0, c)!.getAttribute('aria-haspopup')).toBe('dialog');
        expect(cell(el, 0, c)!.querySelector('.ls-filter-button')).not.toBeNull();
      }
      expect(cell(el, 0, 3)!.hasAttribute('aria-haspopup')).toBe(false);
      expect(cell(el, 0, 1)!.querySelector('.ls-filter-button')!.classList.contains('filtered')).toBe(true);
      expect(cell(el, 0, 0)!.querySelector('.ls-filter-button')!.classList.contains('filtered')).toBe(false);
      expect(shadow(el).querySelector('[data-header-row="1"]')!.classList.contains('filtered-row')).toBe(true);
    });

    it('announces the filter on a header cell', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await select(el, 0, 1);
      expect(status(el)).toBe('B1: Fruit, filtered. Alt+Down for filter options');
      await select(el, 0, 0);
      expect(status(el)).toBe('A1: Name, Alt+Down for filter options');
    });

    it('stacks frozen rows by their visible position', async () => {
      const el = await createSpreadsheet();
      el.freezePanes(4, 0);
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await el.updateComplete;
      // Row 3 is the third visible row (row 2 is hidden)
      expect(cell(el, 3, 0)!.style.top).toBe('calc(var(--_cell-height) * 3)');
      expect(cell(el, 3, 0)!.classList.contains('frozen-row-edge')).toBe(true);
    });
  });

  describe('keyboard', () => {
    it('arrow keys step over hidden rows', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await select(el, 1, 0);
      await press(el, 'ArrowDown');
      expect(active(el)).toEqual({ row: 3, col: 0 });
      await press(el, 'ArrowDown');
      expect(active(el)).toEqual({ row: 6, col: 0 });
      await press(el, 'ArrowUp');
      expect(active(el)).toEqual({ row: 3, col: 0 });
    });

    it('Ctrl+Arrow and Ctrl+End skip hidden rows', async () => {
      const el = await createSpreadsheet(TABLE, 6, 3);
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await select(el, 0, 0);
      await press(el, 'ArrowDown', { ctrlKey: true });
      expect(active(el)).toEqual({ row: 3, col: 0 });
      await press(el, 'End', { ctrlKey: true });
      expect(active(el)).toEqual({ row: 3, col: 2 });
    });

    it('Ctrl+Shift+L adds a filter to the table around the active cell, and removes it', async () => {
      const el = await createSpreadsheet();
      await select(el, 2, 1);
      await press(el, 'L', { ctrlKey: true, shiftKey: true, code: 'KeyL' });
      expect(el.getFilter()).toEqual({ range: TABLE_RANGE, columns: [], hiddenRows: [] });
      expect(announcement(el)).toBe('Filter added to A1:C6. Press Alt+Down on a header cell for filter options');
      await press(el, 'L', { ctrlKey: true, shiftKey: true, code: 'KeyL' });
      expect(el.getFilter()).toBeNull();
      expect(announcement(el)).toBe('Filter removed');
    });

    it('Ctrl+Alt+L reapplies the filter and moves off a row that hides', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      el.fromCSV('Pear', { start: { row: 1, col: 1 } });
      await select(el, 1, 1);
      await press(el, 'l', { ctrlKey: true, altKey: true, code: 'KeyL' });
      expect(el.isRowHidden(1)).toBe(true);
      expect(active(el)).toEqual({ row: 3, col: 1 });
      expect(announcement(el)).toBe('Showing 1 of 5 rows');
    });

    it('works in read-only mode', async () => {
      const el = await createSpreadsheet();
      el.readOnly = true;
      await select(el, 1, 1);
      await press(el, 'L', { ctrlKey: true, shiftKey: true, code: 'KeyL' });
      expect(el.getFilter()).not.toBeNull();
    });

    it('fires a cancelable before-filter for user changes', async () => {
      const el = await createSpreadsheet();
      const before = listen<FilterChangeDetail>(el, 'before-filter');
      el.addEventListener('before-filter', (e) => e.preventDefault());
      await select(el, 1, 1);
      await press(el, 'L', { ctrlKey: true, shiftKey: true, code: 'KeyL' });
      expect(before).toEqual([{ range: TABLE_RANGE, columns: [], hiddenRows: [], source: 'user' }]);
      expect(el.getFilter()).toBeNull();
      expect(announcement(el)).toBe('Filter change cancelled');
    });
  });

  describe('menu', () => {
    afterEach(() => {
      document.body.innerHTML = '';
    });

    async function openMenu(el: Y11nSpreadsheet, col: number) {
      await select(el, 0, col);
      await press(el, 'ArrowDown', { altKey: true });
      await el.updateComplete;
    }

    it('opens with Alt+Down on a header cell, listing the values in order with blanks last', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE);
      await openMenu(el, 1);
      expect(menu(el)).not.toBeNull();
      expect(menu(el)!.getAttribute('aria-label')).toBe('Filter column B, Fruit');
      expect(menuValues(el)).toEqual(['Apple', 'Pear', 'Plum', '(Blanks)']);
      expect(shadow(el).activeElement?.id).toBe('filter-search');
      expect(cell(el, 0, 1)!.getAttribute('aria-expanded')).toBe('true');
    });

    it('sorts numbers numerically', async () => {
      const el = await createSpreadsheet({ '0:0': 'N', '1:0': '10', '2:0': '9', '3:0': '100' });
      el.setFilter({ start: { row: 0, col: 0 }, end: { row: 3, col: 0 } });
      await openMenu(el, 0);
      expect(menuValues(el)).toEqual(['9', '10', '100']);
    });

    it('does not open outside the header row', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE);
      await select(el, 1, 1);
      await press(el, 'ArrowDown', { altKey: true });
      expect(menu(el)).toBeNull();
    });

    it('filters to the checked values on OK and announces the result', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE);
      await openMenu(el, 1);
      await toggle(el, menuCheckbox(el, 'Pear'));
      await toggle(el, menuCheckbox(el, '(Blanks)'));
      await click(el, menuButton(el, 'OK'));
      expect(menu(el)).toBeNull();
      expect(el.getFilter()!.columns).toEqual([{ col: 1, values: ['Apple', 'Plum'] }]);
      expect(el.getFilter()!.hiddenRows).toEqual([2, 4]);
      expect(announcement(el)).toBe('Showing 3 of 5 rows');
    });

    it('lists only values of rows that pass the other columns, and reopens with the criteria checked', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await openMenu(el, 2);
      expect(menuValues(el)).toEqual(['2', '3']);
      menu(el)!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      await el.updateComplete;
      await openMenu(el, 1);
      expect(menuCheckbox(el, 'Apple').checked).toBe(true);
      expect(menuCheckbox(el, 'Pear').checked).toBe(false);
    });

    it('Select all toggles every value, and OK with everything checked clears the column', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await openMenu(el, 1);
      const all = menuCheckbox(el, '(Select all)');
      expect(all.indeterminate).toBe(true);
      await toggle(el, all);
      expect(menuCheckbox(el, 'Pear').checked).toBe(true);
      await click(el, menuButton(el, 'OK'));
      expect(el.getFilter()!.columns).toEqual([]);
      expect(el.getFilter()!.hiddenRows).toEqual([]);
      await openMenu(el, 1);
      await toggle(el, menuCheckbox(el, '(Select all)'));
      expect(menuButton(el, 'OK').disabled).toBe(true);
    });

    it('search narrows the list and OK keeps only the matches', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE);
      await openMenu(el, 1);
      const search = menu(el)!.querySelector<HTMLInputElement>('#filter-search')!;
      search.value = 'p';
      search.dispatchEvent(new Event('input'));
      await el.updateComplete;
      expect(menuValues(el)).toEqual(['Apple', 'Pear', 'Plum']);
      search.value = 'pl';
      search.dispatchEvent(new Event('input'));
      await el.updateComplete;
      expect(menuValues(el)).toEqual(['Apple', 'Plum']);
      search.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
      await el.updateComplete;
      expect(el.getFilter()!.columns).toEqual([{ col: 1, values: ['Apple', 'Plum'] }]);
    });

    it('Escape closes the menu without changes and returns focus to the header cell', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE);
      await openMenu(el, 1);
      await toggle(el, menuCheckbox(el, 'Pear'));
      menu(el)!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      await el.updateComplete;
      await el.updateComplete;
      expect(menu(el)).toBeNull();
      expect(el.getFilter()!.columns).toEqual([]);
      expect(shadow(el).activeElement).toBe(cell(el, 0, 1));
    });

    it('arrow keys move through the value list', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE);
      await openMenu(el, 1);
      const search = menu(el)!.querySelector<HTMLInputElement>('#filter-search')!;
      search.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
      await el.updateComplete;
      await el.updateComplete;
      expect((shadow(el).activeElement as HTMLElement).dataset.filterItem).toBe('0');
      shadow(el).activeElement!.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
      await el.updateComplete;
      await el.updateComplete;
      expect((shadow(el).activeElement as HTMLElement).dataset.filterItem).toBe('4');
      expect(shadow(el).activeElement!.getAttribute('tabindex')).toBe('0');
    });

    it('clears one column from the menu', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }, { col: 2, values: ['3', '2'] }]);
      await openMenu(el, 1);
      await click(el, menuButton(el, 'Clear filter from column B'));
      expect(el.getFilter()!.columns).toEqual([{ col: 2, values: ['3', '2'] }]);
    });

    it('sorts the table from the menu, keeping the header', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE);
      await openMenu(el, 0);
      await click(el, menuButton(el, 'Sort Z to A'));
      expect([1, 2, 3, 4, 5].map((r) => raw(el, r, 0))).toEqual(['Eve', 'Dee', 'Cid', 'Bob', 'Ann']);
      expect(raw(el, 0, 0)).toBe('Name');
    });

    it('opens from the filter button and closes on an outside click', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE);
      await el.updateComplete;
      const button = cell(el, 0, 2)!.querySelector('.ls-filter-button')!;
      button.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, composed: true, button: 0 }));
      await el.updateComplete;
      expect(menu(el)).not.toBeNull();
      expect(active(el)).toEqual({ row: 0, col: 2 });
      document.body.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, composed: true, button: 0 }));
      await el.updateComplete;
      expect(menu(el)).toBeNull();
    });

    it('extends the filter over rows added below the table', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE);
      el.fromCSV('Fay,Pear,4', { start: { row: 6, col: 0 } });
      await openMenu(el, 1);
      await toggle(el, menuCheckbox(el, 'Pear'));
      await click(el, menuButton(el, 'OK'));
      expect(el.getFilter()!.range.end.row).toBe(6);
      expect(el.getFilter()!.hiddenRows).toEqual([2, 6]);
    });
  });

  describe('undo', () => {
    it('undoes and redoes filter changes', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      el.setColumnFilter(2, ['3']);
      const events = listen<FilterChangeDetail>(el, 'filter-change');
      internals(el)._undo();
      expect(el.getFilter()!.hiddenRows).toEqual([2, 4, 5]);
      internals(el)._undo();
      expect(el.getFilter()).toBeNull();
      internals(el)._redo();
      expect(el.getFilter()!.hiddenRows).toEqual([2, 4, 5]);
      expect(events.map((e) => e.source)).toEqual(['undo', 'undo', 'redo']);
    });

    it('keeps the active cell on a visible row when a filter hides it', async () => {
      const el = await createSpreadsheet();
      await select(el, 2, 0);
      grid(el).querySelector<HTMLElement>('[data-row="2"][data-col="0"]')!.focus();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await el.updateComplete;
      expect(active(el)).toEqual({ row: 3, col: 0 });
    });
  });

  describe('writes skip hidden rows', () => {
    it('clearing a selection leaves hidden rows alone', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await select(el, 1, 0, 3, 0);
      await press(el, 'Delete');
      expect([1, 2, 3].map((r) => raw(el, r, 0))).toEqual(['', 'Bob', '']);
    });

    it('fill down writes only visible rows, shifting formulas by the real offset', async () => {
      const el = await createSpreadsheet({ ...TABLE, '1:3': '=C2*2' });
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await select(el, 1, 3, 3, 3);
      await press(el, 'd', { ctrlKey: true });
      expect(raw(el, 2, 3)).toBe('');
      expect(raw(el, 3, 3)).toBe('=C4*2');
    });

    it('fill series continues on visible rows', async () => {
      const el = await createSpreadsheet({ '0:0': 'N', '1:0': '1', '2:0': 'x', '3:0': '2', '4:0': '', '5:0': '' }, 10, 2);
      el.setFilter({ start: { row: 0, col: 0 }, end: { row: 5, col: 0 } }, [{ col: 0, values: ['1', '2', ''] }]);
      expect(el.isRowHidden(2)).toBe(true);
      await select(el, 1, 0, 5, 0);
      await press(el, 'd', { ctrlKey: true, altKey: true, code: 'KeyD' });
      expect([1, 2, 3, 4, 5].map((r) => raw(el, r, 0))).toEqual(['1', 'x', '2', '3', '4']);
    });

    it('Ctrl+Enter fills only visible rows', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await select(el, 1, 3, 3, 3);
      await press(el, 'z');
      const editor = shadow(el).querySelector<HTMLInputElement>('#editor')!;
      editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }));
      await el.updateComplete;
      expect([1, 2, 3].map((r) => raw(el, r, 3))).toEqual(['z', '', 'z']);
    });

    it('sorting moves only visible rows', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple', 'Plum'] }]);
      el.sort(TABLE_RANGE, { col: 2, direction: 'descending' }, { hasHeader: true });
      // Visible rows 1, 3, 5 had Qty 3, 2, 1 already: reversed to ascending order of name
      expect([1, 2, 3, 4, 5].map((r) => raw(el, r, 0))).toEqual(['Ann', 'Bob', 'Cid', 'Dee', 'Eve']);
      el.sort(TABLE_RANGE, { col: 2, direction: 'ascending' }, { hasHeader: true });
      expect([1, 2, 3, 4, 5].map((r) => raw(el, r, 0))).toEqual(['Eve', 'Bob', 'Cid', 'Dee', 'Ann']);
    });

    describe('clipboard', () => {
      let clipboardText = '';
      beforeEach(() => {
        clipboardText = '';
        Object.defineProperty(navigator, 'clipboard', {
          configurable: true,
          value: {
            writeText: async (text: string) => {
              clipboardText = text;
            },
            readText: async () => clipboardText,
            write: async () => {
              throw new Error('unsupported');
            },
            read: async () => {
              throw new Error('unsupported');
            },
          },
        });
      });

      it('copies only visible rows', async () => {
        const el = await createSpreadsheet();
        el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
        await select(el, 0, 0, 5, 1);
        await internals(el)._handleCopy();
        expect(clipboardText).toBe('Name\tFruit\nAnn\tApple\nCid\tApple');
      });

      it('pastes onto the visible rows from the anchor down', async () => {
        const el = await createSpreadsheet();
        el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
        clipboardText = 'x\ny\nz';
        await select(el, 1, 4);
        await internals(el)._handlePaste();
        expect([1, 2, 3, 4, 5, 6].map((r) => raw(el, r, 4))).toEqual(['x', '', 'y', '', '', 'z']);
      });

      it('shifts pasted formulas by how far each row moved', async () => {
        const el = await createSpreadsheet({ ...TABLE, '1:3': '=C2', '3:3': '=C4' });
        el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
        await select(el, 1, 3, 3, 3);
        await internals(el)._handleCopy();
        el.removeFilter();
        await select(el, 10, 3);
        await internals(el)._handlePaste();
        expect(raw(el, 10, 3)).toBe('=C11');
        expect(raw(el, 11, 3)).toBe('=C12');
      });
    });
  });

  describe('rows and columns', () => {
    it('moves with inserted rows and shows the new rows', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      el.insertRows(0, 1);
      expect(el.getFilter()!.range).toEqual({ start: { row: 1, col: 0 }, end: { row: 6, col: 2 } });
      expect(el.getFilter()!.hiddenRows).toEqual([3, 5, 6]);
      el.insertRows(3, 2);
      expect(el.getFilter()!.range.end.row).toBe(8);
      expect(el.getFilter()!.hiddenRows).toEqual([5, 7, 8]);
    });

    it('shrinks with deleted rows, and goes away with its header row; undo brings it back', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      el.deleteRows(4, 2);
      expect(el.getFilter()!.range.end.row).toBe(3);
      expect(el.getFilter()!.hiddenRows).toEqual([2]);
      const events = listen<FilterChangeDetail>(el, 'filter-change');
      el.deleteRows(0, 1);
      expect(el.getFilter()).toBeNull();
      expect(events.map((e) => e.range)).toEqual([null]);
      internals(el)._undo();
      expect(el.getFilter()!.hiddenRows).toEqual([2]);
    });

    it('moves its columns and criteria with inserted and deleted columns', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      el.insertColumns(0, 1);
      expect(el.getFilter()!.range).toEqual({ start: { row: 0, col: 1 }, end: { row: 5, col: 3 } });
      expect(el.getFilter()!.columns).toEqual([{ col: 2, values: ['Apple'] }]);
      el.deleteColumns(2, 1);
      expect(el.getFilter()!.range).toEqual({ start: { row: 0, col: 1 }, end: { row: 5, col: 2 } });
      expect(el.getFilter()!.columns).toEqual([]);
    });

    it('refuses to delete selected rows when some of them are hidden', async () => {
      const el = await createSpreadsheet();
      el.setFilter(TABLE_RANGE, [{ col: 1, values: ['Apple'] }]);
      await select(el, 1, 0, 3, 0);
      await press(el, ' ', { shiftKey: true });
      await press(el, '-', { ctrlKey: true, code: 'Minus' });
      expect(raw(el, 2, 0)).toBe('Bob');
      expect(announcement(el)).toBe('Cannot delete: some of the rows are hidden by the filter');
    });
  });
});
