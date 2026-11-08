// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';
import type { CellData, FreezeChangeDetail, GridData } from '../types.js';
import type { Y11nSpreadsheet } from '../y11n-spreadsheet.js';
import '../y11n-spreadsheet.js';

function data(entries: Record<string, string>): GridData {
  const map: GridData = new Map();
  for (const [key, rawValue] of Object.entries(entries)) {
    map.set(key, { rawValue, displayValue: rawValue, type: 'text' } as CellData);
  }
  return map;
}

async function createSpreadsheet(rows = 30, cols = 8) {
  const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
  el.rows = rows;
  el.cols = cols;
  document.body.appendChild(el);
  await el.updateComplete;
  return el;
}

const shadow = (el: Y11nSpreadsheet) => el.shadowRoot!;
const grid = (el: Y11nSpreadsheet) => shadow(el).querySelector('.ls-grid') as HTMLElement;
const cell = (el: Y11nSpreadsheet, row: number, col: number) =>
  shadow(el).querySelector(`[data-row="${row}"][data-col="${col}"]`) as HTMLElement | null;
const announcement = (el: Y11nSpreadsheet) => shadow(el).querySelector('[role="log"]')!.textContent!.trim();

async function press(el: Y11nSpreadsheet, key: string, mods: KeyboardEventInit = {}) {
  grid(el).dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, composed: true, ...mods }));
  await el.updateComplete;
}

type Internals = {
  _selection: { moveTo(row: number, col: number, extend?: boolean): void; activeCell: { row: number; col: number } };
  _undo(): void;
  _redo(): void;
  _undoStack: unknown[];
  _scrollTop: number;
  _scrollLeft: number;
};
const internals = (el: Y11nSpreadsheet) => el as unknown as Internals;

async function select(el: Y11nSpreadsheet, row: number, col: number) {
  internals(el)._selection.moveTo(row, col);
  el.requestUpdate();
  await el.updateComplete;
}

function listen<T>(el: Y11nSpreadsheet, type: string): T[] {
  const events: T[] = [];
  el.addEventListener(type, (e) => events.push((e as CustomEvent<T>).detail));
  return events;
}

async function scrollTo(el: Y11nSpreadsheet, top: number, left: number) {
  grid(el).scrollTop = top;
  grid(el).scrollLeft = left;
  grid(el).dispatchEvent(new Event('scroll'));
  await el.updateComplete;
}

describe('frozen panes', () => {
  beforeEach(() => {
    document.body.innerHTML = '';
  });

  describe('API and attributes', () => {
    it('reflects frozenRows / frozenCols to frozen-rows / frozen-cols', async () => {
      const el = await createSpreadsheet();
      el.freezePanes(2, 1);
      await el.updateComplete;
      expect(el.getAttribute('frozen-rows')).toBe('2');
      expect(el.getAttribute('frozen-cols')).toBe('1');
      el.setAttribute('frozen-rows', '3');
      await el.updateComplete;
      expect(el.frozenRows).toBe(3);
    });

    it('freezes above and left of the active cell without arguments, and fires freeze-change', async () => {
      const el = await createSpreadsheet();
      const events = listen<FreezeChangeDetail>(el, 'freeze-change');
      await select(el, 3, 2);
      el.freezePanes();
      expect([el.frozenRows, el.frozenCols]).toEqual([3, 2]);
      el.unfreezePanes();
      expect([el.frozenRows, el.frozenCols]).toEqual([0, 0]);
      expect(events).toEqual([
        { rows: 3, cols: 2, source: 'programmatic' },
        { rows: 0, cols: 0, source: 'programmatic' },
      ]);
    });

    it('keeps at least one row and column scrolling, and rejects invalid counts', async () => {
      const el = await createSpreadsheet(5, 3);
      el.freezePanes(10, 10);
      expect([el.frozenRows, el.frozenCols]).toEqual([4, 2]);
      expect(() => el.freezePanes(-1, 0)).toThrow(RangeError);
      expect(() => el.freezePanes(1.5, 0)).toThrow(RangeError);
    });

    it('freezePanes(rows) freezes only rows', async () => {
      const el = await createSpreadsheet();
      await select(el, 4, 3);
      el.freezePanes(1);
      expect([el.frozenRows, el.frozenCols]).toEqual([1, 0]);
    });

    it('is not undoable', async () => {
      const el = await createSpreadsheet();
      el.freezePanes(1, 1);
      expect(internals(el)._undoStack).toHaveLength(0);
    });

    it('ignores invalid property values', async () => {
      const el = await createSpreadsheet(10, 4);
      el.frozenRows = Number.NaN;
      el.frozenCols = 99;
      await el.updateComplete;
      expect(shadow(el).querySelectorAll('.frozen-row')).toHaveLength(0);
      // Clamped to cols - 1
      expect(shadow(el).querySelectorAll('.ls-col-header.frozen-col')).toHaveLength(3);
    });
  });

  describe('rendering', () => {
    it('marks frozen cells and headers sticky, with offsets and freeze lines', async () => {
      const el = await createSpreadsheet();
      el.setColumnWidth(0, 150);
      el.freezePanes(2, 2);
      await el.updateComplete;

      const a1 = cell(el, 0, 0)!;
      expect(a1.className).toContain('frozen-row');
      expect(a1.className).toContain('frozen-col');
      expect(a1.style.top).toBe('calc(var(--_cell-height) * 1)');
      const b2 = cell(el, 1, 1)!;
      expect(b2.className).toContain('frozen-row-edge');
      expect(b2.className).toContain('frozen-col-edge');
      expect(b2.style.left).toBe('calc(var(--_row-header-width) + 150px + 0 * var(--_cell-width))');
      expect(cell(el, 2, 2)!.className).not.toContain('frozen');

      expect(shadow(el).querySelector('[data-header-row="1"]')!.className).toContain('frozen-row-edge');
      expect(shadow(el).querySelector('[data-header-col="1"]')!.className).toContain('frozen-col-edge');
      expect(shadow(el).querySelector<HTMLElement>('[data-header-col="1"]')!.style.left).toContain('150px');
    });

    it('keeps frozen rows and columns rendered when scrolled far away', async () => {
      const el = await createSpreadsheet(500, 60);
      el.freezePanes(2, 1);
      await scrollTo(el, 400 * 28, 40 * 100);
      expect(cell(el, 0, 0)).not.toBeNull();
      expect(cell(el, 1, 0)).not.toBeNull();
      expect(cell(el, 0, 45)).not.toBeNull(); // frozen row, scrolled column
      expect(cell(el, 405, 0)).not.toBeNull(); // scrolled row, frozen column
      expect(cell(el, 50, 10)).toBeNull();
    });

    it('draws no freeze line when nothing is frozen', async () => {
      const el = await createSpreadsheet();
      await el.updateComplete;
      expect(shadow(el).querySelectorAll('.frozen-row-edge, .frozen-col-edge')).toHaveLength(0);
      el.freezePanes(1, 0);
      await el.updateComplete;
      // Only row 1 (cells and header) carries the line, not the column headers
      expect(shadow(el).querySelectorAll('.ls-col-header.frozen-row-edge')).toHaveLength(0);
      expect(shadow(el).querySelectorAll('.frozen-col-edge')).toHaveLength(0);
    });

    it('clips the editor of a cell that scrolled under the frozen rows', async () => {
      const el = await createSpreadsheet(200, 10);
      el.freezePanes(3, 0);
      await select(el, 9, 1);
      await press(el, 'F2');
      const style = () => (shadow(el).querySelector('#editor') as HTMLInputElement).style;
      expect(style().clipPath).toBe('');
      // Row 10's top is now 2 rows above the bottom of the frozen area
      await scrollTo(el, 8 * 28, 0);
      expect(style().clipPath).toMatch(/^inset\(56px 0px 0px/);
    });

    it('keeps ARIA row and column indexes', async () => {
      const el = await createSpreadsheet();
      el.freezePanes(1, 1);
      await el.updateComplete;
      expect(cell(el, 0, 0)!.getAttribute('aria-colindex')).toBe('2');
      expect(cell(el, 0, 0)!.closest('[role="row"]')!.getAttribute('aria-rowindex')).toBe('2');
    });

    it('places the editor over a frozen cell regardless of the scroll position', async () => {
      const el = await createSpreadsheet(200, 40);
      el.freezePanes(1, 1);
      await scrollTo(el, 1000, 800);
      await select(el, 0, 0);
      await press(el, 'F2');
      const style = (shadow(el).querySelector('#editor') as HTMLInputElement).style;
      expect(style.top).toBe('28px');
      expect(style.left).toBe('50px');
    });
  });

  describe('keyboard', () => {
    it('Alt+Shift+F freezes at the active cell and again unfreezes, announcing each', async () => {
      const el = await createSpreadsheet();
      const events = listen<FreezeChangeDetail>(el, 'freeze-change');
      await select(el, 2, 1);
      await press(el, 'F', { altKey: true, shiftKey: true, code: 'KeyF' });
      expect([el.frozenRows, el.frozenCols]).toEqual([2, 1]);
      expect(announcement(el)).toBe('Froze rows 1 to 2 and column A');
      await press(el, 'F', { altKey: true, shiftKey: true, code: 'KeyF' });
      expect([el.frozenRows, el.frozenCols]).toEqual([0, 0]);
      expect(announcement(el)).toBe('Panes unfrozen');
      expect(events.map((e) => e.source)).toEqual(['user', 'user']);
    });

    it('explains why nothing is frozen at A1', async () => {
      const el = await createSpreadsheet();
      await press(el, 'F', { altKey: true, shiftKey: true, code: 'KeyF' });
      expect(el.frozenRows).toBe(0);
      expect(announcement(el)).toMatch(/^Nothing to freeze/);
    });

    it('works in read-only mode', async () => {
      const el = await createSpreadsheet();
      el.readOnly = true;
      await select(el, 1, 0);
      await press(el, 'F', { altKey: true, shiftKey: true, code: 'KeyF' });
      expect(el.frozenRows).toBe(1);
      expect(announcement(el)).toBe('Froze row 1');
    });

    it('Ctrl+Home goes to the first unfrozen cell', async () => {
      const el = await createSpreadsheet();
      el.freezePanes(1, 2);
      await select(el, 10, 5);
      await press(el, 'Home', { ctrlKey: true });
      expect(internals(el)._selection.activeCell).toEqual({ row: 1, col: 2 });
      el.unfreezePanes();
      await press(el, 'Home', { ctrlKey: true });
      expect(internals(el)._selection.activeCell).toEqual({ row: 0, col: 0 });
    });
  });

  describe('inserting and deleting rows and columns', () => {
    it('grows and shrinks the frozen area with the content, and restores it on undo', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'Header' }));
      el.freezePanes(2, 1);
      const events = listen<FreezeChangeDetail>(el, 'freeze-change');

      el.insertRows(1, 2); // inside the frozen rows
      expect(el.frozenRows).toBe(4);
      el.insertRows(4); // at the freeze line: below it
      expect(el.frozenRows).toBe(4);
      el.insertColumns(0); // before the frozen column
      expect(el.frozenCols).toBe(2);
      el.deleteRows(0, 3); // removes three frozen rows, including the header
      expect(el.frozenRows).toBe(1);
      el.deleteRows(5, 2); // below the line
      expect(el.frozenRows).toBe(1);

      expect(events.map((e) => [e.rows, e.cols, e.source])).toEqual([
        [4, 1, 'programmatic'],
        [4, 2, 'programmatic'],
        [1, 2, 'programmatic'],
      ]);

      internals(el)._undo(); // delete below the line
      internals(el)._undo(); // delete of three frozen rows
      expect(el.frozenRows).toBe(4);
      expect(events[events.length - 1]).toEqual({ rows: 4, cols: 2, source: 'undo' });
      internals(el)._redo();
      expect(el.frozenRows).toBe(1);
    });

    it('undo recomputes the freeze line from the current freeze instead of replaying old counts', async () => {
      const el = await createSpreadsheet();
      el.freezePanes(2, 0);
      el.insertRows(0); // 3 frozen
      el.unfreezePanes();
      internals(el)._undo(); // removing row 1 doesn't re-freeze anything
      expect(el.frozenRows).toBe(0);
      internals(el)._redo();
      expect(el.frozenRows).toBe(0);

      const other = await createSpreadsheet();
      other.insertRows(1);
      other.freezePanes(3, 0);
      internals(other)._undo(); // deletes row 2, inside the frozen area
      expect(other.frozenRows).toBe(2);
    });

    it('stores the clamped count, so it cannot come back and freeze other rows', async () => {
      const el = await createSpreadsheet(3, 4);
      el.freezePanes(2, 0);
      const events = listen<FreezeChangeDetail>(el, 'freeze-change');
      el.deleteRows(2); // two rows left: at most one can be frozen
      expect(el.frozenRows).toBe(1);
      expect(events).toEqual([{ rows: 1, cols: 0, source: 'programmatic' }]);
      el.insertRows(1); // below the freeze line
      expect(el.frozenRows).toBe(1);
    });

    it('a deletion spanning the freeze line keeps the rows above it', async () => {
      const el = await createSpreadsheet();
      el.freezePanes(3, 0);
      el.deleteRows(2, 4);
      expect(el.frozenRows).toBe(2);
    });
  });
});
