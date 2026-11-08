// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { CellData, ColumnResizeDetail, GridData } from '../types.js';
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

async function createSpreadsheet(entries: Record<string, string | Partial<CellData>> = {}, rows = 20, cols = 8) {
  const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
  el.rows = rows;
  el.cols = cols;
  document.body.appendChild(el);
  await el.updateComplete;
  if (Object.keys(entries).length > 0) {
    el.setData(data(entries));
    await el.updateComplete;
  }
  return el;
}

const shadow = (el: Y11nSpreadsheet) => el.shadowRoot!;
const grid = (el: Y11nSpreadsheet) => shadow(el).querySelector('.ls-grid') as HTMLElement;
const announcement = (el: Y11nSpreadsheet) => shadow(el).querySelector('[role="log"]')!.textContent!.trim();

async function press(el: Y11nSpreadsheet, key: string, mods: KeyboardEventInit = {}) {
  grid(el).dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, composed: true, ...mods }));
  await el.updateComplete;
}

type Internals = {
  _selection: {
    moveTo(row: number, col: number, extend?: boolean): void;
    selectColumns(from: number, to: number): void;
  };
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

function listen<T>(el: Y11nSpreadsheet, type: string): T[] {
  const events: T[] = [];
  el.addEventListener(type, (e) => events.push((e as CustomEvent<T>).detail));
  return events;
}

const editorStyle = (el: Y11nSpreadsheet) => (shadow(el).querySelector('#editor') as HTMLInputElement).style;

describe('column widths', () => {
  beforeEach(() => {
    document.body.innerHTML = '';
  });

  describe('API', () => {
    it('reports default and custom widths', async () => {
      const el = await createSpreadsheet();
      expect(el.getColumnWidth(0)).toBe(100);
      el.setColumnWidth(1, 150.4);
      expect(el.getColumnWidth(1)).toBe(150);
      expect(el.getColumnWidths()).toEqual({ 1: 150 });
      el.setColumnWidths({ 1: null, 3: 60 });
      expect(el.getColumnWidths()).toEqual({ 3: 60 });
    });

    it('clamps widths and rejects invalid input', async () => {
      const el = await createSpreadsheet();
      el.setColumnWidths({ 0: 1, 1: 99999 });
      expect(el.getColumnWidths()).toEqual({ 0: 20, 1: 2000 });
      expect(() => el.setColumnWidth(8, 50)).toThrow(RangeError);
      expect(() => el.setColumnWidth(-1, 50)).toThrow(RangeError);
      expect(() => el.setColumnWidth(0, 0)).toThrow(RangeError);
      expect(() => el.setColumnWidth(0, NaN)).toThrow(RangeError);
      expect(() => el.autoFitColumns([9])).toThrow(RangeError);
    });

    it('lays custom widths out in the CSS grid, keeping runs of default columns as repeat()', async () => {
      const el = await createSpreadsheet();
      el.setColumnWidths({ 2: 150, 3: 40 });
      await el.updateComplete;
      expect(grid(el).getAttribute('style')).toContain(
        'grid-template-columns: var(--_row-header-width) repeat(2, var(--_cell-width)) 150px 40px repeat(4, var(--_cell-width));'
      );
    });

    it('is undoable and fires column-resize with the effective widths', async () => {
      const el = await createSpreadsheet();
      const events = listen<ColumnResizeDetail>(el, 'column-resize');
      el.setColumnWidths({ 0: 160, 2: 80 });
      expect(events).toEqual([{ columns: [{ col: 0, width: 160 }, { col: 2, width: 80 }], source: 'programmatic' }]);

      internals(el)._undo();
      expect(el.getColumnWidths()).toEqual({});
      expect(events[1]).toEqual({ columns: [{ col: 0, width: 100 }, { col: 2, width: 100 }], source: 'undo' });

      internals(el)._redo();
      expect(el.getColumnWidths()).toEqual({ 0: 160, 2: 80 });
      expect(events[2].source).toBe('redo');
    });

    it('does nothing, and records no history, when no width changes', async () => {
      const el = await createSpreadsheet();
      const events = listen<ColumnResizeDetail>(el, 'column-resize');
      el.setColumnWidths({ 0: null });
      el.setColumnWidth(1, 120);
      el.setColumnWidth(1, 120);
      expect(events).toHaveLength(1);
      expect(internals(el)._undoStack).toHaveLength(1);
    });

    it('does not move focus or the selection when restoring saved widths', async () => {
      const el = await createSpreadsheet();
      const selections = listen(el, 'selection-change');
      const outside = document.createElement('button');
      document.body.appendChild(outside);
      outside.focus();
      el.setColumnWidths({ 0: 200 });
      await el.updateComplete;
      expect(document.activeElement).toBe(outside);
      expect(selections).toHaveLength(0);
    });

    it('moves custom widths with their columns on insert and delete, and back on undo', async () => {
      const el = await createSpreadsheet();
      el.setColumnWidths({ 1: 150, 4: 60 });
      el.insertColumns(2, 2);
      expect(el.getColumnWidths()).toEqual({ 1: 150, 6: 60 });
      el.deleteColumns(0, 2);
      expect(el.getColumnWidths()).toEqual({ 4: 60 });
      internals(el)._undo();
      internals(el)._undo();
      expect(el.getColumnWidths()).toEqual({ 1: 150, 4: 60 });
    });

    it('fires column-resize when widths move with inserted or deleted columns', async () => {
      const el = await createSpreadsheet();
      el.setColumnWidth(0, 180);
      const events = listen<ColumnResizeDetail>(el, 'column-resize');
      el.insertColumns(0);
      expect(events).toEqual([
        { columns: [{ col: 0, width: 100 }, { col: 1, width: 180 }], source: 'programmatic' },
      ]);
      internals(el)._undo();
      expect(events[1]).toEqual({ columns: [{ col: 0, width: 180 }, { col: 1, width: 100 }], source: 'undo' });
    });

    it('does not count as a data change that aborts a pending edit', async () => {
      const el = await createSpreadsheet();
      el.addEventListener('before-change', () => el.setColumnWidth(0, 150));
      await select(el, 0, 0);
      await press(el, 'F2');
      const editor = shadow(el).querySelector('#editor') as HTMLInputElement;
      editor.value = 'hello';
      editor.dispatchEvent(new Event('input', { bubbles: true }));
      await el.updateComplete;
      editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, composed: true }));
      await el.updateComplete;
      expect(el.getData().get('0:0')?.rawValue).toBe('hello');
      expect(el.getColumnWidth(0)).toBe(150);
    });

    it('leaves widths alone when rows are inserted', async () => {
      const el = await createSpreadsheet();
      el.setColumnWidth(1, 150);
      el.insertRows(0, 3);
      expect(el.getColumnWidths()).toEqual({ 1: 150 });
    });
  });

  describe('layout', () => {
    it('positions the inline editor over a column with a custom width', async () => {
      const el = await createSpreadsheet({ '0:2': 'x' });
      el.setColumnWidths({ 0: 150, 1: 40, 2: 222 });
      await select(el, 0, 2);
      await press(el, 'F2');
      // Row header (50) + 150 + 40
      expect(editorStyle(el).left).toBe('240px');
      expect(editorStyle(el).width).toBe('222px');
    });

    it('renders the columns under the scroll position', async () => {
      const el = await createSpreadsheet({}, 10, 60);
      el.setColumnWidths(Object.fromEntries(Array.from({ length: 30 }, (_, c) => [c, 400])));
      await el.updateComplete;
      const g = grid(el);
      g.scrollLeft = 400 * 25;
      g.dispatchEvent(new Event('scroll'));
      await el.updateComplete;
      const rendered = new Set(
        Array.from(shadow(el).querySelectorAll<HTMLElement>('[data-col]'), (c) => Number(c.dataset.col))
      );
      expect(rendered.has(25)).toBe(true);
      // Uniform-width math would have rendered columns around 10000 / 100 = 100 (beyond the grid)
      expect(rendered.has(0)).toBe(true); // the active cell is always rendered
      expect(rendered.has(10)).toBe(false);
    });
  });

  describe('keyboard', () => {
    it('Alt+Shift+Right / Left widen and narrow the active column and announce the width', async () => {
      const el = await createSpreadsheet();
      await select(el, 1, 1);
      await press(el, 'ArrowRight', { altKey: true, shiftKey: true });
      await press(el, 'ArrowRight', { altKey: true, shiftKey: true });
      expect(el.getColumnWidth(1)).toBe(116);
      expect(announcement(el)).toBe('Column B width 116 pixels');
      await press(el, 'ArrowLeft', { altKey: true, shiftKey: true });
      expect(el.getColumnWidth(1)).toBe(108);
      // The active cell didn't move
      expect(shadow(el).querySelector('.active-cell')?.getAttribute('data-key')).toBe('1:1');
    });

    it('resizes every selected column to the same width', async () => {
      const el = await createSpreadsheet();
      el.setColumnWidth(2, 60);
      await select(el, 0, 1, 3, 2);
      await press(el, 'ArrowLeft', { altKey: true, shiftKey: true });
      // From the active column (C, 60px)
      expect(el.getColumnWidths()).toEqual({ 1: 52, 2: 52 });
      expect(announcement(el)).toBe('Columns B to C width 52 pixels');
    });

    it('stops at the minimum width', async () => {
      const el = await createSpreadsheet();
      el.setColumnWidth(0, 24);
      await select(el, 0, 0);
      await press(el, 'ArrowLeft', { altKey: true, shiftKey: true });
      await press(el, 'ArrowLeft', { altKey: true, shiftKey: true });
      expect(el.getColumnWidth(0)).toBe(20);
      expect(announcement(el)).toBe('Column A width 20 pixels');
    });

    it('works in read-only mode, since widths are not data', async () => {
      const el = await createSpreadsheet();
      el.readOnly = true;
      await el.updateComplete;
      await select(el, 0, 0);
      await press(el, 'ArrowRight', { altKey: true, shiftKey: true });
      expect(el.getColumnWidth(0)).toBe(108);
    });

    it('can be vetoed with before-column-resize', async () => {
      const el = await createSpreadsheet();
      const seen: ColumnResizeDetail[] = [];
      el.addEventListener('before-column-resize', (e) => {
        seen.push((e as CustomEvent<ColumnResizeDetail>).detail);
        e.preventDefault();
      });
      await select(el, 0, 3);
      await press(el, 'ArrowRight', { altKey: true, shiftKey: true });
      expect(seen).toEqual([{ columns: [{ col: 3, width: 108 }], source: 'user' }]);
      expect(el.getColumnWidth(3)).toBe(100);
      expect(announcement(el)).toBe('Resize cancelled');
    });

    it('keeps announcing widths after a vetoed resize', async () => {
      const el = await createSpreadsheet();
      el.addEventListener('before-column-resize', (e) => {
        const { columns } = (e as CustomEvent<ColumnResizeDetail>).detail;
        if (columns.some((c) => c.width > 108)) e.preventDefault();
      });
      await select(el, 0, 1);
      await press(el, 'ArrowRight', { altKey: true, shiftKey: true });
      await press(el, 'ArrowRight', { altKey: true, shiftKey: true });
      expect(announcement(el)).toBe('Resize cancelled');
      await press(el, 'ArrowLeft', { altKey: true, shiftKey: true });
      expect(announcement(el)).toBe('Column B width 100 pixels');
    });

    it('Ctrl+Z undoes a resize', async () => {
      const el = await createSpreadsheet();
      await select(el, 0, 0);
      await press(el, 'ArrowRight', { altKey: true, shiftKey: true });
      await press(el, 'z', { ctrlKey: true });
      expect(el.getColumnWidth(0)).toBe(100);
    });
  });

  describe('auto-fit', () => {
    // jsdom has no canvas: measure 7px per character, 8px when bold
    beforeEach(() => {
      vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function () {
        const ctx = {
          font: '',
          measureText: (text: string) => ({ width: text.length * (ctx.font.includes('bold') ? 8 : 7) }),
        };
        return ctx as unknown as CanvasRenderingContext2D;
      } as never);
    });
    afterEach(() => vi.restoreAllMocks());

    it('fits the widest displayed value, including rows outside the rendered window', async () => {
      const el = await createSpreadsheet({ '0:0': 'short', '150:0': 'a much longer value', '1:1': 'x' }, 200);
      el.autoFitColumns([0]);
      // 19 characters x 7px + 15px of padding
      expect(el.getColumnWidth(0)).toBe(19 * 7 + 15);
    });

    it('measures with the cell font and resets an empty column to the default', async () => {
      const el = await createSpreadsheet({ '0:0': { rawValue: 'Total', format: { bold: true } } });
      el.setColumnWidth(1, 300);
      el.autoFitColumns([0, 1]);
      expect(el.getColumnWidth(0)).toBe(5 * 8 + 15);
      expect(el.getColumnWidths()).toEqual({ 0: 55 });
    });

    it('Alt+Shift+0 fits the selected columns', async () => {
      const el = await createSpreadsheet({ '0:0': 'abc', '0:1': 'abcdefghij' });
      await select(el, 0, 0, 0, 1);
      await press(el, ')', { altKey: true, shiftKey: true, code: 'Digit0' });
      expect(el.getColumnWidths()).toEqual({ 0: 3 * 7 + 15, 1: 10 * 7 + 15 });
      expect(announcement(el)).toBe('Columns A to B resized to fit their content');
    });
  });

  describe('pointer', () => {
    const handle = (el: Y11nSpreadsheet, col: number) =>
      shadow(el).querySelector(`[data-resize-col="${col}"]`) as HTMLElement;
    const pointer = (type: string, clientX: number) =>
      new PointerEvent(type, { bubbles: true, composed: true, button: 0, clientX });

    it('dragging a header border previews and then commits the width', async () => {
      const el = await createSpreadsheet();
      const events = listen<ColumnResizeDetail>(el, 'column-resize');
      handle(el, 1).dispatchEvent(pointer('pointerdown', 300));
      document.dispatchEvent(pointer('pointermove', 345));
      await el.updateComplete;
      expect(el.getColumnWidth(1)).toBe(145);
      expect(events).toHaveLength(0); // still a preview
      document.dispatchEvent(pointer('pointerup', 345));
      await el.updateComplete;
      expect(el.getColumnWidths()).toEqual({ 1: 145 });
      expect(events).toHaveLength(1);
      expect(announcement(el)).toBe('Column B width 145 pixels');
      // Dragging the border doesn't select the column
      expect(shadow(el).querySelector('[data-header-col="1"]')?.getAttribute('aria-selected')).toBeNull();
    });

    it('Escape during a drag restores the width', async () => {
      const el = await createSpreadsheet();
      handle(el, 0).dispatchEvent(pointer('pointerdown', 100));
      document.dispatchEvent(pointer('pointermove', 20));
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      document.dispatchEvent(pointer('pointerup', 20));
      await el.updateComplete;
      expect(el.getColumnWidths()).toEqual({});
    });

    it('resizes all selected whole columns together', async () => {
      const el = await createSpreadsheet();
      internals(el)._selection.selectColumns(1, 3);
      handle(el, 2).dispatchEvent(pointer('pointerdown', 0));
      document.dispatchEvent(pointer('pointermove', -30));
      document.dispatchEvent(pointer('pointerup', -30));
      expect(el.getColumnWidths()).toEqual({ 1: 70, 2: 70, 3: 70 });
    });
  });
});
