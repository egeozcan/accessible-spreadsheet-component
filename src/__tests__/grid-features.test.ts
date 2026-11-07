// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';
import type { CellData, DataChangeDetail, GridData } from '../types.js';
import type { Y11nSpreadsheet } from '../y11n-spreadsheet.js';
import '../y11n-spreadsheet.js';

function data(entries: Record<string, string>): GridData {
  const map: GridData = new Map();
  for (const [key, rawValue] of Object.entries(entries)) {
    map.set(key, { rawValue, displayValue: rawValue, type: 'text' } as CellData);
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

function grid(el: Y11nSpreadsheet): HTMLElement {
  return el.shadowRoot!.querySelector('.ls-grid') as HTMLElement;
}

async function press(el: Y11nSpreadsheet, key: string, mods: KeyboardEventInit = {}) {
  grid(el).dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, composed: true, ...mods }));
  await el.updateComplete;
}

async function pressInEditor(el: Y11nSpreadsheet, key: string, mods: KeyboardEventInit = {}) {
  const editor = el.shadowRoot!.querySelector('#editor') as HTMLInputElement;
  editor.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, composed: true, ...mods }));
  await el.updateComplete;
}

async function typeInEditor(el: Y11nSpreadsheet, value: string) {
  const editor = el.shadowRoot!.querySelector('#editor') as HTMLInputElement;
  editor.value = value;
  editor.dispatchEvent(new Event('input', { bubbles: true }));
  await el.updateComplete;
}

function raw(el: Y11nSpreadsheet, key: string): string {
  return el.getData().get(key)?.rawValue ?? '';
}

type Internals = {
  _selection: {
    activeCell: { row: number; col: number };
    range: { start: { row: number; col: number }; end: { row: number; col: number } };
    moveTo(row: number, col: number, extend?: boolean): void;
  };
};

const internals = (el: Y11nSpreadsheet) => el as unknown as Internals;

async function select(el: Y11nSpreadsheet, r1: number, c1: number, r2 = r1, c2 = c1) {
  internals(el)._selection.moveTo(r1, c1);
  internals(el)._selection.moveTo(r2, c2, true);
  el.requestUpdate();
  await el.updateComplete;
}

describe('grid features', () => {
  beforeEach(() => {
    document.body.innerHTML = '';
  });

  describe('Ctrl/Cmd+Arrow', () => {
    it('jumps to data edges and extends with Shift', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '2:0': 'a', '3:0': 'b', '4:0': 'c', '9:0': 'd' }));
      await el.updateComplete;

      await press(el, 'ArrowDown', { ctrlKey: true });
      expect(internals(el)._selection.activeCell).toEqual({ row: 2, col: 0 });
      await press(el, 'ArrowDown', { metaKey: true });
      expect(internals(el)._selection.activeCell).toEqual({ row: 4, col: 0 });
      await press(el, 'ArrowDown', { ctrlKey: true, shiftKey: true });
      expect(internals(el)._selection.range).toEqual({ start: { row: 4, col: 0 }, end: { row: 9, col: 0 } });
      await press(el, 'ArrowRight', { ctrlKey: true });
      expect(internals(el)._selection.activeCell).toEqual({ row: 9, col: 5 });
    });
  });

  describe('row/column selection', () => {
    it('Ctrl+Space selects columns and Shift+Space selects rows without moving the active cell', async () => {
      const el = await createSpreadsheet();
      await select(el, 3, 1, 3, 2);

      await press(el, ' ', { ctrlKey: true });
      expect(internals(el)._selection.range).toEqual({ start: { row: 0, col: 1 }, end: { row: 19, col: 2 } });
      expect(internals(el)._selection.activeCell).toEqual({ row: 3, col: 2 });

      await press(el, ' ', { shiftKey: true });
      expect(internals(el)._selection.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 19, col: 5 } });
    });

    it('Shift+Space does not start editing', async () => {
      const el = await createSpreadsheet();
      await press(el, ' ', { shiftKey: true });
      expect(el.shadowRoot!.querySelector<HTMLInputElement>('#editor')!.style.display).toBe('none');
      expect(internals(el)._selection.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 0, col: 5 } });
    });

    it('clicking a column header selects the column; Shift+click extends', async () => {
      const el = await createSpreadsheet();
      await select(el, 4, 0);
      const header = (c: number) =>
        el.shadowRoot!.querySelector(`[data-header-col="${c}"]`) as HTMLElement;

      header(2).dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0 }));
      document.dispatchEvent(new PointerEvent('pointerup'));
      await el.updateComplete;
      expect(internals(el)._selection.range).toEqual({ start: { row: 0, col: 2 }, end: { row: 19, col: 2 } });
      expect(internals(el)._selection.activeCell).toEqual({ row: 4, col: 2 });
      expect(header(2).getAttribute('aria-selected')).toBe('true');
      expect(header(1).hasAttribute('aria-selected')).toBe(false);

      header(4).dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0, shiftKey: true }));
      document.dispatchEvent(new PointerEvent('pointerup'));
      await el.updateComplete;
      expect(internals(el)._selection.range).toEqual({ start: { row: 0, col: 2 }, end: { row: 19, col: 4 } });
      expect(internals(el)._selection.activeCell).toEqual({ row: 4, col: 2 });
    });

    it('clicking a row header selects the row', async () => {
      const el = await createSpreadsheet();
      await select(el, 0, 3);
      const header = el.shadowRoot!.querySelector('[data-header-row="5"]') as HTMLElement;
      header.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0 }));
      document.dispatchEvent(new PointerEvent('pointerup'));
      await el.updateComplete;
      expect(internals(el)._selection.range).toEqual({ start: { row: 5, col: 0 }, end: { row: 5, col: 5 } });
      expect(internals(el)._selection.activeCell).toEqual({ row: 5, col: 3 });
    });

    it('announces whole-column and whole-row selections', async () => {
      const el = await createSpreadsheet();
      await press(el, ' ', { ctrlKey: true });
      const status = () => el.shadowRoot!.querySelector('[role="status"]')!.textContent!.trim();
      expect(status()).toBe('A1, Column A selected');
      await select(el, 1, 0, 2, 0);
      await press(el, ' ', { shiftKey: true });
      expect(status()).toBe('A3, Rows 2 to 3 selected');
    });
  });

  describe('selection after data changes', () => {
    it('the active cell follows the head again after an edit restores the selection', async () => {
      const el = await createSpreadsheet();
      el.fromCSV('x');
      await el.updateComplete;
      await press(el, 'ArrowRight', { shiftKey: true });
      expect(internals(el)._selection.activeCell).toEqual({ row: 0, col: 1 });
      await press(el, 'ArrowDown');
      expect(internals(el)._selection.activeCell).toEqual({ row: 1, col: 1 });
    });

    it('undo restores a whole-column selection with its active cell', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:2': 'v' }));
      await select(el, 4, 2);
      await press(el, ' ', { ctrlKey: true });
      await press(el, 'd', { ctrlKey: true });
      expect(raw(el, '10:2')).toBe('v');
      await press(el, 'z', { ctrlKey: true });
      expect(raw(el, '10:2')).toBe('');
      expect(internals(el)._selection.activeCell).toEqual({ row: 4, col: 2 });
      expect(internals(el)._selection.range).toEqual({ start: { row: 0, col: 2 }, end: { row: 19, col: 2 } });
    });
  });

  describe('fill down / right', () => {
    it('Ctrl+D copies the top row down, shifting relative references and copying formats', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': '1', '0:1': '=A1*2', '0:2': '=$A$1', '1:0': '2', '2:0': '3' }));
      el.setCellFormat('0:1', { bold: true });
      await select(el, 0, 1, 2, 2);

      const events: DataChangeDetail[] = [];
      el.addEventListener('data-change', (e) => events.push((e as CustomEvent<DataChangeDetail>).detail));
      await press(el, 'd', { ctrlKey: true });

      expect(raw(el, '1:1')).toBe('=A2*2');
      expect(raw(el, '2:1')).toBe('=A3*2');
      expect(raw(el, '2:2')).toBe('=$A$1');
      expect(el.getData().get('2:1')?.displayValue).toBe('6');
      expect(el.getCellFormat('2:1')).toEqual({ bold: true });
      expect(events[events.length - 1]?.operation).toBe('fill');
    });

    it('Ctrl+D on a single row copies from the row above', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'x', '0:1': 'y' }));
      await select(el, 1, 0, 1, 1);
      await press(el, 'd', { ctrlKey: true });
      expect(raw(el, '1:0')).toBe('x');
      expect(raw(el, '1:1')).toBe('y');
    });

    it('Ctrl+R copies the left column right', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': '=B5', '1:0': 'k' }));
      await select(el, 0, 0, 1, 2);
      await press(el, 'r', { ctrlKey: true });
      expect(raw(el, '0:2')).toBe('=D5');
      expect(raw(el, '1:1')).toBe('k');
    });

    it('fill clears target formats when the source has none, and undo restores everything', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'src', '1:0': 'old' }));
      el.setCellFormat('1:0', { italic: true });
      await select(el, 0, 0, 1, 0);
      await press(el, 'd', { ctrlKey: true });
      expect(raw(el, '1:0')).toBe('src');
      expect(el.getCellFormat('1:0')).toBeUndefined();

      await press(el, 'z', { ctrlKey: true });
      expect(raw(el, '1:0')).toBe('old');
      expect(el.getCellFormat('1:0')).toEqual({ italic: true });
    });

    it('does nothing in read-only mode', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'x' }));
      el.readOnly = true;
      await select(el, 0, 0, 2, 0);
      await press(el, 'd', { ctrlKey: true });
      expect(raw(el, '1:0')).toBe('');
    });
  });

  describe('editor shortcuts', () => {
    it('Ctrl+Enter writes the edit into every selected cell with relative formulas', async () => {
      const el = await createSpreadsheet();
      // Anchor at the bottom so the head (the active, edited cell) is B1
      await select(el, 2, 1, 0, 1);
      await press(el, 'F2');
      await typeInEditor(el, '=A1+1');
      await pressInEditor(el, 'Enter', { ctrlKey: true });

      expect(raw(el, '0:1')).toBe('=A1+1');
      expect(raw(el, '1:1')).toBe('=A2+1');
      expect(raw(el, '2:1')).toBe('=A3+1');
      // Selection and active cell stay put
      expect(internals(el)._selection.range).toEqual({ start: { row: 0, col: 1 }, end: { row: 2, col: 1 } });

      await press(el, 'z', { ctrlKey: true });
      expect(raw(el, '1:1')).toBe('');
    });

    it('Shift+Enter commits and moves up', async () => {
      const el = await createSpreadsheet();
      await select(el, 3, 0);
      await press(el, 'F2');
      await typeInEditor(el, 'hello');
      await pressInEditor(el, 'Enter', { shiftKey: true });
      expect(raw(el, '3:0')).toBe('hello');
      expect(internals(el)._selection.activeCell).toEqual({ row: 2, col: 0 });
    });
  });

  describe('CSV', () => {
    it('toCSV exports display values up to the used extent', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'name', '0:1': 'total', '1:0': 'a, b', '1:1': '=2*3', '2:2': 'x' }));
      expect(el.toCSV()).toBe('name,total,\n"a, b",6,\n,,x');
      expect(el.toCSV({ values: 'raw', delimiter: '\t', range: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } } }))
        .toBe('=2*3');
    });

    it('fromCSV rejects a negative or fractional start', async () => {
      const el = await createSpreadsheet();
      expect(() => el.fromCSV('x', { start: { row: -1, col: 0 } })).toThrow(RangeError);
      expect(() => el.fromCSV('x', { start: { row: 0.5, col: 0 } })).toThrow(RangeError);
      expect(el.getData().size).toBe(0);
    });

    it('toCSV of an empty grid is an empty string', async () => {
      const el = await createSpreadsheet();
      expect(el.toCSV()).toBe('');
    });

    it('fromCSV writes values and formulas, clamps to the grid, and is undoable', async () => {
      const el = await createSpreadsheet(3, 2);
      el.setData(data({ '0:0': 'keep' }));
      const events: DataChangeDetail[] = [];
      el.addEventListener('data-change', (e) => events.push((e as CustomEvent<DataChangeDetail>).detail));

      el.fromCSV('1,2,3\n"x,y",=A2*10\n5\n6', { start: { row: 1, col: 0 } });
      expect(raw(el, '0:0')).toBe('keep');
      expect(raw(el, '1:0')).toBe('1');
      expect(raw(el, '1:1')).toBe('2');
      expect(raw(el, '2:0')).toBe('x,y');
      expect(el.getData().get('2:1')?.displayValue).toBe('10');
      expect(el.getData().has('1:2')).toBe(false); // outside cols
      expect(events[events.length - 1]).toMatchObject({ operation: 'import', source: 'programmatic' });

      await press(el, 'z', { ctrlKey: true });
      expect(raw(el, '1:0')).toBe('');
      expect(raw(el, '0:0')).toBe('keep');
    });
  });

  describe('cell format styles', () => {
    it('renders verticalAlign and wrap', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'long text' }));
      el.setCellFormat('0:0', { verticalAlign: 'top', wrap: true });
      await el.updateComplete;
      const cell = el.shadowRoot!.querySelector('[data-row="0"][data-col="0"]') as HTMLElement;
      expect(cell.style.alignItems).toBe('flex-start');
      expect(cell.style.whiteSpace).toBe('normal');
    });

    it('renders the date number format', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': '=DATE(2024, 3, 9)' }));
      el.setCellFormat('0:0', { numberFormat: { type: 'date', datePattern: 'DD.MM.YYYY' } });
      await el.updateComplete;
      expect(el.getData().get('0:0')?.displayValue).toBe('09.03.2024');
    });
  });
});
