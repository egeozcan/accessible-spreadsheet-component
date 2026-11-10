// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { CellData, GridData } from '../types.js';
import { dateToSerial } from '../types.js';
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
const announcement = (el: Y11nSpreadsheet) => shadow(el).querySelector('[role="log"]')!.textContent!.trim();
const statusBar = (el: Y11nSpreadsheet) =>
  [...shadow(el).querySelectorAll('.ls-status span')].map((s) => s.textContent).join(' ');

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

const cell = (el: Y11nSpreadsheet, key: string) => el.getData().get(key);
const raw = (el: Y11nSpreadsheet, key: string) => cell(el, key)?.rawValue ?? '';
const display = (el: Y11nSpreadsheet, key: string) => cell(el, key)?.displayValue ?? '';

type Internals = {
  _selection: { moveTo(row: number, col: number, extend?: boolean): void };
  _undo(): void;
};
const internals = (el: Y11nSpreadsheet) => el as unknown as Internals;

async function select(el: Y11nSpreadsheet, r1: number, c1: number, r2 = r1, c2 = c1) {
  internals(el)._selection.moveTo(r1, c1);
  internals(el)._selection.moveTo(r2, c2, true);
  el.requestUpdate();
  await el.updateComplete;
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

describe('quick wins', () => {
  beforeEach(() => {
    document.body.innerHTML = '';
  });

  describe('typed input', () => {
    it('stores percentages, currency, grouped numbers and dates as numbers with a format', async () => {
      const el = await createSpreadsheet();
      await select(el, 0, 0);
      await enterValue(el, '15%');
      expect(raw(el, '0:0')).toBe('0.15');
      expect(display(el, '0:0')).toBe('15%');
      expect(cell(el, '0:0')?.type).toBe('number');

      await enterValue(el, '$1,200.5');
      expect(raw(el, '1:0')).toBe('1200.5');
      expect(display(el, '1:0')).toBe('$1,200.50');

      await enterValue(el, '2026-01-05');
      expect(raw(el, '2:0')).toBe(String(dateToSerial(2026, 0, 5)));
      expect(display(el, '2:0')).toBe('2026-01-05');

      await select(el, 3, 0);
      await enterValue(el, '=A1*2');
      expect(display(el, '3:0')).toBe('0.3');
    });

    it('edits percent and date cells as displayed, and undo restores value and format together', async () => {
      const el = await createSpreadsheet();
      await enterValue(el, '12.5%');
      await select(el, 0, 0);
      await press(el, 'F2');
      expect(editor(el).value).toBe('12.5%');
      await pressInEditor(el, 'Escape');

      internals(el)._undo();
      await el.updateComplete;
      expect(cell(el, '0:0')).toBeUndefined();
    });

    it('reads a plain number typed into a percent cell as a percentage, like Excel', async () => {
      const el = await createSpreadsheet();
      await enterValue(el, '15%');
      await select(el, 0, 0);
      await enterValue(el, '13');
      expect(raw(el, '0:0')).toBe('0.13');
      expect(display(el, '0:0')).toBe('13%');
      await select(el, 0, 0);
      await enterValue(el, '0.5');
      expect(display(el, '0:0')).toBe('50%');
      await select(el, 0, 0);
      await enterValue(el, '-2.5');
      expect(raw(el, '0:0')).toBe('-0.025');
      // Formulas and text are left alone
      await select(el, 0, 0);
      await enterValue(el, '=1+1');
      expect(display(el, '0:0')).toBe('200%');

      // Ctrl+Enter: only the percent cells convert
      el.setCellFormat('5:0', { numberFormat: { type: 'percent', decimals: 0 } });
      await el.updateComplete;
      await select(el, 5, 0, 5, 1);
      await press(el, 'F2');
      await typeInEditor(el, '20');
      await pressInEditor(el, 'Enter', { ctrlKey: true });
      expect([raw(el, '5:0'), raw(el, '5:1')]).toEqual(['0.2', '20']);
    });

    it('keeps an existing number format of the same type, such as a custom date pattern', async () => {
      const el = await createSpreadsheet();
      el.setCellFormat('0:0', { numberFormat: { type: 'date', datePattern: 'DD/MM/YYYY' }, bold: true });
      await el.updateComplete;
      await enterValue(el, '2026-03-04');
      expect(display(el, '0:0')).toBe('04/03/2026');
      expect(cell(el, '0:0')?.format).toEqual({ numberFormat: { type: 'date', datePattern: 'DD/MM/YYYY' }, bold: true });
    });

    it('validates the parsed value, and raw-input turns parsing off', async () => {
      const el = await createSpreadsheet();
      const seen: string[] = [];
      el.validateCell = (_id, value) => {
        seen.push(value);
        return null;
      };
      await enterValue(el, '50%');
      expect(seen).toEqual(['0.5']);

      el.rawInput = true;
      await el.updateComplete;
      await enterValue(el, '50%');
      expect(raw(el, '1:0')).toBe('50%');
      expect(cell(el, '1:0')?.type).toBe('text');
    });

    it('Ctrl+Enter gives every selected cell the parsed value and format', async () => {
      const el = await createSpreadsheet();
      await select(el, 0, 0, 1, 1);
      await press(el, 'F2');
      await typeInEditor(el, '$5');
      await pressInEditor(el, 'Enter', { ctrlKey: true });
      for (const key of ['0:0', '0:1', '1:0', '1:1']) {
        expect(raw(el, key)).toBe('5');
        expect(display(el, key)).toBe('$5');
      }
    });
  });

  describe('fill series', () => {
    it('continues day names, numbered text and single dates with the fill', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'Mon', '0:1': 'Item 1', '0:2': { rawValue: '46000', type: 'number' } }));
      el.setCellFormat('0:2', { numberFormat: { type: 'date' } });
      await select(el, 0, 0, 2, 2);
      await press(el, 'd', { ctrlKey: true, altKey: true, code: 'KeyD' });
      expect([raw(el, '1:0'), raw(el, '2:0')]).toEqual(['Tue', 'Wed']);
      expect([raw(el, '1:1'), raw(el, '2:1')]).toEqual(['Item 2', 'Item 3']);
      expect([raw(el, '1:2'), raw(el, '2:2')]).toEqual(['46001', '46002']);
    });

    it('Ctrl+Alt+D on filled cells fills down as far as the column to the left goes', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'a', '1:0': 'b', '2:0': 'c', '3:0': 'd', '0:1': '1', '1:1': '2' }));
      await select(el, 0, 1, 1, 1);
      await press(el, 'd', { ctrlKey: true, altKey: true, code: 'KeyD' });
      expect([raw(el, '2:1'), raw(el, '3:1'), raw(el, '4:1')]).toEqual(['3', '4', '']);
    });

    it('a fill-handle double-click stops above data already in the column', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': 'a', '1:0': 'b', '2:0': 'c', '3:0': 'd', '0:1': 'x', '3:1': 'keep' }));
      await select(el, 0, 1);
      const handle = shadow(el).querySelector('.ls-fill-handle') as HTMLElement;
      handle.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, composed: true }));
      await el.updateComplete;
      expect([raw(el, '1:1'), raw(el, '2:1'), raw(el, '3:1')]).toEqual(['x', 'x', 'keep']);
      expect(isEditing(el)).toBe(false);
    });
  });

  describe('shortcuts', () => {
    it('Alt+= on a cell starts a SUM of the numbers above it, or else to its left', async () => {
      const el = await createSpreadsheet();
      el.setData(data({
        '0:0': 'Header',
        '1:0': { rawValue: '1', type: 'number' },
        '2:0': { rawValue: '2', type: 'number' },
        '5:0': { rawValue: '3', type: 'number' },
        '5:1': { rawValue: '4', type: 'number' },
      }));
      await select(el, 3, 0);
      await press(el, '=', { altKey: true, code: 'Equal' });
      expect(isEditing(el)).toBe(true);
      expect(editor(el).value).toBe('=SUM(A2:A3)');
      await pressInEditor(el, 'Enter');
      expect(display(el, '3:0')).toBe('3');

      await select(el, 5, 2);
      await press(el, '=', { altKey: true, code: 'Equal' });
      expect(editor(el).value).toBe('=SUM(A6:B6)');
    });

    it('Alt+= on a range writes a SUM below each column', async () => {
      const el = await createSpreadsheet();
      el.setData(data({
        '0:0': { rawValue: '1', type: 'number' },
        '1:0': { rawValue: '2', type: 'number' },
        '0:1': { rawValue: '3', type: 'number' },
        '1:1': { rawValue: '4', type: 'number' },
      }));
      await select(el, 0, 0, 1, 1);
      await press(el, '=', { altKey: true, code: 'Equal' });
      expect([raw(el, '2:0'), raw(el, '2:1')]).toEqual(['=SUM(A1:A2)', '=SUM(B1:B2)']);
      expect(display(el, '2:1')).toBe('7');

      // With an empty last row selected, the sums go into it
      await select(el, 0, 0, 3, 0);
      el.setData(data({ '0:0': { rawValue: '5', type: 'number' }, '1:0': { rawValue: '6', type: 'number' } }));
      await select(el, 0, 0, 2, 0);
      await press(el, '=', { altKey: true, code: 'Equal' });
      expect(raw(el, '2:0')).toBe('=SUM(A1:A2)');
    });

    it("Ctrl+; enters today's date, which is stored as a date", async () => {
      const el = await createSpreadsheet();
      await press(el, ';', { ctrlKey: true, code: 'Semicolon' });
      expect(isEditing(el)).toBe(true);
      expect(editor(el).value).toMatch(/^\d{4}-\d{2}-\d{2}$/);
      await pressInEditor(el, 'Enter');
      expect(cell(el, '0:0')?.type).toBe('number');
      expect(cell(el, '0:0')?.format?.numberFormat?.type).toBe('date');
    });

    it('Ctrl+Shift+; inserts the time at the caret while editing', async () => {
      const el = await createSpreadsheet();
      await press(el, 'F2');
      await typeInEditor(el, 'at ');
      editor(el).setSelectionRange(3, 3);
      await pressInEditor(el, ':', { ctrlKey: true, shiftKey: true, code: 'Semicolon' });
      expect(editor(el).value).toMatch(/^at \d{2}:\d{2}$/);
    });

    it("Ctrl+' edits with the formula above, Ctrl+Shift+' with its value", async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': { rawValue: '2', type: 'number' }, '1:0': '=A1*3' }));
      await select(el, 2, 0);
      await press(el, "'", { ctrlKey: true, code: 'Quote' });
      expect(editor(el).value).toBe('=A1*3');
      await pressInEditor(el, 'Escape');
      await press(el, '"', { ctrlKey: true, shiftKey: true, code: 'Quote' });
      expect(editor(el).value).toBe('6');
    });

    it('Ctrl+Shift+4/5/1/3/` set and clear number formats on the selection, undoably', async () => {
      const el = await createSpreadsheet();
      el.setData(data({ '0:0': { rawValue: '0.5', type: 'number' }, '1:0': { rawValue: '1234', type: 'number' } }));
      await select(el, 0, 0, 1, 0);
      await press(el, '%', { ctrlKey: true, shiftKey: true, code: 'Digit5' });
      expect(display(el, '0:0')).toBe('50%');
      expect(announcement(el)).toBe('Percent format');
      await press(el, '$', { ctrlKey: true, shiftKey: true, code: 'Digit4' });
      expect(display(el, '1:0')).toBe('$1,234.00');
      await press(el, '!', { ctrlKey: true, shiftKey: true, code: 'Digit1' });
      expect(display(el, '1:0')).toBe('1,234.00');
      await press(el, '#', { ctrlKey: true, shiftKey: true, code: 'Digit3' });
      expect(cell(el, '1:0')?.format?.numberFormat?.type).toBe('date');
      await press(el, '~', { ctrlKey: true, shiftKey: true, code: 'Backquote' });
      expect(display(el, '1:0')).toBe('1234');
      internals(el)._undo();
      await el.updateComplete;
      expect(cell(el, '1:0')?.format?.numberFormat?.type).toBe('date');
    });
  });

  describe('status bar', () => {
    it('shows and announces the Sum, Average and Count of the selection', async () => {
      const el = await createSpreadsheet();
      el.setData(data({
        '0:0': { rawValue: '10', type: 'number' },
        '1:0': { rawValue: '20', type: 'number' },
        '2:0': 'text',
        '3:0': '=A1+A2',
      }));
      await el.updateComplete;
      await select(el, 0, 0);
      expect(statusBar(el)).toBe('');
      await select(el, 0, 0, 4, 0);
      expect(statusBar(el)).toBe('Average: 20 Count: 4 Sum: 60');
      await press(el, 'S', { altKey: true, shiftKey: true, code: 'KeyS' });
      expect(announcement(el)).toBe('Sum 60, Average 20, Count 4');

      await select(el, 2, 0, 2, 3);
      expect(statusBar(el)).toBe('');
    });

    it('uses the first number cell’s format, covers whole-column selections and skips hidden rows', async () => {
      const el = await createSpreadsheet();
      el.setData(data({
        '0:0': 'Name',
        '1:0': { rawValue: '1.5', type: 'number' },
        '2:0': { rawValue: '2', type: 'number' },
        '3:0': { rawValue: '100', type: 'number' },
      }));
      el.setCellFormat('1:0', { numberFormat: { type: 'currency' } });
      await select(el, 0, 0, 19, 0);
      expect(statusBar(el)).toBe('Average: $34.50 Count: 4 Sum: $103.50');
      el.setFilter({ start: { row: 0, col: 0 }, end: { row: 3, col: 0 } }, [{ col: 0, values: ['$1.50', '2'] }]);
      await select(el, 0, 0, 19, 0);
      expect(statusBar(el)).toBe('Average: $1.75 Count: 3 Sum: $3.50');
    });

    it('is hidden with hide-status-bar', async () => {
      const el = await createSpreadsheet();
      el.hideStatusBar = true;
      await el.updateComplete;
      expect(shadow(el).querySelector('.ls-status')).toBeNull();
    });
  });

  describe('paste special', () => {
    let clipboardText = '';
    beforeEach(() => {
      clipboardText = '';
      Object.defineProperty(navigator, 'clipboard', {
        configurable: true,
        value: {
          writeText: async (t: string) => {
            clipboardText = t;
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
    afterEach(() => {
      Reflect.deleteProperty(navigator, 'clipboard');
    });

    const source = () =>
      data({
        '0:0': { rawValue: '2', type: 'number', format: { bold: true } },
        '0:1': '=A1*10',
        '1:0': 'x',
      });

    it('Ctrl+Shift+V pastes values instead of formulas and keeps the target formats', async () => {
      const el = await createSpreadsheet();
      el.setData(source());
      el.setCellFormat('5:0', { italic: true });
      await select(el, 0, 0, 1, 1);
      await press(el, 'c', { ctrlKey: true });
      await settle();
      await select(el, 5, 0);
      await press(el, 'V', { ctrlKey: true, shiftKey: true });
      await settle();
      await el.updateComplete;
      expect([raw(el, '5:0'), raw(el, '5:1'), raw(el, '6:0')]).toEqual(['2', '20', 'x']);
      expect(cell(el, '5:0')?.format).toEqual({ italic: true });
      expect(announcement(el)).toBe('Pasted values');
    });

    it('paste({ transpose }) swaps rows and columns and shifts formulas by how far they moved', async () => {
      const el = await createSpreadsheet();
      el.setData(source());
      await select(el, 0, 0, 1, 1);
      await press(el, 'c', { ctrlKey: true });
      await settle();
      await select(el, 5, 3);
      await el.paste({ transpose: true });
      await el.updateComplete;
      // A1 → D6, B1 → D7 (6 rows down, 2 columns right), A2 → E6
      expect([raw(el, '5:3'), raw(el, '6:3'), raw(el, '5:4')]).toEqual(['2', '=C7*10', 'x']);
    });

    it('paste({ mode: "formats" }) copies only formats, and transposes external text', async () => {
      const el = await createSpreadsheet();
      el.setData(source());
      el.setData(new Map([...source(), ['5:0', { rawValue: 'keep', displayValue: 'keep', type: 'text' }]]));
      await select(el, 0, 0, 0, 0);
      await press(el, 'c', { ctrlKey: true });
      await settle();
      await select(el, 5, 0);
      await el.paste({ mode: 'formats' });
      await el.updateComplete;
      expect(raw(el, '5:0')).toBe('keep');
      expect(cell(el, '5:0')?.format).toEqual({ bold: true });

      clipboardText = 'a\tb\nc\td';
      await select(el, 10, 0);
      await el.paste({ transpose: true });
      await el.updateComplete;
      expect([raw(el, '10:0'), raw(el, '10:1'), raw(el, '11:0'), raw(el, '11:1')]).toEqual(['a', 'c', 'b', 'd']);
    });

    it('Ctrl+Alt+V opens an accessible Paste Special dialog that pastes the chosen options', async () => {
      const el = await createSpreadsheet();
      el.setData(source());
      await select(el, 0, 0, 1, 1);
      await press(el, 'c', { ctrlKey: true });
      await settle();
      await select(el, 5, 0);
      await press(el, 'v', { ctrlKey: true, altKey: true, code: 'KeyV' });
      const dialog = shadow(el).querySelector('.ls-paste-special') as HTMLElement;
      expect(dialog.getAttribute('role')).toBe('dialog');
      expect(shadow(el).activeElement).toBe(dialog.querySelector('input[value="all"]'));
      (dialog.querySelector('input[value="values"]') as HTMLInputElement).click();
      await el.updateComplete;
      dialog.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, composed: true }));
      await settle();
      await el.updateComplete;
      expect(shadow(el).querySelector('.ls-paste-special')).toBeNull();
      expect(raw(el, '5:1')).toBe('20');

      await press(el, 'v', { ctrlKey: true, altKey: true, code: 'KeyV' });
      const again = shadow(el).querySelector('.ls-paste-special') as HTMLElement;
      again.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, composed: true }));
      await el.updateComplete;
      expect(shadow(el).querySelector('.ls-paste-special')).toBeNull();
      expect(announcement(el)).toBe('Paste Special closed');
    });
  });
});
