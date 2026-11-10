// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';
import type { CellData, GridData, SheetsChangeDetail } from '../types.js';
import type { Y11nSpreadsheet } from '../y11n-spreadsheet.js';
import '../y11n-spreadsheet.js';
import { ClipboardManager } from '../controllers/clipboard-manager.js';

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
  el.data = data(entries);
  document.body.appendChild(el);
  await el.updateComplete;
  return el;
}

const shadow = (el: Y11nSpreadsheet) => el.shadowRoot!;
const grid = (el: Y11nSpreadsheet) => shadow(el).querySelector('.ls-grid') as HTMLElement;
const editor = (el: Y11nSpreadsheet) => shadow(el).querySelector('#editor') as HTMLInputElement;
const cellEl = (el: Y11nSpreadsheet, key: string) => shadow(el).querySelector(`[data-key="${key}"]`) as HTMLElement | null;
const announcement = (el: Y11nSpreadsheet) => shadow(el).querySelector('[role="log"]')!.textContent!.trim();
const status = (el: Y11nSpreadsheet) => shadow(el).querySelector('[role="status"][aria-live]')!.textContent!.trim();
const cell = (el: Y11nSpreadsheet, key: string) => el.getData().get(key);
const raw = (el: Y11nSpreadsheet, key: string) => cell(el, key)?.rawValue ?? '';
const display = (el: Y11nSpreadsheet, key: string) => cell(el, key)?.displayValue ?? '';

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

type Internals = {
  _selection: { moveTo(row: number, col: number, extend?: boolean): void };
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

beforeEach(() => {
  document.body.innerHTML = '';
});

describe('spilled array results', () => {
  it('a formula that returns an array fills the cells below it', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '1:0': '2', '2:0': '3', '0:2': '=A1:A3*2' });
    expect(display(el, '0:2')).toBe('2');
    expect(cell(el, '0:2')?.spill).toEqual({ rows: 3, cols: 1 });
    expect(display(el, '1:2')).toBe('4');
    expect(cell(el, '1:2')).toMatchObject({ rawValue: '', spilledFrom: '0:2', type: 'number' });
    expect(display(el, '2:2')).toBe('6');
    expect(cellEl(el, '1:2')!.textContent).toContain('4');
  });

  it('typing into the spill range blocks it, and undo brings it back', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '1:0': '2', '0:2': '=A1:A2*2' });
    await select(el, 1, 2);
    await enterValue(el, 'x');
    expect(display(el, '0:2')).toBe('#SPILL!');
    expect(raw(el, '1:2')).toBe('x');
    internals(el)._undo();
    await el.updateComplete;
    expect(display(el, '0:2')).toBe('2');
    expect(display(el, '1:2')).toBe('4');
    expect(cell(el, '1:2')?.spilledFrom).toBe('0:2');
  });

  it('spilled values follow their inputs and feed other formulas', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '1:0': '2', '0:2': '=A1:A2*2', '0:4': '=SUM(C1:C2)', '1:4': '=C2+1' });
    expect(display(el, '0:4')).toBe('6');
    await select(el, 1, 0);
    await enterValue(el, '10');
    expect(display(el, '1:2')).toBe('20');
    expect(display(el, '0:4')).toBe('22');
    expect(display(el, '1:4')).toBe('21');
  });

  it('number formats apply to spilled cells', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '1:0': '2', '0:2': '=A1:A2/4' });
    el.setCellFormat('1:2', { numberFormat: { type: 'percent', decimals: 0 } });
    await el.updateComplete;
    expect(display(el, '1:2')).toBe('50%');
    expect(cell(el, '1:2')?.spilledFrom).toBe('0:2');
    await select(el, 1, 0);
    await enterValue(el, '3');
    expect(display(el, '1:2')).toBe('75%');
    el.clearCellFormat('1:2');
    await el.updateComplete;
    expect(display(el, '1:2')).toBe('0.75');
  });

  it('a spill that would leave the grid is #SPILL!', async () => {
    const el = await createSpreadsheet({ '0:0': '=SEQUENCE(5)' }, 3, 3);
    expect(display(el, '0:0')).toBe('#SPILL!');
    expect(status(el)).toBe('A1: #SPILL!, the result does not fit on the sheet');
    el.rows = 6;
    await el.updateComplete;
    expect(display(el, '0:0')).toBe('1');
    expect(display(el, '4:0')).toBe('5');
  });

  it('inserting a row above moves the formula and its spilled values', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '1:0': '2', '0:2': '=A1:A2*2' });
    el.insertRows(0, 1);
    await el.updateComplete;
    expect(raw(el, '1:2')).toBe('=A2:A3*2');
    expect(display(el, '1:2')).toBe('2');
    expect(display(el, '2:2')).toBe('4');
    expect(cell(el, '0:2')).toBeUndefined();
    internals(el)._undo();
    await el.updateComplete;
    expect(display(el, '0:2')).toBe('2');
    expect(display(el, '1:2')).toBe('4');
    expect(cell(el, '2:2')).toBeUndefined();
  });

  it('setData does not import stale spilled cells', async () => {
    const el = await createSpreadsheet({ '0:0': '=SEQUENCE(3)' });
    const snapshot = el.getData();
    expect(snapshot.get('2:0')?.spilledFrom).toBe('0:0');
    snapshot.get('0:0')!.rawValue = '=SEQUENCE(2)';
    el.setData(snapshot);
    await el.updateComplete;
    expect(display(el, '1:0')).toBe('2');
    expect(cell(el, '2:0')).toBeUndefined();
  });

  it('sorting leaves spilled values to their formula', async () => {
    const el = await createSpreadsheet({ '0:0': '3', '1:0': '1', '2:0': '2', '0:2': '=A1:A3*10' });
    el.sort({ start: { row: 0, col: 0 }, end: { row: 2, col: 0 } }, [{ col: 0 }]);
    await el.updateComplete;
    expect([display(el, '0:2'), display(el, '1:2'), display(el, '2:2')]).toEqual(['10', '20', '30']);
  });

  it('announces where a spilled value comes from', async () => {
    const el = await createSpreadsheet({ '0:0': '=SEQUENCE(2)' });
    await select(el, 1, 0);
    expect(status(el)).toBe('A2: 2, spilled from A1');
    await select(el, 0, 0);
    expect(status(el)).toBe('A1: 1, spills to A2');
    expect(cellEl(el, '1:0')!.className).toContain('in-spill');
  });

  it('a custom function can return an array, and take arguments element by element', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '1:0': '2', '0:2': '=PAIR(A1)', '0:5': '=TWICE(A1:A2)' });
    el.registerFunction('PAIR', (_ctx, v) => [[v, Number(v) * 10]]);
    el.registerFunction('TWICE', (_ctx, v) => Number(v) * 2, { args: ['number'] });
    await el.updateComplete;
    expect([display(el, '0:2'), display(el, '0:3')]).toEqual(['1', '10']);
    expect([display(el, '0:5'), display(el, '1:5')]).toEqual(['2', '4']);
  });

  it('copying a spilled value and exporting CSV use the values', async () => {
    const el = await createSpreadsheet({ '0:0': '=SEQUENCE(2,2)' });
    expect(el.toCSV()).toBe('1,2\n3,4');
  });
});

describe('several values in one cell', () => {
  it('Ctrl+J turns comma-separated text into a list and back', async () => {
    const el = await createSpreadsheet({ '0:0': 'Ann, Bob, 3' });
    await select(el, 0, 0);
    await press(el, 'j', { ctrlKey: true, code: 'KeyJ' });
    expect(raw(el, '0:0')).toBe('={{"Ann";"Bob";3}}');
    expect(display(el, '0:0')).toBe('Ann, Bob, 3');
    expect(el.getCellValues('0:0')).toEqual({ rows: 3, cols: 1, values: ['Ann', 'Bob', 3] });
    expect(announcement(el)).toBe('Turned 1 cell into a list');
    expect(status(el)).toContain('list of 3 items');
    await press(el, 'j', { ctrlKey: true, code: 'KeyJ' });
    expect(raw(el, '0:0')).toBe('Ann, Bob, 3');
    expect(el.getCellValues('0:0')).toBeNull();
  });

  it('a list is edited as its items', async () => {
    const el = await createSpreadsheet({ '0:0': '={{"Ann";"Bob"}}' });
    await select(el, 0, 0);
    await press(el, 'F2');
    expect(editor(el).value).toBe('Ann, Bob');
    await typeInEditor(el, 'Ann, Bob, Cy');
    await pressInEditor(el, 'Enter');
    expect(raw(el, '0:0')).toBe('={{"Ann";"Bob";"Cy"}}');
  });

  it('typing over a list replaces it with what was typed', async () => {
    const el = await createSpreadsheet({ '0:0': '={{"Ann";"Bob"}}' });
    await select(el, 0, 0);
    await press(el, 'x');
    await typeInEditor(el, 'x, y');
    await pressInEditor(el, 'Enter');
    expect(raw(el, '0:0')).toBe('x, y');
  });

  it('comma-separated text typed below a list becomes a list', async () => {
    const el = await createSpreadsheet({ '0:0': '={{"Ann";"Bob"}}' });
    await select(el, 1, 0);
    await enterValue(el, 'Cy, Di');
    expect(raw(el, '1:0')).toBe('={{"Cy";"Di"}}');
    await select(el, 1, 1);
    await enterValue(el, 'no, list');
    expect(raw(el, '1:1')).toBe('no, list');
  });

  it('Ctrl+J in the editor commits the text as a list', async () => {
    const el = await createSpreadsheet();
    await select(el, 0, 0);
    await press(el, 'F2');
    await typeInEditor(el, 'a, b');
    await pressInEditor(el, 'j', { ctrlKey: true, code: 'KeyJ' });
    expect(raw(el, '0:0')).toBe('={{"a";"b"}}');
  });

  it('formulas see the items', async () => {
    const el = await createSpreadsheet({
      '0:0': '={{1;2;3}}',
      '0:1': '=SUM(A1)',
      '0:2': '=HAS(A1,2)',
      '0:3': '=COUNTA(A1)',
      '2:0': '=A1',
    });
    expect(display(el, '0:1')).toBe('6');
    expect(display(el, '0:2')).toBe('TRUE');
    expect(display(el, '0:3')).toBe('3');
    expect([display(el, '2:0'), display(el, '3:0'), display(el, '4:0')]).toEqual(['1', '2', '3']);
  });

  it('shows the values in a popup with Alt+Down and closes it with Escape', async () => {
    const el = await createSpreadsheet({ '0:0': '={{"Ann";"Bob"}}' });
    await select(el, 0, 0);
    expect(cellEl(el, '0:0')!.getAttribute('aria-haspopup')).toBe('dialog');
    await press(el, 'ArrowDown', { altKey: true });
    await el.updateComplete;
    const card = shadow(el).querySelector('.ls-array-card') as HTMLElement;
    expect(card.getAttribute('role')).toBe('dialog');
    expect([...card.querySelectorAll('li')].map((li) => li.textContent!.trim())).toEqual(['Ann', 'Bob']);
    expect(cellEl(el, '0:0')!.getAttribute('aria-expanded')).toBe('true');
    card.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, composed: true }));
    await el.updateComplete;
    expect(shadow(el).querySelector('.ls-array-card')).toBeNull();
  });

  it('the popup goes into nested arrays', async () => {
    const el = await createSpreadsheet({ '0:0': '={{{1,2};{3,4}}}' });
    expect(await el.showCellValues('0:0')).toBe(true);
    await el.updateComplete;
    const nested = shadow(el).querySelectorAll<HTMLButtonElement>('.ls-array-nested');
    expect(nested).toHaveLength(2);
    nested[1].click();
    await el.updateComplete;
    expect([...shadow(el).querySelectorAll('.ls-array-card li')].map((li) => li.textContent!.trim())).toEqual(['3', '4']);
    expect(shadow(el).querySelector('.ls-array-back')).not.toBeNull();
  });

  it('a filter matches lists by their items', async () => {
    const el = await createSpreadsheet({
      '0:0': 'Owners',
      '1:0': '={{"Ann";"Bob"}}',
      '2:0': '={{"Bob"}}',
      '3:0': 'Cy',
    });
    el.setFilter({ start: { row: 0, col: 0 }, end: { row: 3, col: 0 } }, [{ col: 0, values: ['Ann'] }]);
    await el.updateComplete;
    expect(el.getFilter()?.hiddenRows).toEqual([2, 3]);
    el.setColumnFilter(0, ['Bob']);
    expect(el.getFilter()?.hiddenRows).toEqual([3]);
    await el.openFilterMenu(0);
    const labels = [...shadow(el).querySelectorAll('.ls-filter-menu [role="option"], .ls-filter-menu label')].map((n) =>
      n.textContent!.trim()
    );
    expect(labels.join('|')).toContain('Ann');
    expect(labels.join('|')).toContain('Cy');
    expect(labels.join('|')).not.toContain('Ann, Bob');
  });
});

describe('sheets', () => {
  it('starts with one sheet and adds more', async () => {
    const el = await createSpreadsheet({ '0:0': '5' });
    expect(el.sheets).toEqual(['Sheet1']);
    expect(el.activeSheet).toBe('Sheet1');
    const events: SheetsChangeDetail[] = [];
    el.addEventListener('sheets-change', (e) => events.push((e as CustomEvent<SheetsChangeDetail>).detail));
    expect(el.addSheet()).toBe('Sheet2');
    el.addSheet('Data', { data: data({ '0:0': '7', '1:0': '=A1*2' }) });
    expect(el.sheets).toEqual(['Sheet1', 'Sheet2', 'Data']);
    expect(events.map((e) => `${e.kind}:${e.sheet}`)).toEqual(['add:Sheet2', 'add:Data']);
    expect(el.getSheetData('Data').get('1:0')?.displayValue).toBe('14');
    expect(() => el.addSheet('data')).toThrow(RangeError);
    expect(() => el.addSheet('a/b')).toThrow(RangeError);
  });

  it('formulas refer to other sheets and follow their changes', async () => {
    const el = await createSpreadsheet({ '0:0': "=Data!A1+SUM('My data'!A:A)" });
    expect(display(el, '0:0')).toBe('#REF!');
    el.addSheet('Data', { data: data({ '0:0': '7' }) });
    el.addSheet('My data', { data: data({ '0:0': '1', '1:0': '2' }) });
    await el.updateComplete;
    expect(display(el, '0:0')).toBe('10');
    el.activateSheet('Data');
    await el.updateComplete;
    expect(el.activeSheet).toBe('Data');
    expect(display(el, '0:0')).toBe('7');
    await select(el, 0, 0);
    await enterValue(el, '20');
    expect(el.getSheetData('Sheet1').get('0:0')?.displayValue).toBe('23');
    el.activateSheet('Sheet1');
    await el.updateComplete;
    expect(display(el, '0:0')).toBe('23');
  });

  it('each sheet keeps its own size, selection and filter', async () => {
    const el = await createSpreadsheet({ '0:0': 'h', '1:0': 'a', '2:0': 'b' }, 20, 8);
    el.setFilter({ start: { row: 0, col: 0 }, end: { row: 2, col: 0 } }, [{ col: 0, values: ['a'] }]);
    await select(el, 2, 3);
    el.addSheet('Other', { rows: 5, cols: 3, activate: true });
    await el.updateComplete;
    expect([el.rows, el.cols]).toEqual([5, 3]);
    expect(el.getFilter()).toBeNull();
    expect(status(el)).toBe('A1');
    el.activateSheet('Sheet1');
    await el.updateComplete;
    expect([el.rows, el.cols]).toEqual([20, 8]);
    expect(el.getFilter()?.hiddenRows).toEqual([2]);
    expect(status(el)).toContain('D3');
  });

  it('undo shows the sheet the change was made on', async () => {
    const el = await createSpreadsheet({ '0:0': '1' });
    await select(el, 0, 0);
    await enterValue(el, '2');
    el.addSheet('Other', { activate: true });
    await el.updateComplete;
    await select(el, 0, 0);
    await enterValue(el, 'x');
    internals(el)._undo();
    await el.updateComplete;
    expect(el.activeSheet).toBe('Other');
    expect(raw(el, '0:0')).toBe('');
    internals(el)._undo();
    await el.updateComplete;
    expect(el.activeSheet).toBe('Sheet1');
    expect(raw(el, '0:0')).toBe('1');
    internals(el)._redo();
    await el.updateComplete;
    expect(raw(el, '0:0')).toBe('2');
  });

  it('renaming a sheet rewrites the formulas that refer to it, also in the undo history', async () => {
    const el = await createSpreadsheet({ '0:0': '=Data!A1' });
    el.addSheet('Data', { data: data({ '0:0': '7' }) });
    await select(el, 0, 1);
    await enterValue(el, '=Data!A1*2');
    await select(el, 0, 1);
    await enterValue(el, '');
    expect(raw(el, '0:1')).toBe('');
    el.renameSheet('Data', 'My data');
    await el.updateComplete;
    expect(el.sheets).toEqual(['Sheet1', 'My data']);
    expect(raw(el, '0:0')).toBe("='My data'!A1");
    expect(display(el, '0:0')).toBe('7');
    internals(el)._undo();
    await el.updateComplete;
    expect(raw(el, '0:1')).toBe("='My data'!A1*2");
    expect(display(el, '0:1')).toBe('14');
    expect(() => el.renameSheet('My data', 'Sheet1')).toThrow(RangeError);
  });

  it('removing a sheet leaves #REF! and drops its undo history', async () => {
    const el = await createSpreadsheet({ '0:0': '=Data!A1' });
    el.addSheet('Data', { data: data({ '0:0': '7' }), activate: true });
    await el.updateComplete;
    await select(el, 0, 0);
    await enterValue(el, '8');
    expect(el.removeSheet('Data')).toBe(true);
    await el.updateComplete;
    expect(el.activeSheet).toBe('Sheet1');
    expect(el.sheets).toEqual(['Sheet1']);
    expect(display(el, '0:0')).toBe('#REF!');
    expect(raw(el, '0:0')).toBe('=Data!A1');
    internals(el)._undo();
    await el.updateComplete;
    expect(display(el, '0:0')).toBe('#REF!');
    expect(() => el.removeSheet('Sheet1')).toThrow(RangeError);
  });

  it('inserting rows rewrites references from other sheets and names', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '1:0': '2' });
    el.addSheet('Report', { data: data({ '0:0': '=Sheet1!A2*10', '1:0': '=total' }) });
    el.defineName('total', '=SUM(A1:A2)');
    expect(el.getSheetData('Report').get('1:0')?.displayValue).toBe('3');
    el.insertRows(0, 1);
    await el.updateComplete;
    expect(el.getSheetData('Report').get('0:0')?.rawValue).toBe('=Sheet1!A3*10');
    expect(el.getSheetData('Report').get('0:0')?.displayValue).toBe('20');
    expect(el.getNames()).toEqual([{ name: 'total', refersTo: '=SUM(A2:A3)', sheet: 'Sheet1' }]);
    expect(el.getSheetData('Report').get('1:0')?.displayValue).toBe('3');
    internals(el)._undo();
    await el.updateComplete;
    expect(el.getSheetData('Report').get('0:0')?.rawValue).toBe('=Sheet1!A2*10');
    expect(el.getNames()[0].refersTo).toBe('=SUM(A1:A2)');
  });

  it('Alt+PageDown shows the next sheet', async () => {
    const el = await createSpreadsheet();
    el.addSheet('Two');
    await press(el, 'PageDown', { altKey: true });
    expect(el.activeSheet).toBe('Two');
    expect(announcement(el)).toBe('Two, sheet 2 of 2');
    await press(el, 'PageDown', { altKey: true });
    expect(el.activeSheet).toBe('Sheet1');
  });

  it('moves sheets', async () => {
    const el = await createSpreadsheet();
    el.addSheet('B');
    el.addSheet('C');
    el.moveSheet('C', 0);
    expect(el.sheets).toEqual(['C', 'Sheet1', 'B']);
  });
});

describe('defined names', () => {
  it('formulas use names for ranges, constants and functions', async () => {
    const el = await createSpreadsheet({ '0:0': '10', '1:0': '20', '0:2': '=SUM(sales)*rate', '1:2': '=double(4)', '2:2': '=missing' });
    expect(display(el, '0:2')).toBe('#NAME?');
    el.defineName('sales', 'A1:A2');
    el.defineName('rate', '=0.5');
    el.defineName('double', '=LAMBDA(x, x*2)');
    await el.updateComplete;
    expect(display(el, '0:2')).toBe('15');
    expect(display(el, '1:2')).toBe('8');
    expect(display(el, '2:2')).toBe('#NAME?');
    expect(el.getNames().map((n) => n.name)).toEqual(['sales', 'rate', 'double']);
    expect(el.removeName('RATE')).toBe(true);
    await el.updateComplete;
    expect(display(el, '0:2')).toBe('#NAME?');
    expect(() => el.defineName('A1', '=1')).toThrow(RangeError);
  });

  it('a name keeps pointing at the sheet it was defined on', async () => {
    const el = await createSpreadsheet({ '0:0': '5' });
    el.defineName('first', '=A1');
    el.addSheet('Other', { data: data({ '0:0': '9', '0:1': '=first' }), activate: true });
    await el.updateComplete;
    expect(display(el, '0:1')).toBe('5');
  });
});

describe('review regressions', () => {
  it('a formula that returns a blank cell is 0 to whatever reads it', async () => {
    const el = await createSpreadsheet({ '0:0': '10', '0:1': '=A1', '1:1': '=A2', '0:3': '=COUNT(B1:B2)', '1:3': '=B2&"x"' });
    expect(display(el, '1:1')).toBe('0');
    expect(display(el, '0:3')).toBe('2');
    expect(display(el, '1:3')).toBe('0x');
  });

  it('undoing a row insertion keeps names defined since', async () => {
    const el = await createSpreadsheet({ '4:0': '7', '0:2': '=Total+Rate' });
    el.defineName('Total', '=A5');
    el.insertRows(0, 1);
    el.defineName('Rate', '=10');
    internals(el)._undo();
    await el.updateComplete;
    expect(el.getNames()).toEqual([
      { name: 'Total', refersTo: '=A5', sheet: 'Sheet1' },
      { name: 'Rate', refersTo: '=10', sheet: 'Sheet1' },
    ]);
    expect(display(el, '0:2')).toBe('17');
  });

  it('filling down from a spilled value copies the value', async () => {
    const el = await createSpreadsheet({ '0:0': '=SEQUENCE(1,3)', '1:1': 'keep', '2:1': 'keep' });
    await select(el, 0, 1, 2, 1);
    await press(el, 'd', { ctrlKey: true });
    expect([raw(el, '1:1'), raw(el, '2:1')]).toEqual(['2', '2']);
    expect(display(el, '0:0')).toBe('1');
  });

  it('a whole-column reference still sees spilled rows after the data is replaced', async () => {
    const el = await createSpreadsheet({ '0:0': '=SEQUENCE(5)', '0:1': '=COUNTA(A:A)' });
    expect(display(el, '0:1')).toBe('5');
    const next = el.getData();
    next.set('0:2', { rawValue: '=B1+1', displayValue: '', type: 'text' });
    el.data = next;
    await el.updateComplete;
    expect(display(el, '0:1')).toBe('5');
    expect(display(el, '0:2')).toBe('6');
  });

  it('growing the grid and replacing the data in one step lets a blocked spill through', async () => {
    const el = await createSpreadsheet({ '0:0': '=SEQUENCE(10)' }, 5, 3);
    expect(display(el, '0:0')).toBe('#SPILL!');
    const next = el.getData();
    next.set('0:1', { rawValue: 'x', displayValue: 'x', type: 'text' });
    el.rows = 20;
    el.data = next;
    await el.updateComplete;
    expect(display(el, '0:0')).toBe('1');
    expect(display(el, '9:0')).toBe('10');
  });

  it('the spill that had its range keeps it across a full recalculation', async () => {
    const el = await createSpreadsheet({ '0:3': '5', '1:2': '=SEQUENCE(1,2)' });
    await select(el, 0, 3);
    await enterValue(el, '=SEQUENCE(3)');
    expect(display(el, '0:3')).toBe('#SPILL!');
    el.defineName('anything', '=1');
    await el.updateComplete;
    expect(display(el, '0:3')).toBe('#SPILL!');
    expect(display(el, '1:3')).toBe('2');
  });

  it('data that no longer holds a list loses the list', async () => {
    const el = await createSpreadsheet({ '0:0': '={{"a";"b"}}' });
    const next = el.getData();
    next.get('0:0')!.rawValue = 'plain';
    el.data = next;
    await el.updateComplete;
    expect(el.getCellValues('0:0')).toBeNull();
    expect(cellEl(el, '0:0')!.hasAttribute('aria-haspopup')).toBe(false);
  });

  it('NPV keeps a blank rate as the rate, and TEXTJOIN skips a blank cell', async () => {
    const el = await createSpreadsheet({ '0:0': '10', '2:0': '30', '0:2': '=NPV(A2,A1,A3)', '1:2': '=TEXTJOIN("-",TRUE,E1,A1)' });
    expect(display(el, '0:2')).toBe('40');
    expect(display(el, '1:2')).toBe('10');
  });

  it('a formula is as volatile as the names it uses', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '0:1': '=Stamp' });
    let ticks = 0;
    el.registerFunction('TICK', () => ++ticks, { volatile: true });
    el.defineName('Stamp', '=TICK()');
    await el.updateComplete;
    const before = Number(display(el, '0:1'));
    await select(el, 0, 0);
    await enterValue(el, '2');
    expect(Number(display(el, '0:1'))).toBeGreaterThan(before);
  });
});

describe('name box, go to and picking references', () => {
  const bar = (el: Y11nSpreadsheet) => shadow(el).querySelector('y11n-formula-bar')!;
  const nameBox = (el: Y11nSpreadsheet) => bar(el).shadowRoot!.querySelector<HTMLInputElement>('.name-box')!;

  async function typeInNameBox(el: Y11nSpreadsheet, value: string) {
    const box = nameBox(el);
    box.value = value;
    box.dispatchEvent(new Event('input'));
    box.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
    await el.updateComplete;
    await (bar(el) as unknown as { updateComplete: Promise<unknown> }).updateComplete;
  }

  it('shows the active cell and goes to a typed reference', async () => {
    const el = await createSpreadsheet({ '0:0': '1' });
    await (bar(el) as unknown as { updateComplete: Promise<unknown> }).updateComplete;
    expect(nameBox(el).value).toBe('A1');
    await typeInNameBox(el, 'b3:c4');
    expect(status(el)).toBe('B3, 2 rows, 2 columns selected');
    expect(nameBox(el).value).toBe('B3');
    await typeInNameBox(el, 'ZZ99');
    expect(announcement(el)).toBe('ZZ99 is outside the sheet');
  });

  it('names the selection, shows the name, and goes to it from another sheet', async () => {
    const el = await createSpreadsheet({ '0:0': '10', '1:0': '20', '0:2': '=SUM(sales)' });
    await select(el, 0, 0, 1, 0);
    await typeInNameBox(el, 'sales');
    expect(el.getNames()).toEqual([{ name: 'sales', refersTo: '=Sheet1!$A$1:$A$2', sheet: 'Sheet1' }]);
    expect(display(el, '0:2')).toBe('30');
    expect(announcement(el)).toBe('Named Sheet1!$A$1:$A$2 sales');
    expect(nameBox(el).value).toBe('sales');

    el.addSheet('Other', { activate: true });
    await el.updateComplete;
    expect(el.goTo('sales')).toBe(true);
    await el.updateComplete;
    expect(el.activeSheet).toBe('Sheet1');
    expect(status(el)).toContain('2 rows, 1 column selected');
    expect(el.goTo('Other!B2')).toBe(true);
    expect(el.activeSheet).toBe('Other');
    expect(el.goTo('nonsense(')).toBe(false);
  });

  it('refuses invalid names, constants and read-only, and can be vetoed', async () => {
    const el = await createSpreadsheet({ '0:0': '1' });
    el.defineName('rate', '=0.5');
    await typeInNameBox(el, 'rate');
    expect(announcement(el)).toBe('rate does not refer to cells: it is =0.5');
    await typeInNameBox(el, 'not a name');
    expect(announcement(el)).toBe('not a name is not a cell reference or a valid name');
    el.addEventListener('before-names-change', (e) => e.preventDefault(), { once: true });
    await typeInNameBox(el, 'vetoed');
    el.readOnly = true;
    await el.updateComplete;
    await typeInNameBox(el, 'locked');
    expect(el.getNames().map((n) => n.name)).toEqual(['rate']);
  });

  it('Ctrl+G moves focus to the name box', async () => {
    const el = await createSpreadsheet();
    await press(el, 'g', { ctrlKey: true, code: 'KeyG' });
    expect(bar(el).shadowRoot!.activeElement).toBe(nameBox(el));
  });

  it('a formula stays open on another sheet and takes a reference from it', async () => {
    const el = await createSpreadsheet({ '0:0': '1' });
    el.addSheet('Data', { data: data({ '1:1': '7' }) });
    await select(el, 0, 2);
    await press(el, '=');
    await typeInEditor(el, '=');
    await pressInEditor(el, 'PageDown', { altKey: true });
    expect(el.activeSheet).toBe('Data');
    expect(editor(el).style.display).not.toBe('none');
    expect(announcement(el)).toBe('Data. Pick a cell for the formula in Sheet1!C1; Enter finishes');
    await pressInEditor(el, 'ArrowDown');
    await pressInEditor(el, 'ArrowRight');
    expect(editor(el).value).toBe('=Data!B2');
    await pressInEditor(el, 'Enter');
    expect(el.activeSheet).toBe('Sheet1');
    expect(raw(el, '0:2')).toBe('=Data!B2');
    expect(display(el, '0:2')).toBe('7');
  });

  it('Escape on the other sheet cancels the edit and comes back', async () => {
    const el = await createSpreadsheet({ '0:0': '1' });
    el.addSheet("It's data");
    await select(el, 0, 2);
    await press(el, '=');
    await typeInEditor(el, '=1+');
    await pressInEditor(el, 'PageDown', { altKey: true });
    await pressInEditor(el, 'ArrowDown');
    expect(editor(el).value).toBe("=1+'It''s data'!A2");
    await pressInEditor(el, 'Escape');
    expect(el.activeSheet).toBe('Sheet1');
    expect(raw(el, '0:2')).toBe('');
  });

  it('clicking a cell while typing a formula picks it, and commits when no reference fits', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '1:1': '5' });
    await select(el, 0, 3);
    await press(el, '=');
    await typeInEditor(el, '=A1+');
    editor(el).setSelectionRange(4, 4);
    cellEl(el, '1:1')!.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, composed: true, button: 0 }));
    document.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
    await el.updateComplete;
    expect(editor(el).value).toBe('=A1+B2');
    // Another click replaces the reference being picked
    cellEl(el, '0:0')!.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, composed: true, button: 0 }));
    document.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
    await el.updateComplete;
    expect(editor(el).value).toBe('=A1+A1');
    await pressInEditor(el, 'Enter');
    expect(raw(el, '0:3')).toBe('=A1+A1');

    // A finished formula: the click commits it, as before
    await select(el, 2, 3);
    await press(el, '=');
    await typeInEditor(el, '=A1');
    editor(el).setSelectionRange(3, 3);
    cellEl(el, '1:1')!.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, composed: true, button: 0 }));
    document.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
    await el.updateComplete;
    expect(raw(el, '2:3')).toBe('=A1');
  });

  it('Alt+Shift+V shows the values of a list in a filter header', async () => {
    const el = await createSpreadsheet({ '0:0': '={{"a";"b"}}', '1:0': 'x' });
    el.setFilter({ start: { row: 0, col: 0 }, end: { row: 1, col: 0 } });
    await select(el, 0, 0);
    expect(status(el)).toContain('Alt+Shift+V to show the values');
    await press(el, 'V', { altKey: true, shiftKey: true, code: 'KeyV' });
    await el.updateComplete;
    expect(shadow(el).querySelector('.ls-array-card')).not.toBeNull();
  });
});

describe('name manager', () => {
  const manager = (el: Y11nSpreadsheet) => shadow(el).querySelector('y11n-name-manager') as HTMLElement | null;
  const fire = (el: Y11nSpreadsheet, type: string, detail?: unknown) =>
    manager(el)!.dispatchEvent(new CustomEvent(type, { detail, bubbles: true, composed: true }));

  it('opens with Ctrl+F3 and applies saves, renames and deletes', async () => {
    const el = await createSpreadsheet({ '0:0': '10', '1:0': '20', '0:2': '=SUM(sales)' });
    await press(el, 'F3', { ctrlKey: true });
    await el.updateComplete;
    expect(manager(el)).not.toBeNull();

    fire(el, 'name-save', { name: 'sales', refersTo: '=A1:A2', sheet: 'Sheet1' });
    await el.updateComplete;
    expect(display(el, '0:2')).toBe('30');
    expect((manager(el) as unknown as { status: string }).status).toBe('Name added: sales');

    // Renaming rewrites the formulas that use the name
    fire(el, 'name-save', { name: 'revenue', refersTo: '=A1:A2', sheet: 'Sheet1', previous: 'sales' });
    await el.updateComplete;
    expect(raw(el, '0:2')).toBe('=SUM(revenue)');
    expect(display(el, '0:2')).toBe('30');
    expect(el.getNames().map((n) => n.name)).toEqual(['revenue']);

    fire(el, 'name-remove', { name: 'revenue' });
    await el.updateComplete;
    expect(display(el, '0:2')).toBe('#NAME?');
    expect((manager(el) as unknown as { status: string }).status).toBe('Name deleted: revenue');

    fire(el, 'name-close');
    await el.updateComplete;
    expect(manager(el)).toBeNull();
  });

  it('goes to a name and closes; refuses changes in read-only mode', async () => {
    const el = await createSpreadsheet({ '0:0': '1' });
    el.defineName('spot', '=C3');
    el.defineName('rate', '=0.5');
    await el.openNameManager();
    fire(el, 'name-go', { name: 'rate' });
    await el.updateComplete;
    expect((manager(el) as unknown as { status: string }).status).toBe('rate does not refer to cells');
    fire(el, 'name-go', { name: 'spot' });
    await el.updateComplete;
    expect(manager(el)).toBeNull();
    expect(status(el)).toBe('C3');

    el.readOnly = true;
    await el.openNameManager();
    fire(el, 'name-remove', { name: 'spot' });
    await el.updateComplete;
    expect(el.getNames()).toHaveLength(2);
    expect((manager(el) as unknown as { status: string }).status).toBe('The change was not applied');
  });
});

describe('modals keep Tab inside', () => {
  it('the values popup', async () => {
    const el = await createSpreadsheet({ '0:0': '={{"a";"b"}}' });
    await el.showCellValues('0:0');
    await el.updateComplete;
    const card = shadow(el).querySelector('.ls-array-card') as HTMLElement;
    expect(card.getAttribute('aria-modal')).toBe('true');
    const body = card.querySelector<HTMLElement>('.ls-array-card-body')!;
    const close = card.querySelector<HTMLElement>('.ls-array-close')!;
    body.focus();
    const forward = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, composed: true, cancelable: true });
    body.dispatchEvent(forward);
    expect(forward.defaultPrevented).toBe(true);
    expect(shadow(el).activeElement).toBe(close);
    const back = new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true, composed: true, cancelable: true });
    close.dispatchEvent(back);
    expect(back.defaultPrevented).toBe(true);
    expect(shadow(el).activeElement).toBe(body);
  });

  it('the name manager', async () => {
    const el = await createSpreadsheet();
    el.defineName('one', '=A1');
    await el.openNameManager();
    const manager = shadow(el).querySelector('y11n-name-manager') as HTMLElement & { updateComplete: Promise<unknown> };
    await manager.updateComplete;
    const root = manager.shadowRoot!;
    expect(root.querySelector('[role="dialog"]')!.getAttribute('aria-modal')).toBe('true');
    const buttons = [...root.querySelectorAll<HTMLElement>('button:not(:disabled)')];
    const first = buttons[0];
    const last = buttons[buttons.length - 1];
    last.focus();
    const forward = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, composed: true, cancelable: true });
    last.dispatchEvent(forward);
    expect(forward.defaultPrevented).toBe(true);
    expect(root.activeElement).toBe(first);
    const back = new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true, composed: true, cancelable: true });
    first.dispatchEvent(back);
    expect(back.defaultPrevented).toBe(true);
    expect(root.activeElement).toBe(last);
  });
});

describe('second review regressions', () => {
  const fire = (el: Y11nSpreadsheet, type: string, detail?: unknown) =>
    shadow(el)
      .querySelector('y11n-name-manager')!
      .dispatchEvent(new CustomEvent(type, { detail, bubbles: true, composed: true }));

  it('replacing data keeps a formatted spilled cell marked as spilled', async () => {
    const el = await createSpreadsheet({ '0:0': '=SEQUENCE(2)', '0:3': '1' });
    el.setCellFormat('1:0', { bold: true });
    const next = el.getData();
    next.get('0:3')!.rawValue = '2';
    el.setData(next);
    await el.updateComplete;
    expect(cell(el, '1:0')).toMatchObject({ spilledFrom: '0:0', displayValue: '2' });
    expect(el.toCSV()).toBe('1,,,2\n2,,,');
  });

  it('replacing the shown sheet keeps the undo history of other sheets', async () => {
    const el = await createSpreadsheet({ '0:0': '1' });
    el.addSheet('Data', { activate: true });
    await el.updateComplete;
    await select(el, 0, 0);
    await enterValue(el, 'x');
    el.activateSheet('Sheet1');
    el.setSheetData('Sheet1', data({ '0:0': '9' }));
    await el.updateComplete;
    internals(el)._undo();
    await el.updateComplete;
    expect(el.activeSheet).toBe('Data');
    expect(raw(el, '0:0')).toBe('');
  });

  it('renaming a name leaves local variables and functions alone and keeps recursion', async () => {
    const el = await createSpreadsheet({
      '0:0': '=LET(x,1,Sales+x)',
      '0:1': '=SUM(1,2)+Sum',
      '0:2': '=Fact(5)',
      '0:3': '=Sales*2',
    });
    el.defineName('Sales', '=10');
    el.defineName('Sum', '=100');
    el.defineName('Fact', '=LAMBDA(n,IF(n<=1,1,n*Fact(n-1)))');
    await el.openNameManager();
    expect([display(el, '0:0'), display(el, '0:1'), display(el, '0:2')]).toEqual(['11', '103', '120']);

    // Renaming Sales to x would be captured by the LET's own x: the rename is refused as a whole
    fire(el, 'name-save', { name: 'x', refersTo: '=10', sheet: 'Sheet1', previous: 'Sales' });
    await el.updateComplete;
    const manager = shadow(el).querySelector('y11n-name-manager') as unknown as { status: string };
    expect(manager.status).toBe('Sales was not renamed to x: the formula in Sheet1!A1 would no longer mean the same');
    expect(el.getNames().map((n) => n.name)).toEqual(['Sales', 'Sum', 'Fact']);
    expect([raw(el, '0:0'), raw(el, '0:3'), display(el, '0:0')]).toEqual(['=LET(x,1,Sales+x)', '=Sales*2', '11']);

    // A name nothing binds is fine
    fire(el, 'name-save', { name: 'Turnover', refersTo: '=10', sheet: 'Sheet1', previous: 'Sales' });
    await el.updateComplete;
    expect([raw(el, '0:0'), raw(el, '0:3'), display(el, '0:0')]).toEqual(['=LET(x,1,Turnover+x)', '=Turnover*2', '11']);

    fire(el, 'name-save', { name: 'Revenue', refersTo: '=100', sheet: 'Sheet1', previous: 'Sum' });
    await el.updateComplete;
    expect(raw(el, '0:1')).toBe('=SUM(1,2)+Revenue');
    expect(display(el, '0:1')).toBe('103');

    fire(el, 'name-save', {
      name: 'Factorial',
      refersTo: '=LAMBDA(n,IF(n<=1,1,n*Fact(n-1)))',
      sheet: 'Sheet1',
      previous: 'Fact',
    });
    await el.updateComplete;
    expect(raw(el, '0:2')).toBe('=Factorial(5)');
    expect(display(el, '0:2')).toBe('120');
  });

  it('undoing a row insertion still moves a name that was renamed since', async () => {
    const el = await createSpreadsheet({ '1:0': '7' });
    el.defineName('Foo', '=A2');
    el.insertRows(0, 1);
    await el.openNameManager();
    fire(el, 'name-save', { name: 'Bar', refersTo: '=A3', sheet: 'Sheet1', previous: 'Foo' });
    await el.updateComplete;
    internals(el)._undo();
    await el.updateComplete;
    expect(el.getNames()).toEqual([{ name: 'Bar', refersTo: '=A2', sheet: 'Sheet1' }]);
  });

  it('inserting rows moves sheet-qualified references in conditional formats, on every sheet', async () => {
    const el = await createSpreadsheet({ '0:0': '5' });
    const range = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };
    el.addConditionalFormat({ range, condition: { type: 'formula', formula: '=Sheet1!$A$1>0' }, format: { bold: true } });
    el.addSheet('Other', { activate: true });
    await el.updateComplete;
    el.addConditionalFormat({ range, condition: { type: 'formula', formula: '=Sheet1!$A$1>0' }, format: { bold: true } });
    el.activateSheet('Sheet1');
    await el.updateComplete;
    el.insertRows(0, 1);
    await el.updateComplete;
    expect(el.getConditionalFormats()[0].condition).toMatchObject({ formula: '=Sheet1!$A$2>0' });
    el.activateSheet('Other');
    await el.updateComplete;
    expect(el.getConditionalFormats()[0].condition).toMatchObject({ formula: '=Sheet1!$A$2>0' });
    internals(el)._undo();
    await el.updateComplete;
    expect(el.activeSheet).toBe('Sheet1');
    expect(el.getConditionalFormats()[0].condition).toMatchObject({ formula: '=Sheet1!$A$1>0' });
    el.activateSheet('Other');
    await el.updateComplete;
    expect(el.getConditionalFormats()[0].condition).toMatchObject({ formula: '=Sheet1!$A$1>0' });
  });
});

describe('third review regressions', () => {
  const bar = (el: Y11nSpreadsheet) => shadow(el).querySelector('y11n-formula-bar')!;
  const fire = (el: Y11nSpreadsheet, type: string, detail?: unknown) =>
    shadow(el)
      .querySelector('y11n-name-manager')!
      .dispatchEvent(new CustomEvent(type, { detail, bubbles: true, composed: true }));

  it('text typed into the formula bar is dropped when another sheet is shown', async () => {
    const el = await createSpreadsheet();
    el.addSheet('Other');
    await (bar(el) as unknown as { updateComplete: Promise<unknown> }).updateComplete;
    const input = bar(el).shadowRoot!.querySelector<HTMLInputElement>('input.formula-input')!;
    input.value = '123';
    input.dispatchEvent(new Event('input'));
    el.activateSheet('Other');
    await el.updateComplete;
    await (bar(el) as unknown as { updateComplete: Promise<unknown> }).updateComplete;
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
    await el.updateComplete;
    expect(el.getSheetData('Other').size).toBe(0);
    expect(el.getSheetData('Sheet1').size).toBe(0);
  });

  it('a spill on another sheet that shrinks is seen by whole-column formulas there', async () => {
    const el = await createSpreadsheet({ '0:0': '5' });
    el.addSheet('Data', { data: data({ '0:0': '=SEQUENCE(Sheet1!A1)', '0:2': '=COUNTA(A:A)' }) });
    expect(el.getSheetData('Data').get('0:2')?.displayValue).toBe('5');
    await select(el, 0, 0);
    await enterValue(el, '2');
    expect(el.getSheetData('Data').get('0:2')?.displayValue).toBe('2');
  });

  it('renaming a name keeps another name on a removed sheet pointing there', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '0:1': '=Amount', '0:2': '=Other' });
    el.addSheet('Data', { data: data({ '0:0': '50' }) });
    el.defineName('Amount', '=A1', 'Data');
    el.defineName('Other', '=7');
    await el.updateComplete;
    expect(display(el, '0:1')).toBe('50');
    el.removeSheet('Data');
    await el.updateComplete;
    expect(display(el, '0:1')).toBe('#REF!');
    await el.openNameManager();
    fire(el, 'name-save', { name: 'Seven', refersTo: '=7', sheet: 'Sheet1', previous: 'Other' });
    await el.updateComplete;
    expect(raw(el, '0:2')).toBe('=Seven');
    expect(display(el, '0:1')).toBe('#REF!');
  });

  it('filling a whole spill range down copies the formula, not its values', async () => {
    const el = await createSpreadsheet({ '0:0': '=SEQUENCE(3)' });
    (el as unknown as { _autoFill(a: unknown, b: unknown): void })._autoFill(
      { start: { row: 0, col: 0 }, end: { row: 2, col: 0 } },
      { start: { row: 0, col: 0 }, end: { row: 5, col: 0 } }
    );
    await el.updateComplete;
    expect(raw(el, '3:0')).toBe('=SEQUENCE(3)');
    expect([display(el, '3:0'), display(el, '4:0'), display(el, '5:0')]).toEqual(['1', '2', '3']);
    expect(cell(el, '5:0')?.spilledFrom).toBe('3:0');
  });

  it('undo follows conditional-format rules of another sheet through a sheet rename', async () => {
    const el = await createSpreadsheet({ '0:0': '5' });
    const range = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };
    el.addSheet('Sheet2', { activate: true });
    await el.updateComplete;
    el.addConditionalFormat({ range, condition: { type: 'formula', formula: '=Sheet1!$A$2>0' }, format: { bold: true } });
    el.activateSheet('Sheet1');
    await el.updateComplete;
    el.insertRows(0, 1);
    el.insertRows(0, 1);
    el.renameSheet('Sheet1', 'Data');
    internals(el)._undo();
    internals(el)._undo();
    await el.updateComplete;
    el.activateSheet('Sheet2');
    await el.updateComplete;
    expect(el.getConditionalFormats()[0].condition).toMatchObject({ formula: '=Data!$A$2>0' });
  });

  it('a picked reference replaces the selected text, and stays replaceable on a third sheet', async () => {
    const el = await createSpreadsheet({ '0:0': '1' });
    el.addSheet('Two');
    el.addSheet('Three');
    await select(el, 0, 3);
    await press(el, '=');
    await typeInEditor(el, '=A1');
    editor(el).setSelectionRange(1, 3);
    cellEl(el, '1:1')!.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, composed: true, button: 0 }));
    document.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
    await el.updateComplete;
    expect(editor(el).value).toBe('=B2');

    await pressInEditor(el, 'PageDown', { altKey: true });
    cellEl(el, '1:1')!.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, composed: true, button: 0 }));
    document.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
    await el.updateComplete;
    expect(editor(el).value).toBe('=Two!B2');
    await pressInEditor(el, 'PageDown', { altKey: true });
    cellEl(el, '2:2')!.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, composed: true, button: 0 }));
    document.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
    await el.updateComplete;
    expect(editor(el).value).toBe('=Three!C3');
    expect(el.activeSheet).toBe('Three');
    await pressInEditor(el, 'Enter');
    expect(el.activeSheet).toBe('Sheet1');
    expect(raw(el, '0:3')).toBe('=Three!C3');
  });

  it('a click on a header while a formula is open on another sheet only ends the edit', async () => {
    const el = await createSpreadsheet({ '0:0': '1' });
    el.addSheet('Other');
    await select(el, 0, 2);
    await press(el, '=');
    await typeInEditor(el, '=1');
    await pressInEditor(el, 'PageDown', { altKey: true });
    expect(el.activeSheet).toBe('Other');
    const header = shadow(el).querySelector('[data-header-col="1"]') as HTMLElement;
    header.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, composed: true, button: 0 }));
    await el.updateComplete;
    expect(el.activeSheet).toBe('Sheet1');
    expect(raw(el, '0:2')).toBe('=1');
    expect(status(el)).not.toContain('Column');
  });
});

describe('fourth review regressions', () => {
  const fire = (el: Y11nSpreadsheet, type: string, detail?: unknown) =>
    shadow(el)
      .querySelector('y11n-name-manager')!
      .dispatchEvent(new CustomEvent(type, { detail, bubbles: true, composed: true }));
  const managerStatus = (el: Y11nSpreadsheet) =>
    (shadow(el).querySelector('y11n-name-manager') as unknown as { status: string }).status;

  it('a rename is refused when a formula has a variable of the old name', async () => {
    const el = await createSpreadsheet({ '0:0': '=LET(Sales,Sales*2,Sales)' });
    el.defineName('Sales', '=10');
    await el.openNameManager();
    expect(display(el, '0:0')).toBe('20');
    fire(el, 'name-save', { name: 'Revenue', refersTo: '=10', sheet: 'Sheet1', previous: 'Sales' });
    await el.updateComplete;
    expect(managerStatus(el)).toContain('Sales was not renamed');
    expect(display(el, '0:0')).toBe('20');
    expect(el.getNames().map((n) => n.name)).toEqual(['Sales']);
  });

  it('a rename is refused when undo would bring back a formula it can\'t rewrite', async () => {
    const el = await createSpreadsheet({ '0:0': '=LET(x,1,Sales+x)' });
    el.defineName('Sales', '=10');
    await select(el, 0, 0);
    await enterValue(el, '99');
    await el.openNameManager();
    fire(el, 'name-save', { name: 'x', refersTo: '=10', sheet: 'Sheet1', previous: 'Sales' });
    await el.updateComplete;
    expect(managerStatus(el)).toContain('Sales was not renamed');
    el.closeNameManager();
    internals(el)._undo();
    await el.updateComplete;
    expect(display(el, '0:0')).toBe('11');
  });

  it('a conditional format sees #SPILL! where the grid shows it', async () => {
    const el = await createSpreadsheet({ '0:0': '=SEQUENCE(3)', '1:0': 'x' });
    expect(display(el, '0:0')).toBe('#SPILL!');
    const range = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };
    el.addConditionalFormat({ range, condition: { type: 'formula', formula: '=ISERROR(A1)' }, format: { bold: true } });
    await el.updateComplete;
    expect(el.getConditionalFormatResult('0:0')?.format).toEqual({ bold: true });
  });
});

describe('fifth review regressions', () => {
  const fire = (el: Y11nSpreadsheet, type: string, detail?: unknown) =>
    shadow(el)
      .querySelector('y11n-name-manager')!
      .dispatchEvent(new CustomEvent(type, { detail, bubbles: true, composed: true }));

  it('formatting a volatile cell does not roll it again behind its dependents', async () => {
    const el = await createSpreadsheet({ '0:2': '=ROLL()', '0:3': '=C1*2' });
    let counter = 0;
    el.registerFunction('ROLL', () => ++counter, { volatile: true });
    await el.updateComplete;
    el.setCellFormat('0:2', { numberFormat: { type: 'number', decimals: 0 } });
    await el.updateComplete;
    expect(Number(display(el, '0:3'))).toBe(Number(display(el, '0:2')) * 2);
  });

  it('whole columns of two sheets pair up in SUMIFS', async () => {
    const el = await createSpreadsheet({ '0:0': '10', '1:0': '20', '0:2': '=SUMIFS(A:A,Other!A:A,1)' });
    el.addSheet('Other', { data: data({ '0:0': '1' }) });
    await el.updateComplete;
    expect(display(el, '0:2')).toBe('10');
  });

  it('ROW() in a conditional-format formula is the row of the cell being tested', async () => {
    const el = await createSpreadsheet({ '0:0': 'a', '1:0': 'b', '2:0': 'c', '3:0': 'd' });
    el.addConditionalFormat({
      range: { start: { row: 0, col: 0 }, end: { row: 3, col: 0 } },
      condition: { type: 'formula', formula: '=MOD(ROW(),2)=0' },
      format: { bold: true },
    });
    await el.updateComplete;
    const bold = ['0:0', '1:0', '2:0', '3:0'].map((key) => !!el.getConditionalFormatResult(key)?.format.bold);
    expect(bold).toEqual([false, true, false, true]);
  });

  it('undoing a row insertion leaves a name alone once it is based on another sheet', async () => {
    const el = await createSpreadsheet({ '1:0': '2', '2:0': '3', '0:3': '=Foo' });
    el.addSheet('Other', { data: data({ '1:0': '20', '2:0': '30' }) });
    el.defineName('Foo', '=A2');
    el.insertRows(0, 1);
    await el.openNameManager();
    fire(el, 'name-save', { name: 'Foo', refersTo: '=A3', sheet: 'Other', previous: 'Foo' });
    await el.updateComplete;
    expect(el.getNames()).toEqual([{ name: 'Foo', refersTo: '=A3', sheet: 'Other' }]);
    el.closeNameManager();
    internals(el)._undo();
    await el.updateComplete;
    expect(el.getNames()).toEqual([{ name: 'Foo', refersTo: '=A3', sheet: 'Other' }]);
    expect(display(el, '0:3')).toBe('30');
  });

  it('F4 on a picked reference keeps it replaceable', async () => {
    const el = await createSpreadsheet({ '0:0': '1' });
    await select(el, 0, 3);
    await press(el, '=');
    await typeInEditor(el, '=');
    editor(el).setSelectionRange(1, 1);
    const click = (key: string) => {
      cellEl(el, key)!.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, composed: true, button: 0 }));
      document.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
    };
    click('1:0');
    await el.updateComplete;
    expect(editor(el).value).toBe('=A2');
    editor(el).setSelectionRange(3, 3);
    await pressInEditor(el, 'F4');
    expect(editor(el).value).toBe('=$A$2');
    click('1:1');
    await el.updateComplete;
    expect(editor(el).value).toBe('=B2');
  });
});

describe('sixth review regressions', () => {
  it('a whole column taken through INDEX or OFFSET pairs up with a plain one', async () => {
    const el = await createSpreadsheet(
      { '0:0': 'x', '0:1': '5', '0:2': '=SUMIFS(INDEX(B:B,0),A:A,"x")', '1:2': '=SUMIFS(OFFSET(A:A,0,1),A:A,"x")' },
      10,
      6
    );
    expect(display(el, '0:2')).toBe('5');
    expect(display(el, '1:2')).toBe('5');
  });

  it('formulas over whole columns follow the used part of the workbook', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '0:2': '=COUNTIF(A:A,"")' });
    expect(display(el, '0:2')).toBe('0');
    el.addSheet('Other', { activate: true });
    await el.updateComplete;
    await select(el, 9, 1);
    await enterValue(el, '1');
    const incremental = el.getSheetData('Sheet1').get('0:2')?.displayValue;
    el.activateSheet('Sheet1');
    el.setData(el.getData());
    await el.updateComplete;
    expect(incremental).toBe(display(el, '0:2'));
    expect(incremental).toBe('9');

    // The same on one sheet: a value far down moves the end of the column
    await select(el, 14, 3);
    await enterValue(el, 'far');
    expect(display(el, '0:2')).toBe('14');
  });

  it('a conditional format tests the value a volatile formula shows', async () => {
    const el = await createSpreadsheet({ '0:0': '=ROLL()' });
    let counter = 0;
    el.registerFunction('ROLL', () => ++counter, { volatile: true });
    await el.updateComplete;
    const shown = display(el, '0:0');
    el.addConditionalFormat({
      range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
      condition: { type: 'formula', formula: `=A1=${shown}` },
      format: { bold: true },
    });
    await el.updateComplete;
    expect(display(el, '0:0')).toBe(shown);
    expect(el.getConditionalFormatResult('0:0')?.format).toEqual({ bold: true });
  });
});

describe('a spill range pasted into another spreadsheet', () => {
  const spillData: GridData = new Map([
    ['0:0', { rawValue: '=SEQUENCE(3)', displayValue: '1', type: 'number', spill: { rows: 3, cols: 1 } }],
    ['1:0', { rawValue: '', displayValue: '2', type: 'number', spilledFrom: '0:0' }],
    ['2:0', { rawValue: '', displayValue: '3', type: 'number', spilledFrom: '0:0' }],
  ]);
  const range = { start: { row: 0, col: 0 }, end: { row: 2, col: 0 } };

  async function copyThenPasteElsewhere(options?: { values: boolean }) {
    let html = '';
    const original = globalThis.navigator;
    const globals = globalThis as { ClipboardItem?: unknown; Blob: unknown };
    const originalItem = globals.ClipboardItem;
    const originalBlob = globals.Blob;
    globals.ClipboardItem = class {
      constructor(readonly items: Record<string, Blob>) {}
    };
    // A Blob whose text can be read back in the test environment
    globals.Blob = class {
      constructor(readonly parts: string[]) {}
      text() {
        return Promise.resolve(this.parts.join(''));
      }
    };
    Object.defineProperty(globalThis, 'navigator', {
      value: {
        clipboard: {
          write: async (items: Array<{ items: Record<string, Blob> }>) => {
            html = await items[0].items['text/html'].text();
          },
          writeText: async () => {},
          readText: async () => '1\n2\n3',
          read: async () => [{ types: ['text/html'], getType: async () => ({ text: async () => html }) }],
        },
      },
      configurable: true,
    });
    try {
      await new ClipboardManager().copy(spillData, range);
      // Another instance: it has no memory of what was copied
      return await new ClipboardManager().paste(0, 2, 20, 20, undefined, options);
    } finally {
      Object.defineProperty(globalThis, 'navigator', { value: original, configurable: true });
      globals.ClipboardItem = originalItem;
      globals.Blob = originalBlob;
    }
  }

  it('leaves the spilled cells to the formula', async () => {
    const updates = await copyThenPasteElsewhere();
    expect(updates!.map((u) => u.value)).toEqual(['=SEQUENCE(3)', '', '']);
  });

  it('writes every value for a values-only paste', async () => {
    const updates = await copyThenPasteElsewhere({ values: true });
    expect(updates!.map((u) => u.value)).toEqual(['1', '2', '3']);
  });
});

describe('seventh review regressions', () => {
  it('a whole-column formula does not read its own spill, so it settles anywhere', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '1:0': '2', '1:1': '=A:A*2', '0:3': '=SUM(B:B)' }, 20, 8);
    expect([display(el, '1:1'), display(el, '2:1')]).toEqual(['2', '4']);
    expect(cell(el, '3:1')).toBeUndefined();
    // Another formula does see that spill, to its last row
    expect(display(el, '0:3')).toBe('6');

    const filtered = await createSpreadsheet({ '0:0': '1', '1:0': '2', '2:0': '3', '1:1': '=FILTER(A:A,A:A>0)', '0:3': '=SUM(B:B)' }, 20, 8);
    expect([display(filtered, '1:1'), display(filtered, '2:1'), display(filtered, '3:1')]).toEqual(['1', '2', '3']);
    expect(cell(filtered, '4:1')).toBeUndefined();
    expect(display(filtered, '0:3')).toBe('6');

    // It follows the data both ways: longer, and shorter again
    await select(el, 2, 0);
    await enterValue(el, '3');
    expect([display(el, '3:1'), display(el, '0:3')]).toEqual(['6', '12']);
    await select(el, 2, 0);
    await press(el, 'Delete');
    expect(cell(el, '3:1')).toBeUndefined();
    expect(display(el, '0:3')).toBe('6');
  });

  it('a formatted spilled value copied alone into another spreadsheet stays a number', async () => {
    const source: GridData = new Map([
      ['1:0', { rawValue: '', displayValue: '$2.00', type: 'number', spilledFrom: '0:0', format: { numberFormat: { type: 'currency' } } }],
      ['2:0', { rawValue: '', displayValue: '$3.00', type: 'number', spilledFrom: '0:0', format: { numberFormat: { type: 'currency' } } }],
    ]);
    const plain: Record<string, string> = { '1:0': '2', '2:0': '3' };
    let html = '';
    const original = globalThis.navigator;
    const globals = globalThis as { ClipboardItem?: unknown; Blob: unknown };
    const originalItem = globals.ClipboardItem;
    const originalBlob = globals.Blob;
    globals.ClipboardItem = class {
      constructor(readonly items: Record<string, Blob>) {}
    };
    globals.Blob = class {
      constructor(readonly parts: string[]) {}
      text() {
        return Promise.resolve(this.parts.join(''));
      }
    };
    Object.defineProperty(globalThis, 'navigator', {
      value: {
        clipboard: {
          write: async (items: Array<{ items: Record<string, Blob> }>) => {
            html = await items[0].items['text/html'].text();
          },
          writeText: async () => {},
          readText: async () => '$2.00\n$3.00',
          read: async () => [{ types: ['text/html'], getType: async () => ({ text: async () => html }) }],
        },
      },
      configurable: true,
    });
    try {
      await new ClipboardManager().copy(source, { start: { row: 1, col: 0 }, end: { row: 2, col: 0 } }, undefined, (key) => plain[key]);
      const updates = await new ClipboardManager().paste(0, 0, 20, 20);
      expect(updates!.map((u) => u.value)).toEqual(['2', '3']);
      expect(updates![0].format).toEqual({ numberFormat: { type: 'currency' } });
    } finally {
      Object.defineProperty(globalThis, 'navigator', { value: original, configurable: true });
      globals.ClipboardItem = originalItem;
      globals.Blob = originalBlob;
    }
  });
});

describe('eighth review regressions', () => {
  it('two whole-column formulas that lengthen each other end as #SPILL!, without taking others along', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '1:0': '2', '0:4': '=SORT(A:A)' }, 20, 8);
    await select(el, 1, 1);
    await enterValue(el, '=A:A*2');
    await select(el, 2, 2);
    await enterValue(el, '=B:B*2');
    // B2 reads as far as C3 spills, and C3 as far as B2 spills: neither can settle
    expect([display(el, '1:1'), display(el, '2:2')]).toContain('#SPILL!');
    // The formula in row 1 keeps a result, as long as the data and the spills that remain
    expect(display(el, '0:4')).not.toBe('#SPILL!');
    expect(cell(el, '0:4')?.spill?.rows).toBeLessThan(20);

    // A result with a largest size is fine below row 1
    const capped = await createSpreadsheet({ '0:0': '1', '1:0': '2', '1:1': '=TAKE(A:A,10)' }, 20, 8);
    expect(cell(capped, '1:1')?.spill).toEqual({ rows: 2, cols: 1 });
    expect([display(capped, '1:1'), display(capped, '2:1')]).toEqual(['1', '2']);
  });

  it('a spilled text that looks like a number is copied as text', async () => {
    const manager = new ClipboardManager();
    const source: GridData = new Map([
      ['0:1', { rawValue: '', displayValue: '00456', type: 'text', spilledFrom: '0:0' }],
      ['0:2', { rawValue: '=LEFT("0077",4)', displayValue: '0077', type: 'text' }],
    ]);
    const original = globalThis.navigator;
    Object.defineProperty(globalThis, 'navigator', {
      value: {
        clipboard: {
          write: async () => {},
          writeText: async () => {},
          readText: async () => {
            throw new Error('denied');
          },
        },
      },
      configurable: true,
    });
    try {
      await manager.copy(source, { start: { row: 0, col: 1 }, end: { row: 0, col: 2 } });
      const all = await manager.paste(3, 0, 20, 20);
      expect(all!.map((u) => u.value)).toEqual(['="00456"', '=LEFT("0077",4)']);
      const values = await manager.paste(3, 0, 20, 20, undefined, { values: true });
      expect(values!.map((u) => u.value)).toEqual(['="00456"', '="0077"']);
    } finally {
      Object.defineProperty(globalThis, 'navigator', { value: original, configurable: true });
    }
  });

  it('resizing an empty sheet updates formulas on other sheets that depend on its size', async () => {
    const el = await createSpreadsheet({}, 20, 8);
    el.addSheet('Other', { data: data({ '0:0': '=ROWS(Sheet1!A:A)' }) });
    expect(el.getSheetData('Other').get('0:0')?.displayValue).toBe('20');
    el.rows = 30;
    await el.updateComplete;
    expect(el.getSheetData('Other').get('0:0')?.displayValue).toBe('30');
  });
});

describe('ninth review regressions', () => {
  it('a whole-column spill shortens when content elsewhere stops reaching as far', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '0:1': '=A:A*2', '0:2': '=COUNTA(B:B)', '9:3': '1' }, 20, 8);
    expect(cell(el, '0:1')?.spill).toEqual({ rows: 10, cols: 1 });
    expect(display(el, '0:2')).toBe('10');
    await select(el, 9, 3);
    await press(el, 'Delete');
    expect(cell(el, '0:1')?.spill).toBeUndefined();
    expect(cell(el, '1:1')).toBeUndefined();
    expect(display(el, '0:2')).toBe('1');
  });

  it('copying a cell that holds a nested array keeps its values', async () => {
    const manager = new ClipboardManager();
    const source: GridData = new Map([
      ['1:0', {
        rawValue: '',
        displayValue: '4, 5, 6',
        type: 'text',
        spilledFrom: '0:0',
        array: { rows: 1, cols: 3, values: [4, 5, 6] },
      }],
      ['1:1', {
        rawValue: '={FILTER(A1:A9,A1:A9>1)}',
        displayValue: 'a, 2',
        type: 'text',
        array: { rows: 2, cols: 1, values: ['a', { rows: 1, cols: 2, values: [2, null] }] },
      }],
    ]);
    const original = globalThis.navigator;
    Object.defineProperty(globalThis, 'navigator', {
      value: {
        clipboard: {
          write: async () => {},
          writeText: async () => {},
          readText: async () => {
            throw new Error('denied');
          },
        },
      },
      configurable: true,
    });
    try {
      await manager.copy(source, { start: { row: 1, col: 0 }, end: { row: 1, col: 1 } });
      const values = await manager.paste(5, 0, 20, 20, undefined, { values: true });
      expect(values!.map((u) => u.value)).toEqual(['={{4,5,6}}', '={{"a";{2,""}}}']);
    } finally {
      Object.defineProperty(globalThis, 'navigator', { value: original, configurable: true });
    }

    // And the constant reads back as the same values
    const el = await createSpreadsheet({ '0:0': '={{4,5,6}}', '0:1': '=SUM(A1)' });
    expect(el.getCellValues('0:0')).toEqual({ rows: 1, cols: 3, values: [4, 5, 6] });
    expect(display(el, '0:1')).toBe('15');
  });

  it('a list item that starts like a formula survives an unchanged edit', async () => {
    const el = await createSpreadsheet({ '0:0': '={{"=1+1";"b"}}' });
    await select(el, 0, 0);
    await press(el, 'F2');
    expect(editor(el).value).toBe('"=1+1", b');
    await pressInEditor(el, 'Enter');
    expect(raw(el, '0:0')).toBe('={{"=1+1";"b"}}');
  });
});

describe('tenth review regressions', () => {
  it('a list with text that starts with # is pasted as values intact', async () => {
    const el = await createSpreadsheet({ '0:0': '#TODO, keep' });
    await select(el, 0, 0);
    el.convertToList();
    await el.updateComplete;
    expect(raw(el, '0:0')).toBe('={{"#TODO";"keep"}}');
    const manager = new ClipboardManager();
    const original = globalThis.navigator;
    Object.defineProperty(globalThis, 'navigator', {
      value: {
        clipboard: {
          write: async () => {},
          writeText: async () => {},
          readText: async () => {
            throw new Error('denied');
          },
        },
      },
      configurable: true,
    });
    try {
      await manager.copy(el.getData(), { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
      const values = await manager.paste(3, 0, 20, 20, undefined, { values: true });
      expect(values![0].value).toBe('={{"#TODO";"keep"}}');
    } finally {
      Object.defineProperty(globalThis, 'navigator', { value: original, configurable: true });
    }
    // A real error code inside an array is still written as one
    const withError = await createSpreadsheet({ '0:0': '={{1;#N/A}}', '0:1': '=ISNA(INDEX(A1,2))' });
    expect(display(withError, '0:1')).toBe('TRUE');
  });

  it('filling from a spilled text that looks like a number keeps it text', async () => {
    const el = await createSpreadsheet({ '0:0': '={"seed";"00456"}', '0:2': '={"x","007"}' });
    expect(cell(el, '1:0')).toMatchObject({ displayValue: '00456', type: 'text' });
    await select(el, 1, 0, 2, 0);
    await press(el, 'd', { ctrlKey: true });
    expect(cell(el, '2:0')).toMatchObject({ displayValue: '00456', type: 'text' });
    await select(el, 0, 3, 0, 4);
    await press(el, 'r', { ctrlKey: true });
    expect(cell(el, '0:4')).toMatchObject({ displayValue: '007', type: 'text' });
  });
});

describe('eleventh review regressions', () => {
  it('a name whose sheet was removed keeps working where it can, through a save and restore', async () => {
    const el = await createSpreadsheet({ '0:0': '100', '0:1': '=Amount', '0:2': '=A1*tax', '0:3': '=Total' });
    el.addSheet('Data', { data: data({ '0:0': '50' }) });
    el.defineName('Amount', '=A1', 'Data');
    el.defineName('tax', '=0.19', 'Data');
    el.defineName('Total', '=Sheet1!A1*2', 'Data');
    await el.updateComplete;
    expect([display(el, '0:1'), display(el, '0:2'), display(el, '0:3')]).toEqual(['50', '19', '200']);

    el.removeSheet('Data');
    await el.updateComplete;
    // Only the reference that needed the sheet is lost
    expect([display(el, '0:1'), display(el, '0:2'), display(el, '0:3')]).toEqual(['#REF!', '19', '200']);
    expect(el.getNames().map((n) => n.sheet)).toEqual(['', '', '']);

    el.setNames(el.getNames());
    await el.updateComplete;
    expect([display(el, '0:1'), display(el, '0:2'), display(el, '0:3')]).toEqual(['#REF!', '19', '200']);
    expect(el.getNames().map((n) => n.sheet)).toEqual(['', '', '']);
  });

  it('a series fill uses a visible spilled value whose formula row is filtered out', async () => {
    const el = await createSpreadsheet({
      '0:0': '10',
      '1:0': '=SEQUENCE(2)',
      '0:1': 'header',
      '1:1': 'hide',
      '2:1': 'show',
    });
    el.setFilter({ start: { row: 0, col: 1 }, end: { row: 2, col: 1 } }, [{ col: 1, values: ['show'] }]);
    await el.updateComplete;
    expect(el.getFilter()?.hiddenRows).toEqual([1]);
    await select(el, 0, 0, 4, 0);
    await press(el, 'd', { ctrlKey: true, altKey: true, code: 'KeyD' });
    // The visible seeds are 10 and 2: the series goes on with -6
    expect(raw(el, '3:0')).toBe('-6');
  });
});

describe('twelfth review regressions', () => {
  const fire = (el: Y11nSpreadsheet, type: string, detail?: unknown) =>
    shadow(el)
      .querySelector('y11n-name-manager')!
      .dispatchEvent(new CustomEvent(type, { detail, bubbles: true, composed: true }));
  const managerStatus = (el: Y11nSpreadsheet) =>
    (shadow(el).querySelector('y11n-name-manager') as unknown as { status: string }).status;

  it('a name that is called is not renamed to the name of a function', async () => {
    const el = await createSpreadsheet({ '0:0': '=double(4)', '0:1': '=rate*2' });
    el.registerFunction('TOTAL', (_ctx, v) => Number(v));
    el.defineName('double', '=LAMBDA(x,x*2)');
    el.defineName('rate', '=5');
    await el.openNameManager();
    for (const name of ['SUM', 'total']) {
      fire(el, 'name-save', { name, refersTo: '=LAMBDA(x,x*2)', sheet: 'Sheet1', previous: 'double' });
      await el.updateComplete;
      expect(managerStatus(el)).toContain(`double was not renamed to ${name}`);
      expect([raw(el, '0:0'), display(el, '0:0')]).toEqual(['=double(4)', '8']);
    }
    // A name that is only read can take a function's name: reading it is not a call
    fire(el, 'name-save', { name: 'Sum', refersTo: '=5', sheet: 'Sheet1', previous: 'rate' });
    await el.updateComplete;
    expect([raw(el, '0:1'), display(el, '0:1')]).toEqual(['=Sum*2', '10']);
  });
});

describe('thirteenth review regressions', () => {
  it('names-change fires when a sheet rename or removal changes the names', async () => {
    const el = await createSpreadsheet({ '0:1': '=prices' });
    el.addSheet('Orders', { data: data({ '0:0': '50' }) });
    let saved: Array<{ name: string; refersTo: string; sheet: string }> = [];
    let events = 0;
    el.addEventListener('names-change', (e) => {
      events++;
      saved = (e as CustomEvent<{ names: typeof saved }>).detail.names;
    });
    el.defineName('prices', '=A1', 'Orders');
    el.defineName('other', '=Sheet1!B1');
    expect(events).toBe(2);

    el.renameSheet('Orders', 'Data');
    expect(events).toBe(3);
    expect(saved).toEqual([
      { name: 'prices', refersTo: '=A1', sheet: 'Data' },
      { name: 'other', refersTo: '=Sheet1!B1', sheet: 'Sheet1' },
    ]);

    // A sheet no name has anything to do with changes none
    el.addSheet('Scratch');
    el.renameSheet('Scratch', 'Notes');
    el.removeSheet('Notes');
    expect(events).toBe(3);

    el.removeSheet('Data');
    expect(events).toBe(4);
    expect(saved[0]).toEqual({ name: 'prices', refersTo: '=A1', sheet: '' });

    // What was saved restores the workbook as it is: the name is not bound to a new sheet of that name
    el.addSheet('Data', { data: data({ '0:0': '99' }) });
    el.setNames(saved);
    await el.updateComplete;
    expect(display(el, '0:1')).toBe('#REF!');
  });
});
