// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import type { CellData, ConditionalFormatChangeDetail, GridData, SelectionRange } from '../types.js';
import type { Y11nSpreadsheet } from '../y11n-spreadsheet.js';
import type { Y11nConditionalFormatDialog } from '../components/y11n-conditional-format-dialog.js';
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
 *  5   Dee    Plum    7
 *  6   Eve    Kiwi    1
 */
const TABLE = {
  '0:0': 'Name', '0:1': 'Fruit', '0:2': 'Qty',
  '1:0': 'Ann', '1:1': 'Apple', '1:2': '3',
  '2:0': 'Bob', '2:1': 'Pear', '2:2': '5',
  '3:0': 'Cid', '3:1': 'Apple', '3:2': '2',
  '4:0': 'Dee', '4:1': 'Plum', '4:2': '7',
  '5:0': 'Eve', '5:1': 'Kiwi', '5:2': '1',
};
const QTY: SelectionRange = { start: { row: 1, col: 2 }, end: { row: 5, col: 2 } };
const FRUIT: SelectionRange = { start: { row: 1, col: 1 }, end: { row: 5, col: 1 } };
const RED = { backgroundColor: '#ffc7ce', textColor: '#9c0006' };

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
  shadow(el).querySelector(`[data-row="${row}"][data-col="${col}"]`) as HTMLElement;
const announcement = (el: Y11nSpreadsheet) => shadow(el).querySelector('[role="log"]')!.textContent!.trim();
const status = (el: Y11nSpreadsheet) => shadow(el).querySelector('[role="status"][aria-live]')!.textContent!.trim();
const label = (el: Y11nSpreadsheet, row: number, col: number) =>
  cell(el, row, col).querySelector('.ls-cf-label')?.textContent ?? null;
/** Rows of a column whose cells a conditional format applies to. */
const matched = (el: Y11nSpreadsheet, col: number, rows = [1, 2, 3, 4, 5]) =>
  rows.filter((r) => cell(el, r, col).classList.contains('cf-match'));

async function press(el: Y11nSpreadsheet, key: string, mods: KeyboardEventInit = {}) {
  grid(el).dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, composed: true, ...mods }));
  await el.updateComplete;
}

type Internals = {
  _selection: { moveTo(row: number, col: number, extend?: boolean): void };
  _undo(): void;
  _redo(): void;
  _commitRawValue(key: string, value: string): void;
};
const internals = (el: Y11nSpreadsheet) => el as unknown as Internals;

async function select(el: Y11nSpreadsheet, range: SelectionRange) {
  internals(el)._selection.moveTo(range.start.row, range.start.col);
  internals(el)._selection.moveTo(range.end.row, range.end.col, true);
  await el.updateComplete;
}

function changes(el: Y11nSpreadsheet): ConditionalFormatChangeDetail[] {
  const seen: ConditionalFormatChangeDetail[] = [];
  el.addEventListener('conditional-format-change', (e) => seen.push((e as CustomEvent).detail));
  return seen;
}

afterEach(() => {
  document.body.innerHTML = '';
});

describe('conditional formats: rendering', () => {
  it('merges the winning format over the cell without writing it into the data', async () => {
    const el = await createSpreadsheet();
    el.setCellFormat('2:2', { bold: true, backgroundColor: '#eeeeee' });
    const id = el.addConditionalFormat({
      range: QTY,
      condition: { type: 'cellValue', operator: 'greaterThan', value: 4 },
      format: RED,
    });
    await el.updateComplete;

    expect(id).toBe('cf-1');
    expect(matched(el, 2)).toEqual([2, 4]);
    const style = cell(el, 2, 2).style;
    // The rule's fill replaces the cell's own; the cell's bold stays
    expect(style.backgroundColor).toBe('rgb(255, 199, 206)');
    expect(style.color).toBe('rgb(156, 0, 6)');
    expect(style.fontWeight).toBe('bold');
    expect(cell(el, 1, 2).style.backgroundColor).toBe('');

    // Data, and so copy, sort and export, still see the real formats
    expect(el.getData().get('2:2')!.format).toEqual({ bold: true, backgroundColor: '#eeeeee' });
    expect(el.getData().get('4:2')!.format).toBeUndefined();
    expect(el.toCSV()).toContain('Bob,Pear,5');
  });

  it('exposes a label with the cell and announces it', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'aboveAverage' }, format: RED });
    await select(el, { start: { row: 2, col: 2 }, end: { row: 2, col: 2 } });

    expect(label(el, 2, 2)).toBe(', above average');
    expect(cell(el, 2, 2).textContent!.replace(/\s+/g, ' ').trim()).toBe('5 , above average');
    expect(label(el, 1, 2)).toBeNull();
    expect(status(el)).toBe('C3: 5, above average');

    el.updateConditionalFormat('cf-1', { label: 'strong seller' });
    await el.updateComplete;
    expect(label(el, 2, 2)).toBe(', strong seller');
    expect(status(el)).toBe('C3: 5, strong seller');
  });

  it('draws data bars and keeps the value readable', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'dataBar', color: '#336699' } });
    await el.updateComplete;

    const bar = cell(el, 4, 2).querySelector('.ls-cf-bar') as HTMLElement;
    expect(bar.getAttribute('aria-hidden')).toBe('true');
    expect(bar.style.left).toBe('0%');
    expect(bar.style.width).toBe('100%');
    expect(bar.style.backgroundColor).toBe('rgb(51, 102, 153)');
    expect((cell(el, 5, 2).querySelector('.ls-cf-bar') as HTMLElement).style.width).toBe('14.3%');
    expect(label(el, 5, 2)).toBe(', data bar 14%');
    expect(cell(el, 4, 2).querySelector('.cell-text')!.classList.contains('sr-only')).toBe(false);
    // Text cells get no bar
    expect(cell(el, 0, 2).querySelector('.ls-cf-bar')).toBeNull();
  });

  it('draws icons with a text alternative, and can hide the value visually only', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'iconSet', icons: 'arrows3', showValue: false } });
    await el.updateComplete;

    const icon = cell(el, 4, 2).querySelector('.ls-cf-icon')!;
    expect(icon.getAttribute('aria-hidden')).toBe('true');
    expect(icon.querySelector('svg')).not.toBeNull();
    expect(label(el, 4, 2)).toBe(', up arrow');
    expect(label(el, 5, 2)).toBe(', down arrow');
    // Hidden on screen but still part of the cell's accessible name
    const text = cell(el, 4, 2).querySelector('.cell-text')!;
    expect(text.classList.contains('sr-only')).toBe(true);
    expect(text.textContent).toBe('7');
  });

  it('fills a color scale, with a meter for forced-colors mode', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({
      range: QTY,
      condition: { type: 'colorScale', min: { type: 'min', color: '#000000' }, max: { type: 'max', color: '#ffffff' } },
    });
    await el.updateComplete;

    expect(cell(el, 5, 2).style.backgroundColor).toBe('rgb(0, 0, 0)');
    expect(cell(el, 5, 2).style.color).toBe('rgb(255, 255, 255)');
    expect(cell(el, 4, 2).style.backgroundColor).toBe('rgb(255, 255, 255)');
    expect((cell(el, 4, 2).querySelector('.ls-cf-meter') as HTMLElement).style.width).toBe('100%');
    expect(label(el, 4, 2)).toBe(', color scale 100%');
    // Not a highlight rule: no forced-colors edge marker
    expect(cell(el, 4, 2).classList.contains('cf-highlight')).toBe(false);
  });

  it('marks cells matched by highlight rules', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: FRUIT, condition: { type: 'duplicates' }, format: RED });
    await el.updateComplete;
    expect(matched(el, 1)).toEqual([1, 3]);
    expect(cell(el, 1, 1).classList.contains('cf-highlight')).toBe(true);
  });

  it('reports what the rules give a cell', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'top', rank: 1 }, format: RED, id: 'best' });
    expect(el.getConditionalFormatResult('4:2')).toEqual({ format: RED, labels: ['top 1'], ruleIds: ['best'] });
    expect(el.getConditionalFormatResult('1:2')).toBeNull();
    expect(el.getConditionalFormatResult('20:5')).toBeNull();
  });
});

describe('conditional formats: following the data', () => {
  it('restyles the whole range when one edit moves the average', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'aboveAverage' }, format: RED });
    await el.updateComplete;
    expect(matched(el, 2)).toEqual([2, 4]); // mean 3.6

    internals(el)._commitRawValue('5:2', '100');
    await el.updateComplete;
    expect(matched(el, 2)).toEqual([5]); // mean 23.4

    internals(el)._undo();
    await el.updateComplete;
    expect(matched(el, 2)).toEqual([2, 4]);
  });

  it('evaluates formula rules for each cell with the real engine', async () => {
    const el = await createSpreadsheet();
    // Highlight the whole row of every order with a quantity over 2
    el.addConditionalFormat({
      range: { start: { row: 1, col: 0 }, end: { row: 5, col: 2 } },
      condition: { type: 'formula', formula: '=$C2>2' },
      format: RED,
      label: 'big order',
    });
    await el.updateComplete;
    for (const col of [0, 1, 2]) expect(matched(el, col)).toEqual([1, 2, 4]);
    expect(label(el, 1, 0)).toBe(', big order');

    internals(el)._commitRawValue('3:2', '9');
    await el.updateComplete;
    expect(matched(el, 0)).toEqual([1, 2, 3, 4]);
  });

  it('follows formula cells when their precedents change', async () => {
    const el = await createSpreadsheet({ '0:0': '1', '0:1': '=A1*10', '1:1': '=A1+1' });
    el.addConditionalFormat({
      range: { start: { row: 0, col: 1 }, end: { row: 1, col: 1 } },
      condition: { type: 'cellValue', operator: 'greaterThan', value: 15 },
      format: RED,
    });
    await el.updateComplete;
    expect(matched(el, 1, [0, 1])).toEqual([]);

    internals(el)._commitRawValue('0:0', '2');
    await el.updateComplete;
    expect(matched(el, 1, [0, 1])).toEqual([0]);
  });

  it('compares the number behind a formatted cell, and operands that are formulas', async () => {
    const el = await createSpreadsheet({ '0:0': '0.5', '1:0': '0.05', '0:3': '0.2' });
    el.setRangeFormat({ start: { row: 0, col: 0 }, end: { row: 1, col: 0 } }, { numberFormat: { type: 'percent', decimals: 0 } });
    el.addConditionalFormat({
      range: { start: { row: 0, col: 0 }, end: { row: 1, col: 0 } },
      condition: { type: 'cellValue', operator: 'greaterThan', value: '=$D$1' },
      format: RED,
    });
    await el.updateComplete;
    expect(cell(el, 0, 0).querySelector('.cell-text')!.textContent).toBe('50%');
    expect(matched(el, 0, [0, 1])).toEqual([0]);

    internals(el)._commitRawValue('0:3', '0.01');
    await el.updateComplete;
    expect(matched(el, 0, [0, 1])).toEqual([0, 1]);
  });

  it('keeps rules when the data is replaced, and stays on its range when rows are sorted', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'top', rank: 1 }, format: RED });
    el.setData(data({ '1:2': '9', '2:2': '1' }));
    await el.updateComplete;
    expect(el.getConditionalFormats()).toHaveLength(1);
    expect(matched(el, 2)).toEqual([1]);

    el.sort({ start: { row: 1, col: 2 }, end: { row: 2, col: 2 } }, { col: 2 });
    await el.updateComplete;
    expect(matched(el, 2)).toEqual([2]);
    expect(el.getConditionalFormats()[0].range).toEqual(QTY);
  });
});

describe('conditional formats: rule API', () => {
  it('adds at the top by default, and returns copies', async () => {
    const el = await createSpreadsheet();
    const a = el.addConditionalFormat({ range: QTY, condition: { type: 'blanks' } });
    const b = el.addConditionalFormat({ range: QTY, condition: { type: 'errors' } });
    const c = el.addConditionalFormat({ range: QTY, condition: { type: 'unique' }, id: 'mine' }, 99);
    expect(el.getConditionalFormats().map((r) => r.id)).toEqual([b, a, c]);
    expect(c).toBe('mine');

    const rules = el.getConditionalFormats();
    rules[0].range.start.row = 20;
    expect(el.getConditionalFormats()[0].range.start.row).toBe(1);
  });

  it('applies precedence and stop-if-true in priority order', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'notBlanks' }, format: { backgroundColor: 'blue', italic: true }, id: 'low' });
    el.addConditionalFormat({ range: QTY, condition: { type: 'top', rank: 1 }, format: { backgroundColor: 'red' }, id: 'high' });
    expect(el.getConditionalFormatResult('4:2')!.format).toEqual({ backgroundColor: 'red', italic: true });

    el.updateConditionalFormat('high', { stopIfTrue: true });
    expect(el.getConditionalFormatResult('4:2')!.format).toEqual({ backgroundColor: 'red' });
    expect(el.getConditionalFormatResult('1:2')!.ruleIds).toEqual(['low']);

    expect(el.moveConditionalFormat('high', 1)).toBe(true);
    expect(el.getConditionalFormats().map((r) => r.id)).toEqual(['low', 'high']);
    expect(el.getConditionalFormatResult('4:2')!.format).toEqual({ backgroundColor: 'blue', italic: true });
  });

  it('updates, removes, replaces and clears rules', async () => {
    const el = await createSpreadsheet();
    const seen = changes(el);
    const id = el.addConditionalFormat({ range: QTY, condition: { type: 'blanks' } });
    expect(el.updateConditionalFormat(id, { range: FRUIT })).toBe(true);
    expect(el.getConditionalFormats()[0]).toMatchObject({ id, range: FRUIT, condition: { type: 'blanks' } });
    expect(el.updateConditionalFormat('nope', { range: FRUIT })).toBe(false);
    expect(el.moveConditionalFormat('nope', 0)).toBe(false);
    expect(el.removeConditionalFormat('nope')).toBe(false);
    expect(el.removeConditionalFormat(id)).toBe(true);
    expect(el.getConditionalFormats()).toEqual([]);

    el.setConditionalFormats([
      { range: QTY, condition: { type: 'duplicates' } },
      { range: FRUIT, condition: { type: 'unique' } },
    ]);
    expect(el.getConditionalFormats().map((r) => r.condition.type)).toEqual(['duplicates', 'unique']);
    el.clearConditionalFormats(FRUIT);
    expect(el.getConditionalFormats().map((r) => r.condition.type)).toEqual(['duplicates']);
    el.clearConditionalFormats();
    expect(el.getConditionalFormats()).toEqual([]);

    expect(seen.map((d) => [d.source, d.rules.length])).toEqual([
      ['programmatic', 1],
      ['programmatic', 1],
      ['programmatic', 0],
      ['programmatic', 2],
      ['programmatic', 1],
      ['programmatic', 0],
    ]);
  });

  it('rejects invalid rules without changing anything', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'blanks' }, id: 'a' });
    expect(() => el.addConditionalFormat({ range: QTY, condition: { type: 'nope' } as never })).toThrow(TypeError);
    expect(() => el.addConditionalFormat({ range: QTY, condition: { type: 'blanks' }, id: 'a' })).toThrow(/already exists/);
    expect(() =>
      el.addConditionalFormat({ range: { start: { row: 0, col: 0 }, end: { row: 99, col: 0 } }, condition: { type: 'blanks' } })
    ).toThrow(RangeError);
    expect(() =>
      el.setConditionalFormats([
        { range: QTY, condition: { type: 'blanks' }, id: 'x' },
        { range: QTY, condition: { type: 'errors' }, id: 'x' },
      ])
    ).toThrow(/unique/);
    expect(el.getConditionalFormats().map((r) => r.id)).toEqual(['a']);
  });

  it('undoes and redoes rule changes', async () => {
    const el = await createSpreadsheet();
    const seen = changes(el);
    el.addConditionalFormat({ range: QTY, condition: { type: 'top', rank: 1 }, format: RED });
    await el.updateComplete;
    expect(matched(el, 2)).toEqual([4]);

    internals(el)._undo();
    await el.updateComplete;
    expect(el.getConditionalFormats()).toEqual([]);
    expect(matched(el, 2)).toEqual([]);

    internals(el)._redo();
    await el.updateComplete;
    expect(matched(el, 2)).toEqual([4]);
    expect(seen.map((d) => d.source)).toEqual(['programmatic', 'undo', 'redo']);
  });
});

describe('conditional formats: rows and columns', () => {
  it('moves rules with their cells, in the same undo step', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({
      range: { start: { row: 1, col: 0 }, end: { row: 5, col: 2 } },
      condition: { type: 'formula', formula: '=$C2>4' },
      format: RED,
    });
    const seen = changes(el);
    el.insertRows(0, 2);
    await el.updateComplete;

    const [rule] = el.getConditionalFormats();
    expect(rule.range).toEqual({ start: { row: 3, col: 0 }, end: { row: 7, col: 2 } });
    expect(rule.condition).toEqual({ type: 'formula', formula: '=$C4>4' });
    expect(matched(el, 0, [3, 4, 5, 6, 7])).toEqual([4, 6]);
    expect(seen.map((d) => d.source)).toEqual(['programmatic']);

    internals(el)._undo();
    await el.updateComplete;
    expect(el.getConditionalFormats()[0].range).toEqual({ start: { row: 1, col: 0 }, end: { row: 5, col: 2 } });
    expect(matched(el, 0)).toEqual([2, 4]);
    expect(seen.map((d) => d.source)).toEqual(['programmatic', 'undo']);
  });

  it('follows inserted columns and drops rules whose cells are deleted', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'top', rank: 1 }, format: RED });
    el.insertColumns(0, 1);
    await el.updateComplete;
    expect(el.getConditionalFormats()[0].range).toEqual({ start: { row: 1, col: 3 }, end: { row: 5, col: 3 } });
    expect(matched(el, 3)).toEqual([4]);

    el.deleteColumns(3, 1);
    expect(el.getConditionalFormats()).toEqual([]);
    internals(el)._undo();
    expect(el.getConditionalFormats()).toHaveLength(1);
  });

  it('does not report a change when no rule is affected', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'top', rank: 1 }, format: RED });
    const seen = changes(el);
    el.insertRows(10, 1);
    expect(seen).toEqual([]);
  });
});

describe('conditional formats: user actions', () => {
  it('adds a preset rule to the selection and announces it', async () => {
    const el = await createSpreadsheet();
    const seen = changes(el);
    await select(el, QTY);
    const id = el.applyConditionalFormatPreset('aboveAverage');
    await el.updateComplete;

    expect(id).toBe('cf-1');
    expect(el.getConditionalFormats()[0]).toMatchObject({ range: QTY, condition: { type: 'aboveAverage' } });
    expect(matched(el, 2)).toEqual([2, 4]);
    expect(announcement(el)).toBe('Conditional format added to C2:C6: above average');
    expect(seen.map((d) => d.source)).toEqual(['user']);
  });

  it('lets a listener veto a user change', async () => {
    const el = await createSpreadsheet();
    let proposed: ConditionalFormatChangeDetail | null = null;
    el.addEventListener('before-conditional-format-change', (e) => {
      proposed = (e as CustomEvent).detail;
      e.preventDefault();
    });
    await select(el, QTY);
    expect(el.applyConditionalFormatPreset('dataBar')).toBeNull();
    await el.updateComplete;
    expect(proposed!.rules).toHaveLength(1);
    expect(el.getConditionalFormats()).toEqual([]);
    expect(announcement(el)).toBe('Conditional format change cancelled');
    // Programmatic changes are not vetoed
    el.addConditionalFormat({ range: QTY, condition: { type: 'dataBar' } });
    expect(el.getConditionalFormats()).toHaveLength(1);
  });

  it('refuses user changes in read-only mode but still shows rules', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'top', rank: 1 }, format: RED });
    el.readOnly = true;
    await el.updateComplete;
    expect(el.applyConditionalFormatPreset('dataBar')).toBeNull();
    expect(el.clearSelectionConditionalFormats()).toBe(false);
    expect(el.getConditionalFormats()).toHaveLength(1);
    expect(matched(el, 2)).toEqual([4]);
  });

  it('clears rules from the selection', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'top', rank: 1 }, id: 'qty' });
    el.addConditionalFormat({ range: { start: { row: 0, col: 0 }, end: { row: 5, col: 2 } }, condition: { type: 'blanks' }, id: 'all' });

    await select(el, { start: { row: 2, col: 1 }, end: { row: 3, col: 1 } });
    expect(el.clearSelectionConditionalFormats()).toBe(false);
    await el.updateComplete;
    expect(announcement(el)).toBe('Nothing cleared. 1 rule only partly in the selection was kept');

    await select(el, { start: { row: 0, col: 2 }, end: { row: 5, col: 2 } });
    expect(el.clearSelectionConditionalFormats()).toBe(true);
    await el.updateComplete;
    expect(el.getConditionalFormats()).toEqual([
      { id: 'all', range: { start: { row: 0, col: 0 }, end: { row: 5, col: 1 } }, condition: { type: 'blanks' } },
    ]);
    expect(announcement(el)).toBe('Conditional formats cleared from the selection: 1 rule removed, 1 rule trimmed');

    await select(el, QTY);
    expect(el.clearSelectionConditionalFormats()).toBe(false);
    await el.updateComplete;
    expect(announcement(el)).toBe('No conditional formats in the selection');
  });
});

describe('conditional formats: rule manager', () => {
  const dialog = (el: Y11nSpreadsheet) =>
    shadow(el).querySelector('y11n-conditional-format-dialog') as Y11nConditionalFormatDialog | null;
  const inDialog = <T extends HTMLElement>(el: Y11nSpreadsheet, selector: string) =>
    dialog(el)!.shadowRoot!.querySelector(selector) as T;
  const dialogStatus = (el: Y11nSpreadsheet) => inDialog(el, '[role="status"]').textContent!.trim();

  async function settle(el: Y11nSpreadsheet) {
    await el.updateComplete;
    await dialog(el)?.updateComplete;
    // Focus moves after the update
    await Promise.resolve();
  }

  async function type(el: Y11nSpreadsheet, selector: string, value: string) {
    const input = inDialog<HTMLInputElement>(el, selector);
    input.value = value;
    input.dispatchEvent(new Event(input.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
    await settle(el);
  }

  async function click(el: Y11nSpreadsheet, selector: string) {
    inDialog(el, selector).click();
    await settle(el);
  }

  async function submit(el: Y11nSpreadsheet) {
    inDialog(el, 'form').dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
    await settle(el);
  }

  async function open(el: Y11nSpreadsheet) {
    await press(el, 'C', { code: 'KeyC', altKey: true, shiftKey: true });
    await settle(el);
  }

  it('opens with Alt+Shift+C, lists the rules and closes with Escape', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'aboveAverage' }, format: RED });
    await open(el);

    const root = dialog(el)!.shadowRoot!;
    expect(root.querySelector('[role="dialog"]')!.getAttribute('aria-modal')).toBe('true');
    expect(root.querySelector('h2')!.textContent).toBe('Conditional formatting');
    expect(root.querySelector('.rule-desc')!.textContent).toBe('Above average');
    expect(root.querySelector('.rule-range')!.textContent).toBe('C2:C6');
    expect(root.activeElement?.id).toBe('cf-new');

    root.activeElement!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, composed: true }));
    await el.updateComplete;
    await Promise.resolve();
    expect(dialog(el)).toBeNull();
    expect(shadow(el).activeElement).toBe(cell(el, 0, 0));
  });

  it('adds a rule for the selection', async () => {
    const el = await createSpreadsheet();
    const seen = changes(el);
    await select(el, QTY);
    await open(el);
    expect(inDialog(el, '.empty').textContent).toBe('No rules yet.');

    await click(el, '#cf-new');
    const root = dialog(el)!.shadowRoot!;
    expect(root.querySelector('h2')!.textContent).toBe('New rule');
    expect(root.activeElement?.id).toBe('cf-range');
    expect(inDialog<HTMLInputElement>(el, '#cf-range').value).toBe('C2:C6');

    // Nothing to compare with yet
    await submit(el);
    expect(inDialog(el, '[role="alert"]').textContent).toBe('Enter a value to compare with');
    expect(el.getConditionalFormats()).toEqual([]);

    await type(el, '#cf-value', '4');
    await type(el, '#cf-label', 'restock soon');
    await submit(el);

    expect(el.getConditionalFormats()).toEqual([
      {
        id: 'cf-1',
        range: QTY,
        condition: { type: 'cellValue', operator: 'greaterThan', value: 4 },
        format: RED,
        label: 'restock soon',
      },
    ]);
    expect(matched(el, 2)).toEqual([2, 4]);
    expect(seen.map((d) => d.source)).toEqual(['user']);
    // Back in the list, with the outcome announced inside the dialog
    expect(inDialog(el, 'h2').textContent).toBe('Conditional formatting');
    expect(dialogStatus(el)).toBe('Rule added: greater than 4, C2:C6');
    expect(dialog(el)!.shadowRoot!.activeElement?.id).toBe('cf-new');
  });

  it('builds each kind of rule from the form', async () => {
    const el = await createSpreadsheet();
    await select(el, QTY);
    await open(el);

    await click(el, '#cf-new');
    await type(el, '#cf-kind', 'top');
    await type(el, '#cf-rank', '2');
    await type(el, '#cf-style', 'custom');
    inDialog<HTMLInputElement>(el, '#cf-bold').click();
    await settle(el);
    inDialog<HTMLInputElement>(el, '#cf-stop').click();
    await settle(el);
    await submit(el);

    await click(el, '#cf-new');
    await type(el, '#cf-kind', 'formula');
    await type(el, '#cf-range', 'a2:c6');
    await type(el, '#cf-formula', '$C2=1');
    await type(el, '#cf-style', 'green');
    await submit(el);

    await click(el, '#cf-new');
    await type(el, '#cf-kind', 'iconSet');
    await type(el, '#cf-icons', 'shapes3');
    await submit(el);

    await click(el, '#cf-new');
    await type(el, '#cf-kind', 'cellValue');
    await type(el, '#cf-comparison', 'between');
    await type(el, '#cf-value', '10%');
    await type(el, '#cf-value2', '=$A$1');
    await submit(el);

    expect(el.getConditionalFormats().map(({ id: _id, ...rule }) => rule)).toEqual([
      {
        range: QTY,
        condition: { type: 'cellValue', operator: 'between', value: 0.1, value2: '=$A$1' },
        format: RED,
      },
      { range: QTY, condition: { type: 'iconSet', icons: 'shapes3' } },
      {
        range: { start: { row: 1, col: 0 }, end: { row: 5, col: 2 } },
        condition: { type: 'formula', formula: '=$C2=1' },
        format: { backgroundColor: '#c6efce', textColor: '#006100' },
      },
      { range: QTY, condition: { type: 'top', rank: 2 }, format: { bold: true }, stopIfTrue: true },
    ]);
  });

  it('rejects a range that is not on the grid', async () => {
    const el = await createSpreadsheet();
    await open(el);
    await click(el, '#cf-new');
    await type(el, '#cf-kind', 'duplicates');
    await type(el, '#cf-range', 'nonsense');
    await submit(el);
    expect(inDialog(el, '[role="alert"]').textContent).toBe('Enter a range such as A1:C10');
    await type(el, '#cf-range', 'A1:Z99');
    await submit(el);
    expect(inDialog(el, '[role="alert"]').textContent).toBe('The range is outside the grid');
    expect(el.getConditionalFormats()).toEqual([]);

    // Escape goes back to the list, not out of the dialog
    inDialog(el, '#cf-range').dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, composed: true }));
    await settle(el);
    expect(inDialog(el, 'h2').textContent).toBe('Conditional formatting');
  });

  it('edits a rule, keeping settings the form has no fields for', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({
      range: QTY,
      condition: {
        type: 'iconSet',
        icons: 'arrows3',
        thresholds: [
          { type: 'number', value: 2 },
          { type: 'number', value: 5 },
        ],
      },
      id: 'icons',
    });
    await open(el);
    await click(el, '[data-rule="icons"][data-action="edit"]');
    expect(inDialog(el, 'h2').textContent).toBe('Edit rule');
    inDialog<HTMLInputElement>(el, '#cf-reverse').click();
    await settle(el);
    await submit(el);

    expect(el.getConditionalFormats()[0].condition).toEqual({
      type: 'iconSet',
      icons: 'arrows3',
      thresholds: [
        { type: 'number', value: 2 },
        { type: 'number', value: 5 },
      ],
      reverse: true,
    });
    expect(dialogStatus(el)).toBe('Rule saved: icon set, 3 arrows, C2:C6');
    expect((dialog(el)!.shadowRoot!.activeElement as HTMLElement).dataset.action).toBe('edit');
  });

  it('reorders, toggles stop-if-true and deletes, keeping focus in the dialog', async () => {
    const el = await createSpreadsheet();
    el.setConditionalFormats([
      { range: QTY, condition: { type: 'duplicates' }, id: 'a' },
      { range: QTY, condition: { type: 'unique' }, id: 'b' },
      { range: QTY, condition: { type: 'dataBar' }, id: 'c' },
    ]);
    await open(el);
    const root = dialog(el)!.shadowRoot!;
    const order = () => el.getConditionalFormats().map((r) => r.id);
    const focused = () => {
      const active = root.activeElement as HTMLElement | null;
      return active ? `${active.dataset.rule}:${active.dataset.action}` : null;
    };

    // The first rule can't move up; graphics have no stop-if-true
    expect(inDialog(el, '[data-rule="a"][data-action="up"]').getAttribute('aria-disabled')).toBe('true');
    expect(root.querySelector('[data-rule="c"][data-action="stop"]')).toBeNull();
    await click(el, '[data-rule="a"][data-action="up"]');
    expect(order()).toEqual(['a', 'b', 'c']);

    await click(el, '[data-rule="b"][data-action="up"]');
    expect(order()).toEqual(['b', 'a', 'c']);
    expect(dialogStatus(el)).toBe('Rule moved to position 1 of 3');
    // "Move up" is now disabled for it, so focus goes to "Move down"
    expect(focused()).toBe('b:down');

    await click(el, '[data-rule="b"][data-action="down"]');
    expect(order()).toEqual(['a', 'b', 'c']);
    expect(focused()).toBe('b:down');

    inDialog<HTMLInputElement>(el, '[data-rule="a"][data-action="stop"]').click();
    await settle(el);
    expect(el.getConditionalFormats()[0].stopIfTrue).toBe(true);

    await click(el, '[data-rule="b"][data-action="delete"]');
    expect(order()).toEqual(['a', 'c']);
    expect(dialogStatus(el)).toBe('Rule deleted: unique, C2:C6');
    expect(focused()).toBe('c:edit');

    // Every step is its own undo step
    internals(el)._undo();
    expect(order()).toEqual(['a', 'b', 'c']);
  });

  it('only shows the rules in read-only mode', async () => {
    const el = await createSpreadsheet();
    el.addConditionalFormat({ range: QTY, condition: { type: 'duplicates' }, id: 'a' });
    el.readOnly = true;
    expect(await el.openConditionalFormatManager()).toBe(true);
    await settle(el);
    expect(inDialog<HTMLButtonElement>(el, '#cf-new').disabled).toBe(true);
    expect(inDialog<HTMLButtonElement>(el, '[data-rule="a"][data-action="delete"]').disabled).toBe(true);
    expect(dialog(el)!.shadowRoot!.activeElement?.id).toBe('cf-close');
    await click(el, '#cf-close');
    expect(dialog(el)).toBeNull();
  });

  it('keeps Tab inside the dialog', async () => {
    const el = await createSpreadsheet();
    await open(el);
    const root = dialog(el)!.shadowRoot!;
    const close = inDialog(el, '#cf-close');
    close.focus();
    const tab = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, composed: true, cancelable: true });
    close.dispatchEvent(tab);
    expect(tab.defaultPrevented).toBe(true);
    expect(root.activeElement?.id).toBe('cf-new');
  });
});
