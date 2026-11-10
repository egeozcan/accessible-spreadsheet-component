import { describe, expect, it } from 'vitest';
import type {
  ConditionalFormatCondition,
  ConditionalFormatRule,
  ConditionalFormatRuleInput,
  SelectionRange,
} from '../../types.js';
import { dateToSerial } from '../../types.js';
import {
  ConditionalFormatEvaluator,
  HIGHLIGHT_STYLES,
  adjustRulesForStructureChange,
  clearRulesInRange,
  conditionalFormatPreset,
  describeCondition,
  normalizeConditionalRule,
  parseA1Range,
  rangeToA1,
  readableTextColor,
  type ConditionalFormatSource,
} from '../conditional-format.js';
import type { SortValue } from '../sort.js';

type Cell = number | string | boolean | null | { error: true };

function toValue(cell: Cell | undefined): SortValue {
  if (cell === null || cell === undefined || cell === '') return { kind: 'blank' };
  if (typeof cell === 'number') return { kind: 'number', value: cell };
  if (typeof cell === 'boolean') return { kind: 'boolean', value: cell };
  if (typeof cell === 'object') return { kind: 'error' };
  return { kind: 'text', value: cell };
}

const TODAY = dateToSerial(2026, 10, 10); // Tuesday 2026-11-10

/** A source over a 2D array, counting range scans and evaluating a few formulas by hand. */
function sheet(rows: Cell[][]) {
  const stats = { scans: 0, formulas: [] as string[] };
  const source: ConditionalFormatSource = {
    valueAt: (row, col) => toValue(rows[row]?.[col]),
    forEachValue: (range, visit) => {
      stats.scans++;
      for (let r = range.start.row; r <= range.end.row; r++) {
        for (let c = range.start.col; c <= range.end.col; c++) {
          const value = toValue(rows[r]?.[c]);
          if (value.kind !== 'blank') visit(value);
        }
      }
    },
    evaluate: (formula) => {
      stats.formulas.push(formula);
      // "=B<n>>5" and "=B<n>": enough to check that references are shifted
      const cmp = /^=B(\d+)>5$/.exec(formula);
      if (cmp) {
        const v = rows[Number(cmp[1]) - 1]?.[1];
        return { kind: 'boolean', value: typeof v === 'number' && v > 5 };
      }
      const ref = /^=B(\d+)$/.exec(formula);
      if (ref) return toValue(rows[Number(ref[1]) - 1]?.[1]);
      return { kind: 'error' };
    },
    today: () => TODAY,
  };
  return { source, stats, rows };
}

const range = (r0: number, c0: number, r1: number, c1: number): SelectionRange => ({
  start: { row: r0, col: c0 },
  end: { row: r1, col: c1 },
});

function rule(
  condition: ConditionalFormatCondition,
  r: SelectionRange,
  extra: Partial<ConditionalFormatRuleInput> = {},
  id = 'r1'
): ConditionalFormatRule {
  return normalizeConditionalRule({ range: r, condition, format: { bold: true }, ...extra }, id);
}

/** Rows (in column 0) that a single rule over A1:A<n> matches. */
function matches(values: Cell[], condition: ConditionalFormatCondition): number[] {
  const { source } = sheet(values.map((v) => [v]));
  const evaluator = new ConditionalFormatEvaluator(source);
  evaluator.setRules([rule(condition, range(0, 0, values.length - 1, 0))]);
  return values.map((_, i) => i).filter((i) => evaluator.resultFor(i, 0) !== null);
}

describe('normalizeConditionalRule', () => {
  it('normalizes the range and keeps only the visual format properties', () => {
    const r = normalizeConditionalRule(
      {
        range: { start: { row: 4, col: 2 }, end: { row: 1, col: 0 } },
        condition: { type: 'duplicates' },
        format: { bold: true, italic: false, textColor: '#f00', fontSize: 20 } as never,
        label: '  dupe  ',
        stopIfTrue: true,
      },
      'x'
    );
    expect(r).toEqual({
      id: 'x',
      range: range(1, 0, 4, 2),
      condition: { type: 'duplicates' },
      format: { bold: true, textColor: '#f00' },
      stopIfTrue: true,
      label: 'dupe',
    });
  });

  it('prefixes formulas with = and drops format and stop-if-true from graphics', () => {
    expect(rule({ type: 'formula', formula: 'A1>5' }, range(0, 0, 0, 0)).condition).toEqual({
      type: 'formula',
      formula: '=A1>5',
    });
    const bar = rule({ type: 'dataBar' }, range(0, 0, 0, 0), { stopIfTrue: true });
    expect(bar.format).toBeUndefined();
    expect(bar.stopIfTrue).toBeUndefined();
  });

  it('fills in color scale defaults', () => {
    const scale = rule(
      {
        type: 'colorScale',
        min: { color: '#F00' } as never,
        mid: { color: '#ffff00' } as never,
        max: { color: '#00ff00' } as never,
      },
      range(0, 0, 0, 0)
    );
    expect(scale.condition).toEqual({
      type: 'colorScale',
      min: { type: 'min', color: '#f00' },
      mid: { type: 'percentile', value: 50, color: '#ffff00' },
      max: { type: 'max', color: '#00ff00' },
    });
  });

  it.each<[string, unknown]>([
    ['no condition', { range: range(0, 0, 0, 0) }],
    ['a bad range', { range: { start: { row: -1, col: 0 }, end: { row: 0, col: 0 } }, condition: { type: 'blanks' } }],
    ['an unknown type', { range: range(0, 0, 0, 0), condition: { type: 'sparkle' } }],
    ['between without value2', { range: range(0, 0, 0, 0), condition: { type: 'cellValue', operator: 'between', value: 1 } }],
    ['empty text', { range: range(0, 0, 0, 0), condition: { type: 'text', operator: 'contains', text: '' } }],
    ['rank 0', { range: range(0, 0, 0, 0), condition: { type: 'top', rank: 0 } }],
    ['over 100 percent', { range: range(0, 0, 0, 0), condition: { type: 'top', rank: 101, percent: true } }],
    ['an empty formula', { range: range(0, 0, 0, 0), condition: { type: 'formula', formula: ' = ' } }],
    [
      'a named color in a scale',
      {
        range: range(0, 0, 0, 0),
        condition: { type: 'colorScale', min: { type: 'min', color: 'red' }, max: { type: 'max', color: '#0f0' } },
      },
    ],
    [
      'the wrong number of icon thresholds',
      {
        range: range(0, 0, 0, 0),
        condition: { type: 'iconSet', icons: 'arrows3', thresholds: [{ type: 'percent', value: 50 }] },
      },
    ],
    [
      'a percent threshold over 100',
      { range: range(0, 0, 0, 0), condition: { type: 'dataBar', max: { type: 'percent', value: 150 } } },
    ],
  ])('rejects %s', (_name, input) => {
    expect(() => normalizeConditionalRule(input as ConditionalFormatRuleInput, 'x')).toThrow(TypeError);
  });
});

describe('value rules', () => {
  it('compares numbers, and never matches blanks, errors or other kinds', () => {
    const values: Cell[] = [1, 5, 9, '', 'abc', { error: true }, true];
    expect(matches(values, { type: 'cellValue', operator: 'greaterThan', value: 4 })).toEqual([1, 2]);
    expect(matches(values, { type: 'cellValue', operator: 'lessThanOrEqual', value: 5 })).toEqual([0, 1]);
    expect(matches(values, { type: 'cellValue', operator: 'equal', value: '5' })).toEqual([1]);
    // Text and booleans are "not equal" to a number; blanks and errors still don't match
    expect(matches(values, { type: 'cellValue', operator: 'notEqual', value: 5 })).toEqual([0, 2, 4, 6]);
  });

  it('handles between and not between, with the bounds in either order', () => {
    const values: Cell[] = [1, 3, 5, 7, 'x', ''];
    expect(matches(values, { type: 'cellValue', operator: 'between', value: 5, value2: 3 })).toEqual([1, 2]);
    expect(matches(values, { type: 'cellValue', operator: 'notBetween', value: 3, value2: 5 })).toEqual([0, 3, 4]);
  });

  it('compares text without regard to case', () => {
    const values: Cell[] = ['apple', 'Banana', 'cherry', 3];
    expect(matches(values, { type: 'cellValue', operator: 'equal', value: 'BANANA' })).toEqual([1]);
    expect(matches(values, { type: 'cellValue', operator: 'greaterThan', value: 'b' })).toEqual([1, 2]);
  });

  it('evaluates formula operands shifted for each cell', () => {
    // A is compared with B on the same row
    const { source, stats } = sheet([
      [1, 0],
      [5, 9],
      [7, 2],
    ]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([rule({ type: 'cellValue', operator: 'greaterThan', value: '=B1' }, range(0, 0, 2, 0))]);
    expect([0, 1, 2].filter((r) => evaluator.resultFor(r, 0))).toEqual([0, 2]);
    expect(stats.formulas).toEqual(['=B1', '=B2', '=B3']);
  });

  it('tests text: contains, does not contain, begins with, ends with', () => {
    const values: Cell[] = ['Overdue', 'due soon', 'paid', 120, '', { error: true }];
    expect(matches(values, { type: 'text', operator: 'contains', text: 'DUE' })).toEqual([0, 1]);
    expect(matches(values, { type: 'text', operator: 'notContains', text: 'due' })).toEqual([2, 3]);
    expect(matches(values, { type: 'text', operator: 'beginsWith', text: 'due' })).toEqual([1]);
    expect(matches(values, { type: 'text', operator: 'endsWith', text: '20' })).toEqual([3]);
  });

  it('tests blanks and errors', () => {
    const values: Cell[] = [1, '', '  ', { error: true }, 'x'];
    expect(matches(values, { type: 'blanks' })).toEqual([1, 2]);
    expect(matches(values, { type: 'notBlanks' })).toEqual([0, 3, 4]);
    expect(matches(values, { type: 'errors' })).toEqual([3]);
    expect(matches(values, { type: 'notErrors' })).toEqual([0, 4]);
  });

  it('tests dates relative to today, with weeks starting on Sunday', () => {
    // Today is Tuesday 2026-11-10
    const d = (month: number, day: number) => dateToSerial(2026, month - 1, day);
    const values: Cell[] = [d(11, 9), d(11, 10) + 0.5, d(11, 11), d(11, 8), d(11, 14), d(11, 7), d(11, 15), d(10, 31), d(12, 1), 'x'];
    expect(matches(values, { type: 'date', period: 'yesterday' })).toEqual([0]);
    expect(matches(values, { type: 'date', period: 'today' })).toEqual([1]);
    expect(matches(values, { type: 'date', period: 'tomorrow' })).toEqual([2]);
    expect(matches(values, { type: 'date', period: 'thisWeek' })).toEqual([0, 1, 2, 3, 4]);
    expect(matches(values, { type: 'date', period: 'lastWeek' })).toEqual([5]);
    expect(matches(values, { type: 'date', period: 'nextWeek' })).toEqual([6]);
    expect(matches(values, { type: 'date', period: 'last7Days' })).toEqual([0, 1, 3, 5]);
    expect(matches(values, { type: 'date', period: 'thisMonth' })).toEqual([0, 1, 2, 3, 4, 5, 6]);
    expect(matches(values, { type: 'date', period: 'lastMonth' })).toEqual([7]);
    expect(matches(values, { type: 'date', period: 'nextMonth' })).toEqual([8]);
  });

  it('finds duplicates and unique values, ignoring case and blanks', () => {
    const values: Cell[] = ['a', 'A', 'b', 2, 2, 3, '', ''];
    expect(matches(values, { type: 'duplicates' })).toEqual([0, 1, 3, 4]);
    expect(matches(values, { type: 'unique' })).toEqual([2, 5]);
  });
});

describe('range-statistic rules', () => {
  const values: Cell[] = [10, 40, 20, 40, 30, 'n/a', ''];

  it('highlights the top and bottom N, including ties', () => {
    expect(matches(values, { type: 'top', rank: 1 })).toEqual([1, 3]);
    expect(matches(values, { type: 'top', rank: 3 })).toEqual([1, 3, 4]);
    expect(matches(values, { type: 'bottom', rank: 2 })).toEqual([0, 2]);
    expect(matches(values, { type: 'top', rank: 99 })).toEqual([0, 1, 2, 3, 4]);
  });

  it('highlights the top and bottom percent, at least one cell', () => {
    expect(matches(values, { type: 'top', rank: 40, percent: true })).toEqual([1, 3]);
    expect(matches(values, { type: 'bottom', rank: 1, percent: true })).toEqual([0]);
  });

  it('highlights above and below average', () => {
    // Mean of 10, 40, 20, 40, 30 is 28
    expect(matches(values, { type: 'aboveAverage' })).toEqual([1, 3, 4]);
    expect(matches(values, { type: 'belowAverage' })).toEqual([0, 2]);
    expect(matches([1, 2, 3], { type: 'aboveAverage' })).toEqual([2]);
    expect(matches([1, 2, 3], { type: 'aboveAverage', orEqual: true })).toEqual([1, 2]);
    expect(matches([1, 2, 3], { type: 'belowAverage', orEqual: true })).toEqual([0, 1]);
  });

  it('scans a range once, and again only after a cell inside it changes', () => {
    const { source, stats, rows } = sheet([[1, 100], [2, 100], [3, 100], [4, 100]]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([rule({ type: 'aboveAverage' }, range(0, 0, 3, 0))]);
    for (let r = 0; r < 4; r++) evaluator.resultFor(r, 0);
    expect(stats.scans).toBe(1);

    // A change outside the range keeps the statistics (cell results are recomputed)
    rows[0][1] = 5;
    evaluator.invalidate(['0:1']);
    for (let r = 0; r < 4; r++) evaluator.resultFor(r, 0);
    expect(stats.scans).toBe(1);

    // A change inside it drops them: one edit restyles the whole range
    rows[3][0] = 100;
    evaluator.invalidate(['3:0']);
    expect([0, 1, 2, 3].filter((r) => evaluator.resultFor(r, 0))).toEqual([3]);
    expect(stats.scans).toBe(2);

    evaluator.invalidate();
    evaluator.resultFor(0, 0);
    expect(stats.scans).toBe(3);
  });
});

describe('formula rules', () => {
  it('shifts the formula from the top-left cell, like a filled formula', () => {
    const { source, stats } = sheet([
      ['a', 9],
      ['b', 1],
      ['c', 6],
    ]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([rule({ type: 'formula', formula: '=B1>5' }, range(0, 0, 2, 0))]);
    expect([0, 1, 2].filter((r) => evaluator.resultFor(r, 0))).toEqual([0, 2]);
    expect(stats.formulas).toEqual(['=B1>5', '=B2>5', '=B3>5']);
    // Cached until something changes
    evaluator.resultFor(0, 0);
    expect(stats.formulas).toHaveLength(3);
    evaluator.invalidate(['5:5']);
    evaluator.resultFor(0, 0);
    expect(stats.formulas).toHaveLength(4);
  });

  it('treats non-zero numbers as true and text or errors as false', () => {
    const { source } = sheet([[1, 3], [1, 0], [1, 'x'], [1, { error: true }]]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([rule({ type: 'formula', formula: '=B1' }, range(0, 0, 3, 0))]);
    expect([0, 1, 2, 3].filter((r) => evaluator.resultFor(r, 0))).toEqual([0]);
  });
});

describe('precedence', () => {
  const red = { backgroundColor: 'red', textColor: 'white' };
  const blue = { backgroundColor: 'blue', italic: true };

  it('merges formats, the earlier rule winning conflicts', () => {
    const { source } = sheet([[10]]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([
      rule({ type: 'cellValue', operator: 'greaterThan', value: 5 }, range(0, 0, 0, 0), { format: red }, 'a'),
      rule({ type: 'notBlanks' }, range(0, 0, 0, 0), { format: blue, label: 'filled' }, 'b'),
    ]);
    expect(evaluator.resultFor(0, 0)).toEqual({
      format: { backgroundColor: 'red', textColor: 'white', italic: true },
      labels: ['greater than 5', 'filled'],
      ruleIds: ['a', 'b'],
    });
  });

  it('stops at a matching stop-if-true rule, but not at one that does not match', () => {
    const { source } = sheet([[10], [1]]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([
      rule({ type: 'cellValue', operator: 'greaterThan', value: 5 }, range(0, 0, 1, 0), { format: red, stopIfTrue: true }, 'a'),
      rule({ type: 'notBlanks' }, range(0, 0, 1, 0), { format: blue }, 'b'),
      rule({ type: 'dataBar' }, range(0, 0, 1, 0), {}, 'c'),
    ]);
    expect(evaluator.resultFor(0, 0)?.ruleIds).toEqual(['a']);
    expect(evaluator.resultFor(1, 0)?.ruleIds).toEqual(['b', 'c']);
  });

  it('only applies rules to cells in their range', () => {
    const { source } = sheet([[10, 10]]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([rule({ type: 'notBlanks' }, range(0, 0, 0, 0))]);
    expect(evaluator.resultFor(0, 0)).not.toBeNull();
    expect(evaluator.resultFor(0, 1)).toBeNull();
  });
});

describe('graphics', () => {
  it('colors a 2-color scale and reports each cell’s position', () => {
    const { source } = sheet([[0], [50], [100], ['text']]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([
      rule(
        { type: 'colorScale', min: { type: 'min', color: '#000000' }, max: { type: 'max', color: '#ffffff' } },
        range(0, 0, 3, 0)
      ),
    ]);
    expect(evaluator.resultFor(0, 0)).toMatchObject({
      format: { backgroundColor: '#000000', textColor: '#ffffff' },
      scale: 0,
      labels: ['color scale 0%'],
    });
    expect(evaluator.resultFor(1, 0)).toMatchObject({ format: { backgroundColor: '#808080' }, scale: 0.5 });
    expect(evaluator.resultFor(2, 0)).toMatchObject({
      format: { backgroundColor: '#ffffff', textColor: '#000000' },
      labels: ['color scale 100%'],
    });
    expect(evaluator.resultFor(3, 0)).toBeNull();
  });

  it('runs a 3-color scale through its midpoint', () => {
    const { source } = sheet([[0], [10], [100]]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([
      rule(
        {
          type: 'colorScale',
          min: { type: 'min', color: '#ff0000' },
          mid: { type: 'number', value: 10, color: '#ffff00' },
          max: { type: 'max', color: '#00ff00' },
        },
        range(0, 0, 2, 0),
        { label: 'heat' }
      ),
    ]);
    expect(evaluator.resultFor(1, 0)).toMatchObject({
      format: { backgroundColor: '#ffff00' },
      scale: 0.5,
      labels: ['heat 50%'],
    });
  });

  it('does not override the fill of an earlier highlight rule', () => {
    const { source } = sheet([[0], [100]]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([
      rule({ type: 'cellValue', operator: 'equal', value: 100 }, range(0, 0, 1, 0), { format: { backgroundColor: 'gold' } }, 'a'),
      rule(
        { type: 'colorScale', min: { type: 'min', color: '#000' }, max: { type: 'max', color: '#fff' } },
        range(0, 0, 1, 0),
        {},
        'b'
      ),
    ]);
    expect(evaluator.resultFor(1, 0)?.format).toEqual({ backgroundColor: 'gold' });
  });

  it('draws data bars from zero, with negative bars to the left of the axis', () => {
    const { source } = sheet([[-50], [0], [50], [100]]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([rule({ type: 'dataBar', negativeColor: 'crimson' }, range(0, 0, 3, 0))]);
    const third = 1 / 3;
    expect(evaluator.resultFor(0, 0)?.dataBar).toEqual({
      start: 0,
      size: third,
      negative: true,
      color: 'crimson',
      showValue: true,
    });
    expect(evaluator.resultFor(1, 0)?.dataBar).toMatchObject({ size: 0, negative: false });
    const positive = evaluator.resultFor(2, 0)!;
    expect(positive.dataBar!.start).toBeCloseTo(third);
    expect(positive.dataBar!.size).toBeCloseTo(third);
    expect(positive.labels).toEqual(['data bar 67%']);
    expect(evaluator.resultFor(3, 0)?.dataBar?.size).toBeCloseTo(2 * third);
  });

  it('fills the bar for the largest value when all are positive', () => {
    const { source } = sheet([[25], [100]]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([rule({ type: 'dataBar', showValue: false }, range(0, 0, 1, 0))]);
    expect(evaluator.resultFor(0, 0)?.dataBar).toEqual({ start: 0, size: 0.25, negative: false, showValue: false });
    expect(evaluator.resultFor(1, 0)?.dataBar).toMatchObject({ start: 0, size: 1 });
  });

  it('picks icons by threshold and names them', () => {
    const { source } = sheet([[0], [50], [100], ['x']]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([rule({ type: 'iconSet', icons: 'arrows3' }, range(0, 0, 3, 0))]);
    expect(evaluator.resultFor(0, 0)).toMatchObject({
      icon: { set: 'arrows3', index: 0, label: 'down arrow', showValue: true },
      labels: ['down arrow'],
    });
    expect(evaluator.resultFor(1, 0)?.icon).toMatchObject({ index: 1, label: 'sideways arrow' });
    expect(evaluator.resultFor(2, 0)?.icon).toMatchObject({ index: 2, label: 'up arrow' });
    expect(evaluator.resultFor(3, 0)).toBeNull();
  });

  it('supports reversed icons, custom thresholds and custom text alternatives', () => {
    const { source } = sheet([[1], [5], [9]]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([
      rule(
        {
          type: 'iconSet',
          icons: 'symbols3',
          reverse: true,
          thresholds: [
            { type: 'number', value: 3 },
            { type: 'number', value: 8 },
          ],
          labels: ['late', 'at risk', 'on time'],
        },
        range(0, 0, 2, 0),
        { label: 'Status' }
      ),
    ]);
    expect(evaluator.resultFor(0, 0)).toMatchObject({ icon: { index: 2, label: 'on time' }, labels: ['Status: on time'] });
    expect(evaluator.resultFor(1, 0)?.icon).toMatchObject({ index: 1, label: 'at risk' });
    expect(evaluator.resultFor(2, 0)?.icon).toMatchObject({ index: 0, label: 'late' });
  });

  it('gives a cell at most one bar, one icon and one scale', () => {
    const { source } = sheet([[1], [2]]);
    const evaluator = new ConditionalFormatEvaluator(source);
    evaluator.setRules([
      rule({ type: 'dataBar', color: 'navy' }, range(0, 0, 1, 0), {}, 'a'),
      rule({ type: 'dataBar', color: 'teal' }, range(0, 0, 1, 0), {}, 'b'),
      rule({ type: 'iconSet', icons: 'bars5' }, range(0, 0, 1, 0), {}, 'c'),
    ]);
    const result = evaluator.resultFor(1, 0)!;
    expect(result.ruleIds).toEqual(['a', 'c']);
    expect(result.dataBar?.color).toBe('navy');
    expect(result.icon?.label).toBe('4 of 4 bars');
  });
});

describe('readableTextColor', () => {
  it('picks black on light colors and white on dark ones', () => {
    expect(readableTextColor([255, 235, 132])).toBe('#000000');
    expect(readableTextColor([99, 190, 123])).toBe('#000000');
    expect(readableTextColor([20, 40, 120])).toBe('#ffffff');
  });
});

describe('describeCondition', () => {
  it.each<[ConditionalFormatCondition, string]>([
    [{ type: 'cellValue', operator: 'greaterThan', value: 100 }, 'greater than 100'],
    [{ type: 'cellValue', operator: 'between', value: 1, value2: '=B1' }, 'between 1 and =B1'],
    [{ type: 'text', operator: 'notContains', text: 'x' }, 'does not contain "x"'],
    [{ type: 'date', period: 'last7Days' }, 'dated in the last 7 days'],
    [{ type: 'top', rank: 10, percent: true }, 'top 10%'],
    [{ type: 'bottom', rank: 3 }, 'bottom 3'],
    [{ type: 'aboveAverage' }, 'above average'],
    [{ type: 'belowAverage', orEqual: true }, 'average or below'],
    [{ type: 'duplicates' }, 'duplicate'],
    [{ type: 'formula', formula: '=A1>5' }, 'formula =A1>5'],
    [{ type: 'iconSet', icons: 'arrows5' }, 'icon set, 5 arrows'],
  ])('%j', (condition, text) => {
    expect(describeCondition(condition)).toBe(text);
  });
});

describe('presets', () => {
  it('pairs each highlight preset with a readable style', () => {
    expect(conditionalFormatPreset('aboveAverage')).toEqual({
      condition: { type: 'aboveAverage' },
      format: HIGHLIGHT_STYLES.green,
    });
    expect(conditionalFormatPreset('top10').condition).toEqual({ type: 'top', rank: 10 });
    expect(conditionalFormatPreset('colorScale').condition.type).toBe('colorScale');
    expect(() => conditionalFormatPreset('nope' as never)).toThrow(TypeError);
  });

  it('every preset is a valid rule', () => {
    for (const preset of ['duplicates', 'unique', 'aboveAverage', 'belowAverage', 'top10', 'bottom10', 'colorScale', 'dataBar', 'iconSet'] as const) {
      expect(() => normalizeConditionalRule({ range: range(0, 0, 1, 1), ...conditionalFormatPreset(preset) }, 'x')).not.toThrow();
    }
  });
});

describe('A1 ranges', () => {
  it('formats and parses ranges', () => {
    expect(rangeToA1(range(1, 1, 9, 3))).toBe('B2:D10');
    expect(rangeToA1(range(0, 0, 0, 0))).toBe('A1');
    expect(parseA1Range(' d10 : $B$2 ')).toEqual(range(1, 1, 9, 3));
    expect(parseA1Range('AA5')).toEqual(range(4, 26, 4, 26));
    expect(parseA1Range('A0')).toBeNull();
    expect(parseA1Range('A1:')).toBeNull();
    expect(parseA1Range('hello')).toBeNull();
  });
});

describe('adjustRulesForStructureChange', () => {
  const base = [
    rule({ type: 'formula', formula: '=$B2>5' }, range(1, 0, 4, 0), {}, 'f'),
    rule({ type: 'cellValue', operator: 'greaterThan', value: '=$D$1' }, range(1, 2, 4, 2), {}, 'v'),
  ];

  it('returns the same array when nothing moves', () => {
    expect(adjustRulesForStructureChange(base, { axis: 'row', kind: 'insert', index: 9, count: 1 })).toBe(base);
  });

  it('moves ranges and formulas down when rows are inserted above', () => {
    const out = adjustRulesForStructureChange(base, { axis: 'row', kind: 'insert', index: 0, count: 2 });
    expect(out[0].range).toEqual(range(3, 0, 6, 0));
    expect(out[0].condition).toEqual({ type: 'formula', formula: '=$B4>5' });
    expect(out[1].range).toEqual(range(3, 2, 6, 2));
    expect(out[1].condition).toMatchObject({ value: '=$D$3' });
  });

  it('grows a range when rows are inserted inside it', () => {
    const out = adjustRulesForStructureChange(base, { axis: 'row', kind: 'insert', index: 2, count: 3 });
    expect(out[0].range).toEqual(range(1, 0, 7, 0));
    expect(out[0].condition).toEqual({ type: 'formula', formula: '=$B2>5' });
  });

  it('shrinks ranges and drops rules whose cells are all deleted', () => {
    const shrunk = adjustRulesForStructureChange(base, { axis: 'row', kind: 'delete', index: 3, count: 5 });
    expect(shrunk.map((r) => r.range)).toEqual([range(1, 0, 2, 0), range(1, 2, 2, 2)]);
    const dropped = adjustRulesForStructureChange(base, { axis: 'col', kind: 'delete', index: 0, count: 1 });
    expect(dropped.map((r) => r.id)).toEqual(['v']);
    expect(dropped[0].range).toEqual(range(1, 1, 4, 1));
    expect(dropped[0].condition).toMatchObject({ value: '=$C$1' });
  });

  it('re-anchors the formula when the first row of the range is deleted', () => {
    const out = adjustRulesForStructureChange(base, { axis: 'row', kind: 'delete', index: 0, count: 3 });
    // Rows 1-3 are gone: the range is now A1:A2 and its first cell is the old row 4
    expect(out[0].range).toEqual(range(0, 0, 1, 0));
    expect(out[0].condition).toEqual({ type: 'formula', formula: '=$B1>5' });
  });

  it('moves columns the same way', () => {
    const out = adjustRulesForStructureChange(base, { axis: 'col', kind: 'insert', index: 1, count: 1 });
    expect(out[0].range).toEqual(range(1, 0, 4, 0));
    expect(out[0].condition).toEqual({ type: 'formula', formula: '=$C2>5' });
    expect(out[1].range).toEqual(range(1, 3, 4, 3));
  });
});

describe('clearRulesInRange', () => {
  const rules = [rule({ type: 'formula', formula: '=A1>5' }, range(0, 0, 9, 2), {}, 'a')];

  it('removes rules inside the range', () => {
    expect(clearRulesInRange(rules, range(0, 0, 20, 5))).toMatchObject({ rules: [], removed: 1, trimmed: 0, kept: 0 });
  });

  it('leaves rules it does not touch', () => {
    const out = clearRulesInRange(rules, range(12, 0, 20, 5));
    expect(out.rules).toBe(rules);
    expect(out).toMatchObject({ removed: 0, trimmed: 0, kept: 0 });
  });

  it('trims a rule cut across a whole edge, keeping its formula on the same cells', () => {
    const top = clearRulesInRange(rules, range(0, 0, 3, 2));
    expect(top).toMatchObject({ removed: 0, trimmed: 1 });
    expect(top.rules[0].range).toEqual(range(4, 0, 9, 2));
    expect(top.rules[0].condition).toEqual({ type: 'formula', formula: '=A5>5' });

    const right = clearRulesInRange(rules, range(0, 2, 9, 4));
    expect(right.rules[0].range).toEqual(range(0, 0, 9, 1));
    expect(right.rules[0].condition).toEqual({ type: 'formula', formula: '=A1>5' });
  });

  it('keeps a rule that would not stay a rectangle', () => {
    const out = clearRulesInRange(rules, range(3, 1, 4, 1));
    expect(out.rules).toBe(rules);
    expect(out).toMatchObject({ removed: 0, trimmed: 0, kept: 1 });
    // The middle rows, full width: two rectangles would remain
    expect(clearRulesInRange(rules, range(3, 0, 4, 2)).kept).toBe(1);
  });
});
