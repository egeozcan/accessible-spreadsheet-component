/**
 * Conditional formatting: rule validation, evaluation and the helpers the
 * rule manager uses. Rules are plain data (`ConditionalFormatRule`); the
 * `ConditionalFormatEvaluator` works out what they give each cell.
 */
import {
  type ConditionalColorStop,
  type ConditionalDatePeriod,
  type ConditionalFormatCondition,
  type ConditionalFormatPreset,
  type ConditionalFormatResult,
  type ConditionalFormatRule,
  type ConditionalFormatRuleInput,
  type ConditionalFormatStyle,
  type ConditionalIconSetName,
  type ConditionalOperand,
  type ConditionalThreshold,
  type SelectionRange,
  cellKey,
  coordToRef,
  dateSerialToDate,
  dateToSerial,
  parseKey,
  refToCoord,
} from '../types.js';
import { adjustFormulaForStructureChange, shiftFormulaReferences, type StructureChange } from './references.js';
import { compareSortValues, type SortValue } from './sort.js';

/** Where the evaluator reads cell values from. */
export interface ConditionalFormatSource {
  /** The value of one cell */
  valueAt(row: number, col: number): SortValue;
  /** Call `visit` for every non-empty cell of a range, in any order */
  forEachValue(range: SelectionRange, visit: (value: SortValue) => void): void;
  /**
   * Evaluate formula text (starting with `=`) on its own. `at` is the cell
   * the rule is being tested for, which `ROW()` and `COLUMN()` report.
   */
  evaluate(formula: string, at?: { row: number; col: number }): SortValue;
  /** Today's date as a date serial number */
  today(): number;
}

// ─── Icon sets ──────────────────────────────────────────

/** Text alternatives of each icon set's icons, lowest first. */
export const ICON_SET_LABELS: Record<ConditionalIconSetName, readonly string[]> = {
  arrows3: ['down arrow', 'sideways arrow', 'up arrow'],
  arrows5: ['down arrow', 'diagonal down arrow', 'sideways arrow', 'diagonal up arrow', 'up arrow'],
  shapes3: ['red diamond', 'yellow triangle', 'green circle'],
  symbols3: ['cross', 'exclamation mark', 'check mark'],
  bars5: ['0 of 4 bars', '1 of 4 bars', '2 of 4 bars', '3 of 4 bars', '4 of 4 bars'],
};

/** Names of the icon sets, for menus. */
export const ICON_SET_NAMES: Record<ConditionalIconSetName, string> = {
  arrows3: '3 arrows',
  arrows5: '5 arrows',
  shapes3: '3 shapes',
  symbols3: '3 symbols',
  bars5: '5 bars',
};

// ─── Presets ────────────────────────────────────────────

/** Excel's stock highlight styles; each pairs a fill with a text color that keeps 4.5:1 contrast. */
export const HIGHLIGHT_STYLES = {
  red: { backgroundColor: '#ffc7ce', textColor: '#9c0006' },
  yellow: { backgroundColor: '#ffeb9c', textColor: '#7a4400' },
  green: { backgroundColor: '#c6efce', textColor: '#006100' },
} as const satisfies Record<string, ConditionalFormatStyle>;

/** The condition and format of a one-step preset. */
export function conditionalFormatPreset(
  preset: ConditionalFormatPreset
): Pick<ConditionalFormatRule, 'condition' | 'format'> {
  switch (preset) {
    case 'duplicates':
      return { condition: { type: 'duplicates' }, format: { ...HIGHLIGHT_STYLES.red } };
    case 'unique':
      return { condition: { type: 'unique' }, format: { ...HIGHLIGHT_STYLES.yellow } };
    case 'aboveAverage':
      return { condition: { type: 'aboveAverage' }, format: { ...HIGHLIGHT_STYLES.green } };
    case 'belowAverage':
      return { condition: { type: 'belowAverage' }, format: { ...HIGHLIGHT_STYLES.red } };
    case 'top10':
      return { condition: { type: 'top', rank: 10 }, format: { ...HIGHLIGHT_STYLES.green } };
    case 'bottom10':
      return { condition: { type: 'bottom', rank: 10 }, format: { ...HIGHLIGHT_STYLES.red } };
    case 'colorScale':
      return {
        condition: {
          type: 'colorScale',
          min: { type: 'min', color: '#f8696b' },
          mid: { type: 'percentile', value: 50, color: '#ffeb84' },
          max: { type: 'max', color: '#63be7b' },
        },
      };
    case 'dataBar':
      return { condition: { type: 'dataBar' } };
    case 'iconSet':
      return { condition: { type: 'iconSet', icons: 'arrows3' } };
    default:
      throw new TypeError(`conditional format: unknown preset "${String(preset)}"`);
  }
}

// ─── Validation ─────────────────────────────────────────

const COMPARISONS = new Set([
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
  'equal',
  'notEqual',
  'between',
  'notBetween',
]);
const TEXT_OPERATORS = new Set(['contains', 'notContains', 'beginsWith', 'endsWith']);
const DATE_PERIODS = new Set<ConditionalDatePeriod>([
  'yesterday',
  'today',
  'tomorrow',
  'last7Days',
  'lastWeek',
  'thisWeek',
  'nextWeek',
  'lastMonth',
  'thisMonth',
  'nextMonth',
]);
const THRESHOLD_TYPES = new Set(['min', 'max', 'number', 'percent', 'percentile']);
const STYLE_KEYS = ['bold', 'italic', 'underline', 'strikethrough', 'textColor', 'backgroundColor'] as const;

function fail(message: string): never {
  throw new TypeError(`conditional format: ${message}`);
}

/** Whether a condition draws something in the cell instead of applying a format. */
export function isGraphicCondition(condition: ConditionalFormatCondition): boolean {
  return condition.type === 'colorScale' || condition.type === 'dataBar' || condition.type === 'iconSet';
}

/** `#rgb` / `#rrggbb` to `[r, g, b]`, or null. */
export function parseHexColor(color: string): [number, number, number] | null {
  const m = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(color.trim());
  if (!m) return null;
  const hex = m[1].length === 3 ? [...m[1]].map((ch) => ch + ch).join('') : m[1];
  return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
}

function normalizeOperand(value: unknown, name: string): ConditionalOperand {
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) fail(`${name} must be a finite number`);
    return value;
  }
  if (typeof value !== 'string' || value.trim() === '') fail(`${name} must be a number or a non-empty string`);
  return value;
}

function normalizeThreshold(t: unknown, name: string, fallback?: ConditionalThreshold): ConditionalThreshold {
  if (t === undefined && fallback) return { ...fallback };
  const threshold = t as ConditionalThreshold;
  if (!threshold || !THRESHOLD_TYPES.has(threshold.type)) {
    fail(`${name} needs a type: min, max, number, percent or percentile`);
  }
  if (threshold.type === 'min' || threshold.type === 'max') return { type: threshold.type };
  if (typeof threshold.value !== 'number' || !Number.isFinite(threshold.value)) {
    fail(`${name} of type ${threshold.type} needs a numeric value`);
  }
  if (threshold.type !== 'number' && (threshold.value < 0 || threshold.value > 100)) {
    fail(`${name}: a ${threshold.type} must be between 0 and 100`);
  }
  return { type: threshold.type, value: threshold.value };
}

function normalizeColorStop(stop: unknown, name: string, fallbackType: 'min' | 'max'): ConditionalColorStop {
  const s = stop as ConditionalColorStop;
  if (!s || typeof s.color !== 'string' || !parseHexColor(s.color)) {
    fail(`${name} needs a color as #rgb or #rrggbb`);
  }
  const threshold = normalizeThreshold(s.type === undefined ? undefined : s, name, { type: fallbackType });
  return { ...threshold, color: s.color.trim().toLowerCase() };
}

function normalizeStyle(format: unknown): ConditionalFormatStyle | undefined {
  if (format === undefined || format === null) return undefined;
  if (typeof format !== 'object') fail('format must be an object');
  const style: ConditionalFormatStyle = {};
  const source = format as Record<string, unknown>;
  for (const key of STYLE_KEYS) {
    const value = source[key];
    if (value === undefined || value === null || value === false || value === '') continue;
    if (key === 'textColor' || key === 'backgroundColor') {
      if (typeof value !== 'string') fail(`format.${key} must be a CSS color string`);
      style[key] = value;
    } else {
      style[key] = true;
    }
  }
  return Object.keys(style).length > 0 ? style : undefined;
}

function normalizeCondition(input: unknown): ConditionalFormatCondition {
  const c = input as ConditionalFormatCondition;
  if (!c || typeof c !== 'object') fail('a rule needs a condition');
  switch (c.type) {
    case 'cellValue': {
      if (!COMPARISONS.has(c.operator)) fail(`unknown cellValue operator "${String(c.operator)}"`);
      const value = normalizeOperand(c.value, 'value');
      if (c.operator === 'between' || c.operator === 'notBetween') {
        return { type: 'cellValue', operator: c.operator, value, value2: normalizeOperand(c.value2, 'value2') };
      }
      return { type: 'cellValue', operator: c.operator, value };
    }
    case 'text':
      if (!TEXT_OPERATORS.has(c.operator)) fail(`unknown text operator "${String(c.operator)}"`);
      if (typeof c.text !== 'string' || c.text === '') fail('a text rule needs the text to look for');
      return { type: 'text', operator: c.operator, text: c.text };
    case 'date':
      if (!DATE_PERIODS.has(c.period)) fail(`unknown date period "${String(c.period)}"`);
      return { type: 'date', period: c.period };
    case 'blanks':
    case 'notBlanks':
    case 'errors':
    case 'notErrors':
    case 'duplicates':
    case 'unique':
      return { type: c.type };
    case 'top':
    case 'bottom': {
      if (!Number.isInteger(c.rank) || c.rank < 1) fail(`${c.type}: rank must be a positive integer`);
      if (c.percent && c.rank > 100) fail(`${c.type}: a percentage can be at most 100`);
      return c.percent ? { type: c.type, rank: c.rank, percent: true } : { type: c.type, rank: c.rank };
    }
    case 'aboveAverage':
    case 'belowAverage':
      return c.orEqual ? { type: c.type, orEqual: true } : { type: c.type };
    case 'formula': {
      if (typeof c.formula !== 'string' || c.formula.trim() === '' || c.formula.trim() === '=') {
        fail('a formula rule needs a formula');
      }
      const formula = c.formula.trim();
      return { type: 'formula', formula: formula.startsWith('=') ? formula : `=${formula}` };
    }
    case 'colorScale': {
      const scale: ConditionalFormatCondition = {
        type: 'colorScale',
        min: normalizeColorStop(c.min, 'colorScale.min', 'min'),
        max: normalizeColorStop(c.max, 'colorScale.max', 'max'),
      };
      if (c.mid) {
        const type = c.mid.type ?? 'percentile';
        scale.mid = normalizeColorStop({ ...c.mid, type, value: c.mid.value ?? 50 }, 'colorScale.mid', 'min');
      }
      return scale;
    }
    case 'dataBar': {
      const bar: ConditionalFormatCondition = { type: 'dataBar' };
      if (c.color !== undefined) {
        if (typeof c.color !== 'string' || c.color === '') fail('dataBar.color must be a CSS color string');
        bar.color = c.color;
      }
      if (c.negativeColor !== undefined) {
        if (typeof c.negativeColor !== 'string' || c.negativeColor === '') {
          fail('dataBar.negativeColor must be a CSS color string');
        }
        bar.negativeColor = c.negativeColor;
      }
      if (c.min !== undefined) bar.min = normalizeThreshold(c.min, 'dataBar.min');
      if (c.max !== undefined) bar.max = normalizeThreshold(c.max, 'dataBar.max');
      if (c.showValue === false) bar.showValue = false;
      return bar;
    }
    case 'iconSet': {
      const labels = ICON_SET_LABELS[c.icons];
      if (!labels) fail(`unknown icon set "${String(c.icons)}"`);
      const icons: ConditionalFormatCondition = { type: 'iconSet', icons: c.icons };
      if (c.thresholds !== undefined) {
        if (!Array.isArray(c.thresholds) || c.thresholds.length !== labels.length - 1) {
          fail(`iconSet.thresholds: ${c.icons} needs ${labels.length - 1} thresholds`);
        }
        icons.thresholds = c.thresholds.map((t, i) => normalizeThreshold(t, `iconSet.thresholds[${i}]`));
      }
      if (c.labels !== undefined) {
        if (
          !Array.isArray(c.labels) ||
          c.labels.length !== labels.length ||
          c.labels.some((l) => typeof l !== 'string' || l.trim() === '')
        ) {
          fail(`iconSet.labels: ${c.icons} needs ${labels.length} non-empty labels`);
        }
        icons.labels = c.labels.map((l) => l.trim());
      }
      if (c.reverse) icons.reverse = true;
      if (c.showValue === false) icons.showValue = false;
      return icons;
    }
    default:
      return fail(`unknown condition type "${String((c as { type?: unknown }).type)}"`);
  }
}

/**
 * Validate a rule and return a clean copy of it with the given id. Throws a
 * `TypeError` describing what is wrong. The range is normalized (start is the
 * top-left corner) but not checked against a grid size.
 */
export function normalizeConditionalRule(input: ConditionalFormatRuleInput, id: string): ConditionalFormatRule {
  if (!input || typeof input !== 'object') fail('a rule must be an object');
  const start = input.range?.start;
  const end = input.range?.end;
  const isIndex = (n: unknown): n is number => Number.isInteger(n) && (n as number) >= 0;
  if (!start || !end || ![start.row, start.col, end.row, end.col].every(isIndex)) {
    fail('range must have non-negative integer start and end coordinates');
  }
  const condition = normalizeCondition(input.condition);
  const rule: ConditionalFormatRule = {
    id,
    range: {
      start: { row: Math.min(start.row, end.row), col: Math.min(start.col, end.col) },
      end: { row: Math.max(start.row, end.row), col: Math.max(start.col, end.col) },
    },
    condition,
  };
  if (!isGraphicCondition(condition)) {
    const format = normalizeStyle(input.format);
    if (format) rule.format = format;
    if (input.stopIfTrue) rule.stopIfTrue = true;
  }
  if (input.label !== undefined && input.label !== null) {
    if (typeof input.label !== 'string') fail('label must be a string');
    if (input.label.trim() !== '') rule.label = input.label.trim();
  }
  return rule;
}

/** A deep copy of a rule (rules are plain data). */
export function cloneConditionalRule(rule: ConditionalFormatRule): ConditionalFormatRule {
  return JSON.parse(JSON.stringify(rule)) as ConditionalFormatRule;
}

/** Whether two rule lists are the same rules in the same order. */
export function conditionalRulesEqual(
  a: readonly ConditionalFormatRule[],
  b: readonly ConditionalFormatRule[]
): boolean {
  if (a === b) return true;
  if (a.length !== b.length) return false;
  return a.every((rule, i) => rule === b[i] || JSON.stringify(rule) === JSON.stringify(b[i]));
}

// ─── Descriptions ───────────────────────────────────────

const COMPARISON_TEXT: Record<string, string> = {
  greaterThan: 'greater than',
  greaterThanOrEqual: 'greater than or equal to',
  lessThan: 'less than',
  lessThanOrEqual: 'less than or equal to',
  equal: 'equal to',
  notEqual: 'not equal to',
  between: 'between',
  notBetween: 'not between',
};

const TEXT_OPERATOR_TEXT: Record<string, string> = {
  contains: 'contains',
  notContains: 'does not contain',
  beginsWith: 'begins with',
  endsWith: 'ends with',
};

/** Date periods as read in a sentence ("in the last 7 days"). */
export const DATE_PERIOD_TEXT: Record<ConditionalDatePeriod, string> = {
  yesterday: 'yesterday',
  today: 'today',
  tomorrow: 'tomorrow',
  last7Days: 'in the last 7 days',
  lastWeek: 'last week',
  thisWeek: 'this week',
  nextWeek: 'next week',
  lastMonth: 'last month',
  thisMonth: 'this month',
  nextMonth: 'next month',
};

/**
 * A short description of a condition, in lower case: "greater than 100",
 * "above average", "color scale". It is the default label of a highlight
 * rule, exposed with the cells that match it.
 */
export function describeCondition(condition: ConditionalFormatCondition): string {
  const c = condition;
  switch (c.type) {
    case 'cellValue': {
      const text = COMPARISON_TEXT[c.operator];
      return c.operator === 'between' || c.operator === 'notBetween'
        ? `${text} ${c.value} and ${c.value2}`
        : `${text} ${c.value}`;
    }
    case 'text':
      return `${TEXT_OPERATOR_TEXT[c.operator]} "${c.text}"`;
    case 'date':
      return `dated ${DATE_PERIOD_TEXT[c.period]}`;
    case 'blanks':
      return 'blank';
    case 'notBlanks':
      return 'not blank';
    case 'errors':
      return 'error';
    case 'notErrors':
      return 'no error';
    case 'duplicates':
      return 'duplicate';
    case 'unique':
      return 'unique';
    case 'top':
    case 'bottom':
      return `${c.type} ${c.rank}${c.percent ? '%' : ''}`;
    case 'aboveAverage':
      return c.orEqual ? 'average or above' : 'above average';
    case 'belowAverage':
      return c.orEqual ? 'average or below' : 'below average';
    case 'formula':
      return `formula ${c.formula}`;
    case 'colorScale':
      return c.mid ? '3-color scale' : '2-color scale';
    case 'dataBar':
      return 'data bar';
    case 'iconSet':
      return `icon set, ${ICON_SET_NAMES[c.icons]}`;
  }
}

// ─── Ranges ─────────────────────────────────────────────

/** A range in A1 notation: `B2:D10`, or `B2` for a single cell. */
export function rangeToA1(range: SelectionRange): string {
  const start = coordToRef(range.start);
  const end = coordToRef(range.end);
  return start === end ? start : `${start}:${end}`;
}

/** Parse `B2:D10` or `B2` (absolute markers allowed); null if it isn't a range. */
export function parseA1Range(text: string): SelectionRange | null {
  const m = /^\s*(\$?[A-Za-z]+\$?\d+)\s*(?::\s*(\$?[A-Za-z]+\$?\d+))?\s*$/.exec(text);
  if (!m) return null;
  try {
    const a = refToCoord(m[1]);
    const b = m[2] ? refToCoord(m[2]) : a;
    if (a.row < 0 || b.row < 0) return null;
    return {
      start: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
      end: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
    };
  } catch {
    return null;
  }
}

function inRange(range: SelectionRange, row: number, col: number): boolean {
  return row >= range.start.row && row <= range.end.row && col >= range.start.col && col <= range.end.col;
}

/**
 * The rules after rows or columns are inserted or deleted. A rule's range
 * moves with its cells like a range reference and grows when rows/columns are
 * inserted inside it; a rule whose cells are all deleted is dropped. Formulas
 * (formula rules and `=` operands) are rewritten like cell formulas; pass the
 * name of the sheet the rules are on so sheet-qualified references to it move too. Returns
 * the same array when nothing changes.
 */
export function adjustRulesForStructureChange(
  rules: readonly ConditionalFormatRule[],
  change: StructureChange,
  sheet?: string
): readonly ConditionalFormatRule[] {
  const { axis, kind, index, count } = change;
  let changed = false;
  const out: ConditionalFormatRule[] = [];
  for (const rule of rules) {
    const start = axis === 'row' ? rule.range.start.row : rule.range.start.col;
    const end = axis === 'row' ? rule.range.end.row : rule.range.end.col;
    let first = start;
    let newStart = start;
    let newEnd = end;
    if (kind === 'insert') {
      if (index <= start) newStart = start + count;
      if (index <= end) newEnd = end + count;
    } else {
      const delEnd = index + count - 1;
      // The first cell of the range that survives, in old coordinates
      if (start >= index && start <= delEnd) first = delEnd + 1;
      if (first > end) {
        changed = true;
        continue;
      }
      newStart = first > delEnd ? first - count : first;
      newEnd = end > delEnd ? end - count : Math.min(end, index - 1);
    }
    const rewrite = (formula: string): string => {
      // The formula is written for the range's first cell: when that cell is
      // deleted, re-anchor it on the first one that remains.
      const offset = first - start;
      const anchored =
        offset === 0 ? formula : shiftFormulaReferences(formula, axis === 'row' ? offset : 0, axis === 'col' ? offset : 0);
      // With the sheet's name, references that spell it out (Sheet1!A1) move too
      return adjustFormulaForStructureChange(anchored, change, sheet === undefined ? {} : { formulaSheet: sheet });
    };
    let condition = rule.condition;
    if (condition.type === 'formula') {
      const formula = rewrite(condition.formula);
      if (formula !== condition.formula) condition = { ...condition, formula };
    } else if (condition.type === 'cellValue') {
      const fix = (v: ConditionalOperand | undefined) => (typeof v === 'string' && v.startsWith('=') ? rewrite(v) : v);
      const value = fix(condition.value)!;
      const value2 = fix(condition.value2);
      if (value !== condition.value || value2 !== condition.value2) {
        condition = { ...condition, value };
        if (value2 !== undefined) condition.value2 = value2;
      }
    }
    if (newStart === start && newEnd === end && condition === rule.condition) {
      out.push(rule);
      continue;
    }
    changed = true;
    const range: SelectionRange =
      axis === 'row'
        ? { start: { row: newStart, col: rule.range.start.col }, end: { row: newEnd, col: rule.range.end.col } }
        : { start: { row: rule.range.start.row, col: newStart }, end: { row: rule.range.end.row, col: newEnd } };
    out.push({ ...rule, range, condition });
  }
  return changed ? out : rules;
}

/**
 * The rules after clearing conditional formatting from `range`: rules inside
 * it are dropped, and rules it cuts across a whole edge are trimmed. A rule
 * that would not stay a rectangle is kept whole. Returns the same array when
 * nothing changes.
 */
export function clearRulesInRange(
  rules: readonly ConditionalFormatRule[],
  range: SelectionRange
): { rules: readonly ConditionalFormatRule[]; removed: number; trimmed: number; kept: number } {
  let removed = 0;
  let trimmed = 0;
  let kept = 0;
  const out: ConditionalFormatRule[] = [];
  for (const rule of rules) {
    const r = rule.range;
    const overlaps =
      r.start.row <= range.end.row &&
      r.end.row >= range.start.row &&
      r.start.col <= range.end.col &&
      r.end.col >= range.start.col;
    if (!overlaps) {
      out.push(rule);
      continue;
    }
    const coversRows = range.start.row <= r.start.row && range.end.row >= r.end.row;
    const coversCols = range.start.col <= r.start.col && range.end.col >= r.end.col;
    if (coversRows && coversCols) {
      removed++;
      continue;
    }
    let next: SelectionRange | null = null;
    if (coversCols) {
      // Whole rows of the rule are cleared: it keeps a rectangle if they are at its top or bottom
      if (range.start.row <= r.start.row) next = { start: { row: range.end.row + 1, col: r.start.col }, end: r.end };
      else if (range.end.row >= r.end.row) next = { start: r.start, end: { row: range.start.row - 1, col: r.end.col } };
    } else if (coversRows) {
      if (range.start.col <= r.start.col) next = { start: { row: r.start.row, col: range.end.col + 1 }, end: r.end };
      else if (range.end.col >= r.end.col) next = { start: r.start, end: { row: r.end.row, col: range.start.col - 1 } };
    }
    if (!next) {
      kept++;
      out.push(rule);
      continue;
    }
    trimmed++;
    let condition = rule.condition;
    const dRow = next.start.row - r.start.row;
    const dCol = next.start.col - r.start.col;
    if (dRow !== 0 || dCol !== 0) {
      // Keep formulas pointing at the same cells from the new first cell
      const shift = (f: string) => shiftFormulaReferences(f, dRow, dCol);
      if (condition.type === 'formula') {
        condition = { ...condition, formula: shift(condition.formula) };
      } else if (condition.type === 'cellValue') {
        const fix = (v: ConditionalOperand | undefined) => (typeof v === 'string' && v.startsWith('=') ? shift(v) : v);
        condition = { ...condition, value: fix(condition.value)! };
        if (condition.value2 !== undefined) condition.value2 = fix(condition.value2);
      }
    }
    out.push({
      ...rule,
      range: { start: { ...next.start }, end: { ...next.end } },
      condition,
    });
  }
  return { rules: removed + trimmed > 0 ? out : rules, removed, trimmed, kept };
}

// ─── Evaluation ─────────────────────────────────────────

interface RuleStats {
  /** The numbers in the range, ascending */
  sorted?: Float64Array;
  sum?: number;
  /** How often each value occurs in the range */
  counts?: Map<string, number>;
}

function clamp01(n: number): number {
  return n < 0 ? 0 : n > 1 ? 1 : n;
}

function textOf(value: SortValue): string {
  switch (value.kind) {
    case 'number':
      return String(value.value);
    case 'text':
      return value.value;
    case 'boolean':
      return value.value ? 'TRUE' : 'FALSE';
    default:
      return '';
  }
}

function isBlank(value: SortValue): boolean {
  return value.kind === 'blank' || (value.kind === 'text' && value.value.trim() === '');
}

/** Key under which equal values are counted for duplicate rules; null for values that never count. */
function duplicateKey(value: SortValue): string | null {
  switch (value.kind) {
    case 'number':
      return `n:${value.value}`;
    case 'text':
      return value.value.trim() === '' ? null : `t:${value.value.toLowerCase()}`;
    case 'boolean':
      return `b:${value.value}`;
    default:
      return null;
  }
}

/** A literal operand as a value: `'5'` is a number, `'TRUE'` a boolean, anything else text. */
function literalOperand(operand: ConditionalOperand): SortValue {
  if (typeof operand === 'number') return { kind: 'number', value: operand };
  const trimmed = operand.trim();
  const num = Number(trimmed);
  if (trimmed !== '' && Number.isFinite(num)) return { kind: 'number', value: num };
  const upper = trimmed.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') return { kind: 'boolean', value: upper === 'TRUE' };
  return { kind: 'text', value: operand };
}

/** Excel's PERCENTILE.INC over an ascending array. */
function percentile(sorted: Float64Array, p: number): number {
  const pos = (sorted.length - 1) * p;
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
}

function resolveThreshold(threshold: ConditionalThreshold, sorted: Float64Array): number {
  const min = sorted[0];
  const max = sorted[sorted.length - 1];
  switch (threshold.type) {
    case 'min':
      return min;
    case 'max':
      return max;
    case 'number':
      return threshold.value ?? 0;
    case 'percent':
      return min + ((max - min) * (threshold.value ?? 50)) / 100;
    case 'percentile':
      return percentile(sorted, (threshold.value ?? 50) / 100);
  }
}

function mixColors(a: string, b: string, t: number): [number, number, number] {
  const ca = parseHexColor(a) ?? [255, 255, 255];
  const cb = parseHexColor(b) ?? [255, 255, 255];
  return [0, 1, 2].map((i) => Math.round(ca[i] + (cb[i] - ca[i]) * t)) as [number, number, number];
}

function toHex(rgb: [number, number, number]): string {
  return `#${rgb.map((n) => n.toString(16).padStart(2, '0')).join('')}`;
}

/** Black or white, whichever contrasts more with the given color (WCAG relative luminance). */
export function readableTextColor(rgb: [number, number, number]): string {
  const [r, g, b] = rgb.map((v) => {
    const s = v / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  });
  const luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  // Contrast with black is (L + 0.05) / 0.05, with white 1.05 / (L + 0.05)
  return luminance > 0.179 ? '#000000' : '#ffffff';
}

/** First and last day (date serials) of a period relative to `today`. */
function datePeriodBounds(period: ConditionalDatePeriod, today: number): [number, number] {
  // Serial 0 (1899-12-30) was a Saturday; weeks start on Sunday
  const weekStart = today - ((today + 6) % 7);
  const d = dateSerialToDate(today);
  const month = (offset: number): [number, number] => [
    dateToSerial(d.getUTCFullYear(), d.getUTCMonth() + offset, 1),
    dateToSerial(d.getUTCFullYear(), d.getUTCMonth() + offset + 1, 1) - 1,
  ];
  switch (period) {
    case 'yesterday':
      return [today - 1, today - 1];
    case 'today':
      return [today, today];
    case 'tomorrow':
      return [today + 1, today + 1];
    case 'last7Days':
      return [today - 6, today];
    case 'lastWeek':
      return [weekStart - 7, weekStart - 1];
    case 'thisWeek':
      return [weekStart, weekStart + 6];
    case 'nextWeek':
      return [weekStart + 7, weekStart + 13];
    case 'lastMonth':
      return month(-1);
    case 'thisMonth':
      return month(0);
    case 'nextMonth':
      return month(1);
  }
}

/**
 * Works out what the rules give each cell. Results are cached per cell and
 * dropped on every change; the statistics of a rule's range (average,
 * top/bottom cut-off, duplicate counts, scale bounds) are cached per rule and
 * dropped only when a cell inside that range changes, since they cost a pass
 * over the whole range.
 */
export class ConditionalFormatEvaluator {
  private _rules: readonly ConditionalFormatRule[] = [];
  private _stats = new Map<string, RuleStats>();
  private _cells = new Map<string, ConditionalFormatResult | null>();
  private _today = NaN;

  /** Most cell results kept; scrolling a large sheet would otherwise grow the cache without bound. */
  private static readonly MAX_CACHED_CELLS = 20_000;

  constructor(private readonly _source: ConditionalFormatSource) {}

  get rules(): readonly ConditionalFormatRule[] {
    return this._rules;
  }

  setRules(rules: readonly ConditionalFormatRule[]): void {
    this._rules = rules;
    this._stats.clear();
    this._cells.clear();
  }

  /**
   * Forget results that depend on the given cells (keys in `"row:col"`
   * format), or everything when called without keys.
   */
  invalidate(keys?: Iterable<string>): void {
    this._cells.clear();
    if (!keys) {
      this._stats.clear();
      return;
    }
    if (this._stats.size === 0) return;
    const cached = this._rules.filter((rule) => this._stats.has(rule.id));
    for (const key of keys) {
      const { row, col } = parseKey(key);
      for (let i = cached.length - 1; i >= 0; i--) {
        if (inRange(cached[i].range, row, col)) {
          this._stats.delete(cached[i].id);
          cached.splice(i, 1);
        }
      }
      if (cached.length === 0) return;
    }
  }

  /** What the rules give a cell, or null when none applies to it. */
  resultFor(row: number, col: number): ConditionalFormatResult | null {
    if (this._rules.length === 0) return null;
    const today = this._source.today();
    if (today !== this._today) {
      // Date rules compare with today
      this._today = today;
      this._cells.clear();
    }
    const key = cellKey(row, col);
    let result = this._cells.get(key);
    if (result === undefined) {
      if (this._cells.size >= ConditionalFormatEvaluator.MAX_CACHED_CELLS) this._cells.clear();
      result = this._compute(row, col);
      this._cells.set(key, result);
    }
    return result;
  }

  private _compute(row: number, col: number): ConditionalFormatResult | null {
    let value: SortValue | undefined;
    const r: ConditionalFormatResult = { format: {}, labels: [], ruleIds: [] };

    for (const rule of this._rules) {
      if (!inRange(rule.range, row, col)) continue;
      value ??= this._source.valueAt(row, col);
      const c = rule.condition;
      if (c.type === 'colorScale') {
        if (r.scale !== undefined) continue;
        const scale = this._colorScale(rule, c, value);
        if (!scale) continue;
        r.scale = scale.position;
        if (r.format.backgroundColor === undefined) {
          r.format.backgroundColor = toHex(scale.color);
          // Keep the value readable on whatever color the scale lands on
          if (r.format.textColor === undefined) r.format.textColor = readableTextColor(scale.color);
        }
        r.labels.push(`${rule.label ?? 'color scale'} ${Math.round(scale.position * 100)}%`);
        r.ruleIds.push(rule.id);
      } else if (c.type === 'dataBar') {
        if (r.dataBar) continue;
        const bar = this._dataBar(rule, c, value);
        if (!bar) continue;
        r.dataBar = bar.bar;
        r.labels.push(`${rule.label ?? 'data bar'} ${Math.round(bar.position * 100)}%`);
        r.ruleIds.push(rule.id);
      } else if (c.type === 'iconSet') {
        if (r.icon) continue;
        const icon = this._icon(rule, c, value);
        if (!icon) continue;
        r.icon = icon;
        r.labels.push(rule.label ? `${rule.label}: ${icon.label}` : icon.label);
        r.ruleIds.push(rule.id);
      } else {
        if (!this._matches(rule, c, value, row, col)) continue;
        // Earlier rules win where formats conflict
        r.format = { ...rule.format, ...r.format };
        r.labels.push(rule.label ?? describeCondition(c));
        r.ruleIds.push(rule.id);
        if (rule.stopIfTrue) break;
      }
    }
    return r.ruleIds.length > 0 ? r : null;
  }

  private _operand(rule: ConditionalFormatRule, operand: ConditionalOperand, row: number, col: number): SortValue {
    if (typeof operand === 'string' && operand.startsWith('=')) {
      return this._source.evaluate(
        shiftFormulaReferences(operand, row - rule.range.start.row, col - rule.range.start.col),
        { row, col }
      );
    }
    return literalOperand(operand);
  }

  private _matches(
    rule: ConditionalFormatRule,
    c: ConditionalFormatCondition,
    value: SortValue,
    row: number,
    col: number
  ): boolean {
    switch (c.type) {
      case 'cellValue': {
        if (isBlank(value) || value.kind === 'error') return false;
        const a = this._operand(rule, c.value, row, col);
        if (a.kind === 'error' || a.kind === 'blank') return false;
        if (c.operator === 'between' || c.operator === 'notBetween') {
          const b = this._operand(rule, c.value2 as ConditionalOperand, row, col);
          if (b.kind !== a.kind) return false;
          if (value.kind !== a.kind) return c.operator === 'notBetween';
          const [lo, hi] = compareSortValues(a, b) <= 0 ? [a, b] : [b, a];
          const inside = compareSortValues(value, lo) >= 0 && compareSortValues(value, hi) <= 0;
          return c.operator === 'between' ? inside : !inside;
        }
        // Values of different kinds (a number and text) are never equal and have no order
        if (value.kind !== a.kind) return c.operator === 'notEqual';
        const cmp = compareSortValues(value, a);
        switch (c.operator) {
          case 'greaterThan':
            return cmp > 0;
          case 'greaterThanOrEqual':
            return cmp >= 0;
          case 'lessThan':
            return cmp < 0;
          case 'lessThanOrEqual':
            return cmp <= 0;
          case 'equal':
            return cmp === 0;
          default:
            return cmp !== 0;
        }
      }
      case 'text': {
        if (isBlank(value) || value.kind === 'error') return false;
        const hay = textOf(value).toLowerCase();
        const needle = c.text.toLowerCase();
        switch (c.operator) {
          case 'contains':
            return hay.includes(needle);
          case 'notContains':
            return !hay.includes(needle);
          case 'beginsWith':
            return hay.startsWith(needle);
          default:
            return hay.endsWith(needle);
        }
      }
      case 'date': {
        if (value.kind !== 'number') return false;
        const day = Math.floor(value.value);
        const [from, to] = datePeriodBounds(c.period, this._today);
        return day >= from && day <= to;
      }
      case 'blanks':
        return isBlank(value);
      case 'notBlanks':
        return !isBlank(value);
      case 'errors':
        return value.kind === 'error';
      case 'notErrors':
        return value.kind !== 'error' && !isBlank(value);
      case 'duplicates':
      case 'unique': {
        const key = duplicateKey(value);
        if (key === null) return false;
        const count = this._countsFor(rule).get(key) ?? 0;
        return c.type === 'duplicates' ? count > 1 : count === 1;
      }
      case 'top':
      case 'bottom': {
        if (value.kind !== 'number') return false;
        const { sorted } = this._numbersFor(rule);
        const n = sorted.length;
        if (n === 0) return false;
        const count = c.percent ? Math.max(1, Math.floor((n * c.rank) / 100)) : Math.min(c.rank, n);
        // Ties with the cut-off value are included, as in Excel
        return c.type === 'top' ? value.value >= sorted[n - count] : value.value <= sorted[count - 1];
      }
      case 'aboveAverage':
      case 'belowAverage': {
        if (value.kind !== 'number') return false;
        const { sorted, sum } = this._numbersFor(rule);
        if (sorted.length === 0) return false;
        const mean = sum / sorted.length;
        if (value.value === mean) return !!c.orEqual;
        return c.type === 'aboveAverage' ? value.value > mean : value.value < mean;
      }
      case 'formula': {
        const result = this._source.evaluate(
          shiftFormulaReferences(c.formula, row - rule.range.start.row, col - rule.range.start.col),
          { row, col }
        );
        return result.kind === 'boolean' ? result.value : result.kind === 'number' && result.value !== 0;
      }
      default:
        return false;
    }
  }

  private _colorScale(
    rule: ConditionalFormatRule,
    c: Extract<ConditionalFormatCondition, { type: 'colorScale' }>,
    value: SortValue
  ): { color: [number, number, number]; position: number } | null {
    if (value.kind !== 'number') return null;
    const { sorted } = this._numbersFor(rule);
    if (sorted.length === 0) return null;
    const v = value.value;
    const lo = resolveThreshold(c.min, sorted);
    const hi = resolveThreshold(c.max, sorted);
    if (!(hi > lo)) return { color: mixColors(c.min.color, c.max.color, 0.5), position: 0.5 };
    if (!c.mid) {
      const t = clamp01((v - lo) / (hi - lo));
      return { color: mixColors(c.min.color, c.max.color, t), position: t };
    }
    const mid = Math.min(hi, Math.max(lo, resolveThreshold(c.mid, sorted)));
    if (v <= mid) {
      const t = mid > lo ? clamp01((v - lo) / (mid - lo)) : 1;
      return { color: mixColors(c.min.color, c.mid.color, t), position: t / 2 };
    }
    const t = hi > mid ? clamp01((v - mid) / (hi - mid)) : 1;
    return { color: mixColors(c.mid.color, c.max.color, t), position: 0.5 + t / 2 };
  }

  private _dataBar(
    rule: ConditionalFormatRule,
    c: Extract<ConditionalFormatCondition, { type: 'dataBar' }>,
    value: SortValue
  ): { bar: NonNullable<ConditionalFormatResult['dataBar']>; position: number } | null {
    if (value.kind !== 'number') return null;
    const { sorted } = this._numbersFor(rule);
    if (sorted.length === 0) return null;
    const v = value.value;
    // Automatic bounds include zero, so bars grow from a zero axis
    const lo = c.min ? resolveThreshold(c.min, sorted) : Math.min(0, sorted[0]);
    const hi = c.max ? resolveThreshold(c.max, sorted) : Math.max(0, sorted[sorted.length - 1]);
    const at = (n: number) => (hi > lo ? clamp01((n - lo) / (hi - lo)) : n >= hi && n !== 0 ? 1 : 0);
    const position = at(v);
    const zero = hi > lo ? at(0) : 0;
    const negative = v < 0;
    const bar: NonNullable<ConditionalFormatResult['dataBar']> = {
      start: Math.min(zero, position),
      size: Math.abs(position - zero),
      negative,
      showValue: c.showValue !== false,
    };
    const color = negative ? c.negativeColor : c.color;
    if (color) bar.color = color;
    return { bar, position };
  }

  private _icon(
    rule: ConditionalFormatRule,
    c: Extract<ConditionalFormatCondition, { type: 'iconSet' }>,
    value: SortValue
  ): NonNullable<ConditionalFormatResult['icon']> | null {
    if (value.kind !== 'number') return null;
    const { sorted } = this._numbersFor(rule);
    if (sorted.length === 0) return null;
    const labels = c.labels ?? ICON_SET_LABELS[c.icons];
    const count = labels.length;
    const thresholds: ConditionalThreshold[] =
      c.thresholds ??
      Array.from({ length: count - 1 }, (_, i) => ({ type: 'percent' as const, value: ((i + 1) * 100) / count }));
    let index = 0;
    for (const t of thresholds) {
      if (value.value >= resolveThreshold(t, sorted)) index++;
    }
    if (c.reverse) index = count - 1 - index;
    return { set: c.icons, index, label: labels[index], showValue: c.showValue !== false };
  }

  private _numbersFor(rule: ConditionalFormatRule): { sorted: Float64Array; sum: number } {
    let stats = this._stats.get(rule.id);
    if (!stats) this._stats.set(rule.id, (stats = {}));
    if (!stats.sorted) {
      const numbers: number[] = [];
      let sum = 0;
      this._source.forEachValue(rule.range, (value) => {
        if (value.kind === 'number') {
          numbers.push(value.value);
          sum += value.value;
        }
      });
      stats.sorted = Float64Array.from(numbers).sort();
      stats.sum = sum;
    }
    return { sorted: stats.sorted, sum: stats.sum ?? 0 };
  }

  private _countsFor(rule: ConditionalFormatRule): Map<string, number> {
    let stats = this._stats.get(rule.id);
    if (!stats) this._stats.set(rule.id, (stats = {}));
    if (!stats.counts) {
      const counts = new Map<string, number>();
      this._source.forEachValue(rule.range, (value) => {
        const key = duplicateKey(value);
        if (key !== null) counts.set(key, (counts.get(key) ?? 0) + 1);
      });
      stats.counts = counts;
    }
    return stats.counts;
  }
}
