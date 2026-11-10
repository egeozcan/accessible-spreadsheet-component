/**
 * Built-in formula functions that only need their argument values. Functions
 * that work on references, evaluate lazily or call lambdas live in
 * `formula-engine.ts`.
 */
import { type FormulaContext, dateSerialToDate, dateToSerial } from '../types.js';
import {
  RangeValue,
  flattenDeep,
  isErrorCode,
  scalarsEqual,
  toBoolean,
  toNumber,
  toText,
} from './values.js';

/** A captured argument error; the function rethrows it only if it uses the argument. */
export class ArgumentError {
  constructor(readonly code: string) {}
}

/**
 * Registers built-ins. `kinds` has one letter per argument, the last one
 * repeating (or the group after `|` repeating):
 * `n` single value, blank as 0 · `t` single value, blank as "" ·
 * `v` single value, blank as undefined and errors as their code ·
 * `a` array or value as is (a blank cell as 0) · `b` like `a`, but a blank
 * cell stays undefined · `r` reference as is · `x` like `a`, but an
 * error in the argument is captured as an `ArgumentError`.
 * Arrays given for `n`/`t`/`v` arguments call the function per element.
 */
export interface Registrar {
  def(
    name: string,
    kinds: string,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    fn: (ctx: FormulaContext, ...args: any[]) => unknown,
    opts?: { arity?: [number, number]; volatile?: boolean }
  ): void;
  /**
   * An aggregate: receives every value of its arguments in one list (arrays
   * and nested arrays flattened, blanks dropped). `propagate` makes an error
   * anywhere in the arguments the result.
   */
  agg(name: string, fn: (values: unknown[]) => unknown, opts?: { propagate?: boolean }): void;
}

/**
 * Ordering used by approximate-match lookups: numbers compare numerically,
 * text compares case-insensitively, and values of different kinds are
 * incomparable (NaN), so they never match.
 */
function lookupCompare(a: unknown, b: unknown): number {
  const isNum = (v: unknown) =>
    typeof v === 'number' || (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v)));
  if (a === undefined || b === undefined) return NaN;
  if (isNum(a) && isNum(b)) return Number(a) - Number(b);
  if (isNum(a) || isNum(b) || typeof a === 'boolean' || typeof b === 'boolean') return NaN;
  const as = String(a).toLowerCase();
  const bs = String(b).toLowerCase();
  return as < bs ? -1 : as > bs ? 1 : 0;
}

/**
 * Index of the last value <= `lookup` in an ascending list, stopping at the
 * first larger value; values of a different kind are skipped. -1 if none.
 */
function approximateMatch(values: unknown[], lookup: unknown): number {
  let match = -1;
  for (let i = 0; i < values.length; i++) {
    const cmp = lookupCompare(values[i], lookup);
    if (isNaN(cmp)) continue;
    if (cmp > 0) break;
    match = i;
  }
  return match;
}

/**
 * Helper to check if a value matches a criteria string.
 * Supports operator prefixes: <>, >=, <=, >, <, =
 * Without operator prefix, does exact match (case-insensitive for strings).
 */
function matchesCriteria(value: unknown, criteria: string): boolean {
  let op = '=';
  let target = criteria;

  if (criteria.startsWith('<>')) {
    op = '<>';
    target = criteria.substring(2);
  } else if (criteria.startsWith('>=')) {
    op = '>=';
    target = criteria.substring(2);
  } else if (criteria.startsWith('<=')) {
    op = '<=';
    target = criteria.substring(2);
  } else if (criteria.startsWith('>')) {
    op = '>';
    target = criteria.substring(1);
  } else if (criteria.startsWith('<')) {
    op = '<';
    target = criteria.substring(1);
  } else if (criteria.startsWith('=')) {
    op = '=';
    target = criteria.substring(1);
  }

  const numTarget = Number(target);
  const numValue = Number(value);
  const targetNumeric = !isNaN(numTarget) && target.trim() !== '';
  const valueNumeric = !isNaN(numValue) && String(value).trim() !== '';
  const bothNumeric = targetNumeric && valueNumeric;

  // Relational operators only compare like with like (Excel: ">5" never matches text)
  if (op !== '=' && op !== '<>' && targetNumeric !== valueNumeric) {
    return false;
  }

  const valLower = String(value).toLowerCase();
  const targetLower = target.toLowerCase();

  // Text equality supports Excel wildcards: * (any run), ? (one char), ~ escapes.
  const wildcard = (op === '=' || op === '<>') && !bothNumeric ? compileWildcard(target) : null;

  // Wildcards only ever match text (Excel: COUNTIF(range, "*") counts text
  // cells) — not numbers, booleans, or error values carried as strings.
  const wildcardMatch = () =>
    typeof value === 'string' && !isErrorCode(value) && wildcardMatches(wildcard!, value);

  switch (op) {
    case '=':
      if (bothNumeric) return numValue === numTarget;
      if (wildcard) return wildcardMatch();
      return valLower === targetLower;
    case '<>':
      if (bothNumeric) return numValue !== numTarget;
      if (wildcard) return !wildcardMatch();
      return valLower !== targetLower;
    case '>':
      return bothNumeric ? numValue > numTarget : valLower > targetLower;
    case '<':
      return bothNumeric ? numValue < numTarget : valLower < targetLower;
    case '>=':
      return bothNumeric ? numValue >= numTarget : valLower >= targetLower;
    case '<=':
      return bothNumeric ? numValue <= numTarget : valLower <= targetLower;
    default:
      return false;
  }
}

type WildcardToken = { kind: 'lit'; ch: string } | { kind: 'any' } | { kind: 'star' };

/**
 * Compile an Excel wildcard pattern (`*` any run, `?` one character, `~`
 * escapes the next `*`, `?` or `~`) into tokens for `wildcardMatches`.
 * Returns null when the pattern has no wildcard syntax at all.
 */
function compileWildcard(pattern: string): WildcardToken[] | null {
  if (!/[*?~]/.test(pattern)) return null;
  const tokens: WildcardToken[] = [];
  // Split into code points before lowercasing each one, so characters whose
  // lowercase form is longer (e.g. "İ") still count as one for `?`.
  const chars = [...pattern];
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i];
    if (ch === '~' && i + 1 < chars.length && '*?~'.includes(chars[i + 1])) {
      tokens.push({ kind: 'lit', ch: chars[++i].toLowerCase() });
    } else if (ch === '*') {
      // Consecutive stars are equivalent to one
      if (tokens[tokens.length - 1]?.kind !== 'star') tokens.push({ kind: 'star' });
    } else if (ch === '?') {
      tokens.push({ kind: 'any' });
    } else {
      tokens.push({ kind: 'lit', ch: ch.toLowerCase() });
    }
  }
  return tokens;
}

/**
 * Case-insensitive whole-string wildcard match. Greedy matching that
 * backtracks only to the most recent star, so the worst case is
 * O(pattern × text) — unlike a RegExp, which can backtrack exponentially
 * on patterns such as "*a*a*a*b".
 */
function wildcardMatches(tokens: WildcardToken[], value: string): boolean {
  const text = [...value];
  let p = 0;
  let t = 0;
  let starP = -1;
  let starT = 0;
  while (t < text.length) {
    const tok = tokens[p];
    if (tok && (tok.kind === 'any' || (tok.kind === 'lit' && tok.ch === text[t].toLowerCase()))) {
      p++;
      t++;
    } else if (tok && tok.kind === 'star') {
      starP = p++;
      starT = t;
    } else if (starP !== -1) {
      p = starP + 1;
      t = ++starT;
    } else {
      return false;
    }
  }
  while (tokens[p]?.kind === 'star') p++;
  return p === tokens.length;
}

/** Criteria match that also handles empty cells (which only match "" or "="). */
function cellMatchesCriteria(value: unknown, criteria: string): boolean {
  if (value === undefined) return criteria === '' || criteria === '=';
  return matchesCriteria(value, criteria);
}

/** [rows, cols] of a range argument; arrays are a column, scalars 1×1. */
function rangeShape(arg: unknown): [number, number] {
  if (arg instanceof RangeValue) return [arg.rows, arg.cols];
  if (Array.isArray(arg)) return [arg.length, 1];
  return [1, 1];
}

/** Flat list of values from a range, array, or single value. */
function toValueList(arg: unknown): unknown[] {
  if (arg instanceof RangeValue) return arg.values;
  if (Array.isArray(arg)) return arg;
  return [arg];
}

/** Any value as a `RangeValue` (a scalar becomes 1×1). */
export function asArray(arg: unknown): RangeValue {
  if (arg instanceof RangeValue) return arg;
  if (Array.isArray(arg)) return RangeValue.column(arg);
  return new RangeValue([arg], 1, 1);
}

/**
 * Evaluate `(criteria_range, criteria)` pairs as used by SUMIFS/COUNTIFS/
 * AVERAGEIFS. Returns a mask of positions that satisfy every pair. Every
 * criteria range must have the same shape as `shapeOf`, otherwise #VALUE!.
 */
function matchAllCriteria(pairs: unknown[], shapeOf: unknown): boolean[] {
  if (pairs.length === 0 || pairs.length % 2 !== 0) throw new Error('#VALUE!');
  const [rows, cols] = rangeShape(shapeOf);
  const size = rows * cols;
  const mask = new Array<boolean>(size).fill(true);
  for (let p = 0; p < pairs.length; p += 2) {
    const values = toValueList(pairs[p]);
    const [r, c] = rangeShape(pairs[p]);
    if (r !== rows || c !== cols) throw new Error('#VALUE!');
    const criteria = String(pairs[p + 1]);
    for (let i = 0; i < size; i++) {
      if (mask[i] && !cellMatchesCriteria(values[i], criteria)) mask[i] = false;
    }
  }
  return mask;
}

/** Numbers (or numeric strings) among flattened aggregate arguments; text and booleans are skipped. */
function numericArgs(args: unknown[]): number[] {
  const nums: number[] = [];
  for (const v of args) {
    if (typeof v === 'number') nums.push(v);
    else if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) nums.push(Number(v));
  }
  return nums;
}

/**
 * Round `n` to `digits` decimal places with the given rounding function,
 * applied to the magnitude (so ceil rounds away from zero, floor toward it).
 */
function roundMagnitude(n: number, digits: number, round: (x: number) => number): number {
  const factor = Math.pow(10, digits);
  // Neutralise IEEE 754 drift (e.g. 1.005*100 → 100.4999…) before rounding.
  const shifted = parseFloat((Math.abs(n) * factor).toPrecision(15));
  return (Math.sign(n) * round(shifted)) / factor;
}

/** Date serial number argument → Date (UTC fields), rejecting negatives. */
function serialArg(v: unknown): Date {
  const serial = toNumber(v);
  if (serial < 0) throw new Error('#NUM!');
  return dateSerialToDate(Math.floor(serial));
}

/** Exact match as the classic lookups do it: numbers by value, everything else as text ignoring case. */
function looseEquals(cell: unknown, lookup: unknown): boolean {
  const numLookup = Number(lookup);
  const numCell = Number(cell);
  const bothNum =
    !isNaN(numLookup) && !isNaN(numCell) && String(lookup).trim() !== '' && String(cell).trim() !== '';
  return bothNum ? numLookup === numCell : String(cell).toLowerCase() === String(lookup).toLowerCase();
}

/**
 * Sort order of values in SORT / SORTBY / XMATCH: numbers, then text
 * (ignoring case), then FALSE, TRUE, then errors; blanks always last.
 */
function sortRank(v: unknown): number {
  if (v === undefined || v === null || v === '') return 4;
  if (typeof v === 'number') return 0;
  if (typeof v === 'boolean') return 2;
  if (isErrorCode(v)) return 3;
  return 1;
}

function compareForSort(a: unknown, b: unknown, direction: number): number {
  const ra = sortRank(a);
  const rb = sortRank(b);
  // Blanks stay last in either direction
  if (ra === 4 || rb === 4) return ra - rb;
  if (ra !== rb) return (ra - rb) * direction;
  if (ra === 0) return ((a as number) - (b as number)) * direction;
  if (ra === 2) return (Number(a) - Number(b)) * direction;
  const as = String(a).toLowerCase();
  const bs = String(b).toLowerCase();
  return (as < bs ? -1 : as > bs ? 1 : 0) * direction;
}

/** A whole number argument; text that is not a number is #VALUE!. */
function intArg(v: unknown, fallback?: number): number {
  if (v === undefined) {
    if (fallback === undefined) throw new Error('#VALUE!');
    return fallback;
  }
  return Math.trunc(toNumber(v));
}

/** Key that is equal for values UNIQUE and HAS treat as the same. */
function identityKey(v: unknown): string {
  if (v instanceof RangeValue) return 'a:' + v.rows + 'x' + v.cols + ':' + v.values.map(identityKey).join('\u0001');
  if (v === undefined || v === null) return 'b:';
  if (typeof v === 'number') return 'n:' + v;
  if (typeof v === 'boolean') return 'l:' + v;
  return 's:' + String(v).toLowerCase();
}

/** Pick rows (or columns) of an array by index. */
function pickRows(arr: RangeValue, indices: number[]): RangeValue {
  const values: unknown[] = [];
  for (const r of indices) for (let c = 0; c < arr.cols; c++) values.push(arr.get(r, c));
  return new RangeValue(values, indices.length, arr.cols);
}

function pickCols(arr: RangeValue, indices: number[]): RangeValue {
  const values: unknown[] = [];
  for (let r = 0; r < arr.rows; r++) for (const c of indices) values.push(arr.get(r, c));
  return new RangeValue(values, arr.rows, indices.length);
}

export function transpose(arr: RangeValue): RangeValue {
  const values: unknown[] = [];
  for (let c = 0; c < arr.cols; c++) for (let r = 0; r < arr.rows; r++) values.push(arr.get(r, c));
  return new RangeValue(values, arr.cols, arr.rows);
}

/** 1-based positions (negative counts from the end) → 0-based indices; out of range is #VALUE!. */
function positions(args: unknown[], size: number): number[] {
  const out: number[] = [];
  for (const arg of args) {
    for (const v of flattenDeep(arg)) {
      const n = intArg(v);
      const i = n > 0 ? n - 1 : size + n;
      if (n === 0 || i < 0 || i >= size) throw new Error('#VALUE!');
      out.push(i);
    }
  }
  return out;
}

/**
 * Remove one level of nesting: every element that is an array is laid out
 * in place, as a block. Rows and columns of blocks are padded to a common size.
 */
function flattenOnce(arr: RangeValue, pad: unknown): RangeValue {
  const rowHeights = new Array<number>(arr.rows).fill(1);
  const colWidths = new Array<number>(arr.cols).fill(1);
  for (let r = 0; r < arr.rows; r++) {
    for (let c = 0; c < arr.cols; c++) {
      const v = arr.get(r, c);
      if (v instanceof RangeValue) {
        rowHeights[r] = Math.max(rowHeights[r], v.rows);
        colWidths[c] = Math.max(colWidths[c], v.cols);
      }
    }
  }
  const rowStart: number[] = [];
  const colStart: number[] = [];
  let rows = 0;
  let cols = 0;
  for (const h of rowHeights) {
    rowStart.push(rows);
    rows += h;
  }
  for (const w of colWidths) {
    colStart.push(cols);
    cols += w;
  }
  const values = new Array<unknown>(rows * cols).fill(pad);
  for (let r = 0; r < arr.rows; r++) {
    for (let c = 0; c < arr.cols; c++) {
      const v = arr.get(r, c);
      const block = v instanceof RangeValue ? v : new RangeValue([v], 1, 1);
      for (let br = 0; br < block.rows; br++) {
        for (let bc = 0; bc < block.cols; bc++) {
          values[(rowStart[r] + br) * cols + colStart[c] + bc] = block.get(br, bc);
        }
      }
    }
  }
  return new RangeValue(values, rows, cols);
}

/** Stack arrays below (or beside) each other, padding with #N/A. */
function stack(arrays: unknown[], vertical: boolean): RangeValue {
  const parts = arrays.map(asArray);
  if (parts.length === 0) throw new Error('#VALUE!');
  if (!vertical) return transpose(stack(parts.map(transpose), true));
  const cols = Math.max(...parts.map((p) => p.cols));
  const values: unknown[] = [];
  let rows = 0;
  for (const p of parts) {
    for (let r = 0; r < p.rows; r++) {
      for (let c = 0; c < cols; c++) values.push(c < p.cols ? p.get(r, c) : '#N/A');
    }
    rows += p.rows;
  }
  return new RangeValue(values, rows, cols);
}

/** TAKE / DROP along one axis. `count` is rows/columns from the start (positive) or the end (negative). */
function takeIndices(size: number, count: number | undefined, drop: boolean): number[] {
  const all = Array.from({ length: size }, (_, i) => i);
  if (count === undefined) return all;
  if (drop) {
    const kept = count >= 0 ? all.slice(count) : all.slice(0, Math.max(0, size + count));
    if (kept.length === 0) throw new Error('#CALC!');
    return kept;
  }
  if (count === 0) throw new Error('#CALC!');
  return count > 0 ? all.slice(0, count) : all.slice(Math.max(0, size + count));
}

/** Split text at any of the delimiters. */
function splitText(text: string, delimiters: string[], ignoreCase: boolean): string[] {
  const delims = delimiters.filter((d) => d !== '');
  if (delims.length === 0) return [text];
  const out: string[] = [];
  const hay = ignoreCase ? text.toLowerCase() : text;
  const needles = ignoreCase ? delims.map((d) => d.toLowerCase()) : delims;
  let start = 0;
  let i = 0;
  while (i < text.length) {
    const hit = needles.find((d) => hay.startsWith(d, i));
    if (hit !== undefined) {
      out.push(text.slice(start, i));
      i += hit.length;
      start = i;
    } else {
      i++;
    }
  }
  out.push(text.slice(start));
  return out;
}

export function registerPureBuiltins(reg: Registrar): void {
  const { def, agg } = reg;

  // ─── Aggregates ─────────────────────────────────────

  agg('SUM', (args) => args.reduce((sum: number, v) => sum + (Number(v) || 0), 0), { propagate: true });

  agg(
    'AVERAGE',
    (args) => {
      const nums = args.filter((v) => typeof v === 'number' || !isNaN(Number(v)));
      if (nums.length === 0) throw new Error('#DIV/0!');
      const sum = nums.reduce((s: number, v) => s + Number(v), 0);
      return sum / nums.length;
    },
    { propagate: true }
  );

  const extreme = (pick: (a: number, b: number) => number) => (args: unknown[]) => {
    let result: number | undefined;
    for (const v of args) {
      const n = Number(v);
      if (!isNaN(n)) result = result === undefined ? n : pick(result, n);
    }
    return result ?? 0;
  };
  agg('MIN', extreme(Math.min), { propagate: true });
  agg('MAX', extreme(Math.max), { propagate: true });

  agg('COUNT', (args) =>
    args.filter((v) => typeof v === 'number' || (!isNaN(Number(v)) && String(v).trim() !== '')).length
  );

  agg('COUNTA', (args) => args.filter((v) => v !== '' && v !== null && v !== undefined).length);

  agg('CONCAT', (args) => args.map(toText).join(''), { propagate: true });

  agg(
    'PRODUCT',
    (args) => {
      const nums = numericArgs(args);
      return nums.length ? nums.reduce((p, n) => p * n, 1) : 0;
    },
    { propagate: true }
  );

  agg(
    'MEDIAN',
    (args) => {
      const nums = numericArgs(args).sort((a, b) => a - b);
      if (nums.length === 0) throw new Error('#NUM!');
      const mid = Math.floor(nums.length / 2);
      return nums.length % 2 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
    },
    { propagate: true }
  );

  // Sample (n - 1) and population (n) variance / standard deviation.
  // Text and booleans in ranges are ignored, as in Excel.
  const variance = (args: unknown[], sample: boolean): number => {
    const nums = numericArgs(args);
    const n = nums.length;
    if (n < (sample ? 2 : 1)) throw new Error('#DIV/0!');
    const mean = nums.reduce((a, b) => a + b, 0) / n;
    const ss = nums.reduce((a, b) => a + (b - mean) ** 2, 0);
    return ss / (sample ? n - 1 : n);
  };
  for (const name of ['VAR', 'VAR.S']) agg(name, (args) => variance(args, true), { propagate: true });
  agg('VAR.P', (args) => variance(args, false), { propagate: true });
  for (const name of ['STDEV', 'STDEV.S']) agg(name, (args) => Math.sqrt(variance(args, true)), { propagate: true });
  agg('STDEV.P', (args) => Math.sqrt(variance(args, false)), { propagate: true });

  // NPV: discounted values starting one period from now; text and blanks skipped.
  def('NPV', 'nb', (_ctx, rate, ...args) => {
    const r = toNumber(rate);
    if (r === -1) throw new Error('#DIV/0!');
    const values = args.flatMap((arg) => flattenDeep(arg)).filter((v) => v !== undefined);
    const error = values.find(isErrorCode);
    if (error !== undefined) throw new Error(error);
    const nums = numericArgs(values);
    if (nums.length === 0) throw new Error('#VALUE!');
    return nums.reduce((sum, v, i) => sum + v / Math.pow(1 + r, i + 1), 0);
  });

  // ─── Basic ──────────────────────────────────────────

  def('ABS', 'n', (_ctx, val) => Math.abs(Number(val)));

  def('ROUND', 'n', (_ctx, val, digits) => {
    const d = Number(digits) || 0;
    const n = Number(val);
    const factor = Math.pow(10, d);
    // Neutralise IEEE 754 drift (e.g. 1.005*100 → 100.4999… → 100.5)
    // then round half away from zero (Excel convention).
    const shifted = parseFloat((Math.abs(n) * factor).toPrecision(15));
    return (Math.sign(n) * Math.round(shifted)) / factor;
  });

  def('UPPER', 't', (_ctx, val) => toText(val).toUpperCase());
  def('LOWER', 't', (_ctx, val) => toText(val).toLowerCase());
  def('LEN', 't', (_ctx, val) => toText(val).length);
  def('TRIM', 't', (_ctx, val) => toText(val).trim().replace(/ +/g, ' '));

  // ─── Logic ──────────────────────────────────────────

  const everyValue = (args: unknown[], test: (v: unknown) => boolean): void => {
    for (const arg of args) {
      if (arg instanceof RangeValue) {
        for (const v of flattenDeep(arg)) {
          if (v !== undefined && test(v)) return;
        }
      } else if (test(arg)) {
        return;
      }
    }
  };

  def('AND', 'a', (_ctx, ...args) => {
    let result = true;
    everyValue(args, (v) => {
      if (isErrorCode(v)) throw new Error(v);
      if (!v) result = false;
      return !result;
    });
    return result;
  });

  def('OR', 'a', (_ctx, ...args) => {
    let result = false;
    everyValue(args, (v) => {
      if (isErrorCode(v)) throw new Error(v);
      if (v) result = true;
      return result;
    });
    return result;
  });

  def('NOT', 'n', (_ctx, val) => !val);
  def('TRUE', 'n', () => true, { arity: [0, 0] });
  def('FALSE', 'n', () => false, { arity: [0, 0] });
  def('NA', 'n', () => {
    throw new Error('#N/A');
  }, { arity: [0, 0] });

  // ─── Conditional aggregation ────────────────────────

  def('SUMIF', 'ana', (_ctx, range, criteria, sumRange?) => {
    const criteriaStr = String(criteria);
    const rangeArr = toValueList(range);
    const sumArr: unknown[] | undefined =
      sumRange instanceof RangeValue ? sumRange.values : Array.isArray(sumRange) ? sumRange : undefined;

    let total = 0;
    for (let i = 0; i < rangeArr.length; i++) {
      const val = rangeArr[i];
      if (cellMatchesCriteria(val, criteriaStr)) {
        total += Number(sumArr ? sumArr[i] : val) || 0;
      }
    }
    return total;
  });

  def('COUNTIF', 'an', (_ctx, range, criteria) => {
    const criteriaStr = String(criteria);
    let count = 0;
    for (const val of toValueList(range)) {
      if (cellMatchesCriteria(val, criteriaStr)) count++;
    }
    return count;
  });

  def('AVERAGEIF', 'ana', (_ctx, range, criteria, avgRange?) => {
    const criteriaStr = String(criteria);
    const rangeArr = toValueList(range);
    const avgArr: unknown[] | undefined =
      avgRange instanceof RangeValue ? avgRange.values : Array.isArray(avgRange) ? avgRange : undefined;

    let total = 0;
    let count = 0;
    for (let i = 0; i < rangeArr.length; i++) {
      const val = rangeArr[i];
      if (val !== undefined && matchesCriteria(val, criteriaStr)) {
        const numVal = avgArr ? Number(avgArr[i]) : Number(val);
        if (!isNaN(numVal)) {
          total += numVal;
          count++;
        }
      }
    }
    if (count === 0) throw new Error('#DIV/0!');
    return total / count;
  });

  def('SUMIFS', 'a|an', (_ctx, sumRange, ...pairs) => {
    const sumValues = toValueList(sumRange);
    const mask = matchAllCriteria(pairs, sumRange);
    let total = 0;
    for (let i = 0; i < sumValues.length; i++) {
      if (!mask[i]) continue;
      // An error in a summed cell is the result (Excel), not a silent 0
      const v = sumValues[i];
      if (isErrorCode(v)) throw new Error(v);
      // Text (even numeric-looking) and booleans in the sum range are skipped
      if (typeof v === 'number') total += v;
    }
    return total;
  });

  def('COUNTIFS', '|an', (_ctx, ...pairs) => {
    if (pairs.length === 0) throw new Error('#VALUE!');
    const mask = matchAllCriteria(pairs, pairs[0]);
    return mask.filter(Boolean).length;
  });

  def('AVERAGEIFS', 'a|an', (_ctx, avgRange, ...pairs) => {
    const avgValues = toValueList(avgRange);
    const mask = matchAllCriteria(pairs, avgRange);
    let total = 0;
    let count = 0;
    for (let i = 0; i < avgValues.length; i++) {
      const v = avgValues[i];
      if (!mask[i]) continue;
      if (isErrorCode(v)) throw new Error(v);
      if (typeof v === 'number') {
        total += v;
        count++;
      }
    }
    if (count === 0) throw new Error('#DIV/0!');
    return total / count;
  });

  // ─── Lookup ─────────────────────────────────────────

  def('XLOOKUP', 'naaxnn', (_ctx, lookupValue, lookupRange, returnRange, ifNotFound?, matchMode?, searchMode?) => {
    if (!(lookupRange instanceof RangeValue) || !(returnRange instanceof RangeValue)) {
      throw new Error('#VALUE!');
    }
    const vertical = lookupRange.cols === 1;
    if (!vertical && lookupRange.rows !== 1) throw new Error('#VALUE!');
    const keys = lookupRange.values;
    if ((vertical ? returnRange.rows : returnRange.cols) !== keys.length) {
      throw new Error('#VALUE!');
    }

    // 0 = exact, -1 = exact or next smaller, 1 = exact or next larger.
    // Wildcard matching (2) is not supported.
    const mode = matchMode === undefined ? 0 : Math.trunc(Number(matchMode));
    if (mode !== 0 && mode !== -1 && mode !== 1) throw new Error('#VALUE!');
    // 1 / 2 = first to last, -1 / -2 = last to first (binary search modes
    // give the same answer on sorted data, so they search linearly).
    const search = searchMode === undefined ? 1 : Math.trunc(Number(searchMode));
    if (![1, -1, 2, -2].includes(search)) throw new Error('#VALUE!');

    let best = -1;
    for (let step = 0; step < keys.length; step++) {
      const i = search < 0 ? keys.length - 1 - step : step;
      const key = keys[i];
      // lookupCompare orders numbers and text only; booleans match exactly
      if (typeof lookupValue === 'boolean' || typeof key === 'boolean') {
        if (key === lookupValue) {
          best = i;
          break;
        }
        continue;
      }
      const cmp = lookupCompare(key, lookupValue);
      if (isNaN(cmp)) continue;
      if (cmp === 0) {
        best = i;
        break;
      }
      const closer =
        best === -1 ||
        (mode === -1 ? lookupCompare(keys[i], keys[best]) > 0 : lookupCompare(keys[i], keys[best]) < 0);
      if (((mode === -1 && cmp < 0) || (mode === 1 && cmp > 0)) && closer) {
        best = i;
      }
    }

    if (best === -1) {
      if (ifNotFound instanceof ArgumentError) throw new Error(ifNotFound.code);
      if (ifNotFound !== undefined) return ifNotFound;
      throw new Error('#N/A');
    }
    // A return range with several columns (rows) gives the whole matching row (column)
    if (vertical) {
      return returnRange.cols === 1 ? returnRange.get(best, 0) : RangeValue.row(returnRange.getRow(best));
    }
    return returnRange.rows === 1 ? returnRange.get(0, best) : RangeValue.column(returnRange.getCol(best));
  });

  def('VLOOKUP', 'nann', (_ctx, lookupValue, tableRange, colIndex, exactMatch?) => {
    if (!(tableRange instanceof RangeValue)) {
      throw new Error('#VALUE!');
    }
    const colIdx = Math.trunc(Number(colIndex));
    if (colIdx < 1 || colIdx > tableRange.cols) {
      throw new Error('#REF!');
    }
    // Excel convention: omitted or TRUE/1 = approximate match, FALSE/0 = exact match
    const isExact = exactMatch === false || exactMatch === 0;

    const firstCol = tableRange.getCol(0);
    if (isExact) {
      for (let r = 0; r < firstCol.length; r++) {
        if (looseEquals(firstCol[r], lookupValue)) return tableRange.get(r, colIdx - 1);
      }
    } else {
      // Approximate match: largest value <= lookupValue (data assumed sorted ascending)
      const r = approximateMatch(firstCol, lookupValue);
      if (r !== -1) return tableRange.get(r, colIdx - 1);
    }
    throw new Error('#N/A');
  });

  def('HLOOKUP', 'nann', (_ctx, lookupValue, tableRange, rowIndex, exactMatch?) => {
    if (!(tableRange instanceof RangeValue)) {
      throw new Error('#VALUE!');
    }
    const rowIdx = Math.trunc(Number(rowIndex));
    if (rowIdx < 1 || rowIdx > tableRange.rows) {
      throw new Error('#REF!');
    }
    const isExact = exactMatch === false || exactMatch === 0;

    const firstRow = tableRange.getRow(0);
    if (isExact) {
      for (let c = 0; c < firstRow.length; c++) {
        if (looseEquals(firstRow[c], lookupValue)) return tableRange.get(rowIdx - 1, c);
      }
    } else {
      const c = approximateMatch(firstRow, lookupValue);
      if (c !== -1) return tableRange.get(rowIdx - 1, c);
    }
    throw new Error('#N/A');
  });

  def('MATCH', 'nan', (_ctx, lookupValue, rangeArg, matchType?) => {
    let arr: unknown[];
    if (rangeArg instanceof RangeValue) {
      arr = rangeArg.values;
    } else if (Array.isArray(rangeArg)) {
      arr = rangeArg;
    } else {
      throw new Error('#VALUE!');
    }

    const mt = matchType !== undefined ? Number(matchType) : 1;

    if (mt === 0) {
      for (let i = 0; i < arr.length; i++) {
        if (looseEquals(arr[i], lookupValue)) return i + 1;
      }
      throw new Error('#N/A');
    } else if (mt === 1) {
      // Largest value <= lookupValue (data assumed sorted ascending)
      const match = approximateMatch(arr, lookupValue);
      if (match === -1) throw new Error('#N/A');
      return match + 1;
    } else {
      // mt === -1: Smallest value >= lookupValue (data assumed sorted descending)
      let lastMatch = -1;
      for (let i = 0; i < arr.length; i++) {
        if (lookupCompare(arr[i], lookupValue) >= 0) {
          lastMatch = i;
        }
      }
      if (lastMatch === -1) throw new Error('#N/A');
      return lastMatch + 1;
    }
  });

  def('XMATCH', 'nann', (_ctx, lookupValue, lookupArray, matchMode?, searchMode?) => {
    const keys = flattenDeep(lookupArray);
    const mode = intArg(matchMode, 0);
    if (mode !== 0 && mode !== -1 && mode !== 1) throw new Error('#VALUE!');
    const search = intArg(searchMode, 1);
    if (![1, -1, 2, -2].includes(search)) throw new Error('#VALUE!');
    let best = -1;
    for (let step = 0; step < keys.length; step++) {
      const i = search < 0 ? keys.length - 1 - step : step;
      const key = keys[i];
      if (typeof lookupValue === 'boolean' || typeof key === 'boolean') {
        if (key === lookupValue) return i + 1;
        continue;
      }
      const cmp = lookupCompare(key, lookupValue);
      if (isNaN(cmp)) continue;
      if (cmp === 0) return i + 1;
      const closer =
        best === -1 ||
        (mode === -1 ? lookupCompare(keys[i], keys[best]) > 0 : lookupCompare(keys[i], keys[best]) < 0);
      if (((mode === -1 && cmp < 0) || (mode === 1 && cmp > 0)) && closer) best = i;
    }
    if (best === -1) throw new Error('#N/A');
    return best + 1;
  });

  // ─── Math ───────────────────────────────────────────

  def('MOD', 'n', (_ctx, num, divisor) => {
    const n = Number(num);
    const d = Number(divisor);
    if (d === 0) throw new Error('#DIV/0!');
    // Excel MOD: n - d * FLOOR(n/d). Result sign follows the divisor,
    // unlike JavaScript's % which follows the dividend.
    return n - d * Math.floor(n / d);
  });

  def('POWER', 'n', (_ctx, base, exp) => power(Number(base), Number(exp)));

  def('CEILING', 'n', (_ctx, num, significance?) => {
    const n = Number(num);
    const sig = significance !== undefined ? Number(significance) : 1;
    if (sig === 0) return 0;
    // Round the quotient to 10 significant digits to neutralise IEEE 754
    // drift (e.g. 0.07/0.01 → 7.000000000000001 → 7) before ceiling.
    const q = parseFloat((n / sig).toPrecision(10));
    return Math.ceil(q) * sig;
  });

  def('FLOOR', 'n', (_ctx, num, significance?) => {
    const n = Number(num);
    const sig = significance !== undefined ? Number(significance) : 1;
    if (sig === 0) return 0;
    const q = parseFloat((n / sig).toPrecision(10));
    return Math.floor(q) * sig;
  });

  // Omitted precision is 0; anything non-numeric is #VALUE! (not a silent 0)
  const precision = (digits: unknown) => (digits === undefined ? 0 : Math.trunc(toNumber(digits)));

  def('ROUNDUP', 'n', (_ctx, val, digits) => roundMagnitude(toNumber(val), precision(digits), Math.ceil));
  def('ROUNDDOWN', 'n', (_ctx, val, digits) => roundMagnitude(toNumber(val), precision(digits), Math.floor));
  def('INT', 'n', (_ctx, val) => Math.floor(toNumber(val)));

  def('SQRT', 'n', (_ctx, val) => {
    const n = toNumber(val);
    if (n < 0) throw new Error('#NUM!');
    return Math.sqrt(n);
  });

  // SUMPRODUCT: equal-shaped arrays multiplied element-wise; non-numbers are 0.
  def('SUMPRODUCT', 'a', (_ctx, ...arrays) => {
    if (arrays.length === 0) throw new Error('#VALUE!');
    const [rows, cols] = rangeShape(arrays[0]);
    const lists = arrays.map((a) => {
      const [r, c] = rangeShape(a);
      if (r !== rows || c !== cols) throw new Error('#VALUE!');
      return toValueList(a);
    });
    let total = 0;
    for (let i = 0; i < rows * cols; i++) {
      let product = 1;
      for (const list of lists) {
        const v = list[i];
        if (isErrorCode(v)) throw new Error(v);
        product *= typeof v === 'number' ? v : 0;
      }
      total += product;
    }
    return total;
  });

  // ─── Financial ──────────────────────────────────────
  // Excel's sign convention: money paid out is negative, received positive.
  // type 1 means payments at the start of each period, 0 (default) at the end.

  const finArgs = (rate: unknown, nper: unknown, type: unknown) => {
    const r = toNumber(rate);
    const n = toNumber(nper);
    const t = type === undefined ? 0 : toNumber(type) ? 1 : 0;
    return { r, n, t };
  };

  def('PMT', 'n', (_ctx, rate, nper, pv, fv?, type?) => {
    const { r, n, t } = finArgs(rate, nper, type);
    const p = toNumber(pv);
    const f = fv === undefined ? 0 : toNumber(fv);
    if (n === 0) throw new Error('#NUM!');
    if (r === 0) return -(p + f) / n;
    const g = Math.pow(1 + r, n);
    return -(r * (p * g + f)) / ((1 + r * t) * (g - 1));
  });

  def('FV', 'n', (_ctx, rate, nper, pmt, pv?, type?) => {
    const { r, n, t } = finArgs(rate, nper, type);
    const m = toNumber(pmt);
    const p = pv === undefined ? 0 : toNumber(pv);
    if (r === 0) return -(p + m * n);
    const g = Math.pow(1 + r, n);
    return -(p * g + (m * (1 + r * t) * (g - 1)) / r);
  });

  def('PV', 'n', (_ctx, rate, nper, pmt, fv?, type?) => {
    const { r, n, t } = finArgs(rate, nper, type);
    const m = toNumber(pmt);
    const f = fv === undefined ? 0 : toNumber(fv);
    if (r === 0) return -(f + m * n);
    const g = Math.pow(1 + r, n);
    return -(f + (m * (1 + r * t) * (g - 1)) / r) / g;
  });

  // ─── Information ────────────────────────────────────
  // Errors arrive as their code, empty cells as undefined.

  const info = (name: string, test: (v: unknown) => boolean) => def(name, 'v', (_ctx, v) => test(v), { arity: [1, 1] });
  info('ISBLANK', (v) => v === undefined);
  info('ISNUMBER', (v) => typeof v === 'number');
  info('ISLOGICAL', (v) => typeof v === 'boolean');
  info('ISTEXT', (v) => typeof v === 'string' && !isErrorCode(v));
  info('ISERROR', (v) => isErrorCode(v));
  info('ISERR', (v) => isErrorCode(v) && v !== '#N/A');
  info('ISNA', (v) => v === '#N/A');

  // ─── Text ───────────────────────────────────────────

  def('TEXTJOIN', 'tnb', (_ctx, delimiter, ignoreEmpty, ...texts) => {
    const parts: string[] = [];
    for (const arg of texts) {
      for (const v of flattenDeep(arg)) {
        if (isErrorCode(v)) throw new Error(v);
        if (v === undefined || v === '') {
          if (!ignoreEmpty) parts.push('');
        } else {
          parts.push(toText(v));
        }
      }
    }
    return parts.join(toText(delimiter));
  });

  def('LEFT', 'tn', (_ctx, text, n?) => {
    const count = n !== undefined ? Number(n) : 1;
    if (!(count >= 0)) throw new Error('#VALUE!');
    return toText(text).substring(0, count);
  });

  def('RIGHT', 'tn', (_ctx, text, n?) => {
    const s = toText(text);
    const count = n !== undefined ? Number(n) : 1;
    if (!(count >= 0)) throw new Error('#VALUE!');
    return s.substring(s.length - count);
  });

  def('MID', 'tn', (_ctx, text, start, n) => {
    const s = toText(text);
    const startNum = Number(start);
    const count = Number(n);
    if (!(startNum >= 1) || !(count >= 0)) throw new Error('#VALUE!');
    return s.substring(startNum - 1, startNum - 1 + count);
  });

  def('SUBSTITUTE', 'tttn', (_ctx, text, oldStr, newStr, instance?) => {
    const s = toText(text);
    const old = toText(oldStr);
    const replacement = toText(newStr);

    // Excel: empty search string → return text unchanged
    if (old === '') return s;

    if (instance !== undefined) {
      const nth = Number(instance);
      if (!(nth >= 1)) throw new Error('#VALUE!');
      let count = 0;
      let searchFrom = 0;
      while (searchFrom < s.length) {
        const idx = s.indexOf(old, searchFrom);
        if (idx === -1) break;
        count++;
        if (count === nth) {
          return s.substring(0, idx) + replacement + s.substring(idx + old.length);
        }
        searchFrom = idx + 1;
      }
      return s; // nth occurrence not found, return unchanged
    }

    return s.split(old).join(replacement);
  });

  def('FIND', 'ttn', (_ctx, search, text, start?) => {
    const s = toText(text);
    const searchStr = toText(search);
    const startNum = start !== undefined ? Number(start) : 1;
    if (!(startNum >= 1)) throw new Error('#VALUE!');
    const idx = s.indexOf(searchStr, startNum - 1);
    if (idx === -1) throw new Error('#VALUE!');
    return idx + 1; // 1-indexed
  });

  def('TEXT', 'nt', (_ctx, value, format) => {
    const num = Number(value);
    const fmt = String(format);

    if (isNaN(num)) return String(value);

    if (fmt === '0') {
      // Round half away from zero to match Excel (Math.round rounds toward +Infinity)
      return (Math.sign(num) * Math.round(Math.abs(num))).toString();
    }
    if (fmt === '0.00' || fmt === '0.0') {
      const decimals = (fmt.split('.')[1] || '').length;
      return num.toFixed(decimals);
    }
    if (fmt === '#,##0' || fmt === '#,##0.00') {
      const decimals = fmt.includes('.') ? (fmt.split('.')[1] || '').length : 0;
      return num.toLocaleString('en-US', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
      });
    }

    return num.toString();
  });

  def('VALUE', 'n', (_ctx, text) => {
    if (typeof text === 'boolean') return text ? 1 : 0;
    const num = Number(String(text));
    if (isNaN(num)) throw new Error('#VALUE!');
    return num;
  });

  /**
   * TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]).
   * Given a range of texts, every text gets its own array (one array per cell).
   */
  def('TEXTSPLIT', 'taannv', (_ctx, text, colDelim, rowDelim?, ignoreEmpty?, matchMode?, padWith?) => {
    const delims = (arg: unknown): string[] =>
      arg === undefined ? [] : flattenDeep(arg).filter((d) => d !== undefined).map(toText);
    const colDelims = delims(colDelim);
    const rowDelims = delims(rowDelim);
    if (colDelims.length === 0 && rowDelims.length === 0) throw new Error('#VALUE!');
    const ignoreCase = matchMode !== undefined && toNumber(matchMode) === 1;
    const skipEmpty = ignoreEmpty !== undefined && toBoolean(ignoreEmpty);
    const keep = (parts: string[]) => (skipEmpty ? parts.filter((p) => p !== '') : parts);

    const lines = keep(splitText(toText(text), rowDelims, ignoreCase)).map((line) =>
      keep(splitText(line, colDelims, ignoreCase))
    );
    const rows = lines.filter((l) => l.length > 0 || !skipEmpty);
    if (rows.length === 0) throw new Error('#CALC!');
    const cols = Math.max(1, ...rows.map((l) => l.length));
    const pad = padWith === undefined ? '#N/A' : padWith;
    const values: unknown[] = [];
    for (const line of rows) {
      for (let c = 0; c < cols; c++) values.push(c < line.length ? line[c] : pad);
    }
    return new RangeValue(values, rows.length, cols);
  });

  // ─── Date ───────────────────────────────────────────

  def('DATE', 'n', (_ctx, year, month, day) => {
    // Excel serial number: days since 1899-12-30
    let y = Number(year);
    const m = Number(month);
    const d = Number(day);
    // Excel two-digit year windowing: 0-29 → 2000+, 30-99 → 1900+
    if (y >= 0 && y <= 29) y += 2000;
    else if (y >= 30 && y <= 99) y += 1900;
    // Use setFullYear to bypass JS Date's own 0-99 → 1900+ mapping
    const date = new Date(0);
    date.setFullYear(y, m - 1, d);
    date.setHours(0, 0, 0, 0);
    const epoch = new Date(0);
    epoch.setFullYear(1899, 11, 30);
    epoch.setHours(0, 0, 0, 0);
    const diff = date.getTime() - epoch.getTime();
    return Math.round(diff / (1000 * 60 * 60 * 24));
  });

  def(
    'NOW',
    'n',
    () => {
      // Serial of the local wall-clock time. Computed from date fields rather
      // than a timestamp difference, so the timezone offset in effect in 1899
      // (often a non-whole-hour local mean time) doesn't skew the result.
      const now = new Date();
      return (
        dateToSerial(
          now.getFullYear(), now.getMonth(), now.getDate(),
          now.getHours(), now.getMinutes(), now.getSeconds()
        ) + now.getMilliseconds() / 86_400_000
      );
    },
    { volatile: true }
  );

  def(
    'TODAY',
    'n',
    () => {
      const now = new Date();
      return dateToSerial(now.getFullYear(), now.getMonth(), now.getDate());
    },
    { volatile: true }
  );

  def('YEAR', 'n', (_ctx, serial) => serialArg(serial).getUTCFullYear());
  def('MONTH', 'n', (_ctx, serial) => serialArg(serial).getUTCMonth() + 1);
  def('DAY', 'n', (_ctx, serial) => serialArg(serial).getUTCDate());

  // EDATE: same day n months later, clamped to the end of a shorter month.
  def('EDATE', 'n', (_ctx, start, months) => {
    const d = serialArg(start);
    const m = d.getUTCMonth() + Math.trunc(toNumber(months));
    const lastDay = new Date(Date.UTC(d.getUTCFullYear(), m + 1, 0)).getUTCDate();
    const serial = dateToSerial(d.getUTCFullYear(), m, Math.min(d.getUTCDate(), lastDay));
    if (serial < 0) throw new Error('#NUM!');
    return serial;
  });

  // EOMONTH: last day of the month n months later.
  def('EOMONTH', 'n', (_ctx, start, months) => {
    const d = serialArg(start);
    const serial = dateToSerial(d.getUTCFullYear(), d.getUTCMonth() + Math.trunc(toNumber(months)) + 1, 0);
    if (serial < 0) throw new Error('#NUM!');
    return serial;
  });

  def('WEEKDAY', 'n', (_ctx, serial, returnType?) => {
    const day = serialArg(serial).getUTCDay(); // 0 = Sunday
    const type = returnType === undefined ? 1 : Math.trunc(Number(returnType));
    switch (type) {
      case 1: return day + 1;             // Sunday = 1 … Saturday = 7
      case 2: return ((day + 6) % 7) + 1; // Monday = 1 … Sunday = 7
      case 3: return (day + 6) % 7;       // Monday = 0 … Sunday = 6
      default: throw new Error('#NUM!');
    }
  });

  // ─── Dynamic arrays ─────────────────────────────────

  def('SEQUENCE', 'n', (_ctx, rows, cols?, start?, step?) => {
    const r = intArg(rows);
    const c = intArg(cols, 1);
    if (r < 1 || c < 1) throw new Error('#CALC!');
    if (r * c > 10_000_000) throw new Error('#NUM!');
    const first = start === undefined ? 1 : toNumber(start);
    const by = step === undefined ? 1 : toNumber(step);
    const values = new Array<unknown>(r * c);
    for (let i = 0; i < values.length; i++) values[i] = first + i * by;
    return new RangeValue(values, r, c);
  });

  def('TRANSPOSE', 'a', (_ctx, array) => transpose(asArray(array)));

  def('FILTER', 'aax', (_ctx, array, include, ifEmpty?) => {
    const arr = asArray(array);
    const inc = asArray(include);
    const byRow = inc.cols === 1 && inc.rows === arr.rows;
    const byCol = !byRow && inc.rows === 1 && inc.cols === arr.cols;
    if (!byRow && !byCol) throw new Error('#VALUE!');
    const kept: number[] = [];
    inc.values.forEach((v, i) => {
      if (isErrorCode(v)) throw new Error(v);
      if (toBoolean(v)) kept.push(i);
    });
    if (kept.length === 0) {
      if (ifEmpty instanceof ArgumentError) throw new Error(ifEmpty.code);
      if (ifEmpty !== undefined) return ifEmpty;
      throw new Error('#CALC!');
    }
    return byRow ? pickRows(arr, kept) : pickCols(arr, kept);
  });

  def('SORT', 'annn', (_ctx, array, sortIndex?, sortOrder?, byCol?) => {
    const columns = byCol !== undefined && toBoolean(byCol);
    const arr = columns ? transpose(asArray(array)) : asArray(array);
    const index = intArg(sortIndex, 1) - 1;
    if (index < 0 || index >= arr.cols) throw new Error('#VALUE!');
    const direction = intArg(sortOrder, 1);
    if (direction !== 1 && direction !== -1) throw new Error('#VALUE!');
    const order = Array.from({ length: arr.rows }, (_, i) => i);
    order.sort((a, b) => compareForSort(arr.get(a, index), arr.get(b, index), direction) || a - b);
    const sorted = pickRows(arr, order);
    return columns ? transpose(sorted) : sorted;
  });

  def('SORTBY', 'a|an', (_ctx, array, ...keys) => {
    const arr = asArray(array);
    if (keys.length === 0) throw new Error('#VALUE!');
    const specs: Array<{ values: unknown[]; direction: number }> = [];
    let byRow = true;
    for (let i = 0; i < keys.length; i += 2) {
      const by = asArray(keys[i]);
      const direction = intArg(keys[i + 1], 1);
      if (direction !== 1 && direction !== -1) throw new Error('#VALUE!');
      // A column of keys orders the rows; a row of keys orders the columns
      if (i === 0) byRow = by.cols === 1 && by.rows === arr.rows;
      if (byRow ? by.cols !== 1 || by.rows !== arr.rows : by.rows !== 1 || by.cols !== arr.cols) {
        throw new Error('#VALUE!');
      }
      specs.push({ values: by.values, direction });
    }
    const size = byRow ? arr.rows : arr.cols;
    const order = Array.from({ length: size }, (_, i) => i);
    order.sort((a, b) => {
      for (const s of specs) {
        const cmp = compareForSort(s.values[a], s.values[b], s.direction);
        if (cmp !== 0) return cmp;
      }
      return a - b;
    });
    return byRow ? pickRows(arr, order) : pickCols(arr, order);
  });

  def('UNIQUE', 'ann', (_ctx, array, byCol?, exactlyOnce?) => {
    const columns = byCol !== undefined && toBoolean(byCol);
    const arr = columns ? transpose(asArray(array)) : asArray(array);
    const once = exactlyOnce !== undefined && toBoolean(exactlyOnce);
    const seen = new Map<string, { row: number; count: number }>();
    for (let r = 0; r < arr.rows; r++) {
      const key = arr.getRow(r).map(identityKey).join('\u0000');
      const entry = seen.get(key);
      if (entry) entry.count++;
      else seen.set(key, { row: r, count: 1 });
    }
    const rows = [...seen.values()].filter((e) => !once || e.count === 1).map((e) => e.row);
    if (rows.length === 0) throw new Error('#CALC!');
    const result = pickRows(arr, rows);
    return columns ? transpose(result) : result;
  });

  def('VSTACK', 'a', (_ctx, ...arrays) => stack(arrays, true));
  def('HSTACK', 'a', (_ctx, ...arrays) => stack(arrays, false));

  const takeDrop = (drop: boolean) => (_ctx: FormulaContext, array: unknown, rows?: unknown, cols?: unknown) => {
    const arr = asArray(array);
    const r = rows === undefined ? undefined : intArg(rows);
    const c = cols === undefined ? undefined : intArg(cols);
    return pickCols(pickRows(arr, takeIndices(arr.rows, r, drop)), takeIndices(arr.cols, c, drop));
  };
  def('TAKE', 'ann', takeDrop(false));
  def('DROP', 'ann', takeDrop(true));

  def('CHOOSECOLS', 'a', (_ctx, array, ...cols) => {
    const arr = asArray(array);
    return pickCols(arr, positions(cols, arr.cols));
  });

  def('CHOOSEROWS', 'a', (_ctx, array, ...rows) => {
    const arr = asArray(array);
    return pickRows(arr, positions(rows, arr.rows));
  });

  // TOCOL / TOROW(array, [ignore], [scan_by_column]); ignore: 1 blanks, 2 errors, 3 both
  const toVector = (array: unknown, ignore: unknown, byColumn: unknown): unknown[] => {
    let arr = asArray(array);
    if (byColumn !== undefined && toBoolean(byColumn)) arr = transpose(arr);
    const mode = intArg(ignore, 0);
    const values = arr.values.filter(
      (v) => !((mode & 1) !== 0 && v === undefined) && !((mode & 2) !== 0 && isErrorCode(v))
    );
    if (values.length === 0) throw new Error('#CALC!');
    return values;
  };
  def('TOCOL', 'ann', (_ctx, array, ignore?, byColumn?) => RangeValue.column(toVector(array, ignore, byColumn)));
  def('TOROW', 'ann', (_ctx, array, ignore?, byColumn?) => RangeValue.row(toVector(array, ignore, byColumn)));

  // ─── Several values in one cell ─────────────────────

  /**
   * FLATTEN(array, [pad_value], [levels]): lay nested arrays out in place.
   * Arrays of different sizes are padded with `pad_value` (default #N/A).
   * `levels` is how many levels of nesting to remove (default: all).
   */
  def('FLATTEN', 'avn', (_ctx, array, padValue?, levels?) => {
    let arr = asArray(array);
    const pad = padValue === undefined ? '#N/A' : padValue;
    let remaining = levels === undefined ? Infinity : intArg(levels);
    if (remaining < 1) throw new Error('#VALUE!');
    while (remaining-- > 0 && arr.values.some((v) => v instanceof RangeValue)) {
      arr = flattenOnce(arr, pad);
    }
    return arr;
  });

  const itemSet = (array: unknown): Set<string> => {
    const set = new Set<string>();
    for (const v of flattenDeep(array)) {
      if (v !== undefined) set.add(identityKey(v));
    }
    return set;
  };
  const wanted = (values: unknown): string[] => {
    const keys = flattenDeep(values).filter((v) => v !== undefined).map(identityKey);
    if (keys.length === 0) throw new Error('#VALUE!');
    return keys;
  };

  /** HAS(array, value): whether the value is one of the array's items, at any depth. */
  def('HAS', 'an', (_ctx, array, value) => flattenDeep(array).some((v) => v !== undefined && scalarsEqual(v, value)));

  /** HASANY(array, values): whether at least one of the values is in the array. */
  def('HASANY', 'aa', (_ctx, array, values) => {
    const items = itemSet(array);
    return wanted(values).some((k) => items.has(k));
  });

  /** HASALL(array, values): whether every one of the values is in the array. */
  def('HASALL', 'aa', (_ctx, array, values) => {
    const items = itemSet(array);
    return wanted(values).every((k) => items.has(k));
  });
}

/** `base ^ exp` with Excel's errors: 0^(-n) is #DIV/0!, negative^fraction and overflow are #NUM!. */
export function power(b: number, e: number): number {
  const result = Math.pow(b, e);
  if (!Number.isFinite(result) && Number.isFinite(b) && Number.isFinite(e)) {
    // Only a zero base divides by zero; anything else too large (or not real) is #NUM!
    throw new Error(b === 0 && e < 0 ? '#DIV/0!' : '#NUM!');
  }
  return result;
}
