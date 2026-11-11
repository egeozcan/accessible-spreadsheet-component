/** Logical, information and lookup built-ins (values only; lazy or reference-taking functions live in formula-engine.ts). */
import { type Registrar, ArgumentError, asArray, checkArraySize } from './functions.js';
import { RangeValue, flattenDeep, isErrorCode, toBoolean, toNumber } from './values.js';

/** A captured argument error: thrown only by the function that uses the argument. */
function used(value: unknown): unknown {
  if (value instanceof ArgumentError) throw new Error(value.code);
  return value;
}

/**
 * Excel's approximate-match order: numbers, text (ignoring case) and booleans
 * compare within their own kind; anything else is incomparable (NaN).
 */
function orderOf(a: unknown, b: unknown): number {
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'string' && typeof b === 'string') {
    const x = a.toLowerCase();
    const y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b);
  return NaN;
}

/** Error numbers returned by ERROR.TYPE. */
const ERROR_NUMBERS = new Map<string, number>([
  ['#DIV/0!', 2],
  ['#VALUE!', 3],
  ['#REF!', 4],
  ['#NAME?', 5],
  ['#NUM!', 6],
  ['#N/A', 7],
]);

/** Excel's `=` for a SWITCH match: text ignoring case, numbers and booleans by value, nothing across kinds. */
export function sameValue(a: unknown, b: unknown): boolean {
  if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
  return typeof a === typeof b && a === b;
}

/** A vector of values, row-major, as WRAPROWS and WRAPCOLS read it. */
function wrap(vector: unknown, count: unknown, pad: unknown, byRows: boolean): RangeValue {
  const n = Math.trunc(toNumber(count));
  if (n < 1) throw new Error('#VALUE!');
  const items = asArray(vector).values;
  if (pad instanceof RangeValue) throw new Error('#VALUE!'); // pad_with must be a single value
  const padValue = pad === undefined ? '#N/A' : pad;
  const groups = Math.max(1, Math.ceil(items.length / n));
  const rows = byRows ? groups : n;
  const cols = byRows ? n : groups;
  checkArraySize(rows, cols);
  const values: unknown[] = new Array(rows * cols);
  for (let i = 0; i < rows * cols; i++) {
    const r = Math.floor(i / cols);
    const c = i % cols;
    // Position of the i-th cell in the wrapped order
    const index = byRows ? r * n + c : c * n + r;
    values[i] = index < items.length ? items[index] : padValue;
  }
  return new RangeValue(values, rows, cols);
}

/**
 * Result of CHOOSE / SWITCH over an array of selectors: each element picks an alternative, and
 * alternatives that are arrays are broadcast (a 1-wide dimension repeats) to the common shape.
 */
function pickBroadcast(
  selector: RangeValue,
  resolve: (element: unknown) => { alt: unknown } | { error: string },
  alternatives: unknown[]
): RangeValue {
  let rows = selector.rows;
  let cols = selector.cols;
  for (const alt of alternatives) {
    if (alt instanceof RangeValue) {
      rows = Math.max(rows, alt.rows);
      cols = Math.max(cols, alt.cols);
    }
  }
  checkArraySize(rows, cols);
  // A blank element reads as 0, as it does in a cell
  const at = (a: RangeValue, r: number, c: number): unknown =>
    (a.rows === 1 || r < a.rows) && (a.cols === 1 || c < a.cols)
      ? (a.get(a.rows === 1 ? 0 : r, a.cols === 1 ? 0 : c) ?? 0)
      : '#N/A';
  const values: unknown[] = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const element = at(selector, r, c);
      const picked = isErrorCode(element) ? { error: element } : resolve(element);
      if ('error' in picked) {
        values.push(picked.error);
        continue;
      }
      const alt = picked.alt;
      if (alt instanceof ArgumentError) values.push(alt.code);
      else values.push(alt instanceof RangeValue ? at(alt, r, c) : (alt ?? 0));
    }
  }
  return new RangeValue(values, rows, cols);
}

/** SWITCH where the expression and/or the match values are arrays: every cell is matched on its own. */
function switchBroadcast(expression: unknown, rest: unknown[], hasDefault: boolean): RangeValue {
  const pairsEnd = rest.length - (hasDefault ? 1 : 0);
  let rows = 1;
  let cols = 1;
  for (const v of [expression, ...rest]) {
    if (v instanceof RangeValue) {
      rows = Math.max(rows, v.rows);
      cols = Math.max(cols, v.cols);
    }
  }
  checkArraySize(rows, cols);
  const at = (v: unknown, r: number, c: number): unknown => {
    if (v instanceof ArgumentError) return v.code;
    if (!(v instanceof RangeValue)) return v ?? 0;
    return (v.rows === 1 || r < v.rows) && (v.cols === 1 || c < v.cols)
      ? (v.get(v.rows === 1 ? 0 : r, v.cols === 1 ? 0 : c) ?? 0)
      : '#N/A';
  };
  const values: unknown[] = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const value = at(expression, r, c);
      let result: unknown;
      if (isErrorCode(value)) result = value;
      else {
        let found = false;
        for (let i = 0; i < pairsEnd && !found; i += 2) {
          const candidate = at(rest[i], r, c);
          if (isErrorCode(candidate)) {
            result = candidate;
            found = true;
          } else if (sameValue(value, candidate)) {
            result = at(rest[i + 1], r, c);
            found = true;
          }
        }
        if (!found) result = hasDefault ? at(rest[rest.length - 1], r, c) : '#N/A';
      }
      values.push(result);
    }
  }
  return new RangeValue(values, rows, cols);
}

export function registerLogicalInfoBuiltins(reg: Registrar): void {
  const { def } = reg;

  // ─── Logical ────────────────────────────────────────

  // CHOOSE(index, value1, ...): only the chosen value's error is returned. An array index
  // picks element by element, and array alternatives are broadcast against it.
  def('CHOOSE', 'xx', (_ctx, index, ...values) => {
    if (index instanceof RangeValue) {
      return pickBroadcast(index, (e) => {
        if (isErrorCode(e)) return { error: e };
        const i = typeof e === 'number' || typeof e === 'boolean' || (typeof e === 'string' && e.trim() !== '' && !isNaN(Number(e))) ? Math.trunc(toNumber(e)) : NaN;
        if (isNaN(i)) return { error: '#VALUE!' };
        return i < 1 || i > values.length ? { error: '#VALUE!' } : { alt: values[i - 1] };
      }, values);
    }
    const i = Math.trunc(toNumber(used(index)));
    if (i < 1 || i > values.length) throw new Error('#VALUE!');
    return used(values[i - 1]);
  }, { arity: [2, 254] });

  // SWITCH(expression, value1, result1, ..., [default]): results are only used when chosen.
  def('SWITCH', 'x', (_ctx, expression, ...rest) => {
    const hasDefault = rest.length % 2 === 1;
    const pairsEnd = rest.length - (hasDefault ? 1 : 0);
    const resolve = (value: unknown): { alt: unknown } | { error: string } => {
      if (isErrorCode(value)) return { error: value };
      for (let i = 0; i < pairsEnd; i += 2) {
        // An error in a match value is only reached (and only then returned) by the elements that get this far
        let candidate: unknown;
        try {
          candidate = used(rest[i]);
        } catch (e) {
          return { error: (e as Error).message };
        }
        if (sameValue(value, candidate)) return { alt: rest[i + 1] };
      }
      return hasDefault ? { alt: rest[rest.length - 1] } : { error: '#N/A' };
    };
    // An array expression, or array match values, are matched element by element
    const caseIsArray = rest.slice(0, pairsEnd).some((v, i) => i % 2 === 0 && v instanceof RangeValue);
    if (caseIsArray) return switchBroadcast(expression, rest, hasDefault);
    if (expression instanceof RangeValue) return pickBroadcast(expression, resolve, rest);
    const picked = resolve(used(expression));
    if ('error' in picked) throw new Error(picked.error);
    return used(picked.alt);
  }, { arity: [3, 254] });

  // XOR: TRUE when an odd number of the logical values are TRUE. Text in a range is ignored.
  def('XOR', 'a', (_ctx, ...args) => {
    let logicals = 0;
    let trues = 0;
    const count = (v: unknown) => {
      if (isErrorCode(v)) throw new Error(v);
      logicals++;
      if (toBoolean(v)) trues++;
    };
    for (const arg of args) {
      if (arg instanceof RangeValue) {
        for (const v of flattenDeep(arg)) {
          if (isErrorCode(v)) throw new Error(v);
          if (typeof v === 'boolean' || typeof v === 'number') count(v);
        }
      } else if (isErrorCode(arg)) {
        throw new Error(arg);
      } else if (typeof arg === 'string') {
        // Text given directly must read as a logical value or a number
        const numeric = arg.trim() !== '' && !isNaN(Number(arg));
        // Excel ignores text that is not a logical value or a number
        if (numeric || /^(true|false)$/i.test(arg)) count(arg);
      } else if (arg !== undefined) {
        count(arg);
      }
    }
    if (logicals === 0) throw new Error('#VALUE!');
    return trues % 2 === 1;
  });

  // ─── Information ────────────────────────────────────

  def('ISEVEN', 'n', (_ctx, num) => Math.abs(Math.trunc(toNumber(num))) % 2 === 0, { arity: [1, 1] });
  def('ISODD', 'n', (_ctx, num) => Math.abs(Math.trunc(toNumber(num))) % 2 === 1, { arity: [1, 1] });
  def('ISNONTEXT', 'v', (_ctx, value) => !(typeof value === 'string' && !isErrorCode(value)), { arity: [1, 1] });


  // TYPE: 1 number (and blank), 2 text, 4 logical, 16 error, 64 array
  def('TYPE', 'x', (_ctx, value) => {
    if (value instanceof ArgumentError) return 16;
    if (value instanceof RangeValue) return 64;
    if (typeof value === 'boolean') return 4;
    if (typeof value === 'string') return 2;
    return 1;
  }, { arity: [1, 1] });

  // ERROR.TYPE: the number of an error value, #N/A for anything else
  def('ERROR.TYPE', 'x', (_ctx, value) => {
    const code = value instanceof ArgumentError ? value.code : undefined;
    const n = code === undefined ? undefined : ERROR_NUMBERS.get(code);
    if (n === undefined) throw new Error('#N/A');
    return n;
  }, { arity: [1, 1] });

  // ─── Lookup ─────────────────────────────────────────

  // LOOKUP(value, vector, [result]) and LOOKUP(value, array): the largest value
  // <= value in an ascending vector. The array form searches the first row when
  // the array is wider than it is tall, the first column otherwise.
  def('LOOKUP', 'naa', (_ctx, lookupValue, vector, resultArg?) => {
    const source = asArray(vector);
    let keys: unknown[];
    let results: unknown[];
    if (resultArg === undefined) {
      const wide = source.cols > source.rows;
      keys = wide ? source.getRow(0) : source.getCol(0);
      results = wide ? source.getRow(source.rows - 1) : source.getCol(source.cols - 1);
    } else {
      keys = source.values;
      results = asArray(resultArg).values;
      if (results.length !== keys.length) throw new Error('#N/A');
    }
    let found = -1;
    for (let i = 0; i < keys.length; i++) {
      if (isErrorCode(keys[i])) continue; // errors in the lookup vector are skipped
      const cmp = orderOf(keys[i], lookupValue);
      if (Number.isNaN(cmp)) continue;
      if (cmp > 0) break;
      found = i;
    }
    if (found === -1) throw new Error('#N/A');
    const result = results[found];
    if (isErrorCode(result)) throw new Error(result);
    return result;
  }, { arity: [2, 3] });

  // ─── Array shaping ──────────────────────────────────

  // EXPAND(array, [rows], [cols], [pad_with]): pads with #N/A unless told otherwise
  def('EXPAND', 'annx', (_ctx, array, rows?, cols?, pad?) => {
    const source = asArray(array);
    const r = rows === undefined ? source.rows : Math.trunc(toNumber(rows));
    const c = cols === undefined ? source.cols : Math.trunc(toNumber(cols));
    if (r < source.rows || c < source.cols) throw new Error('#VALUE!');
    checkArraySize(r, c);
    if (pad instanceof RangeValue) throw new Error('#VALUE!'); // pad_with must be a single value
    const padValue = pad === undefined ? '#N/A' : pad instanceof ArgumentError ? pad.code : pad;
    const values: unknown[] = new Array(r * c);
    for (let i = 0; i < r * c; i++) {
      const row = Math.floor(i / c);
      const col = i % c;
      values[i] = row < source.rows && col < source.cols ? source.get(row, col) : padValue;
    }
    return new RangeValue(values, r, c);
  }, { arity: [1, 4] });

  // WRAPROWS(vector, wrap_count, [pad_with]): fills rows of wrap_count, in order
  def('WRAPROWS', 'anx', (_ctx, vector, count, pad?) => wrap(vector, count, pad instanceof ArgumentError ? pad.code : pad, true), {
    arity: [2, 3],
  });

  // WRAPCOLS(vector, wrap_count, [pad_with]): fills columns of wrap_count, in order
  def('WRAPCOLS', 'anx', (_ctx, vector, count, pad?) => wrap(vector, count, pad instanceof ArgumentError ? pad.code : pad, false), {
    arity: [2, 3],
  });
}
