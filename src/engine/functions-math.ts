/** Math and trigonometry built-ins (values only; lazy or reference-taking functions live in formula-engine.ts). */
import { ArgumentError, type Registrar, asArray, checkArraySize, numericArgs, power } from './functions.js';
import { RangeValue, flattenDeep, isErrorCode, toBoolean, toNumber, toText } from './values.js';

/** An optional numeric argument: omitted (undefined) means `fallback`. */
const optional = (v: unknown, fallback: number): number => (v === undefined ? fallback : toNumber(v));

/** A finite result; overflow is #NUM!. */
const finite = (n: number): number => {
  if (!Number.isFinite(n)) throw new Error('#NUM!');
  return n;
};

/** A quotient cleaned to 15 significant digits (Excel's precision), so IEEE drift (0.07/0.01 = 7.000000000000001) does not move a rounding. */
const clean = (x: number): number => parseFloat(x.toPrecision(15));

/** Greatest common divisor (Euclid). */
function gcd(a: number, b: number): number {
  while (b) [a, b] = [b, a % b];
  return a;
}

/** C(n, k) for integers 0 <= k <= n. */
function choose(n: number, k: number): number {
  const j = Math.min(k, n - k);
  let r = 1;
  for (let i = 1; i <= j; i++) {
    r = (r * (n - j + i)) / i;
    if (!Number.isFinite(r)) throw new Error('#NUM!'); // stop at overflow, however large k is
  }
  return finite(Math.round(r));
}

/** Integer values of the arguments for GCD / LCM (blanks skipped, negatives #NUM!). */
function nonNegativeIntegers(args: unknown[]): number[] {
  const out: number[] = [];
  for (const v of args.flatMap((arg) => flattenDeep(arg))) {
    const n = Math.trunc(toNumber(v === undefined ? 0 : v)); // a blank is 0 (so LCM with a blank is 0)
    if (n < 0) throw new Error('#NUM!');
    out.push(n);
  }
  return out;
}

/**
 * CEILING.MATH / FLOOR.MATH style rounding of `num` to a multiple of
 * |significance|. `towardCeiling` picks the direction for the magnitude.
 */
function roundToMultiple(num: number, significance: number, towardCeiling: boolean): number {
  const s = Math.abs(significance);
  if (s === 0 || num === 0) return 0;
  const q = clean(num / s);
  return (towardCeiling ? Math.ceil(q) : Math.floor(q)) * s;
}

/**
 * Direction for CEILING.MATH / FLOOR.MATH: positive numbers round as their name
 * says; negative ones round toward +∞ (CEILING) or −∞ (FLOOR) only when mode is 0
 * for CEILING and non-zero for FLOOR, otherwise they round the other way.
 */
function mathDirection(num: number, mode: number, ceiling: boolean): boolean {
  if (num >= 0) return ceiling;
  return ceiling ? mode === 0 : mode !== 0;
}

const ROMAN_TABLE: Array<[number, string]> = [
  [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
  [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
];

const ROMAN_VALUE: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };

/** Roman numeral text (classic form) to a number; anything else is #VALUE!. */
function fromRoman(text: string): number {
  let s = text.trim().toUpperCase();
  if (s === '') return 0;
  let sign = 1;
  if (s.startsWith('-')) {
    sign = -1;
    s = s.slice(1);
  }
  if (!/^[IVXLCDM]+$/.test(s)) throw new Error('#VALUE!');
  let total = 0;
  for (let i = 0; i < s.length; i++) {
    const v = ROMAN_VALUE[s[i]];
    const next = i + 1 < s.length ? ROMAN_VALUE[s[i + 1]] : 0;
    total += next > v ? -v : v;
  }
  return sign * total;
}

/** Digits of `text` in `radix` to a number; a digit outside the radix is #NUM!. */
function fromRadix(text: string, radix: number): number {
  let total = 0;
  for (const ch of text.trim().toUpperCase()) {
    const digit = parseInt(ch, 36);
    if (Number.isNaN(digit) || digit >= radix) throw new Error('#NUM!');
    total = total * radix + digit;
  }
  return total;
}

/** Sum of `term(x, y)` over paired elements of two arrays; non-numeric pairs are skipped, mismatched shapes #N/A. */
function pairedSum(xs: unknown, ys: unknown, term: (x: number, y: number) => number): number {
  const a = asArray(xs);
  const b = asArray(ys);
  // Same number of elements: a row against a column of equal length pairs up too
  if (a.values.length !== b.values.length) throw new Error('#N/A');
  let total = 0;
  let pairs = 0;
  for (let i = 0; i < a.values.length; i++) {
    const x = a.values[i];
    const y = b.values[i];
    if (isErrorCode(x)) throw new Error(x);
    if (isErrorCode(y)) throw new Error(y);
    if (typeof x === 'number' && typeof y === 'number') {
      total += term(x, y);
      pairs++;
    }
  }
  // Excel: no pair of numbers at all is #DIV/0!
  if (pairs === 0) throw new Error('#DIV/0!');
  return total;
}

/** A matrix element: numbers as they are; a blank or anything else is #VALUE! (errors propagate). */
function matrixNumber(v: unknown): number {
  if (typeof v === 'number') return v;
  if (isErrorCode(v)) throw new Error(v);
  throw new Error('#VALUE!');
}

/** Rows of numbers from a matrix argument. */
function matrixRows(arr: RangeValue): number[][] {
  const rows: number[][] = [];
  for (let r = 0; r < arr.rows; r++) rows.push(arr.getRow(r).map(matrixNumber));
  return rows;
}

/** A square matrix argument's rows, or #VALUE!. */
function squareMatrix(arg: unknown): number[][] {
  const arr = asArray(arg);
  if (arr.rows !== arr.cols) throw new Error('#VALUE!');
  return matrixRows(arr);
}

/** Determinant by Gaussian elimination with partial pivoting. */
/** The most elementary operations a matrix function does before giving up with #NUM!, so one cell cannot freeze the sheet. */
const MAX_MATRIX_WORK = 300_000_000;
function checkMatrixWork(n: number, cubes = 1): void {
  if (n * n * n * cubes > MAX_MATRIX_WORK) throw new Error('#NUM!');
}

function determinant(m: number[][]): number {
  const n = m.length;
  checkMatrixWork(n);
  const a = m.map((row) => [...row]);
  let det = 1;
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r;
    if (a[p][c] === 0) return 0;
    if (p !== c) {
      [a[p], a[c]] = [a[c], a[p]];
      det = -det;
    }
    det *= a[c][c];
    for (let r = c + 1; r < n; r++) {
      if (a[r][c] === 0) continue; // nothing to eliminate
      const f = a[r][c] / a[c][c];
      for (let k = c; k < n; k++) a[r][k] -= f * a[c][k];
    }
  }
  return det;
}

/** Inverse by Gauss-Jordan elimination; a singular matrix is #NUM!. */
function inverse(m: number[][]): number[][] {
  const n = m.length;
  checkMatrixWork(n, 2);
  let scale = 0;
  for (const row of m) for (const v of row) scale = Math.max(scale, Math.abs(v));
  if (scale === 0) throw new Error('#NUM!');
  const a = m.map((row, i) => [...row, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r;
    if (Math.abs(a[p][c]) <= 1e-12 * scale) throw new Error('#NUM!');
    [a[p], a[c]] = [a[c], a[p]];
    const pivot = a[c][c];
    for (let k = 0; k < 2 * n; k++) a[c][k] /= pivot;
    for (let r = 0; r < n; r++) {
      if (r === c || a[r][c] === 0) continue;
      const f = a[r][c];
      for (let k = 0; k < 2 * n; k++) a[r][k] -= f * a[c][k];
    }
  }
  return a.map((row) => row.slice(n));
}

/** Sample (n - 1) or population (n) variance; too few values is #DIV/0!. */
function variance(nums: number[], sample: boolean): number {
  const n = nums.length;
  if (n < (sample ? 2 : 1)) throw new Error('#DIV/0!');
  const mean = nums.reduce((a, b) => a + b, 0) / n;
  return nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (sample ? n - 1 : n);
}

/** One item of an aggregate's arguments: `direct` is an argument typed as a value, not a value taken from a range. */
interface AggItem {
  v: unknown;
  direct: boolean;
}

/** Items of aggregate arguments: ranges flattened with blanks dropped, direct values kept. */
function aggItems(args: unknown[]): AggItem[] {
  const items: AggItem[] = [];
  for (const arg of args) {
    if (arg instanceof RangeValue) {
      for (const v of flattenDeep(arg)) if (v !== undefined) items.push({ v, direct: false });
    } else if (arg instanceof ArgumentError) {
      items.push({ v: arg.code, direct: true });
    } else if (arg !== undefined) {
      items.push({ v: arg, direct: true });
    }
  }
  return items;
}

/** Numbers among aggregate items: numbers, and direct booleans and numeric text (text in ranges is ignored). */
function numbersOf(items: AggItem[]): number[] {
  const nums: number[] = [];
  for (const { v, direct } of items) {
    if (typeof v === 'number') nums.push(v);
    else if (direct && typeof v === 'boolean') nums.push(v ? 1 : 0);
    else if (direct && typeof v === 'string' && v.trim() !== '' && !isErrorCode(v) && !isNaN(Number(v))) {
      nums.push(Number(v));
    }
  }
  return nums;
}

/** The numbers of aggregate arguments; an error among them is thrown unless errors are ignored. */
function collectNumbers(args: unknown[], ignoreErrors: boolean): number[] {
  const items = aggItems(args);
  if (!ignoreErrors) {
    const error = items.find((i) => isErrorCode(i.v));
    if (error) throw new Error(error.v as string);
  }
  return numbersOf(items);
}

/** SUBTOTAL / AGGREGATE functions 1-11 over the arguments. */
function summarize(fn: number, args: unknown[], ignoreErrors: boolean): unknown {
  const items = aggItems(args);
  // COUNT never fails on an error; COUNTA counts errors unless they are ignored
  if (fn === 2) return numbersOf(items).length;
  if (fn === 3) return (ignoreErrors ? items.filter((i) => !isErrorCode(i.v)) : items).length;
  const nums = collectNumbers(args, ignoreErrors);
  switch (fn) {
    case 1:
      if (nums.length === 0) throw new Error('#DIV/0!');
      return nums.reduce((a, b) => a + b, 0) / nums.length;
    case 4:
      return nums.length ? nums.reduce((a, b) => Math.max(a, b)) : 0;
    case 5:
      return nums.length ? nums.reduce((a, b) => Math.min(a, b)) : 0;
    case 6:
      return nums.length ? nums.reduce((p, n) => p * n, 1) : 0;
    case 7:
      return Math.sqrt(variance(nums, true));
    case 8:
      return Math.sqrt(variance(nums, false));
    case 9:
      return nums.reduce((a, b) => a + b, 0);
    case 10:
      return variance(nums, true);
    case 11:
      return variance(nums, false);
    default:
      throw new Error('#VALUE!');
  }
}

function medianOf(nums: number[]): number {
  if (nums.length === 0) throw new Error('#NUM!');
  const s = [...nums].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

/** Most frequent value (the first to reach the top count); no repeats is #N/A. */
function modeOf(nums: number[]): number {
  const counts = new Map<number, number>();
  for (const n of nums) counts.set(n, (counts.get(n) ?? 0) + 1);
  let top = 0;
  for (const c of counts.values()) top = Math.max(top, c);
  if (top < 2) throw new Error('#N/A');
  return nums.find((n) => counts.get(n) === top) as number;
}

/**
 * Percentile of the numbers. Inclusive uses rank k(n-1); exclusive uses k(n+1)
 * (1-based) and needs k between 1/(n+1) and n/(n+1). Otherwise #NUM!.
 */
function percentile(nums: number[], k: number, exclusive: boolean): number {
  const s = [...nums].sort((a, b) => a - b);
  const n = s.length;
  if (n === 0) throw new Error('#NUM!');
  let idx: number;
  if (exclusive) {
    if (k < 1 / (n + 1) || k > n / (n + 1)) throw new Error('#NUM!');
    idx = k * (n + 1) - 1;
  } else {
    if (k < 0 || k > 1) throw new Error('#NUM!');
    idx = k * (n - 1);
  }
  idx = Math.max(0, Math.min(n - 1, idx));
  const lo = Math.floor(idx);
  const frac = idx - lo;
  return lo + 1 < n ? s[lo] + frac * (s[lo + 1] - s[lo]) : s[lo];
}

export function registerMathBuiltins(reg: Registrar): void {
  const { def } = reg;

  /** A one-argument numeric function. */
  const unary = (name: string, fn: (x: number) => number) =>
    def(name, 'n', (_ctx, v) => fn(toNumber(v)), { arity: [1, 1] });

  // ─── Sign, powers and logarithms ─────────────────────

  unary('SIGN', (x) => (x > 0 ? 1 : x < 0 ? -1 : 0));
  unary('EXP', (x) => finite(Math.exp(x)));
  unary('LN', (x) => {
    if (x <= 0) throw new Error('#NUM!');
    return Math.log(x);
  });
  unary('LOG10', (x) => {
    if (x <= 0) throw new Error('#NUM!');
    return Math.log10(x);
  });
  def(
    'LOG',
    'nn',
    (_ctx, number, base) => {
      const n = toNumber(number);
      const b = optional(base, 10);
      if (n <= 0 || b <= 0) throw new Error('#NUM!');
      if (b === 1) throw new Error('#DIV/0!');
      if (b === 10) return Math.log10(n);
      if (b === 2) return Math.log2(n);
      return Math.log(n) / Math.log(b);
    },
    { arity: [1, 2] }
  );
  unary('SQRTPI', (x) => {
    if (x < 0) throw new Error('#NUM!');
    return Math.sqrt(x * Math.PI);
  });

  // ─── Constants and angles ────────────────────────────

  def('PI', 'n', () => Math.PI, { arity: [0, 0] });
  unary('DEGREES', (x) => (x * 180) / Math.PI);
  unary('RADIANS', (x) => (x * Math.PI) / 180);

  // ─── Trigonometry ────────────────────────────────────

  unary('SIN', Math.sin);
  unary('COS', Math.cos);
  unary('TAN', Math.tan);
  unary('COT', (x) => {
    const t = Math.tan(x);
    if (t === 0) throw new Error('#DIV/0!');
    return 1 / t;
  });
  unary('ASIN', (x) => {
    if (Math.abs(x) > 1) throw new Error('#NUM!');
    return Math.asin(x);
  });
  unary('ACOS', (x) => {
    if (Math.abs(x) > 1) throw new Error('#NUM!');
    return Math.acos(x);
  });
  unary('ATAN', Math.atan);
  // ATAN2(x, y): the angle of the point (x, y) from the x-axis
  def(
    'ATAN2',
    'nn',
    (_ctx, x, y) => {
      const xv = toNumber(x);
      const yv = toNumber(y);
      if (xv === 0 && yv === 0) throw new Error('#DIV/0!');
      return Math.atan2(yv, xv);
    },
    { arity: [2, 2] }
  );
  unary('SINH', (x) => finite(Math.sinh(x)));
  unary('COSH', (x) => finite(Math.cosh(x)));
  unary('TANH', Math.tanh);
  unary('ASINH', Math.asinh);
  unary('ACOSH', (x) => {
    if (x < 1) throw new Error('#NUM!');
    return Math.acosh(x);
  });
  unary('ATANH', (x) => {
    if (Math.abs(x) >= 1) throw new Error('#NUM!');
    return Math.atanh(x);
  });

  // ─── Factorials and combinatorics ────────────────────

  unary('FACT', (x) => {
    const k = Math.trunc(x);
    if (k < 0 || k > 170) throw new Error('#NUM!');
    let r = 1;
    for (let i = 2; i <= k; i++) r *= i;
    return finite(r);
  });
  unary('FACTDOUBLE', (x) => {
    const k = Math.trunc(x);
    if (k < 0 || k > 300) throw new Error('#NUM!');
    let r = 1;
    for (let i = k; i > 1; i -= 2) r *= i;
    return finite(r);
  });
  def(
    'COMBIN',
    'nn',
    (_ctx, number, chosen) => {
      const n = Math.trunc(toNumber(number));
      const k = Math.trunc(toNumber(chosen));
      if (n < 0 || k < 0 || k > n) throw new Error('#NUM!');
      return choose(n, k);
    },
    { arity: [2, 2] }
  );
  // COMBINA(n, k) = C(n + k - 1, k): combinations with repetition
  def(
    'COMBINA',
    'nn',
    (_ctx, number, chosen) => {
      const n = Math.trunc(toNumber(number));
      const k = Math.trunc(toNumber(chosen));
      if (n < 0 || k < 0 || (n === 0 && k > 0)) throw new Error('#NUM!');
      if (n === 0 && k === 0) return 1;
      return choose(n + k - 1, k);
    },
    { arity: [2, 2] }
  );
  def(
    'PERMUT',
    'nn',
    (_ctx, number, chosen) => {
      const n = Math.trunc(toNumber(number));
      const k = Math.trunc(toNumber(chosen));
      if (n < 0 || k < 0 || k > n) throw new Error('#NUM!');
      let r = 1;
      for (let i = 0; i < k; i++) {
        r *= n - i;
        if (!Number.isFinite(r)) throw new Error('#NUM!');
      }
      return finite(r);
    },
    { arity: [2, 2] }
  );

  // ─── Integers, rounding and multiples ────────────────

  def('GCD', 'a', (_ctx, ...args) => nonNegativeIntegers(args).reduce((g, n) => gcd(g, n), 0), {
    arity: [1, 255],
  });
  def(
    'LCM',
    'a',
    (_ctx, ...args) => {
      const nums = nonNegativeIntegers(args);
      if (nums.some((n) => n === 0)) return 0;
      return finite(nums.reduce((l, n) => (l / gcd(l, n)) * n, 1));
    },
    { arity: [1, 255] }
  );
  unary('EVEN', (x) => {
    if (x === 0) return 0;
    return Math.sign(x) * Math.ceil(clean(Math.abs(x) / 2)) * 2;
  });
  unary('ODD', (x) => {
    const a = Math.abs(x);
    const sign = x < 0 ? -1 : 1;
    // smallest odd integer at or above |x|
    return sign * (2 * Math.ceil(clean((a - 1) / 2)) + 1);
  });
  def(
    'TRUNC',
    'nn',
    (_ctx, num, digits) => {
      const n = toNumber(num);
      const factor = Math.pow(10, Math.trunc(optional(digits, 0)));
      const shifted = parseFloat((Math.abs(n) * factor).toPrecision(15));
      return (Math.sign(n) * Math.floor(shifted)) / factor;
    },
    { arity: [1, 2] }
  );
  def(
    'MROUND',
    'nn',
    (_ctx, num, multiple) => {
      const n = toNumber(num);
      const m = toNumber(multiple);
      if (m === 0 || n === 0) return 0;
      if ((n < 0) !== (m < 0)) throw new Error('#NUM!');
      const mm = Math.abs(m);
      // halves round away from zero, as Excel does, but only exact halves
      return Math.sign(n) * Math.round(Math.abs(n) / mm) * mm; // the plain quotient: 6.05 / 0.1 is just under 60.5
    },
    { arity: [2, 2] }
  );
  def(
    'QUOTIENT',
    'nn',
    (_ctx, numerator, denominator) => {
      const d = toNumber(denominator);
      if (d === 0) throw new Error('#DIV/0!');
      return Math.trunc(toNumber(numerator) / d); // plain quotient: 0.3 / 0.1 is just under 3, so 2
    },
    { arity: [2, 2] }
  );
  def(
    'CEILING.MATH',
    'nnn',
    (_ctx, num, significance, mode) => {
      const n = toNumber(num);
      const m = optional(mode, 0);
      return roundToMultiple(n, optional(significance, 1), mathDirection(n, m, true));
    },
    { arity: [1, 3] }
  );
  def(
    'FLOOR.MATH',
    'nnn',
    (_ctx, num, significance, mode) => {
      const n = toNumber(num);
      const m = optional(mode, 0);
      return roundToMultiple(n, optional(significance, 1), mathDirection(n, m, false));
    },
    { arity: [1, 3] }
  );
  // The PRECISE forms always round toward +∞ / −∞, whatever the sign
  def(
    'CEILING.PRECISE',
    'nn',
    (_ctx, num, significance) => roundToMultiple(toNumber(num), optional(significance, 1), true),
    { arity: [1, 2] }
  );
  def(
    'FLOOR.PRECISE',
    'nn',
    (_ctx, num, significance) => roundToMultiple(toNumber(num), optional(significance, 1), false),
    { arity: [1, 2] }
  );

  // ─── Sums of squares and products of pairs ───────────

  reg.agg('SUMSQ', (args) => numericArgs(args).reduce((s, n) => s + n * n, 0), { propagate: true });
  def('SUMX2MY2', 'aa', (_ctx, xs, ys) => pairedSum(xs, ys, (x, y) => x * x - y * y), { arity: [2, 2] });
  def('SUMX2PY2', 'aa', (_ctx, xs, ys) => pairedSum(xs, ys, (x, y) => x * x + y * y), { arity: [2, 2] });
  def('SUMXMY2', 'aa', (_ctx, xs, ys) => pairedSum(xs, ys, (x, y) => (x - y) ** 2), { arity: [2, 2] });

  // ─── Roman numerals, bases ───────────────────────────

  // ROMAN(number, form): only the classic form (0) is implemented; the concise forms 1-4 are #VALUE! for now.
  def(
    'ROMAN',
    'nn',
    (_ctx, number, form) => {
      const n = Math.trunc(toNumber(number));
      const f = optional(form, 0);
      if (f !== 0 || n < 0 || n > 3999) throw new Error('#VALUE!');
      let rest = n;
      let out = '';
      for (const [value, letters] of ROMAN_TABLE) {
        while (rest >= value) {
          out += letters;
          rest -= value;
        }
      }
      return out;
    },
    { arity: [1, 2] }
  );
  def('ARABIC', 't', (_ctx, text) => fromRoman(toText(text)), { arity: [1, 1] });
  def(
    'BASE',
    'nnn',
    (_ctx, number, radix, minLength) => {
      const n = Math.trunc(toNumber(number));
      const r = Math.trunc(toNumber(radix));
      const width = optional(minLength, 0);
      if (n < 0 || n >= 2 ** 53) throw new Error('#NUM!');
      if (r < 2 || r > 36) throw new Error('#NUM!');
      if (width < 0 || width > 255) throw new Error('#NUM!');
      return n.toString(r).toUpperCase().padStart(Math.trunc(width), '0');
    },
    { arity: [2, 3] }
  );
  def(
    'DECIMAL',
    'tn',
    (_ctx, text, radix) => {
      const r = Math.trunc(toNumber(radix));
      if (r < 2 || r > 36) throw new Error('#NUM!');
      return fromRadix(toText(text), r);
    },
    { arity: [2, 2] }
  );

  // ─── Matrices ────────────────────────────────────────

  def(
    'MMULT',
    'aa',
    (_ctx, a, b) => {
      const A = asArray(a);
      const B = asArray(b);
      if (A.cols !== B.rows) throw new Error('#VALUE!');
      checkArraySize(A.rows, B.cols);
      if (A.rows * A.cols * B.cols > MAX_MATRIX_WORK) throw new Error('#NUM!');
      const left = matrixRows(A);
      const right = matrixRows(B);
      const values: unknown[] = [];
      for (let r = 0; r < A.rows; r++) {
        for (let c = 0; c < B.cols; c++) {
          let sum = 0;
          for (let k = 0; k < A.cols; k++) sum += left[r][k] * right[k][c];
          values.push(sum);
        }
      }
      return new RangeValue(values, A.rows, B.cols);
    },
    { arity: [2, 2] }
  );
  def('MDETERM', 'a', (_ctx, array) => determinant(squareMatrix(array)), { arity: [1, 1] });
  def(
    'MINVERSE',
    'a',
    (_ctx, array) => {
      const inv = inverse(squareMatrix(array));
      return new RangeValue(inv.flat(), inv.length, inv.length);
    },
    { arity: [1, 1] }
  );
  def(
    'MUNIT',
    'n',
    (_ctx, dimension) => {
      const d = Math.trunc(toNumber(dimension));
      if (d < 1) throw new Error('#VALUE!');
      checkArraySize(d, d);
      const values: unknown[] = [];
      for (let r = 0; r < d; r++) for (let c = 0; c < d; c++) values.push(r === c ? 1 : 0);
      return new RangeValue(values, d, d);
    },
    { arity: [1, 1] }
  );

  // SERIESSUM(x, n, m, coefficients) = Σ a_i · x^(n + i·m)
  def(
    'SERIESSUM',
    'nnna',
    (_ctx, x, n, m, coefficients) => {
      const base = toNumber(x);
      const start = toNumber(n);
      const step = toNumber(m);
      let total = 0;
      flattenDeep(coefficients).forEach((c, i) => {
        if (c === undefined) return;
        if (isErrorCode(c)) throw new Error(c);
        if (typeof c !== 'number') throw new Error('#VALUE!');
        // 0 to a negative power is #NUM! here, not #DIV/0!
        const exponent = start + i * step;
        if (base === 0 && exponent < 0) throw new Error('#NUM!');
        total += c * power(base, exponent);
      });
      return total;
    },
    { arity: [4, 4] }
  );

  // ─── Random (volatile) ───────────────────────────────

  def('RAND', 'n', () => Math.random(), { arity: [0, 0], volatile: true });
  def(
    'RANDBETWEEN',
    'nn',
    (_ctx, bottom, top) => {
      const lo = Math.ceil(toNumber(bottom));
      const hi = Math.floor(toNumber(top));
      if (lo > hi) throw new Error('#NUM!');
      return lo + Math.floor(Math.random() * (hi - lo + 1));
    },
    { arity: [2, 2], volatile: true }
  );
  def(
    'RANDARRAY',
    'nnnnn',
    (_ctx, rows, cols, min, max, integer) => {
      const r = Math.trunc(optional(rows, 1));
      const c = Math.trunc(optional(cols, 1));
      const lo = optional(min, 0);
      const hi = optional(max, 1);
      if (r < 1 || c < 1 || lo > hi) throw new Error('#VALUE!');
      const whole = integer !== undefined && toBoolean(integer);
      const a = whole ? Math.ceil(lo) : lo;
      const b = whole ? Math.floor(hi) : hi;
      if (a > b) throw new Error('#VALUE!');
      checkArraySize(r, c);
      const values: unknown[] = [];
      for (let i = 0; i < r * c; i++) {
        values.push(whole ? a + Math.floor(Math.random() * (b - a + 1)) : a + Math.random() * (b - a));
      }
      return new RangeValue(values, r, c);
    },
    { arity: [0, 5], volatile: true }
  );

  // ─── SUBTOTAL and AGGREGATE ──────────────────────────

  // SUBTOTAL: 1-11 and 101-111 are the same here; no row is hidden by hand, so nothing is skipped.
  def(
    'SUBTOTAL',
    'nb',
    (_ctx, functionNum, ...args) => {
      const fn = Math.trunc(toNumber(functionNum));
      if (!((fn >= 1 && fn <= 11) || (fn >= 101 && fn <= 111))) throw new Error('#VALUE!');
      return summarize(fn % 100, args, false);
    },
    { arity: [2, 254] }
  );

  // AGGREGATE(function_num, options, ...): options 2, 3, 6, 7 ignore error values.
  // Hidden rows are treated as visible, so options 1, 3, 5, 7 ignore nothing more.
  def(
    'AGGREGATE',
    'nnx',
    (_ctx, functionNum, option, ...args) => {
      const fn = Math.trunc(toNumber(functionNum));
      const opt = Math.trunc(toNumber(option));
      if (fn < 1 || fn > 19 || opt < 0 || opt > 7) throw new Error('#VALUE!');
      const ignoreErrors = opt === 2 || opt === 3 || opt === 6 || opt === 7;
      if (fn <= 11) return summarize(fn, args, ignoreErrors);
      if (fn === 12) return medianOf(collectNumbers(args, ignoreErrors));
      if (fn === 13) return modeOf(collectNumbers(args, ignoreErrors));

      // 14-19 take an array and one more argument (k, or quart)
      if (args.length !== 2) throw new Error('#VALUE!');
      const [array, kArg] = args;
      if (kArg instanceof ArgumentError) throw new Error(kArg.code);
      const nums = collectNumbers([array], ignoreErrors);
      const pick = (kValue: unknown): unknown => {
      const k = toNumber(kValue);
      switch (fn) {
        case 14: {
          const kk = Math.trunc(k);
          if (kk < 1 || kk > nums.length) throw new Error('#NUM!');
          return [...nums].sort((a, b) => b - a)[kk - 1];
        }
        case 15: {
          const kk = Math.trunc(k);
          if (kk < 1 || kk > nums.length) throw new Error('#NUM!');
          return [...nums].sort((a, b) => a - b)[kk - 1];
        }
        case 16:
          return percentile(nums, k, false);
        case 17: {
          const q = Math.trunc(k);
          if (q < 0 || q > 4) throw new Error('#NUM!');
          return percentile(nums, q / 4, false);
        }
        case 18:
          return percentile(nums, k, true);
        case 19: {
          const q = Math.trunc(k);
          if (q < 1 || q > 3) throw new Error('#NUM!');
          return percentile(nums, q / 4, true);
        }
        default:
          throw new Error('#VALUE!');
      }
      };
      // An array of k (or quart) values gives one result per element
      if (kArg instanceof RangeValue) {
        return new RangeValue(
          kArg.values.map((kv) => {
            try {
              return pick(kv);
            } catch (e) {
              return (e as Error).message;
            }
          }),
          kArg.rows,
          kArg.cols
        );
      }
      return pick(kArg);
    },
    { arity: [3, 255] }
  );
}
