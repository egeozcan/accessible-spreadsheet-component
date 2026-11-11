/** Statistics built-ins (values only; lazy or reference-taking functions live in formula-engine.ts). */
import { type Registrar, asArray, cellMatchesCriteria, numericArgs } from './functions.js';
import { RangeValue, flattenDeep, isErrorCode, toNumber } from './values.js';

/** Numbers in range or array arguments: text, booleans and blanks are skipped, an error is thrown. */
function rangeNumbers(...args: unknown[]): number[] {
  const out: number[] = [];
  for (const arg of args) {
    for (const v of flattenDeep(asArray(arg))) {
      if (isErrorCode(v)) throw new Error(v);
      if (typeof v === 'number') out.push(v);
    }
  }
  return out;
}

/** What the A-variants count: numbers as is, TRUE as 1, text and FALSE as 0. */
function aValue(v: unknown): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  return 0;
}

function ascending(xs: number[]): number[] {
  return [...xs].sort((a, b) => a - b);
}

/**
 * The numbers of a reference or array in ascending order. A range value is immutable and shared
 * within one evaluation, so queries lifted over an array (LARGE(r, SEQUENCE(n))) sort it once.
 */
const sortedCache = new WeakMap<object, number[]>();
function sortedNumbers(arg: unknown): number[] {
  if (arg instanceof RangeValue) {
    const cached = sortedCache.get(arg);
    if (cached) return cached;
    const sorted = ascending(rangeNumbers(arg));
    sortedCache.set(arg, sorted);
    return sorted;
  }
  return ascending(rangeNumbers(arg));
}

/** Number of elements of the ascending list that are < x (lower bound) or <= x (upper bound). */
function countBelow(sorted: number[], x: number, inclusive: boolean): number {
  let lo = 0;
  let hi = sorted.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (inclusive ? sorted[mid] <= x : sorted[mid] < x) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

function mean(xs: number[]): number {
  return xs.reduce((a, b) => a + b, 0) / xs.length;
}

function sumOf(xs: number[], f: (x: number) => number): number {
  return xs.reduce((s, x) => s + f(x), 0);
}

/** Variance around the mean: sample (n - 1) or population (n). */
function varianceOf(xs: number[], sample: boolean): number {
  if (xs.length < (sample ? 2 : 1)) throw new Error('#DIV/0!');
  const m = mean(xs);
  return sumOf(xs, (x) => (x - m) ** 2) / (sample ? xs.length - 1 : xs.length);
}

/** Pairs of numbers from two arrays of equal size; a position with a non-number in either is dropped. */
function pairedNumbers(first: unknown, second: unknown): [number[], number[]] {
  const a = flattenDeep(asArray(first));
  const b = flattenDeep(asArray(second));
  if (a.length !== b.length) throw new Error('#N/A');
  const as: number[] = [];
  const bs: number[] = [];
  for (let i = 0; i < a.length; i++) {
    if (isErrorCode(a[i])) throw new Error(a[i] as string);
    if (isErrorCode(b[i])) throw new Error(b[i] as string);
    if (typeof a[i] === 'number' && typeof b[i] === 'number') {
      as.push(a[i] as number);
      bs.push(b[i] as number);
    }
  }
  return [as, bs];
}

/** Least-squares line through (xs, ys). */
function regression(ys: number[], xs: number[]): { slope: number; intercept: number } {
  if (xs.length === 0) throw new Error('#DIV/0!');
  const mx = mean(xs);
  const my = mean(ys);
  const sxx = sumOf(xs, (x) => (x - mx) ** 2);
  if (sxx === 0) throw new Error('#DIV/0!');
  let sxy = 0;
  for (let i = 0; i < xs.length; i++) sxy += (xs[i] - mx) * (ys[i] - my);
  const slope = sxy / sxx;
  return { slope, intercept: my - slope * mx };
}

function covariance(a: number[], b: number[], sample: boolean): number {
  const n = a.length;
  if (n === 0 || (sample && n < 2)) throw new Error('#DIV/0!');
  const ma = mean(a);
  const mb = mean(b);
  let s = 0;
  for (let i = 0; i < n; i++) s += (a[i] - ma) * (b[i] - mb);
  return s / (sample ? n - 1 : n);
}

function pearson(a: number[], b: number[]): number {
  if (a.length < 2) throw new Error('#DIV/0!');
  const ma = mean(a);
  const mb = mean(b);
  let sab = 0;
  let saa = 0;
  let sbb = 0;
  for (let i = 0; i < a.length; i++) {
    sab += (a[i] - ma) * (b[i] - mb);
    saa += (a[i] - ma) ** 2;
    sbb += (b[i] - mb) ** 2;
  }
  if (saa === 0 || sbb === 0) throw new Error('#DIV/0!');
  return sab / Math.sqrt(saa * sbb);
}

/** Rank of a number among the numbers of a reference; `avg` gives tied values the average of their ranks. */
function rankOf(value: unknown, ref: unknown, order: unknown, avg: boolean): number {
  const x = toNumber(value);
  const sorted = sortedNumbers(ref);
  const below = countBelow(sorted, x, false);
  const upTo = countBelow(sorted, x, true);
  const ties = upTo - below;
  if (ties === 0) throw new Error('#N/A');
  const ascendingOrder = order !== undefined && toNumber(order) !== 0;
  const beats = ascendingOrder ? below : sorted.length - upTo;
  return avg ? beats + (ties + 1) / 2 : beats + 1;
}

/** Inclusive percentile (PERCENTILE.INC) of ascending values, k in [0, 1]. */
function percentileInc(sorted: number[], k: number): number {
  const rank = k * (sorted.length - 1);
  const lo = Math.floor(rank);
  if (lo >= sorted.length - 1) return sorted[sorted.length - 1];
  return sorted[lo] + (rank - lo) * (sorted[lo + 1] - sorted[lo]);
}

/** Exclusive percentile (PERCENTILE.EXC) of ascending values; the rank must fall inside the list. */
function percentileExc(sorted: number[], k: number): number {
  const rank = k * (sorted.length + 1);
  if (rank < 1 || rank > sorted.length) throw new Error('#NUM!');
  const lo = Math.floor(rank) - 1;
  if (lo >= sorted.length - 1) return sorted[sorted.length - 1];
  return sorted[lo] + (rank - 1 - lo) * (sorted[lo + 1] - sorted[lo]);
}

function percentRank(sorted: number[], x: number, significance: number, exclusive: boolean): number {
  if (sorted.length === 0) throw new Error('#N/A');
  const n = sorted.length;
  if (x < sorted[0] || x > sorted[n - 1]) throw new Error('#N/A');
  const less = (v: number) => countBelow(sorted, v, false);
  const rankOfValue = (v: number): number => {
    if (exclusive) return (less(v) + 1) / (n + 1);
    return n === 1 ? 1 : less(v) / (n - 1);
  };
  // Below a gap the rank is that of the last of any tied values
  const rankOfLast = (v: number): number => {
    const upTo = countBelow(sorted, v, true);
    if (exclusive) return upTo / (n + 1);
    return n === 1 ? 1 : (upTo - 1) / (n - 1);
  };
  let rank: number;
  const first = less(x);
  if (first < n && sorted[first] === x) {
    rank = rankOfValue(x);
  } else {
    const lo = sorted[first - 1]; // the largest value below x
    const hi = sorted[countBelow(sorted, x, true)]; // the smallest above
    rank = rankOfLast(lo) + ((x - lo) / (hi - lo)) * (rankOfValue(hi) - rankOfLast(lo));
  }
  // Truncated (not rounded) to `significance` digits; the epsilon keeps 0.29 * 100 from losing a digit.
  const factor = 10 ** significance;
  return Math.floor(rank * factor + 1e-9) / factor;
}

/** Values that occur most often (numbers only), in order of first appearance; #N/A if none repeats. */
function modesOf(xs: number[]): number[] {
  const counts = new Map<number, number>();
  for (const x of xs) counts.set(x, (counts.get(x) ?? 0) + 1);
  let best = 0;
  for (const c of counts.values()) best = Math.max(best, c);
  if (best < 2) throw new Error('#N/A');
  return [...counts].filter(([, c]) => c === best).map(([v]) => v);
}

/** Values of a range (or a single value) that satisfy every (criteria_range, criteria) pair. */
function matchingValues(target: unknown, pairs: unknown[]): unknown[] {
  if (pairs.length === 0 || pairs.length % 2 !== 0) throw new Error('#VALUE!');
  const shape = asArray(target);
  const keep = shape.values.map(() => true);
  for (let p = 0; p < pairs.length; p += 2) {
    const criteriaRange = asArray(pairs[p]);
    if (criteriaRange.rows !== shape.rows || criteriaRange.cols !== shape.cols) throw new Error('#VALUE!');
    const criteria = String(pairs[p + 1]);
    criteriaRange.values.forEach((v, i) => {
      if (keep[i] && !cellMatchesCriteria(v, criteria)) keep[i] = false;
    });
  }
  return shape.values.filter((_, i) => keep[i]);
}

/** Largest (or smallest) numeric value among the values matching the criteria; 0 if none. */
function conditionalExtreme(pick: (a: number, b: number) => number) {
  return (_ctx: unknown, target: unknown, ...pairs: unknown[]): number => {
    let result: number | undefined;
    for (const v of matchingValues(target, pairs)) {
      if (isErrorCode(v)) throw new Error(v);
      if (typeof v === 'number') result = result === undefined ? v : pick(result, v);
    }
    return result ?? 0;
  };
}

export function registerStatsBuiltins(reg: Registrar): void {
  const { def, agg } = reg;

  // ─── Counting and extremes ──────────────────────────

  // COUNTBLANK is a special form of the engine: it counts the blanks of a whole column or row too

  def('MAXIFS', 'b|bn', conditionalExtreme(Math.max));
  def('MINIFS', 'b|bn', conditionalExtreme(Math.min));

  // A-variants: text counts as 0, TRUE as 1, FALSE as 0 (in ranges and direct arguments alike).
  agg('MAXA', (args) => (args.length ? args.map(aValue).reduce((a, b) => Math.max(a, b)) : 0), { propagate: true });
  agg('MINA', (args) => (args.length ? args.map(aValue).reduce((a, b) => Math.min(a, b)) : 0), { propagate: true });
  agg(
    'AVERAGEA',
    (args) => {
      if (args.length === 0) throw new Error('#DIV/0!');
      return mean(args.map(aValue));
    },
    { propagate: true }
  );

  // ─── Rank and order statistics ──────────────────────

  def('LARGE', 'an', (_ctx, array, k) => {
    const sorted = sortedNumbers(array);
    const index = Math.trunc(toNumber(k));
    if (index < 1 || index > sorted.length) throw new Error('#NUM!');
    return sorted[sorted.length - index];
  }, { arity: [2, 2] });

  def('SMALL', 'an', (_ctx, array, k) => {
    const sorted = sortedNumbers(array);
    const index = Math.trunc(toNumber(k));
    if (index < 1 || index > sorted.length) throw new Error('#NUM!');
    return sorted[index - 1];
  }, { arity: [2, 2] });

  def('RANK', 'nan', (_ctx, number, ref, order) => rankOf(number, ref, order, false), { arity: [2, 3] });
  def('RANK.EQ', 'nan', (_ctx, number, ref, order) => rankOf(number, ref, order, false), { arity: [2, 3] });
  def('RANK.AVG', 'nan', (_ctx, number, ref, order) => rankOf(number, ref, order, true), { arity: [2, 3] });

  const percentileIncOf = (array: unknown, k: unknown): number => {
    const p = toNumber(k);
    const sorted = sortedNumbers(array);
    if (p < 0 || p > 1 || sorted.length === 0) throw new Error('#NUM!');
    return percentileInc(sorted, p);
  };
  const percentileExcOf = (array: unknown, k: unknown): number => {
    const p = toNumber(k);
    const sorted = sortedNumbers(array);
    if (p <= 0 || p >= 1 || sorted.length === 0) throw new Error('#NUM!');
    return percentileExc(sorted, p);
  };
  const quartileIncOf = (array: unknown, quart: unknown): number => {
    const q = Math.trunc(toNumber(quart));
    const sorted = sortedNumbers(array);
    if (q < 0 || q > 4 || sorted.length === 0) throw new Error('#NUM!');
    return percentileInc(sorted, q / 4);
  };
  const quartileExcOf = (array: unknown, quart: unknown): number => {
    const q = Math.trunc(toNumber(quart));
    const sorted = sortedNumbers(array);
    if (q < 1 || q > 3 || sorted.length === 0) throw new Error('#NUM!');
    return percentileExc(sorted, q / 4);
  };

  for (const name of ['PERCENTILE', 'PERCENTILE.INC']) {
    def(name, 'an', (_ctx, array, k) => percentileIncOf(array, k), { arity: [2, 2] });
  }
  def('PERCENTILE.EXC', 'an', (_ctx, array, k) => percentileExcOf(array, k), { arity: [2, 2] });
  for (const name of ['QUARTILE', 'QUARTILE.INC']) {
    def(name, 'an', (_ctx, array, quart) => quartileIncOf(array, quart), { arity: [2, 2] });
  }
  def('QUARTILE.EXC', 'an', (_ctx, array, quart) => quartileExcOf(array, quart), { arity: [2, 2] });

  const percentRankOf = (exclusive: boolean) =>
    (_ctx: unknown, array: unknown, x: unknown, significance?: unknown): number => {
      const digits = significance === undefined ? 3 : Math.trunc(toNumber(significance));
      if (digits < 1) throw new Error('#NUM!');
      return percentRank(sortedNumbers(array), toNumber(x), digits, exclusive);
    };
  def('PERCENTRANK', 'ann', percentRankOf(false), { arity: [2, 3] });
  def('PERCENTRANK.INC', 'ann', percentRankOf(false), { arity: [2, 3] });
  def('PERCENTRANK.EXC', 'ann', percentRankOf(true), { arity: [2, 3] });

  // ─── Mode and frequency ─────────────────────────────

  def('MODE', 'a', (_ctx, ...args) => modesOf(rangeNumbers(...args))[0]);
  def('MODE.SNGL', 'a', (_ctx, ...args) => modesOf(rangeNumbers(...args))[0]);
  def('MODE.MULT', 'a', (_ctx, ...args) => RangeValue.column(modesOf(rangeNumbers(...args))));

  // Counts per bin: one more result than bins, a vertical array.
  def('FREQUENCY', 'aa', (_ctx, data, bins) => {
    const values = rangeNumbers(data);
    const edges = ascending(rangeNumbers(bins));
    const counts = new Array<number>(edges.length + 1).fill(0);
    for (const x of values) {
      // First bin whose edge is >= x (binary search; the edges are sorted)
      let lo = 0;
      let hi = edges.length;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (x > edges[mid]) lo = mid + 1;
        else hi = mid;
      }
      counts[lo]++;
    }
    return RangeValue.column(counts);
  });

  // ─── Descriptive statistics ─────────────────────────

  agg('VARA', (args) => varianceOf(args.map(aValue), true), { propagate: true });
  agg('VARPA', (args) => varianceOf(args.map(aValue), false), { propagate: true });
  agg('STDEVA', (args) => Math.sqrt(varianceOf(args.map(aValue), true)), { propagate: true });
  agg('STDEVPA', (args) => Math.sqrt(varianceOf(args.map(aValue), false)), { propagate: true });

  agg(
    'GEOMEAN',
    (args) => {
      const nums = numericArgs(args);
      if (nums.length === 0 || nums.some((n) => n <= 0)) throw new Error('#NUM!');
      return Math.exp(sumOf(nums, Math.log) / nums.length);
    },
    { propagate: true }
  );

  agg(
    'HARMEAN',
    (args) => {
      const nums = numericArgs(args);
      if (nums.length === 0) throw new Error('#N/A');
      if (nums.some((n) => n <= 0)) throw new Error('#NUM!');
      return nums.length / sumOf(nums, (n) => 1 / n);
    },
    { propagate: true }
  );

  def('TRIMMEAN', 'an', (_ctx, array, percent) => {
    const p = toNumber(percent);
    const sorted = ascending(rangeNumbers(array));
    if (p < 0 || p >= 1 || sorted.length === 0) throw new Error('#NUM!');
    // The excluded count is n * percent rounded down to an even number, split between both ends.
    const excluded = Math.floor(sorted.length * p + 1e-9);
    const each = Math.floor(excluded / 2);
    return mean(sorted.slice(each, sorted.length - each));
  }, { arity: [2, 2] });

  agg(
    'AVEDEV',
    (args) => {
      const nums = numericArgs(args);
      if (nums.length === 0) throw new Error('#NUM!');
      const m = mean(nums);
      return sumOf(nums, (x) => Math.abs(x - m)) / nums.length;
    },
    { propagate: true }
  );

  agg(
    'DEVSQ',
    (args) => {
      const nums = numericArgs(args);
      if (nums.length === 0) throw new Error('#NUM!');
      const m = mean(nums);
      return sumOf(nums, (x) => (x - m) ** 2);
    },
    { propagate: true }
  );

  // Sample-based skew: n / ((n-1)(n-2)) * sum(((x - mean) / s)^3).
  agg(
    'SKEW',
    (args) => {
      const nums = numericArgs(args);
      const n = nums.length;
      if (n < 3) throw new Error('#DIV/0!');
      const s = Math.sqrt(varianceOf(nums, true));
      if (s === 0) throw new Error('#DIV/0!');
      const m = mean(nums);
      return (n / ((n - 1) * (n - 2))) * sumOf(nums, (x) => ((x - m) / s) ** 3);
    },
    { propagate: true }
  );

  agg(
    'SKEW.P',
    (args) => {
      const nums = numericArgs(args);
      const n = nums.length;
      if (n < 3) throw new Error('#DIV/0!');
      const s = Math.sqrt(varianceOf(nums, false));
      if (s === 0) throw new Error('#DIV/0!');
      const m = mean(nums);
      return sumOf(nums, (x) => ((x - m) / s) ** 3) / n;
    },
    { propagate: true }
  );

  // Excess kurtosis from the sample standard deviation.
  agg(
    'KURT',
    (args) => {
      const nums = numericArgs(args);
      const n = nums.length;
      if (n < 4) throw new Error('#DIV/0!');
      const s = Math.sqrt(varianceOf(nums, true));
      if (s === 0) throw new Error('#DIV/0!');
      const m = mean(nums);
      const fourth = sumOf(nums, (x) => ((x - m) / s) ** 4);
      return (
        ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * fourth -
        (3 * (n - 1) ** 2) / ((n - 2) * (n - 3))
      );
    },
    { propagate: true }
  );

  // ─── Probability and standardisation ────────────────

  def('PROB', 'aann', (_ctx, xs, ps, lower, upper?) => {
    const xv = flattenDeep(asArray(xs));
    const pv = flattenDeep(asArray(ps));
    if (xv.length !== pv.length) throw new Error('#N/A');
    const lo = toNumber(lower);
    const hi = upper === undefined ? lo : toNumber(upper);
    let total = 0;
    for (let i = 0; i < pv.length; i++) {
      if (isErrorCode(xv[i])) throw new Error(xv[i] as string);
      if (isErrorCode(pv[i])) throw new Error(pv[i] as string);
      if (typeof xv[i] !== 'number' || typeof pv[i] !== 'number') throw new Error('#VALUE!');
      const p = pv[i] as number;
      if (p < 0 || p > 1) throw new Error('#NUM!');
      total += p;
    }
    if (Math.abs(total - 1) > 1e-7) throw new Error('#NUM!');
    let sum = 0;
    for (let i = 0; i < xv.length; i++) {
      const x = xv[i] as number;
      if (x >= lo && x <= hi) sum += pv[i] as number;
    }
    return sum;
  }, { arity: [3, 4] });

  def('STANDARDIZE', 'nnn', (_ctx, x, m, sd) => {
    const deviation = toNumber(sd);
    if (deviation <= 0) throw new Error('#NUM!');
    return (toNumber(x) - toNumber(m)) / deviation;
  }, { arity: [3, 3] });

  def('PERMUTATIONA', 'nn', (_ctx, number, chosen) => {
    const n = Math.trunc(toNumber(number));
    const k = Math.trunc(toNumber(chosen));
    if (n < 0 || k < 0) throw new Error('#NUM!');
    return n ** k;
  }, { arity: [2, 2] });

  // ─── Two-array statistics and regression ────────────

  def('COVARIANCE.P', 'aa', (_ctx, a, b) => {
    const [as, bs] = pairedNumbers(a, b);
    return covariance(as, bs, false);
  }, { arity: [2, 2] });
  def('COVAR', 'aa', (_ctx, a, b) => {
    const [as, bs] = pairedNumbers(a, b);
    return covariance(as, bs, false);
  }, { arity: [2, 2] });
  def('COVARIANCE.S', 'aa', (_ctx, a, b) => {
    const [as, bs] = pairedNumbers(a, b);
    return covariance(as, bs, true);
  }, { arity: [2, 2] });

  for (const name of ['CORREL', 'PEARSON']) {
    def(name, 'aa', (_ctx, a, b) => {
      const [as, bs] = pairedNumbers(a, b);
      return pearson(as, bs);
    }, { arity: [2, 2] });
  }
  def('RSQ', 'aa', (_ctx, known_y, known_x) => {
    const [ys, xs] = pairedNumbers(known_y, known_x);
    return pearson(ys, xs) ** 2;
  }, { arity: [2, 2] });

  def('SLOPE', 'aa', (_ctx, known_y, known_x) => {
    const [ys, xs] = pairedNumbers(known_y, known_x);
    return regression(ys, xs).slope;
  }, { arity: [2, 2] });
  def('INTERCEPT', 'aa', (_ctx, known_y, known_x) => {
    const [ys, xs] = pairedNumbers(known_y, known_x);
    return regression(ys, xs).intercept;
  }, { arity: [2, 2] });

  // Standard error of the predicted y: sqrt(sum of squared residuals / (n - 2)).
  def('STEYX', 'aa', (_ctx, known_y, known_x) => {
    const [ys, xs] = pairedNumbers(known_y, known_x);
    if (xs.length < 3) throw new Error('#DIV/0!');
    const { slope, intercept } = regression(ys, xs);
    const residuals = ys.reduce((s, y, i) => s + (y - (intercept + slope * xs[i])) ** 2, 0);
    return Math.sqrt(residuals / (xs.length - 2));
  }, { arity: [2, 2] });

  const forecast = (_ctx: unknown, x: unknown, known_y: unknown, known_x: unknown): number => {
    const [ys, xs] = pairedNumbers(known_y, known_x);
    const { slope, intercept } = regression(ys, xs);
    return intercept + slope * toNumber(x);
  };
  def('FORECAST', 'naa', forecast, { arity: [3, 3] });
  def('FORECAST.LINEAR', 'naa', forecast, { arity: [3, 3] });
}
