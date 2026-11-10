/** One source cell of a fill series, as `seriesFor()` needs it. */
export interface SeriesCell {
  /** The cell's raw value ('' for an empty cell) */
  raw: string;
  /** The number the cell holds as a literal (not a formula), or null */
  number: number | null;
  /** Whether the cell has a date number format */
  isDate: boolean;
}

const NAME_LISTS: string[][] = [
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
];

/** Render a computed series value without floating-point noise (0.1+0.2 → 0.3). */
export function formatSeriesNumber(value: number): string {
  return String(parseFloat(value.toPrecision(15)));
}

/** Least-squares line through `values` at x = 0, 1, 2, ... (null if any is missing or fewer than 2). */
export function linearTrend(values: Array<number | null>): { intercept: number; slope: number } | null {
  if (values.length < 2 || values.some((v) => v === null)) return null;
  const ys = values as number[];
  const n = ys.length;
  const meanX = (n - 1) / 2;
  const meanY = ys.reduce((a, b) => a + b, 0) / n;
  let num = 0;
  let den = 0;
  ys.forEach((y, x) => {
    num += (x - meanX) * (y - meanY);
    den += (x - meanX) ** 2;
  });
  const slope = num / den;
  return { intercept: meanY - slope * meanX, slope };
}

/** The step between consecutive values, if it is the same everywhere (1 for a single value). */
function constantStep(values: number[], modulo?: number): number | null {
  if (values.length === 1) return 1;
  const norm = (d: number) => (modulo ? ((d % modulo) + modulo) % modulo : d);
  const step = norm(values[1] - values[0]);
  for (let i = 2; i < values.length; i++) {
    if (norm(values[i] - values[i - 1]) !== step) return null;
  }
  return step;
}

/** Apply the letter case of `sample` (UPPER, lower or Title) to a name. */
function matchCase(name: string, sample: string): string {
  if (sample === sample.toUpperCase()) return name.toUpperCase();
  if (sample === sample.toLowerCase()) return name.toLowerCase();
  return name;
}

function nameSeries(raws: string[]): ((offset: number) => string) | null {
  for (const list of NAME_LISTS) {
    const lower = list.map((n) => n.toLowerCase());
    const indexes = raws.map((r) => lower.indexOf(r.toLowerCase()));
    if (indexes.some((i) => i < 0)) continue;
    const step = constantStep(indexes, list.length);
    if (step === null) return null;
    const len = list.length;
    return (offset) => {
      const i = (((indexes[0] + step * offset) % len) + len) % len;
      return matchCase(list[i], raws[0]);
    };
  }
  return null;
}

function numberedTextSeries(raws: string[]): ((offset: number) => string) | null {
  const parts = raws.map((r) => /^(.*?)(\d+)$/.exec(r));
  if (parts.some((p) => !p || p[1] === '')) return null;
  const prefix = parts[0]![1];
  if (parts.some((p) => p![1] !== prefix)) return null;
  const numbers = parts.map((p) => Number(p![2]));
  const step = constantStep(numbers);
  if (step === null) return null;
  const first = parts[0]![2];
  const width = first.length > 1 && first.startsWith('0') ? first.length : 0;
  return (offset) => prefix + String(Math.abs(numbers[0] + step * offset)).padStart(width, '0');
}

/**
 * How the fill handle continues a line of source cells, as in Excel:
 *
 * - two or more numbers continue their linear trend (`1, 2` → `3, 4`);
 * - a single date steps by one day;
 * - day and month names continue and wrap (`Mon` → `Tue`, `Dec` → `Jan`),
 *   by their step when there are several (`Jan, Mar` → `May`);
 * - text ending in a number counts up (`Item 1` → `Item 2`, `Q09` → `Q10`),
 *   by the step when there are several with the same prefix.
 *
 * Returns the value at `offset` from the first source cell (negative when
 * filling up or left), or null when the cells just repeat (single numbers,
 * formulas, other text, mixed content).
 */
export function seriesFor(cells: SeriesCell[]): ((offset: number) => string) | null {
  if (cells.length === 0) return null;
  const numbers = cells.map((c) => c.number);
  const trend = linearTrend(numbers);
  if (trend) return (offset) => formatSeriesNumber(trend.intercept + trend.slope * offset);
  if (cells.length === 1 && cells[0].number !== null && cells[0].isDate) {
    const start = cells[0].number;
    return (offset) => formatSeriesNumber(start + offset);
  }
  if (cells.some((c) => c.number !== null || c.raw === '' || c.raw.startsWith('='))) return null;
  const raws = cells.map((c) => c.raw);
  return nameSeries(raws) ?? numberedTextSeries(raws);
}
