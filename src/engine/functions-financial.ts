/** Financial and engineering built-ins (values only; lazy or reference-taking functions live in formula-engine.ts). */
import { parseDateText } from './functions-datetime.js';
import { type Registrar, parseNumberText } from './functions.js';
import { flattenDeep, isErrorCode, toNumber, toText } from './values.js';

function fail(code: string): never {
  throw new Error(code);
}

/** Every element of an array argument in order; an error element is raised. */
function listOf(arg: unknown): unknown[] {
  return flattenDeep(arg).map((v) => {
    if (isErrorCode(v)) fail(v);
    return v;
  });
}

/** Numbers of an array argument; text, logicals and blanks are skipped, errors raised. */
function numbersOf(arg: unknown): number[] {
  const out: number[] = [];
  for (const v of listOf(arg)) if (typeof v === 'number') out.push(v);
  return out;
}

/** Excel's `type` argument: any non-zero value means payments at the start of the period. */
function paymentTiming(type: unknown): number {
  return type === undefined ? 0 : toNumber(type) ? 1 : 0;
}

/**
 * Root of `f` near `guess`: Newton's method with a numeric derivative, then
 * bisection over the first sign change on a grid. #NUM! if neither converges.
 */
function solveRate(f: (r: number) => number, guess: number): number {
  const at = (r: number) => {
    const y = f(r);
    return Number.isFinite(y) ? y : NaN;
  };
  let r = guess;
  for (let i = 0; i < 100 && r > -1; i++) {
    const y = at(r);
    if (Number.isNaN(y)) break;
    if (y === 0) return r;
    const h = 1e-6 * Math.max(1, Math.abs(r));
    const canGoLower = r - h > -1;
    const lower = canGoLower ? at(r - h) : y;
    const upper = at(r + h);
    const slope = (upper - lower) / (canGoLower ? 2 * h : h);
    if (!Number.isFinite(slope) || slope === 0) break;
    const next = r - y / slope;
    if (!(next > -1) || !Number.isFinite(next)) break;
    if (Math.abs(next - r) <= 1e-13 * Math.max(1, Math.abs(r))) return next;
    r = next;
  }
  const grid = [-0.999999, -0.99, -0.9, -0.5, -0.2, -0.05, 0, 0.05, 0.1, 0.2, 0.5, 1, 2, 5, 10, 100];
  for (let k = 0; k + 1 < grid.length; k++) {
    let a = grid[k];
    let b = grid[k + 1];
    let fa = at(a);
    const fb = at(b);
    if (Number.isNaN(fa) || Number.isNaN(fb)) continue;
    if (fa === 0) return a;
    if (fa * fb > 0) continue;
    for (let i = 0; i < 200; i++) {
      const m = (a + b) / 2;
      const fm = at(m);
      if (Number.isNaN(fm)) break;
      if (fm === 0 || b - a < 1e-15) return m;
      if (fa * fm < 0) {
        b = m;
      } else {
        a = m;
        fa = fm;
      }
    }
  }
  return fail('#NUM!');
}

// ─── Loan helpers (Excel's sign convention: money paid out is negative) ───

/** Growth factor (1+r)^n. */
const growth = (r: number, n: number) => Math.pow(1 + r, n);

/** Level payment of a loan of `pv` that leaves `fv` after `n` periods. */
function paymentOf(r: number, n: number, pv: number, fv: number, t: number): number {
  if (r === 0) return -(pv + fv) / n;
  const g = growth(r, n);
  return (-(pv * g + fv) * r) / ((1 + r * t) * (g - 1));
}

/** Balance owed after `k` end-of-period payments (Excel sign, positive while owed). */
function balanceEnd(r: number, k: number, pmt: number, pv: number): number {
  if (r === 0) return pv + pmt * k;
  return pv * growth(r, k) + (pmt * (growth(r, k) - 1)) / r;
}

/** Balance owed after `k` start-of-period payments. */
function balanceStart(r: number, k: number, pmt: number, pv: number): number {
  if (r === 0) return pv + pmt * k;
  return pv * growth(r, k) + (pmt * (1 + r) * (growth(r, k) - 1)) / r;
}

/** Interest paid in period `per` of a loan (Excel's IPMT). */
function interestOf(r: number, n: number, per: number, pv: number, fv: number, t: number): number {
  if (per < 1 || per > n || n <= 0) fail('#NUM!');
  const pmt = paymentOf(r, n, pv, fv, t);
  if (t === 0) return -r * balanceEnd(r, per - 1, pmt, pv);
  if (per === 1) return 0;
  // Annuity due: interest runs on the balance left after the payment made at the start of period per - 1
  return (-r * balanceStart(r, per - 1, pmt, pv)) / (1 + r);
}

/** Principal repaid in period `per` (PMT minus the interest). */
function principalOf(r: number, n: number, per: number, pv: number, fv: number, t: number): number {
  return paymentOf(r, n, pv, fv, t) - interestOf(r, n, per, pv, fv, t);
}

/** Sum of a per-period function over periods start..end, validated as CUMIPMT/CUMPRINC are. */
function cumulative(
  rate: unknown,
  nper: unknown,
  pv: unknown,
  start: unknown,
  end: unknown,
  type: unknown,
  per: (r: number, n: number, p: number, pv: number, fv: number, t: number) => number
): number {
  const r = toNumber(rate);
  const n = Math.trunc(toNumber(nper));
  const p = toNumber(pv);
  const s = Math.trunc(toNumber(start));
  const e = Math.trunc(toNumber(end));
  const t = toNumber(type);
  if (r <= 0 || n <= 0 || p <= 0) fail('#NUM!');
  if (s < 1 || e < s || e > n) fail('#NUM!');
  if (t !== 0 && t !== 1) fail('#NUM!');
  if (!Number.isSafeInteger(e) || e - s > 5_000_000) fail('#NUM!'); // bounded work, and k++ must advance
  let total = 0;
  for (let k = s; k <= e; k++) total += per(r, n, k, p, 0, t);
  return total;
}

// ─── Date-weighted cash flows (XNPV, XIRR) ───

/** Values and dates of an XNPV/XIRR call, validated; dates are truncated to whole days. */
function datedFlows(values: unknown, dates: unknown, xirr: boolean): { amounts: number[]; days: number[] } {
  const vs = listOf(values);
  const ds = listOf(dates);
  if (vs.length !== ds.length || vs.length === 0) fail('#NUM!');
  // XIRR takes numeric text amounts and date text (imported cash flows) and a blank amount as 0;
  // XNPV is strict: anything but a number is #NUM!
  const asNumber = (v: unknown): number | undefined => {
    if (typeof v === 'number') return v;
    if (!xirr || typeof v !== 'string' || v.trim() === '') return undefined;
    try {
      return parseNumberText(v.trim()); // 300, $300, 1,000, 30%
    } catch {
      /* not a number: a date text */
    }
    const parsed = parseDateText(v.trim());
    return parsed ? (parsed.date ?? 0) + (parsed.time ?? 0) : undefined;
  };
  const amounts = vs.map((v) => {
    if (v === undefined) {
      if (!xirr) fail('#NUM!');
      return 0;
    }
    const n = asNumber(v);
    if (n === undefined) fail(xirr ? '#VALUE!' : '#NUM!');
    return n;
  });
  const days = ds.map((d) => {
    const n = asNumber(d);
    if (n === undefined) fail(xirr ? '#VALUE!' : '#NUM!');
    return Math.trunc(n);
  });
  if (days.some((d) => d < days[0])) fail('#NUM!');
  return { amounts, days: days.map((d) => (d - days[0]) / 365) };
}

function hasBothSigns(values: number[]): boolean {
  return values.some((v) => v > 0) && values.some((v) => v < 0);
}

// ─── Depreciation ───

/** Fixed-rate depreciation rate of DB, rounded to three decimals as Excel does. */
function dbRate(cost: number, salvage: number, life: number): number {
  return Math.round((1 - Math.pow(salvage / cost, 1 / life)) * 1000) / 1000;
}

// ─── Number bases ───

const TWO_TO_48 = 2 ** 48;

/**
 * Parses a binary, octal or hexadecimal string (at most 10 characters). Ten
 * digits are two's complement: the top bit of the 10-digit width means negative.
 */
function parseBase(text: string, radix: 2 | 8 | 16, bitsPerDigit: number): number {
  if (text.length > 10) fail('#NUM!');
  if (text === '') return 0;
  const valid = radix === 2 ? /^[01]+$/ : radix === 8 ? /^[0-7]+$/ : /^[0-9a-fA-F]+$/;
  if (!valid.test(text)) fail('#NUM!');
  let n = parseInt(text, radix);
  if (text.length === 10) {
    const bits = 10 * bitsPerDigit;
    if (n >= 2 ** (bits - 1)) n -= 2 ** bits;
  }
  return n;
}

/**
 * Formats an integer in `radix` within the signed range of a 10-digit
 * number (`bits` wide). Negative values are 10-digit two's complement and
 * ignore `places`; otherwise `places` pads with zeros and must fit.
 */
function formatBase(value: number, radix: 2 | 8 | 16, bits: number, places?: unknown): string {
  const n = Math.trunc(value);
  const limit = 2 ** (bits - 1);
  if (n < -limit || n >= limit) fail('#NUM!');
  if (n < 0) return (n + 2 ** bits).toString(radix).toUpperCase();
  const text = n.toString(radix).toUpperCase();
  if (places === undefined) return text;
  const p = Math.trunc(toNumber(places));
  if (p < 1 || p > 10 || text.length > p) fail('#NUM!');
  return text.padStart(p, '0');
}

/** Bit-operation operand: a whole number from 0 up to 2^48-1. */
function bitOperand(v: unknown): bigint {
  const n = toNumber(v);
  if (!Number.isInteger(n) || n < 0 || n >= TWO_TO_48) fail('#NUM!');
  return BigInt(n);
}

/** Shift by `amount` bits (negative shifts the other way); results must fit in 48 bits. */
function shiftBits(value: unknown, amount: unknown, left: boolean): number {
  const n = bitOperand(value);
  const s = Math.trunc(toNumber(amount));
  if (Math.abs(s) > 53) fail('#NUM!');
  const toLeft = left === s >= 0;
  const result = toLeft ? n << BigInt(Math.abs(s)) : n >> BigInt(Math.abs(s));
  if (result >= BigInt(TWO_TO_48)) fail('#NUM!');
  return Number(result);
}

export function registerFinancialBuiltins(reg: Registrar): void {
  const { def } = reg;

  // ─── Investment appraisal ───

  // IRR: Newton from the guess (default 0.1); needs a positive and a negative flow.
  def(
    'IRR',
    'bn',
    (_ctx, values, guess?) => {
      const flows = numbersOf(values);
      if (!hasBothSigns(flows)) fail('#NUM!');
      const g = guess === undefined ? 0.1 : toNumber(guess);
      return solveRate(
        (r) => flows.reduce((sum, c, i) => sum + c / Math.pow(1 + r, i), 0),
        g
      );
    },
    { arity: [1, 2] }
  );

  // XIRR: like IRR, with flows dated in days from the first date (365-day years).
  def(
    'XIRR',
    'bbn',
    (_ctx, values, dates, guess?) => {
      const { amounts, days } = datedFlows(values, dates, true);
      if (!hasBothSigns(amounts)) fail('#NUM!');
      const g = guess === undefined ? 0.1 : toNumber(guess);
      return solveRate(
        (r) => amounts.reduce((sum, c, i) => sum + c / Math.pow(1 + r, days[i]), 0),
        g
      );
    },
    { arity: [2, 3] }
  );

  // XNPV: discounts each flow by its day offset from the first date over 365.
  def(
    'XNPV',
    'nbb',
    (_ctx, rate, values, dates) => {
      const r = toNumber(rate);
      if (r <= -1) fail('#NUM!');
      const { amounts, days } = datedFlows(values, dates, false);
      return amounts.reduce((sum, c, i) => sum + c / Math.pow(1 + r, days[i]), 0);
    },
    { arity: [3, 3] }
  );

  // MIRR: negative flows discounted at the finance rate, positive flows compounded at the reinvestment rate.
  def(
    'MIRR',
    'bnn',
    (_ctx, values, financeRate, reinvestRate) => {
      const flows = numbersOf(values);
      const fr = toNumber(financeRate);
      const rr = toNumber(reinvestRate);
      if (!hasBothSigns(flows)) fail('#DIV/0!');
      const n = flows.length;
      const positive = flows.reduce((sum, c, i) => (c > 0 ? sum + c * Math.pow(1 + rr, n - 1 - i) : sum), 0);
      const negative = flows.reduce((sum, c, i) => (c < 0 ? sum + c / Math.pow(1 + fr, i) : sum), 0);
      return Math.pow(-positive / negative, 1 / (n - 1)) - 1;
    },
    { arity: [3, 3] }
  );

  // RATE: the rate at which the annuity balances; guess defaults to 0.1.
  def(
    'RATE',
    'n',
    (_ctx, nper, pmt, pv, fv?, type?, guess?) => {
      const n = toNumber(nper);
      const p = toNumber(pmt);
      const v = toNumber(pv);
      const f = fv === undefined ? 0 : toNumber(fv);
      const t = paymentTiming(type);
      const g = guess === undefined ? 0.1 : toNumber(guess);
      if (n <= 0) fail('#NUM!');
      // The payments exactly repay the loan: the rate is 0 (Newton cancels badly there and can land on a distant root)
      const atZero = v + p * n + f;
      if (guess === undefined && Math.abs(atZero) <= 1e-9 * Math.max(1, Math.abs(v), Math.abs(p * n), Math.abs(f))) return 0;
      return solveRate((r) => {
        if (Math.abs(r) < 1e-12) return v + p * n + f;
        return v * growth(r, n) + (p * (1 + r * t) * (growth(r, n) - 1)) / r + f;
      }, g);
    },
    { arity: [3, 6] }
  );

  // NPER: the number of periods of an annuity.
  def(
    'NPER',
    'n',
    (_ctx, rate, pmt, pv, fv?, type?) => {
      const r = toNumber(rate);
      const p = toNumber(pmt);
      const v = toNumber(pv);
      const f = fv === undefined ? 0 : toNumber(fv);
      const t = paymentTiming(type);
      if (r === 0) {
        if (p === 0) fail('#DIV/0!');
        return -(v + f) / p;
      }
      const k = (p * (1 + r * t)) / r;
      const ratio = (k - f) / (v + k);
      if (!(ratio > 0)) fail('#NUM!');
      return Math.log(ratio) / Math.log(1 + r);
    },
    { arity: [3, 5] }
  );

  // IPMT / PPMT: interest and principal of one period.
  def(
    'IPMT',
    'n',
    (_ctx, rate, per, nper, pv, fv?, type?) => {
      const n = Math.trunc(toNumber(nper));
      return interestOf(
        toNumber(rate),
        n,
        Math.trunc(toNumber(per)),
        toNumber(pv),
        fv === undefined ? 0 : toNumber(fv),
        paymentTiming(type)
      );
    },
    { arity: [4, 6] }
  );

  def(
    'PPMT',
    'n',
    (_ctx, rate, per, nper, pv, fv?, type?) => {
      const n = Math.trunc(toNumber(nper));
      return principalOf(
        toNumber(rate),
        n,
        Math.trunc(toNumber(per)),
        toNumber(pv),
        fv === undefined ? 0 : toNumber(fv),
        paymentTiming(type)
      );
    },
    { arity: [4, 6] }
  );

  // CUMIPMT / CUMPRINC: sums over periods start..end of a loan (rate, nper and pv must be positive).
  def(
    'CUMIPMT',
    'n',
    (_ctx, rate, nper, pv, start, end, type) =>
      cumulative(rate, nper, pv, start, end, type, (r, n, p, v, f, t) => interestOf(r, n, p, v, f, t)),
    { arity: [6, 6] }
  );

  def(
    'CUMPRINC',
    'n',
    (_ctx, rate, nper, pv, start, end, type) =>
      cumulative(rate, nper, pv, start, end, type, (r, n, p, v, f, t) => principalOf(r, n, p, v, f, t)),
    { arity: [6, 6] }
  );

  // ISPMT: interest of a loan repaid in equal principal instalments.
  def(
    'ISPMT',
    'n',
    (_ctx, rate, per, nper, pv) => {
      const n = toNumber(nper);
      if (n === 0) fail('#DIV/0!');
      return toNumber(pv) * toNumber(rate) * (toNumber(per) / n - 1);
    },
    { arity: [4, 4] }
  );

  // EFFECT / NOMINAL: converting between nominal and effective annual rates.
  def(
    'EFFECT',
    'n',
    (_ctx, nominal, npery) => {
      const r = toNumber(nominal);
      const n = Math.trunc(toNumber(npery));
      if (r <= 0 || n < 1) fail('#NUM!');
      return Math.pow(1 + r / n, n) - 1;
    },
    { arity: [2, 2] }
  );

  def(
    'NOMINAL',
    'n',
    (_ctx, effect, npery) => {
      const r = toNumber(effect);
      const n = Math.trunc(toNumber(npery));
      if (r <= 0 || n < 1) fail('#NUM!');
      return n * (Math.pow(1 + r, 1 / n) - 1);
    },
    { arity: [2, 2] }
  );

  // FVSCHEDULE: principal compounded through a list of rates; blanks are skipped.
  def(
    'FVSCHEDULE',
    'nb',
    (_ctx, principal, schedule) => {
      let value = toNumber(principal);
      for (const v of listOf(schedule)) {
        if (v === undefined) continue;
        if (typeof v !== 'number') fail('#VALUE!');
        value *= 1 + v;
      }
      return value;
    },
    { arity: [2, 2] }
  );

  // PDURATION: periods for a present value to grow to a future value at a rate.
  def(
    'PDURATION',
    'n',
    (_ctx, rate, pv, fv) => {
      const r = toNumber(rate);
      const p = toNumber(pv);
      const f = toNumber(fv);
      if (r <= 0 || p <= 0 || f <= 0) fail('#NUM!');
      return (Math.log(f) - Math.log(p)) / Math.log(1 + r);
    },
    { arity: [3, 3] }
  );

  // RRI: the equivalent growth rate per period.
  def(
    'RRI',
    'n',
    (_ctx, nper, pv, fv) => {
      const n = toNumber(nper);
      const p = toNumber(pv);
      const f = toNumber(fv);
      if (n <= 0 || p === 0) fail('#NUM!');
      const result = Math.pow(f / p, 1 / n) - 1;
      if (!Number.isFinite(result)) fail('#NUM!');
      return result;
    },
    { arity: [3, 3] }
  );

  // ISPMT(rate, per, nper, pv): interest of a period of a loan with even principal payments.
  def(
    'ISPMT',
    'n',
    (_ctx, rate, per, nper, pv) => {
      const n = toNumber(nper);
      if (n === 0) fail('#DIV/0!');
      return toNumber(pv) * toNumber(rate) * (toNumber(per) / n - 1);
    },
    { arity: [4, 4] }
  );

  // ─── Depreciation ───

  def(
    'SLN',
    'n',
    (_ctx, cost, salvage, life) => {
      const n = toNumber(life);
      if (n === 0) fail('#DIV/0!');
      return (toNumber(cost) - toNumber(salvage)) / n;
    },
    { arity: [3, 3] }
  );

  // SYD: sum-of-years'-digits depreciation for one period.
  def(
    'SYD',
    'n',
    (_ctx, cost, salvage, life, per) => {
      const n = Math.trunc(toNumber(life));
      const p = Math.trunc(toNumber(per));
      if (n <= 0 || p <= 0 || p > n) fail('#NUM!');
      return ((toNumber(cost) - toNumber(salvage)) * (n - p + 1) * 2) / (n * (n + 1));
    },
    { arity: [4, 4] }
  );

  // DDB: declining balance with `factor` (default 2), never below salvage.
  def(
    'DDB',
    'n',
    (_ctx, cost, salvage, life, per, factor?) => {
      const c = toNumber(cost);
      const s = toNumber(salvage);
      const n = toNumber(life);
      const p = Math.trunc(toNumber(per));
      const f = factor === undefined ? 2 : toNumber(factor);
      if (c < 0 || s < 0 || n <= 0 || f <= 0 || p <= 0 || p > n) fail('#NUM!');
      // Closed form of the period-by-period loop: the book value falls by the factor
      // (1 - f/n) each period and stops at the salvage value
      const keep = 1 - f / n;
      const before = p === 1 ? c : keep <= 0 ? Math.min(c, s) : Math.max(Math.min(c, s), c * Math.pow(keep, p - 1));
      return Math.min((before * f) / n, Math.max(0, before - s));
    },
    { arity: [4, 5] }
  );

  // DB: fixed-declining balance; the first and last periods are prorated by `month`.
  def(
    'DB',
    'n',
    (_ctx, cost, salvage, life, per, month?) => {
      const c = toNumber(cost);
      const s = toNumber(salvage);
      const n = Math.trunc(toNumber(life));
      const p = Math.trunc(toNumber(per));
      const m = month === undefined ? 12 : Math.trunc(toNumber(month));
      if (c < 0 || s < 0 || n <= 0 || p < 1 || p > n + 1 || m < 1 || m > 12) fail('#NUM!');
      if (c === 0) return 0;
      const rate = dbRate(c, s, n);
      // Closed form: after the prorated first period the remainder shrinks by (1 - rate) each period
      if (p === 1) return (c * rate * m) / 12;
      const afterFirst = c - (c * rate * m) / 12;
      if (p <= n) return afterFirst * Math.pow(1 - rate, p - 2) * rate;
      return (afterFirst * Math.pow(1 - rate, n - 1) * rate * (12 - m)) / 12;
    },
    { arity: [4, 5] }
  );

  // ─── Fractional prices ───

  // DOLLARDE: a price written as integer.numerator (over `denominator`) to a decimal price.
  def(
    'DOLLARDE',
    'n',
    (_ctx, fraction, denominator) => {
      const x = toNumber(fraction);
      const d = Math.trunc(toNumber(denominator));
      if (d < 0) fail('#NUM!');
      if (d === 0) fail('#DIV/0!');
      const digits = Math.ceil(Math.log10(d));
      const scale = Math.pow(10, digits);
      const abs = Math.abs(x);
      const whole = Math.floor(abs);
      // The numerator may be fractional (a half tick: 99.015 at 32 is 99 and 1.5/32)
      const numerator = parseFloat(((abs - whole) * scale).toPrecision(12));
      return Math.sign(x) * (whole + numerator / d);
    },
    { arity: [2, 2] }
  );

  // DOLLARFR: the inverse of DOLLARDE.
  def(
    'DOLLARFR',
    'n',
    (_ctx, decimal, denominator) => {
      const x = toNumber(decimal);
      const d = Math.trunc(toNumber(denominator));
      if (d < 0) fail('#NUM!');
      if (d === 0) fail('#DIV/0!');
      const digits = Math.ceil(Math.log10(d));
      const scale = Math.pow(10, digits);
      const abs = Math.abs(x);
      const whole = Math.floor(abs);
      return Math.sign(x) * (whole + ((abs - whole) * d) / scale);
    },
    { arity: [2, 2] }
  );

  // ─── Number bases ───

  // Conversions read their input as text (a number cell 101 is "101").
  type Base = { radix: 2 | 8 | 16; bitsPerDigit: number; bits: number };
  const BIN: Base = { radix: 2, bitsPerDigit: 1, bits: 10 };
  const OCT: Base = { radix: 8, bitsPerDigit: 3, bits: 30 };
  const HEX: Base = { radix: 16, bitsPerDigit: 4, bits: 40 };

  // Text in a base, read as a signed decimal number.
  const toDecimal = (name: string, from: Base) => {
    def(name, 't', (_ctx, text) => parseBase(toText(text), from.radix, from.bitsPerDigit), { arity: [1, 1] });
  };

  // Decimal number in, text in a base; the number is truncated and negatives are two's complement.
  const fromDecimal = (name: string, to: Base) => {
    def(name, 'nn', (_ctx, number, places?) => formatBase(toNumber(number), to.radix, to.bits, places), {
      arity: [1, 2],
    });
  };

  // Text in one base, text in another, through the signed decimal value.
  const convert = (name: string, from: Base, to: Base) => {
    def(
      name,
      'tn',
      (_ctx, text, places?) =>
        formatBase(parseBase(toText(text), from.radix, from.bitsPerDigit), to.radix, to.bits, places),
      { arity: [1, 2] }
    );
  };

  toDecimal('BIN2DEC', BIN);
  toDecimal('OCT2DEC', OCT);
  toDecimal('HEX2DEC', HEX);
  fromDecimal('DEC2BIN', BIN);
  fromDecimal('DEC2OCT', OCT);
  fromDecimal('DEC2HEX', HEX);
  convert('BIN2OCT', BIN, OCT);
  convert('BIN2HEX', BIN, HEX);
  convert('OCT2BIN', OCT, BIN);
  convert('OCT2HEX', OCT, HEX);
  convert('HEX2BIN', HEX, BIN);
  convert('HEX2OCT', HEX, OCT);

  // ─── Bitwise ───

  def(
    'BITAND',
    'nn',
    (_ctx, a, b) => Number(bitOperand(a) & bitOperand(b)),
    { arity: [2, 2] }
  );
  def(
    'BITOR',
    'nn',
    (_ctx, a, b) => Number(bitOperand(a) | bitOperand(b)),
    { arity: [2, 2] }
  );
  def(
    'BITXOR',
    'nn',
    (_ctx, a, b) => Number(bitOperand(a) ^ bitOperand(b)),
    { arity: [2, 2] }
  );
  def('BITLSHIFT', 'nn', (_ctx, n, shift) => shiftBits(n, shift, true), { arity: [2, 2] });
  def('BITRSHIFT', 'nn', (_ctx, n, shift) => shiftBits(n, shift, false), { arity: [2, 2] });

  // ─── Comparison helpers ───

  def('DELTA', 'n', (_ctx, a, b?) => (toNumber(a) === (b === undefined ? 0 : toNumber(b)) ? 1 : 0), {
    arity: [1, 2],
  });
  def('GESTEP', 'n', (_ctx, n, step?) => (toNumber(n) >= (step === undefined ? 0 : toNumber(step)) ? 1 : 0), {
    arity: [1, 2],
  });
}
