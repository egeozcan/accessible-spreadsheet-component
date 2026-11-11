/** Date and time built-ins (values only; lazy or reference-taking functions live in formula-engine.ts). */
import { type Registrar } from './functions.js';
import { dateSerialToDate, dateToSerial } from '../types.js';
import { flattenDeep, isErrorCode, toBoolean, toNumber, toText } from './values.js';

/** 9999-12-31, the last serial Excel accepts. */
const MAX_SERIAL = 2_958_465;

type Ymd = { y: number; m: number; d: number; wd: number }; // m 1-12, wd 0 = Sunday

function ymdOf(day: number): Ymd {
  const date = dateSerialToDate(day);
  return {
    y: date.getUTCFullYear(),
    m: date.getUTCMonth() + 1,
    d: date.getUTCDate(),
    wd: date.getUTCDay(),
  };
}

function isLeap(y: number): boolean {
  return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
}

function daysInMonth(y: number, m: number): number {
  return [31, isLeap(y) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];
}

/** Serial of a calendar date (m is 1-12); UTC, so no timezone or DST shift. */
function serialOf(y: number, m: number, d: number): number {
  return dateToSerial(y, m - 1, d);
}

const NUMBER_TEXT = /^\s*[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?\s*$/;
const MONTHS = [
  'january', 'february', 'march', 'april', 'may', 'june',
  'july', 'august', 'september', 'october', 'november', 'december',
];

function monthNumber(name: string): number {
  const n = name.toLowerCase();
  const i = MONTHS.findIndex((full) => n === full || (n.length === 3 && full.startsWith(n)));
  if (i < 0) throw new Error('#VALUE!');
  return i + 1;
}

/** Two-digit years use Excel's window (00-29 → 2000s, 30-99 → 1900s). */
function yearText(text: string): number {
  const n = Number(text);
  if (text.length === 2) return n <= 29 ? 2000 + n : 1900 + n;
  return n;
}

/**
 * Reads date and time text: ISO `2020-01-05`, `1/5/2020`, `1-Jan-2020`,
 * `January 5, 2020`, `5 January 2020`, optionally followed by `14:30`,
 * `14:30:15` or `2:30 PM`. Null if the text is none of these.
 */
export function parseDateText(text: string): { date?: number; time?: number } | null {
  let s = text.trim();
  let time: number | undefined;
  const t = /(?:^|\s)(\d{1,2})(?::(\d{2})(?::(\d{2}(?:\.\d+)?))?(?:\s*([AaPp][Mm]))?|\s*([AaPp][Mm]))$/.exec(s);
  if (t) {
    let h = Number(t[1]);
    const mi = t[2] === undefined ? 0 : Number(t[2]);
    const se = t[3] === undefined ? 0 : Number(t[3]);
    const ampm = (t[4] ?? t[5])?.toUpperCase();
    if (ampm !== undefined) {
      if (h < 1 || h > 12) return null;
      if (h === 12) h = 0;
      if (ampm === 'PM') h += 12;
    }
    if (mi > 59 || se >= 60) return null;
    s = s.slice(0, t.index).trim();
    // A bare time may run past 24 hours (Excel wraps it); one after a date may not
    if (h > 23 && (s !== '' || ampm !== undefined)) return null;
    time = (h * 3600 + mi * 60 + se) / 86400;
  }
  if (s === '') return time === undefined ? null : { time };

  let y: number, m: number, d: number;
  let r: RegExpExecArray | null;
  if ((r = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(s))) {
    y = Number(r[1]); m = Number(r[2]); d = Number(r[3]);
  } else if ((r = /^(\d{1,2})[-/](\d{1,2})[-/](\d{2}|\d{4})$/.exec(s))) {
    m = Number(r[1]); d = Number(r[2]); y = yearText(r[3]);
  } else if ((r = /^(\d{1,2})-([A-Za-z]+)-(\d{2}|\d{4})$/.exec(s))) {
    d = Number(r[1]); m = monthNumber(r[2]); y = yearText(r[3]);
  } else if ((r = /^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{2}|\d{4})$/.exec(s))) {
    m = monthNumber(r[1]); d = Number(r[2]); y = yearText(r[3]);
  } else if ((r = /^(\d{1,2})\s+([A-Za-z]+),?\s+(\d{2}|\d{4})$/.exec(s))) {
    d = Number(r[1]); m = monthNumber(r[2]); y = yearText(r[3]);
  } else {
    return null;
  }
  if (y < 1900 || y > 9999 || m < 1 || m > 12 || d < 1 || d > daysInMonth(y, m)) return null;
  return { date: serialOf(y, m, d), time };
}

/** Text argument that must be a date or time (or the error it carries). */
function dateTextArg(v: unknown): { date?: number; time?: number } {
  if (typeof v !== 'string') throw new Error('#VALUE!');
  if (isErrorCode(v)) throw new Error(v);
  const parsed = parseDateText(v);
  if (!parsed) throw new Error('#VALUE!');
  return parsed;
}

/**
 * A date or time argument as a serial number, time as the fraction of a day.
 * Blank is 0; numbers are taken as they are; text is read as a date or time.
 */
function serialArg(v: unknown): number {
  if (v === undefined || v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    if (isErrorCode(v)) throw new Error(v);
    if (NUMBER_TEXT.test(v)) return Number(v);
    const parsed = parseDateText(v);
    if (!parsed) throw new Error('#VALUE!');
    return (parsed.date ?? 0) + (parsed.time ?? 0);
  }
  throw new Error('#VALUE!');
}

/** A date argument as a whole day serial (the time is dropped). */
function dayArg(v: unknown): number {
  const day = Math.floor(serialArg(v));
  if (day < 0 || day > MAX_SERIAL) throw new Error('#NUM!');
  return day;
}

/** A time argument as whole seconds into the day (rounded, like Excel). */
function secondOfDay(v: unknown): number {
  const serial = serialArg(v);
  if (serial < 0 || serial > MAX_SERIAL + 1) throw new Error('#NUM!');
  return Math.round((serial - Math.floor(serial)) * 86400) % 86400;
}

/** Text argument; blank is "". */
function textArg(v: unknown): string {
  if (v === undefined || v === null) return '';
  if (isErrorCode(v)) throw new Error(v);
  return toText(v);
}

/** Weekend codes of NETWORKDAYS.INTL / WORKDAY.INTL: the weekdays (Sunday = 0) that are off. */
const WEEKEND_CODES: Record<number, number[]> = {
  1: [6, 0], 2: [0, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6],
  11: [0], 12: [1], 13: [2], 14: [3], 15: [4], 16: [5], 17: [6],
};

/**
 * The weekend argument as a flag per weekday (Sunday = 0): a code number, or
 * seven characters from Monday to Sunday where 1 is a day off and 0 a workday.
 */
function weekendArg(v: unknown): boolean[] {
  const flags = new Array<boolean>(7).fill(false);
  const value = v === undefined ? 1 : v;
  if (typeof value === 'string' && /^[01]+$/.test(value)) {
    // A string of 0s and 1s is always a weekend pattern, so it must be seven characters long
    if (value.length !== 7) throw new Error('#VALUE!');
    for (let i = 0; i < 7; i++) flags[(i + 1) % 7] = value[i] === '1';
    return flags;
  }
  const days = WEEKEND_CODES[Math.trunc(toNumber(value))];
  if (!days) throw new Error('#NUM!');
  for (const d of days) flags[d] = true;
  return flags;
}

/** The holidays argument (a value or a range) as a set of whole day serials. */
function holidaySet(v: unknown): Set<number> {
  const days = new Set<number>();
  for (const x of flattenDeep(v)) {
    if (x === undefined || x === null) continue;
    days.add(Math.floor(serialArg(x)));
  }
  return days;
}

/** Working days from `from` to `to` inclusive (`from` <= `to`): whole weeks counted, only the rest stepped. */
function workdayCount(from: number, to: number, weekend: boolean[], holidays: Set<number>): number {
  const perWeek = weekend.filter((w) => !w).length;
  const length = to - from + 1;
  const weeks = Math.floor(length / 7);
  let count = weeks * perWeek;
  let wd = ymdOf(from + weeks * 7).wd;
  for (let day = from + weeks * 7; day <= to; day++) {
    if (!weekend[wd]) count++;
    wd = (wd + 1) % 7;
  }
  for (const h of holidays) if (h >= from && h <= to && !weekend[ymdOf(h).wd]) count--;
  return count;
}

function networkDays(start: number, end: number, weekend: boolean[], holidays: Set<number>): number {
  return start <= end
    ? workdayCount(start, end, weekend, holidays)
    : -workdayCount(end, start, weekend, holidays);
}

/** The day `n` (>= 1) working days from `start` in direction `step`, ignoring holidays: whole weeks are skipped at once. */
function advanceWorkdays(start: number, n: number, step: 1 | -1, weekend: boolean[]): number {
  const perWeek = weekend.filter((w) => !w).length;
  let day = start;
  let left = n;
  if (left > perWeek) {
    const weeks = Math.floor((left - 1) / perWeek);
    day += step * 7 * weeks;
    left -= weeks * perWeek;
  }
  let wd = ymdOf(day).wd;
  while (left > 0) {
    day += step;
    wd = (wd + step + 7) % 7;
    if (!weekend[wd]) left--;
  }
  return day;
}

/** The day `days` working days away from `start` (start itself when days is 0). */
function workday(start: number, days: number, weekend: boolean[], holidays: Set<number>): number {
  if (weekend.every(Boolean)) throw new Error('#VALUE!');
  const n = Math.trunc(days);
  if (n === 0) return start;
  if (Math.abs(n) > MAX_SERIAL) throw new Error('#NUM!');
  const step = n > 0 ? 1 : -1;
  // Holidays on working days, ascending, so those passed are counted by bisection
  const working = [...holidays].filter((h) => !weekend[ymdOf(h).wd]).sort((a, b) => a - b);
  const countThrough = (x: number): number => {
    let lo = 0;
    let hi = working.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (working[mid] <= x) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  };
  let day = start;
  let left = Math.abs(n);
  // Each holiday passed pushes the result one working day further; repeat until none is passed anew
  while (left > 0) {
    const next = advanceWorkdays(day, left, step, weekend);
    if (next < 0 || next > MAX_SERIAL) throw new Error('#NUM!');
    const lo = step > 0 ? day + 1 : next;
    const hi = step > 0 ? next : day - 1;
    left = countThrough(hi) - countThrough(lo - 1);
    day = next;
  }
  return day;
}

/** Whole months from `a` to `b` (the day of month decides the last one). */
function monthsBetween(a: Ymd, b: Ymd): number {
  const months = (b.y - a.y) * 12 + (b.m - a.m);
  return b.d < a.d ? months - 1 : months;
}

/** The date `months` months after `a`, the day clamped to the end of a shorter month. */
function addMonths(a: Ymd, months: number): number {
  const total = a.y * 12 + (a.m - 1) + months;
  const y = Math.floor(total / 12);
  const m = (total % 12) + 1;
  return serialOf(y, m, Math.min(a.d, daysInMonth(y, m)));
}

/** Day count on a 30-day-month calendar, US (NASD) or European method. */
function days360(start: number, end: number, european: boolean, yearFrac = false): number {
  const a = ymdOf(start);
  const b = ymdOf(end);
  let d1 = a.d;
  let m2 = b.m;
  let d2 = b.d;
  if (european) {
    if (d1 === 31) d1 = 30;
    if (d2 === 31) d2 = 30;
  } else {
    // NASD rules as Excel applies them (fitted against Excel for 72 start/end pairs): the last
    // day of February starts as the 30th and a 31st start is the 30th. DAYS360 moves a 31st end to
    // the 30th if the start is on the 30th or later, else to the 1st of the next month; an end on
    // the last day of February is left alone. YEARFRAC basis 0 differs: both dates on the last
    // day of February make the end the 30th, and a 31st end is the 30th only if the start was
    // originally the 30th or 31st (otherwise it stays 31).
    const febEnd = (y: number, m: number, d: number) => m === 2 && d === daysInMonth(y, 2);
    const originalStart = d1;
    const startFebEnd = febEnd(a.y, a.m, a.d);
    if (yearFrac && startFebEnd && febEnd(b.y, b.m, d2)) d2 = 30;
    if (startFebEnd) d1 = 30;
    if (d1 === 31) d1 = 30;
    if (d2 === 31) {
      if (yearFrac) {
        if (originalStart >= 30) d2 = 30;
      } else if (d1 >= 30) {
        d2 = 30;
      } else {
        m2 += 1;
        d2 = 1;
      }
    }
  }
  return 360 * (b.y - a.y) + 30 * (m2 - a.m) + (d2 - d1);
}

/** YEARFRAC on actual days (basis 1), `start` <= `end`. */
function actualYearFrac(start: number, end: number): number {
  const a = ymdOf(start);
  const b = ymdOf(end);
  const oneYearLater = serialOf(a.y + 1, a.m, Math.min(a.d, daysInMonth(a.y + 1, a.m)));
  let denominator: number;
  if (end > oneYearLater) {
    let total = 0;
    for (let y = a.y; y <= b.y; y++) total += isLeap(y) ? 366 : 365;
    denominator = total / (b.y - a.y + 1);
  } else if (a.y === b.y) {
    denominator = isLeap(a.y) ? 366 : 365;
  } else {
    const leapDayInSpan =
      (isLeap(a.y) && start <= serialOf(a.y, 2, 29)) ||
      (isLeap(b.y) && end >= serialOf(b.y, 2, 29));
    denominator = leapDayInSpan ? 366 : 365;
  }
  return (end - start) / denominator;
}

/** Week number where the week containing January 1 is week 1; weeks start on `weekStart` (Sunday = 0). */
function weekNumber(day: number, weekStart: number): number {
  const jan1 = serialOf(ymdOf(day).y, 1, 1);
  const offset = (ymdOf(jan1).wd - weekStart + 7) % 7;
  return Math.floor((day - jan1 + offset) / 7) + 1;
}

/** ISO 8601 week: weeks start on Monday, and week 1 holds the year's first Thursday. */
function isoWeekNumber(day: number): number {
  const mondayBased = (ymdOf(day).wd + 6) % 7;
  const thursday = day - mondayBased + 3;
  const jan1 = serialOf(ymdOf(thursday).y, 1, 1);
  return Math.floor((thursday - jan1) / 7) + 1;
}

/** WEEKNUM return types → first day of the week (Sunday = 0). */
const WEEK_START: Record<number, number> = {
  1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0,
};

export function registerDatetimeBuiltins(reg: Registrar): void {
  const { def } = reg;

  // TIME: the fraction of a day for hour, minute and second (arguments truncated).
  def(
    'TIME',
    'n',
    (_ctx, hour, minute, second) => {
      const total =
        Math.trunc(toNumber(hour)) * 3600 +
        Math.trunc(toNumber(minute)) * 60 +
        Math.trunc(toNumber(second));
      if (total < 0) throw new Error('#NUM!');
      return (total % 86400) / 86400;
    },
    { arity: [3, 3] }
  );

  def('HOUR', 'v', (_ctx, time) => Math.floor(secondOfDay(time) / 3600), { arity: [1, 1] });
  def('MINUTE', 'v', (_ctx, time) => Math.floor((secondOfDay(time) % 3600) / 60), { arity: [1, 1] });
  def('SECOND', 'v', (_ctx, time) => secondOfDay(time) % 60, { arity: [1, 1] });

  // TIMEVALUE / DATEVALUE: the time or date a text reads as.
  def(
    'TIMEVALUE',
    'v',
    (_ctx, text) => {
      const time = dateTextArg(text).time ?? 0;
      return time - Math.floor(time);
    },
    { arity: [1, 1] }
  );
  def(
    'DATEVALUE',
    'v',
    (_ctx, text) => {
      const { date, time } = dateTextArg(text);
      // A bare time reads as a date too: day 0
      if (date === undefined) {
        if (time === undefined) throw new Error('#VALUE!');
        return 0;
      }
      return date;
    },
    { arity: [1, 1] }
  );

  // DATEDIF: complete years, months or days between two dates, or the leftover parts.
  def(
    'DATEDIF',
    'vvv',
    (_ctx, startArg, endArg, unitArg) => {
      const start = dayArg(startArg);
      const end = dayArg(endArg);
      if (start > end) throw new Error('#NUM!');
      const unit = textArg(unitArg).toUpperCase();
      const a = ymdOf(start);
      const b = ymdOf(end);
      switch (unit) {
        case 'Y':
          return Math.floor(monthsBetween(a, b) / 12);
        case 'M':
          return monthsBetween(a, b);
        case 'D':
          return end - start;
        case 'MD':
          return end - addMonths(a, monthsBetween(a, b));
        case 'YM':
          return monthsBetween(a, b) % 12;
        case 'YD': {
          let anchor = serialOf(b.y, a.m, Math.min(a.d, daysInMonth(b.y, a.m)));
          if (anchor > end) {
            anchor = serialOf(b.y - 1, a.m, Math.min(a.d, daysInMonth(b.y - 1, a.m)));
          }
          return end - anchor;
        }
        default:
          throw new Error('#NUM!');
      }
    },
    { arity: [3, 3] }
  );

  // DAYS: days from start to end (end first in the argument list).
  def('DAYS', 'vv', (_ctx, end, start) => dayArg(end) - dayArg(start), { arity: [2, 2] });

  def(
    'DAYS360',
    'vvv',
    (_ctx, start, end, method) => days360(dayArg(start), dayArg(end), toBoolean(method)),
    { arity: [2, 3] }
  );

  def(
    'NETWORKDAYS',
    'vvb',
    (_ctx, start, end, holidays) =>
      networkDays(dayArg(start), dayArg(end), weekendArg(1), holidaySet(holidays)),
    { arity: [2, 3] }
  );

  def(
    'NETWORKDAYS.INTL',
    'vvvb',
    (_ctx, start, end, weekend, holidays) =>
      networkDays(dayArg(start), dayArg(end), weekendArg(weekend), holidaySet(holidays)),
    { arity: [2, 4] }
  );

  def(
    'WORKDAY',
    'vvb',
    (_ctx, start, days, holidays) =>
      workday(dayArg(start), toNumber(days), weekendArg(1), holidaySet(holidays)),
    { arity: [2, 3] }
  );

  def(
    'WORKDAY.INTL',
    'vvvb',
    (_ctx, start, days, weekend, holidays) =>
      workday(dayArg(start), toNumber(days), weekendArg(weekend), holidaySet(holidays)),
    { arity: [2, 4] }
  );

  def(
    'WEEKNUM',
    'vv',
    (_ctx, serial, returnType) => {
      const day = dayArg(serial);
      const type = returnType === undefined ? 1 : Math.trunc(toNumber(returnType));
      if (type === 21) return isoWeekNumber(day);
      const weekStart = WEEK_START[type];
      if (weekStart === undefined) throw new Error('#NUM!');
      return weekNumber(day, weekStart);
    },
    { arity: [1, 2] }
  );

  def('ISOWEEKNUM', 'v', (_ctx, serial) => isoWeekNumber(dayArg(serial)), { arity: [1, 1] });

  // YEARFRAC: the part of a year between two dates, by the day-count basis (0-4).
  def(
    'YEARFRAC',
    'vvv',
    (_ctx, startArg, endArg, basisArg) => {
      let start = dayArg(startArg);
      let end = dayArg(endArg);
      if (start > end) [start, end] = [end, start];
      const basis = basisArg === undefined ? 0 : Math.trunc(toNumber(basisArg));
      switch (basis) {
        case 0: return days360(start, end, false, true) / 360;
        case 1: return actualYearFrac(start, end);
        case 2: return (end - start) / 360;
        case 3: return (end - start) / 365;
        case 4: return days360(start, end, true) / 360;
        default: throw new Error('#NUM!');
      }
    },
    { arity: [2, 3] }
  );
}
