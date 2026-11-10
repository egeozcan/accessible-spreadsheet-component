import { type NumberFormatOptions, dateToSerial } from '../types.js';

/** A typed value recognised as a number, with the format Excel would give it. */
export interface ParsedInput {
  /** The plain number to store as the cell's raw value, e.g. `"0.15"` for `15%` */
  value: string;
  /** The number format the input implies */
  numberFormat: NumberFormatOptions;
}

const CURRENCY_SYMBOLS = ['$', '€', '£', '¥'];

/** Digits with optional thousands separators in groups of three, and decimals. */
const GROUPED = /^(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?$/;
const PLAIN = /^(\d+)?(\.\d+)?$/;

/** Count the decimals of a number written as text (`"1.50"` → 2). */
function decimalsOf(fraction: string | undefined): number {
  return fraction ? fraction.length - 1 : 0;
}

/** `text` times 10^exp as a string, without float noise (`1.1`, -2 → `"0.011"`). */
function shiftDecimal(text: string, exp: number): string {
  return String(parseFloat(`${text}e${exp}`));
}

function parseTime(h: string, m: string, s: string | undefined): number | null {
  const hours = Number(h);
  const minutes = Number(m);
  const seconds = s === undefined ? 0 : Number(s);
  if (hours > 23 || minutes > 59 || seconds > 59) return null;
  return (hours * 3600 + minutes * 60 + seconds) / 86400;
}

/** Serial for an ISO-ordered date, or null when the day doesn't exist (2026-02-30). */
function parseDate(y: string, mo: string, d: string): number | null {
  const year = Number(y);
  const month = Number(mo) - 1;
  const day = Number(d);
  if (month < 0 || month > 11 || day < 1) return null;
  const serial = dateToSerial(year, month, day);
  // dateToSerial rolls overflowing days into the next month; reject those
  if (dateToSerial(year, month + 1, 1) <= serial) return null;
  return serial;
}

/**
 * Recognise typed input that Excel would store as a number with a format:
 *
 * - percentages: `15%`, `-2.5 %` → 0.15 / -0.025 as a percent
 * - currency: `$1,200`, `-€3.50`, `£-3` → a currency number with that symbol
 * - grouped numbers: `1,234,567.8` → a number with thousands separators
 * - dates: `2026-01-05`, `2026/1/5` → a date serial shown as `YYYY-MM-DD`
 * - times: `14:30`, `9:05:30` → a fraction of a day shown as `HH:mm(:ss)`
 * - both: `2026-01-05 14:30`
 *
 * Only unambiguous, year-first dates are recognised (`1/5/2026` could be
 * January or May). Returns null for anything else, including plain numbers,
 * which need no format. Leading and trailing spaces are ignored.
 */
export function parseInput(input: string): ParsedInput | null {
  const text = input.trim();
  if (text === '' || text.startsWith('=')) return null;

  // Percent
  const pct = /^([+-]?)\s*(\d[\d,]*(?:\.\d+)?|\.\d+)\s*%$/.exec(text);
  if (pct) {
    const digits = pct[2];
    const m = GROUPED.exec(digits) ?? PLAIN.exec(digits);
    if (!m) return null;
    const value = shiftDecimal(pct[1] + digits.replace(/,/g, ''), -2);
    return { value, numberFormat: { type: 'percent', decimals: decimalsOf(m[2]) } };
  }

  // Currency: sign before or after the symbol
  const cur = /^([+-]?)(\S)\s?([+-]?)(\d[\d,]*(?:\.\d+)?|\.\d+)$/.exec(text);
  if (cur && CURRENCY_SYMBOLS.includes(cur[2]) && !(cur[1] && cur[3])) {
    const digits = cur[4];
    const m = GROUPED.exec(digits) ?? PLAIN.exec(digits);
    if (!m) return null;
    const sign = cur[1] === '-' || cur[3] === '-' ? '-' : '';
    const value = String(Number(sign + digits.replace(/,/g, '')));
    return {
      value,
      numberFormat: { type: 'currency', decimals: decimalsOf(m[2]) > 0 ? 2 : 0, currencySymbol: cur[2] },
    };
  }

  // Grouped number (a comma is required; plain numbers need no format)
  const grouped = /^([+-]?)(\d{1,3}(?:,\d{3})+)(\.\d+)?$/.exec(text);
  if (grouped) {
    const value = String(Number(grouped[1] + grouped[2].replace(/,/g, '') + (grouped[3] ?? '')));
    return { value, numberFormat: { type: 'number', decimals: decimalsOf(grouped[3]), thousandsSep: true } };
  }

  // Date, optionally with a time
  const date = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(text);
  if (date) {
    const day = parseDate(date[1], date[2], date[3]);
    if (day === null) return null;
    if (date[4] === undefined) {
      return { value: String(day), numberFormat: { type: 'date', datePattern: 'YYYY-MM-DD' } };
    }
    const time = parseTime(date[4], date[5], date[6]);
    if (time === null) return null;
    return {
      value: String(day + time),
      numberFormat: { type: 'date', datePattern: date[6] === undefined ? 'YYYY-MM-DD HH:mm' : 'YYYY-MM-DD HH:mm:ss' },
    };
  }

  // Time of day
  const time = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(text);
  if (time) {
    const t = parseTime(time[1], time[2], time[3]);
    if (t === null) return null;
    return {
      value: String(t),
      numberFormat: { type: 'date', datePattern: time[3] === undefined ? 'HH:mm' : 'HH:mm:ss' },
    };
  }

  return null;
}
