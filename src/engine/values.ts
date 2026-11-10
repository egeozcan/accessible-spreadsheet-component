/**
 * The values formulas compute with: scalars (number, string, boolean,
 * `undefined` for a blank cell), error codes (strings such as `#N/A`),
 * arrays (`RangeValue`, whose elements can themselves be arrays), references
 * (`RefValue`) and functions made by LAMBDA (`LambdaValue`).
 *
 * A scalar error is thrown as an `Error` whose message is the code; inside an
 * array it is stored as the code string.
 */
import type { CellArray, CellArrayItem } from '../types.js';
import { ERROR_CODE_LIST, type Node } from './parser.js';

export const ERROR_CODES: ReadonlySet<string> = new Set(ERROR_CODE_LIST);

export function isErrorCode(v: unknown): v is string {
  return typeof v === 'string' && ERROR_CODES.has(v);
}

/** Map a thrown value to a spreadsheet error code. */
export function toErrorCode(e: unknown): string {
  const msg = e instanceof Error ? e.message : '';
  return msg.startsWith('#') && ERROR_CODES.has(msg) ? msg : '#ERROR!';
}

/**
 * A 2D array of values with shape information: what a range reference
 * (`A1:B5`), an array literal (`{1,2;3,4}`) or an array function produces.
 * `values` is row-major. An element can be a number, string, boolean,
 * `undefined` (blank), an error code string, or a nested `RangeValue`.
 */
export class RangeValue {
  readonly values: unknown[];
  readonly rows: number;
  readonly cols: number;

  constructor(values: unknown[], rows: number, cols: number) {
    this.values = values;
    this.rows = rows;
    this.cols = cols;
  }

  get(row: number, col: number): unknown {
    return this.values[row * this.cols + col];
  }

  getRow(row: number): unknown[] {
    const start = row * this.cols;
    return this.values.slice(start, start + this.cols);
  }

  getCol(col: number): unknown[] {
    const result: unknown[] = [];
    for (let r = 0; r < this.rows; r++) {
      result.push(this.values[r * this.cols + col]);
    }
    return result;
  }

  /** A column (n×1) from a list of values. */
  static column(values: unknown[]): RangeValue {
    return new RangeValue(values, values.length, 1);
  }

  /** A row (1×n) from a list of values. */
  static row(values: unknown[]): RangeValue {
    return new RangeValue(values, 1, values.length);
  }
}

/** A sheet as the engine sees it. */
export interface EngineSheet {
  id: number;
  name: string;
  data: Map<string, import('../types.js').CellData>;
  /** Grid size; whole-column/row references span it */
  rows: number;
  cols: number;
  /** Cached extent of the cells that hold anything; null until needed */
  used: { rows: number; cols: number } | null;
}

/**
 * A reference to a rectangle of cells on one sheet. Functions that work on
 * references (ROW, OFFSET, the `:` operator, ...) receive it as is; everything
 * else receives the values it points to.
 */
export class RefValue {
  constructor(
    readonly sheet: EngineSheet,
    readonly top: number,
    readonly left: number,
    readonly bottom: number,
    readonly right: number,
    /** Whole-column (`A:C`) or whole-row (`1:3`) reference; its values are trimmed to the used part of the sheet */
    readonly whole: 'cols' | 'rows' | null = null,
    /** Written as a range (`A1:A1`), so it reads as an array even when it covers one cell */
    readonly area = false
  ) {}

  get rows(): number {
    return this.bottom - this.top + 1;
  }

  get cols(): number {
    return this.right - this.left + 1;
  }

  /** Whether this is a reference to one cell, which reads as that cell's value */
  get isCell(): boolean {
    return !this.area && this.top === this.bottom && this.left === this.right;
  }
}

/** Variable bindings of LET and LAMBDA, innermost first. */
export interface Scope {
  name: string;
  value: unknown;
  parent: Scope | null;
}

/** A function made by LAMBDA. */
export class LambdaValue {
  constructor(
    readonly params: string[],
    readonly body: Node,
    readonly scope: Scope | null,
    /** Runs the body with the given argument values */
    readonly call: (args: unknown[]) => unknown
  ) {}
}

/**
 * Coerce an operand of an arithmetic operator to a number.
 * Empty values are 0 and booleans are 1/0 (Excel semantics); anything else
 * that is not numeric is a #VALUE! error rather than a silent NaN.
 */
export function toNumber(v: unknown): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v === undefined || v === null || v === '') return 0;
  if (typeof v === 'string') {
    if (ERROR_CODES.has(v)) throw new Error(v);
    const n = Number(v);
    if (v.trim() !== '' && !isNaN(n)) return n;
  }
  throw new Error('#VALUE!');
}

/** Text of a scalar as the `&` operator and text functions see it. */
export function toText(v: unknown): string {
  if (v === undefined || v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return numberToText(v);
  if (v instanceof RangeValue) return arrayToText(v);
  return String(v);
}

/** A number as shown in a cell: 15 significant digits, no float noise. */
export function numberToText(n: number): string {
  return Number.isFinite(n) ? String(parseFloat(n.toPrecision(15))) : String(n);
}

/** Truthiness of a condition (IF, AND, FILTER's include, ...). Text other than TRUE/FALSE is #VALUE!. */
export function toBoolean(v: unknown): boolean {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === undefined || v === null || v === '') return false;
  if (typeof v === 'string') {
    if (ERROR_CODES.has(v)) throw new Error(v);
    const upper = v.toUpperCase();
    if (upper === 'TRUE') return true;
    if (upper === 'FALSE') return false;
    const n = Number(v);
    if (v.trim() !== '' && !isNaN(n)) return n !== 0;
  }
  return !!v;
}

/** Interpret a literal cell's raw text as a value (number, boolean or text). */
export function literalValue(raw: string): unknown {
  if (raw === '') return undefined;
  const num = Number(raw);
  if (!isNaN(num) && raw.trim() !== '') return num;
  const upper = raw.toUpperCase();
  if (upper === 'TRUE') return true;
  if (upper === 'FALSE') return false;
  return raw;
}

/** Every scalar in a value, going into arrays and nested arrays (row-major). */
export function flattenDeep(value: unknown, out: unknown[] = []): unknown[] {
  if (value instanceof RangeValue) {
    for (const v of value.values) flattenDeep(v, out);
  } else {
    out.push(value);
  }
  return out;
}

/** Whether any element of an array is itself an array. */
export function hasNested(value: RangeValue): boolean {
  return value.values.some((v) => v instanceof RangeValue);
}

/**
 * Text shown for an array kept in one cell: items separated by `, `, rows
 * by `; `, nested arrays in braces.
 */
export function arrayToText(value: RangeValue): string {
  const item = (v: unknown): string => (v instanceof RangeValue ? `{${arrayToText(v)}}` : toText(v));
  const rows: string[] = [];
  for (let r = 0; r < value.rows; r++) {
    rows.push(value.getRow(r).map(item).join(', '));
  }
  // A column reads as a plain list
  return value.cols === 1 ? rows.join(', ') : rows.join('; ');
}

/** The plain-data form of an array, as exposed on `CellData.array`. */
export function toCellArray(value: RangeValue): CellArray {
  const values: CellArrayItem[] = value.values.map((v) => {
    if (v instanceof RangeValue) return toCellArray(v);
    if (v === undefined || v === null) return null;
    if (typeof v === 'number' || typeof v === 'boolean' || typeof v === 'string') return v;
    return String(v);
  });
  return { rows: value.rows, cols: value.cols, values };
}

/** Whether two plain-data arrays are the same. */
export function cellArraysEqual(a: CellArray | undefined, b: CellArray | undefined): boolean {
  if (a === b) return true;
  if (!a || !b || a.rows !== b.rows || a.cols !== b.cols) return false;
  for (let i = 0; i < a.values.length; i++) {
    const x = a.values[i];
    const y = b.values[i];
    if (x === y) continue;
    if (x && y && typeof x === 'object' && typeof y === 'object') {
      if (!cellArraysEqual(x, y)) return false;
    } else {
      return false;
    }
  }
  return true;
}

/** Excel's `=` for two scalars: numbers by value, text ignoring case, blank equal to 0 and "". */
export function scalarsEqual(a: unknown, b: unknown): boolean {
  if (a === undefined || a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === undefined || b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
  if (typeof a === 'boolean' || typeof b === 'boolean') return a === b;
  const isNum = (v: unknown) => typeof v === 'number' || (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v)));
  if (isNum(a) && isNum(b)) return Number(a) === Number(b);
  return false;
}
