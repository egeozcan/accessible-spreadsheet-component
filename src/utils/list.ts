/**
 * Lists: several values kept in one cell. A list is stored as a formula
 * whose result is an array that stays in the cell, `={{"Ann";"Bob";3}}`,
 * so it needs nothing beyond `rawValue` to survive undo, copy and export.
 * These helpers convert between that formula and the comma-separated text
 * a user types and edits ("Ann, Bob, 3").
 */
import type { CellArray, CellArrayItem } from '../types.js';
import { ERROR_CODE_LIST, type Node, parseFormula } from '../engine/parser.js';

/** A list item: text, a number or a boolean. */
export type ListItem = string | number | boolean;

function itemToFormula(item: ListItem): string {
  if (typeof item === 'number') return Number.isFinite(item) ? String(item) : '0';
  if (typeof item === 'boolean') return item ? 'TRUE' : 'FALSE';
  return `"${item.replace(/"/g, '""')}"`;
}

/** One value of an array as it is written in an array constant. */
function constantFor(item: CellArrayItem): string {
  if (item === null) return '""';
  if (typeof item === 'object') return arrayConstant(item);
  // Only the engine's own error codes are written bare; "#TODO" is text
  if (typeof item === 'string' && ERROR_CODE_LIST.includes(item)) return item;
  return itemToFormula(item);
}

function arrayConstant(array: CellArray): string {
  const rows: string[] = [];
  for (let r = 0; r < array.rows; r++) {
    rows.push(array.values.slice(r * array.cols, (r + 1) * array.cols).map(constantFor).join(','));
  }
  return `{${rows.join(';')}}`;
}

/**
 * The raw cell value that holds the given values in one cell, nested arrays
 * included: what `CellData.array` of a computed cell is, written as a
 * constant. `{ rows: 1, cols: 3, values: [4, 5, 6] }` gives `={{4,5,6}}`.
 */
export function arrayToFormula(array: CellArray): string {
  return `={${arrayConstant(array)}}`;
}

/**
 * The raw cell value that holds the given items as a list.
 *
 * @returns e.g. `={{"Ann";"Bob";3}}`; an empty string for no items
 */
export function listToFormula(items: ListItem[]): string {
  if (items.length === 0) return '';
  return `={{${items.map(itemToFormula).join(';')}}}`;
}

function literalOf(node: Node): ListItem | undefined {
  if (node.t === 'num' || node.t === 'str' || node.t === 'bool') return node.v;
  if (node.t === 'unary' && node.op === '-' && node.x.t === 'num') return -node.x.v;
  return undefined;
}

/**
 * The items of a list, when a raw cell value is one: a formula of the form
 * `={{a;b;c}}` (or `={{a,b,c}}`) with only constants in it.
 *
 * @returns The items, or `null` when the value is anything else
 */
export function listItemsOf(rawValue: string): ListItem[] | null {
  if (!rawValue.startsWith('={{')) return null;
  let ast: Node;
  try {
    ast = parseFormula(rawValue.slice(1));
  } catch {
    return null;
  }
  if (ast.t !== 'array' || ast.rows.length !== 1 || ast.rows[0].length !== 1) return null;
  const inner = ast.rows[0][0];
  if (inner.t !== 'array') return null;
  // One column or one row of constants
  if (inner.rows.length !== 1 && inner.rows[0].length !== 1) return null;
  const items: ListItem[] = [];
  for (const row of inner.rows) {
    for (const node of row) {
      const item = literalOf(node);
      if (item === undefined) return null;
      items.push(item);
    }
  }
  return items;
}

/** Whether an item must be quoted in list text to read back as the same text. */
function needsQuotes(text: string, separator: string): boolean {
  if (text === '' || text !== text.trim() || text.includes(separator) || text.includes('"')) return true;
  // Text that starts like a formula would be taken for one when the list is edited
  if (text.startsWith('=')) return true;
  // Text that would read back as a number or boolean
  return typeof parseListItem(text) !== 'string';
}

/**
 * The text a list is typed and edited as: items separated by `", "`, with
 * items quoted where needed (`"Smith, J"`, `"007"`).
 */
export function listToText(items: ListItem[], separator = ','): string {
  return items
    .map((item) => {
      if (typeof item === 'boolean') return item ? 'TRUE' : 'FALSE';
      if (typeof item === 'number') return String(item);
      return needsQuotes(item, separator) ? `"${item.replace(/"/g, '""')}"` : item;
    })
    .join(`${separator} `);
}

function parseListItem(text: string): ListItem {
  const num = Number(text);
  if (text !== '' && !isNaN(num) && Number.isFinite(num)) return num;
  const upper = text.toUpperCase();
  if (upper === 'TRUE') return true;
  if (upper === 'FALSE') return false;
  return text;
}

/**
 * Split typed text into list items at the separator (default `,`). Items are
 * trimmed; numbers and TRUE/FALSE become values; an item in double quotes is
 * kept as text exactly as written (`""` for a quote). Empty items are dropped.
 */
export function parseListText(text: string, separator = ','): ListItem[] {
  const items: ListItem[] = [];
  let i = 0;
  const n = text.length;
  while (i <= n) {
    while (i < n && (text[i] === ' ' || text[i] === '\t')) i++;
    if (text[i] === '"') {
      let value = '';
      i++;
      while (i < n) {
        if (text[i] === '"') {
          if (text[i + 1] === '"') {
            value += '"';
            i += 2;
            continue;
          }
          break;
        }
        value += text[i++];
      }
      i++;
      // Anything between the closing quote and the separator is ignored
      while (i < n && text[i] !== separator) i++;
      items.push(value);
    } else {
      let end = text.indexOf(separator, i);
      if (end === -1) end = n;
      const part = text.slice(i, end).trim();
      if (part !== '') items.push(parseListItem(part));
      i = end;
    }
    i++;
  }
  return items;
}

/**
 * Whether typed text reads as a list: it has the separator outside of quotes
 * and at least two items.
 */
export function looksLikeList(text: string, separator = ','): boolean {
  if (text.startsWith('=') || !text.includes(separator)) return false;
  return parseListText(text, separator).length >= 2;
}
