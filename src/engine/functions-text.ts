/** Text built-ins (values only; lazy or reference-taking functions live in formula-engine.ts). */
import {
  type Registrar,
  type WildcardToken,
  ArgumentError,
  compileWildcard,
  parseNumberText,
  findSegment,
  wildcardSegments,
} from './functions.js';
import { RangeValue, flattenDeep, isErrorCode, toBoolean, toNumber, toText } from './values.js';

/** Excel's longest cell text; ARRAYTOTEXT stops building past it. */
const MAX_TEXT = 32767;

/** Text of an argument; an error (as a scalar or inside an array) is raised instead. */
function textOf(v: unknown): string {
  if (isErrorCode(v)) throw new Error(v);
  return toText(v);
}

/** Excel text for a quoted value in strict formats. */
const quoted = (s: string): string => `"${s.replace(/"/g, '""')}"`;

/** Windows-1252 characters for codes 128-159 (the rest of 1-255 are Latin-1 code points). */
const CP1252_HIGH = [
  0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152, 0x8d, 0x17d,
  0x8f, 0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122, 0x161, 0x203a, 0x153, 0x9d,
  0x17e, 0x178,
];

function charOfCode(code: number): string {
  return String.fromCodePoint(code >= 0x80 && code <= 0x9f ? CP1252_HIGH[code - 0x80] : code);
}

function codeOfChar(ch: string): number {
  const cp = ch.codePointAt(0)!;
  const high = CP1252_HIGH.indexOf(cp);
  if (high >= 0) return 0x80 + high;
  return cp <= 255 ? cp : 63; // "?" for characters outside the code page
}

/**
 * Index of the leftmost place where the wildcard pattern matches a run of `text`, at or after
 * `from`; #VALUE! if none. The first star-separated segment decides the start (its first match),
 * each later one is taken leftmost after the previous (greedy is optimal), and if that chain
 * fails no later start can do better. Plain segments are searched in linear time.
 */
function wildcardSearch(tokens: WildcardToken[], text: string[], from: number): number {
  const { segments, leading } = wildcardSegments(tokens);
  if (segments.length === 0) {
    if (from > text.length) throw new Error('#VALUE!');
    return from; // empty pattern or only stars
  }
  const firstAt = findSegment(segments[0], text, from);
  if (firstAt < 0) throw new Error('#VALUE!');
  let position = firstAt + segments[0].length;
  for (let k = 1; k < segments.length; k++) {
    const at = findSegment(segments[k], text, position);
    if (at < 0) throw new Error('#VALUE!');
    position = at + segments[k].length;
  }
  return leading ? from : firstAt; // a leading star absorbs everything before the first segment
}

/** SEARCH: case-insensitive, wildcards allowed; 1-based position or #VALUE!. */
function search(findV: unknown, withinV: unknown, startV: unknown): number {
  const within = textOf(withinV);
  const find = textOf(findV);
  const from = startV === undefined ? 1 : Math.trunc(toNumber(startV));
  if (from < 1 || from > Math.max(within.length, 1)) throw new Error('#VALUE!');
  const tokens = compileWildcard(find);
  const lowerWithin = within.toLowerCase();
  if (!tokens) {
    const at = lowerWithin.indexOf(find.toLowerCase(), from - 1);
    if (at < 0) throw new Error('#VALUE!');
    return at + 1;
  }
  return wildcardSearch(tokens, [...lowerWithin], from - 1) + 1;
}

/** FIND: case-sensitive, literal; 1-based position or #VALUE!. */
function find(findV: unknown, withinV: unknown, startV: unknown): number {
  const within = textOf(withinV);
  const needle = textOf(findV);
  const from = startV === undefined ? 1 : Math.trunc(toNumber(startV));
  if (from < 1 || from > Math.max(within.length, 1)) throw new Error('#VALUE!');
  const idx = within.indexOf(needle, from - 1);
  if (idx === -1) throw new Error('#VALUE!');
  return idx + 1;
}

/** REPLACE(old_text, start_num, num_chars, new_text). */
function replaceChars(oldV: unknown, startV: unknown, numV: unknown, newV: unknown): string {
  const s = textOf(oldV);
  const start = Math.trunc(toNumber(startV));
  const count = Math.trunc(toNumber(numV));
  if (start < 1 || count < 0) throw new Error('#VALUE!');
  return s.slice(0, start - 1) + textOf(newV) + s.slice(start - 1 + count);
}

/** Occurrences of a delimiter (left to right, non-overlapping); empty delimiters match at every position. */
function occurrences(text: string, delim: string, ignoreCase: boolean, matchEnd: boolean, fromEnd: boolean) {
  const found: Array<{ pos: number; len: number }> = [];
  if (delim === '') {
    for (let p = 0; p <= text.length; p++) found.push({ pos: p, len: 0 });
  } else {
    const hay = ignoreCase ? text.toLowerCase() : text;
    const needle = ignoreCase ? delim.toLowerCase() : delim;
    for (let p = hay.indexOf(needle); p !== -1; p = hay.indexOf(needle, p + needle.length)) {
      found.push({ pos: p, len: needle.length });
    }
  }
  // match_end: the end of the text counts as a delimiter (the start, when searching from the end)
  if (matchEnd) {
    if (fromEnd) found.unshift({ pos: 0, len: 0 });
    else if (!(found.length && found[found.length - 1].pos === text.length)) found.push({ pos: text.length, len: 0 });
  }
  return found;
}

/** TEXTBEFORE / TEXTAFTER: the text before or after the chosen delimiter, or undefined if there is none. */
function splitAtDelimiter(
  text: string,
  delim: string,
  instance: number,
  ignoreCase: boolean,
  matchEnd: boolean
): { pos: number; len: number } | undefined {
  if (instance === 0) throw new Error('#VALUE!');
  const found = occurrences(text, delim, ignoreCase, matchEnd, instance < 0);
  const idx = instance > 0 ? instance - 1 : found.length + instance;
  return idx >= 0 && idx < found.length ? found[idx] : undefined;
}

/** Digits rounded half away from zero to `d` places (negative `d` rounds to tens, hundreds, ...). */
function roundDigits(n: number, d: number): number {
  const factor = Math.pow(10, d);
  const rounded = Math.round(parseFloat((Math.abs(n) * factor).toPrecision(15)));
  const magnitude = d >= 0 ? rounded / factor : rounded * Math.pow(10, -d);
  return Math.sign(n) * magnitude;
}

/** Number text for FIXED and DOLLAR: rounded, fixed decimals, optional thousands separators. */
function fixedParts(value: unknown, decimalsArg: unknown, noCommas: boolean) {
  if (typeof value === 'string' && value.trim() === '') throw new Error('#VALUE!'); // "" is not 0 here
  // Number text as typed: 12%, 1,200.50, $1,200.50
  let n: number;
  if (typeof value === 'string') {
    try {
      n = parseNumberText(value.trim());
    } catch {
      n = toNumber(value);
    }
  } else {
    n = toNumber(value);
  }
  const decimals = decimalsArg === undefined ? 2 : Math.trunc(toNumber(decimalsArg));
  const rounded = roundDigits(n, decimals);
  const [intPart, frac] = Math.abs(rounded).toFixed(Math.max(decimals, 0)).split('.');
  const grouped = noCommas ? intPart : intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return { negative: rounded < 0 && /[1-9]/.test(intPart + (frac ?? '')), body: frac === undefined ? grouped : `${grouped}.${frac}` };
}

/** Text of a number as a plain scalar or array element (booleans are TRUE/FALSE). */
function arrayCell(v: unknown, strict: boolean): string {
  if (typeof v === 'string' && !isErrorCode(v)) return strict ? quoted(v) : v;
  return toText(v);
}

export function registerTextBuiltins(reg: Registrar): void {
  const { def } = reg;

  def('CONCATENATE', 't', (_ctx, ...parts) => parts.map(textOf).join(''), { arity: [1, 255] });

  def('CHAR', 'n', (_ctx, value) => {
    const code = Math.trunc(toNumber(value));
    if (code < 1 || code > 255) throw new Error('#VALUE!');
    return charOfCode(code);
  });

  def('UNICHAR', 'n', (_ctx, value) => {
    const code = Math.trunc(toNumber(value));
    if (code < 1 || code > 1114111) throw new Error('#VALUE!');
    return String.fromCodePoint(code);
  });

  def('CODE', 't', (_ctx, value) => {
    const first = Array.from(textOf(value))[0];
    if (first === undefined) throw new Error('#VALUE!');
    return codeOfChar(first);
  });

  def('UNICODE', 't', (_ctx, value) => {
    const s = textOf(value);
    if (s === '') throw new Error('#VALUE!');
    return s.codePointAt(0)!;
  });

  // Removes the non-printing characters 0-31
  def('CLEAN', 't', (_ctx, value) => textOf(value).replace(/[\u0000-\u001f]/g, ''));

  def('EXACT', 'tt', (_ctx, a, b) => textOf(a) === textOf(b));

  // Capital letter at the start of each word (a run of letters), lower case elsewhere
  def('PROPER', 't', (_ctx, value) => {
    let out = '';
    let inWord = false;
    for (const ch of textOf(value)) {
      const letter = /\p{L}/u.test(ch);
      out += letter ? (inWord ? ch.toLowerCase() : ch.toUpperCase()) : ch;
      inWord = letter;
    }
    return out;
  });

  def(
    'REPT',
    'tn',
    (_ctx, value, times) => {
      const s = textOf(value);
      const count = Math.trunc(toNumber(times));
      if (count < 0) throw new Error('#VALUE!');
      if (s === '') return '';
      if (s.length * count > 32767) throw new Error('#VALUE!');
      return s.repeat(count);
    }
  );

  def('REPLACE', 'tnnt', (_ctx, oldText, start, num, newText) => replaceChars(oldText, start, num, newText), {
    arity: [4, 4],
  });
  def('REPLACEB', 'tnnt', (_ctx, oldText, start, num, newText) => replaceChars(oldText, start, num, newText), {
    arity: [4, 4],
  });

  def('SEARCH', 'ttn', (_ctx, findText, within, start) => search(findText, within, start), { arity: [2, 3] });
  def('SEARCHB', 'ttn', (_ctx, findText, within, start) => search(findText, within, start), { arity: [2, 3] });
  def('FINDB', 'ttn', (_ctx, findText, within, start) => find(findText, within, start), { arity: [2, 3] });

  // Single-byte character sets: the B functions count characters, as the non-B ones do
  def('LENB', 't', (_ctx, value) => textOf(value).length);
  const leftChars = (_ctx: unknown, value: unknown, n?: unknown) => {
    const count = n === undefined ? 1 : Math.trunc(toNumber(n));
    if (count < 0) throw new Error('#VALUE!');
    return textOf(value).slice(0, count);
  };
  def('LEFTB', 'tn', leftChars, { arity: [1, 2] });
  def(
    'RIGHTB',
    'tn',
    (_ctx, value, n?) => {
      const s = textOf(value);
      const count = n === undefined ? 1 : Math.trunc(toNumber(n));
      if (count < 0) throw new Error('#VALUE!');
      return s.slice(Math.max(0, s.length - count));
    },
    { arity: [1, 2] }
  );
  def(
    'MIDB',
    'tnn',
    (_ctx, value, startV, countV) => {
      const s = textOf(value);
      const start = Math.trunc(toNumber(startV));
      const count = Math.trunc(toNumber(countV));
      if (start < 1 || count < 0) throw new Error('#VALUE!');
      return s.slice(start - 1, start - 1 + count);
    },
    { arity: [3, 3] }
  );

  // TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]) and TEXTAFTER
  const textAround = (after: boolean) => (_ctx: unknown, textV: unknown, delimV: unknown, ...rest: unknown[]) => {
    const [instanceV, modeV, endV, notFound] = rest;
    const text = textOf(textV);
    const delim = textOf(delimV);
    const instance = instanceV === undefined ? 1 : Math.trunc(toNumber(instanceV));
    const ignoreCase = modeV !== undefined && Math.trunc(toNumber(modeV)) === 1;
    const matchEnd = endV !== undefined && Math.trunc(toNumber(endV)) === 1;
    const hit = splitAtDelimiter(text, delim, instance, ignoreCase, matchEnd);
    if (!hit) {
      if (notFound === undefined) throw new Error('#N/A');
      if (notFound instanceof ArgumentError) throw new Error(notFound.code);
      return notFound;
    }
    return after ? text.slice(hit.pos + hit.len) : text.slice(0, hit.pos);
  };
  def('TEXTBEFORE', 'ttnnnv', textAround(false), { arity: [2, 6] });
  def('TEXTAFTER', 'ttnnnv', textAround(true), { arity: [2, 6] });

  // NUMBERVALUE(text, [decimal_separator], [group_separator]): locale-independent number parsing
  def(
    'NUMBERVALUE',
    'ttt',
    (_ctx, textV, decV, grpV) => {
      const dec = decV === undefined || decV === '' ? '.' : textOf(decV);
      const grp = grpV === undefined || grpV === '' ? ',' : textOf(grpV);
      if (dec.length !== 1 || grp.length !== 1 || dec === grp) throw new Error('#VALUE!');
      let s = textOf(textV).replace(/\s+/g, '');
      let percents = 0;
      while (s.endsWith('%')) {
        s = s.slice(0, -1);
        percents++;
      }
      if (s === '') return 0;
      s = s.split(grp).join('').split(dec).join('.');
      if (!/^[+-]?(\d+(\.\d*)?|\.\d+)$/.test(s)) throw new Error('#VALUE!');
      return Number(s) / Math.pow(100, percents);
    },
    { arity: [1, 3] }
  );

  // VALUETOTEXT(value, [format]): 0 concise, 1 strict (text quoted)
  def(
    'VALUETOTEXT',
    'vn',
    (_ctx, value, format) => {
      if (isErrorCode(value)) return value;
      const strict = format !== undefined && Math.trunc(toNumber(format)) === 1;
      return arrayCell(value, strict);
    },
    { arity: [1, 2] }
  );

  // ARRAYTOTEXT(array, [format]): 0 concise (comma-separated), 1 strict (array constant syntax)
  def(
    'ARRAYTOTEXT',
    'an',
    (_ctx, value, format) => {
      const strict = format !== undefined && Math.trunc(toNumber(format)) === 1;
      if (!(value instanceof RangeValue)) return arrayCell(value, strict);
      if (strict) {
        const rows: string[] = [];
        let size = 0;
        for (let r = 0; r < value.rows; r++) {
          const parts: string[] = [];
          for (const v of value.getRow(r)) {
            const cell = arrayCell(v, true);
            size += cell.length + 1;
            if (size > MAX_TEXT) throw new Error('#VALUE!'); // stop before building gigabytes
            parts.push(cell);
          }
          rows.push(parts.join(','));
        }
        return `{${rows.join(';')}}`;
      }
      const parts: string[] = [];
      let size = 0;
      for (const v of flattenDeep(value)) {
        const cell = arrayCell(v, false);
        size += cell.length + 2;
        if (size > MAX_TEXT) throw new Error('#VALUE!');
        parts.push(cell);
      }
      return parts.join(', ');
    },
    { arity: [1, 2] }
  );

  // T and N are special forms of the engine: a reference range gives the value of its first cell

  def(
    'FIXED',
    'nnv',
    (_ctx, value, decimals, noCommas) => {
      const parts = fixedParts(value, decimals, noCommas !== undefined && toBoolean(noCommas));
      return (parts.negative ? '-' : '') + parts.body;
    },
    { arity: [1, 3] }
  );

  // DOLLAR(number, [decimals]): currency text, negative amounts in parentheses
  def(
    'DOLLAR',
    'nn',
    (_ctx, value, decimals) => {
      const parts = fixedParts(value, decimals, false);
      return parts.negative ? `($${parts.body})` : `$${parts.body}`;
    },
    { arity: [1, 2] }
  );
}
