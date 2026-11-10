/**
 * Formula tokenizer and parser. Formulas are parsed once into an AST
 * (`parseFormula`) and evaluated many times by the engine; the token stream
 * (`tokenizeFormula`) is also what `utils/references.ts` rewrites references on.
 *
 * Grammar (highest precedence last):
 *   expression = comparison
 *   comparison = concat (("=" | "<>" | "<" | ">" | "<=" | ">=") concat)*
 *   concat     = additive ("&" additive)*
 *   additive   = term (("+" | "-") term)*
 *   term       = power (("*" | "/") power)*
 *   power      = unary ("^" unary)*                    left-associative, as in Excel
 *   unary      = ("-" | "+" | "@") unary | postfix     so -2^2 is 4, as in Excel
 *   postfix    = range ("%")*
 *   range      = primary (":" primary)*                A1:INDEX(...)
 *   primary    = NUMBER | STRING | BOOLEAN | ERROR | REF ["#"] | NAME
 *              | FUNC "(" args ")" | "(" expression ")" | "{" array "}"
 *              each optionally followed by "(" args ")" to call a LAMBDA
 */

/** Largest row/column counts a reference can address (Excel's sheet size). */
export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

export const ERROR_CODE_LIST = [
  '#ERROR!', '#REF!', '#DIV/0!', '#NAME?', '#CIRC!', '#VALUE!', '#N/A', '#NUM!', '#SPILL!', '#CALC!',
];

export type TokenType =
  | 'NUMBER'
  | 'STRING'
  | 'BOOLEAN'
  | 'REF'
  | 'NAME'
  | 'FUNC'
  | 'OPERATOR'
  | 'LPAREN'
  | 'RPAREN'
  | 'LBRACE'
  | 'RBRACE'
  | 'COMMA'
  | 'SEMICOLON'
  | 'COLON'
  | 'HASH'
  | 'ERROR'
  | 'EOF';

/** One end of a reference. `row`/`col` are 0-based; -1 when the reference has no such part (`A:A`, `1:1`). */
export interface RefEnd {
  row: number;
  col: number;
  rowAbs: boolean;
  colAbs: boolean;
}

/** A parsed reference: `A1`, `A1:B2`, `A:C`, `1:3`, each optionally `Sheet!`-qualified. */
export interface RefInfo {
  /** `cell` (A1), `range` (A1:B2), `cols` (A:C) or `rows` (1:3) */
  kind: 'cell' | 'range' | 'cols' | 'rows';
  /** Sheet name without quotes, when the reference is sheet-qualified */
  sheet?: string;
  start: RefEnd;
  /** Same as `start` for a single cell */
  end: RefEnd;
}

export interface Token {
  type: TokenType;
  value: string;
  /** Offsets of the token in the tokenized text */
  start: number;
  end: number;
  /** For REF tokens */
  ref?: RefInfo;
  /** For REF tokens: offset where the part after `Sheet!` starts */
  refStart?: number;
}

export type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: string }
  | { t: 'ref'; ref: RefInfo }
  | { t: 'name'; name: string }
  | { t: 'unary'; op: '-' | '+' | '%' | '@' | '#'; x: Node }
  | { t: 'bin'; op: string; l: Node; r: Node }
  | { t: 'call'; name: string; args: Array<Node | null> }
  | { t: 'invoke'; callee: Node; args: Array<Node | null> }
  | { t: 'array'; rows: Node[][] };

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$/;
const COL_RE = /^(\$?)([A-Za-z]{1,3})$/;
const ROW_RANGE_RE = /(\$?)([0-9]+):(\$?)([0-9]+)(?![A-Za-z0-9_.$])/y;
const NUMBER_RE = /(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?/y;
const IDENT_RE = /[A-Za-z_$\\][A-Za-z0-9_$.]*/y;
const QUOTED_SHEET_RE = /'((?:[^']|'')+)'!/y;
const SHEET_RE = /([A-Za-z_][A-Za-z0-9_.]*)!/y;

function lettersToCol(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) & 31);
  return col - 1;
}

function parseCell(text: string): RefEnd | null {
  const m = CELL_RE.exec(text);
  if (!m) return null;
  return { colAbs: m[1] === '$', col: lettersToCol(m[2]), rowAbs: m[3] === '$', row: parseInt(m[4], 10) - 1 };
}

function parseCol(text: string): RefEnd | null {
  const m = COL_RE.exec(text);
  if (!m) return null;
  return { colAbs: m[1] === '$', col: lettersToCol(m[2]), rowAbs: false, row: -1 };
}

function matchAt(re: RegExp, input: string, at: number): RegExpExecArray | null {
  re.lastIndex = at;
  return re.exec(input);
}

/**
 * Split formula text (without the leading `=`) into tokens.
 *
 * @throws {Error} On text that is not part of the formula language
 */
export function tokenizeFormula(input: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const push = (type: TokenType, value: string, start: number, end: number): Token => {
    const token: Token = { type, value, start, end };
    tokens.push(token);
    return token;
  };
  /** Whether a `#` here can be the spill operator (`A1#`) rather than an error literal. */
  const afterOperand = () => {
    const prev = tokens[tokens.length - 1];
    return (
      !!prev &&
      (prev.type === 'REF' || prev.type === 'RPAREN' || prev.type === 'NAME' || prev.type === 'ERROR') &&
      prev.end === i
    );
  };

  while (i < input.length) {
    const ch = input[i];

    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r' || ch === ' ') {
      i++;
      continue;
    }

    // String literal ("" is an escaped quote)
    if (ch === '"') {
      const start = i;
      let str = '';
      i++;
      for (;;) {
        if (i >= input.length) throw new Error('Unterminated string literal');
        if (input[i] === '"') {
          if (input[i + 1] === '"') {
            str += '"';
            i += 2;
            continue;
          }
          break;
        }
        str += input[i++];
      }
      i++;
      push('STRING', str, start, i);
      continue;
    }

    // Sheet-qualified reference: Sheet2!A1, 'My sheet'!A1:B2, Sheet2!A:A
    let sheet: string | undefined;
    const refTokenStart = i;
    if (ch === "'") {
      const m = matchAt(QUOTED_SHEET_RE, input, i);
      if (!m) throw new Error("Unexpected character: '");
      sheet = m[1].replace(/''/g, "'");
      i += m[0].length;
    } else if (/[A-Za-z_]/.test(ch)) {
      const m = matchAt(SHEET_RE, input, i);
      if (m) {
        sheet = m[1];
        i += m[0].length;
      }
    }

    if (sheet !== undefined) {
      // Only a reference (or an error left by a deleted one) may follow
      if (input[i] === '#') {
        const rest = input.slice(i, i + 8).toUpperCase();
        const code = ERROR_CODE_LIST.find((c) => rest.startsWith(c));
        if (!code) throw new Error('Expected a reference after the sheet name');
        i += code.length;
        push('ERROR', code, refTokenStart, i);
        continue;
      }
    }

    const c = input[i] ?? '';

    // Row range (1:3, $1:$3) — must be tried before number literals
    if (/[0-9$]/.test(c)) {
      const m = matchAt(ROW_RANGE_RE, input, i);
      if (m) {
        const token = push('REF', m[0].toUpperCase(), refTokenStart, i + m[0].length);
        token.refStart = i;
        token.ref = {
          kind: 'rows',
          sheet,
          start: { row: parseInt(m[2], 10) - 1, col: -1, rowAbs: m[1] === '$', colAbs: false },
          end: { row: parseInt(m[4], 10) - 1, col: -1, rowAbs: m[3] === '$', colAbs: false },
        };
        i += m[0].length;
        continue;
      }
    }

    if (sheet === undefined && (/[0-9]/.test(c) || (c === '.' && /[0-9]/.test(input[i + 1] ?? '')))) {
      const m = matchAt(NUMBER_RE, input, i)!;
      push('NUMBER', m[0], i, i + m[0].length);
      i += m[0].length;
      continue;
    }

    // Identifiers: references, function names, names, booleans
    if (/[A-Za-z_$\\]/.test(c)) {
      const m = matchAt(IDENT_RE, input, i)!;
      let ident = m[0];
      const identStart = i;
      i += ident.length;
      const hasDollar = ident.includes('$');

      const cell = parseCell(ident);
      if (cell) {
        // A1:B2 (both ends literal cells) is a single reference token
        let end = cell;
        let kind: RefInfo['kind'] = 'cell';
        if (input[i] === ':') {
          const m2 = matchAt(IDENT_RE, input, i + 1);
          const other = m2 ? parseCell(m2[0]) : null;
          if (other && !isCallAhead(input, i + 1 + m2![0].length)) {
            end = other;
            kind = 'range';
            i += 1 + m2![0].length;
          } else if (m2 && m2[0].includes('$')) {
            throw new Error(`Invalid range: ${ident}:${m2[0]}`);
          }
        }
        if (!isCallAhead(input, i) || kind === 'range') {
          const token = push('REF', input.slice(identStart, i).toUpperCase(), refTokenStart, i);
          token.refStart = identStart;
          token.ref = { kind, sheet, start: cell, end };
          continue;
        }
      }

      // Column range (A:C, $A:$C)
      const colStart = parseCol(ident);
      if (colStart && input[i] === ':') {
        const m2 = matchAt(IDENT_RE, input, i + 1);
        const colEnd = m2 ? parseCol(m2[0]) : null;
        if (colEnd && !isCallAhead(input, i + 1 + m2![0].length)) {
          i += 1 + m2![0].length;
          const token = push('REF', input.slice(identStart, i).toUpperCase(), refTokenStart, i);
          token.refStart = identStart;
          token.ref = { kind: 'cols', sheet, start: colStart, end: colEnd };
          continue;
        }
      }

      if (sheet !== undefined) throw new Error('Expected a reference after the sheet name');
      if (hasDollar) throw new Error(`Unexpected identifier: ${ident}`);

      // Trailing dots belong to what follows, not to the name
      while (ident.endsWith('.')) {
        ident = ident.slice(0, -1);
        i--;
      }
      const upper = ident.toUpperCase();
      if (isCallAhead(input, i)) {
        push('FUNC', upper, identStart, i);
      } else if (upper === 'TRUE' || upper === 'FALSE') {
        push('BOOLEAN', upper, identStart, i);
      } else {
        push('NAME', ident, identStart, i);
      }
      continue;
    }

    if (sheet !== undefined) throw new Error('Expected a reference after the sheet name');

    if (ch === '#') {
      const rest = input.slice(i, i + 8).toUpperCase();
      const code = ERROR_CODE_LIST.find((c2) => rest.startsWith(c2));
      if (code) {
        push('ERROR', code, i, i + code.length);
        i += code.length;
        continue;
      }
      if (afterOperand()) {
        push('HASH', '#', i, i + 1);
        i++;
        continue;
      }
      throw new Error('Unexpected character: #');
    }

    if (ch === '<' || ch === '>') {
      const next = input[i + 1];
      if (next === '=' || (ch === '<' && next === '>')) {
        push('OPERATOR', ch + next, i, i + 2);
        i += 2;
      } else {
        push('OPERATOR', ch, i, i + 1);
        i++;
      }
      continue;
    }

    const single: Record<string, TokenType> = {
      '+': 'OPERATOR', '-': 'OPERATOR', '*': 'OPERATOR', '/': 'OPERATOR', '&': 'OPERATOR',
      '^': 'OPERATOR', '%': 'OPERATOR', '=': 'OPERATOR', '@': 'OPERATOR',
      '(': 'LPAREN', ')': 'RPAREN', '{': 'LBRACE', '}': 'RBRACE',
      ',': 'COMMA', ';': 'SEMICOLON', ':': 'COLON',
    };
    const type = single[ch];
    if (type) {
      push(type, ch, i, i + 1);
      i++;
      continue;
    }

    throw new Error(`Unexpected character: ${ch}`);
  }

  push('EOF', '', input.length, input.length);
  return tokens;
}

/** Whether the next non-space character at `at` opens a call. */
function isCallAhead(input: string, at: number): boolean {
  let j = at;
  while (j < input.length && (input[j] === ' ' || input[j] === '\t')) j++;
  return input[j] === '(';
}

interface ParserState {
  tokens: Token[];
  pos: number;
  depth: number;
}

/** Nesting limit for the parser, so pathological input fails cleanly instead of overflowing the stack. */
const MAX_PARSE_DEPTH = 256;

function peek(s: ParserState): Token {
  return s.tokens[s.pos];
}

function isOp(s: ParserState, ...ops: string[]): boolean {
  const t = s.tokens[s.pos];
  return t.type === 'OPERATOR' && ops.includes(t.value);
}

function expect(s: ParserState, type: TokenType): Token {
  const token = s.tokens[s.pos];
  if (token.type !== type) {
    throw new Error(`Expected ${type} but got ${token.type} (${token.value})`);
  }
  s.pos++;
  return token;
}

/**
 * Parse formula text (without the leading `=`) into an AST.
 *
 * @throws {Error} When the text is not a valid formula
 */
export function parseFormula(input: string): Node {
  const s: ParserState = { tokens: tokenizeFormula(input), pos: 0, depth: 0 };
  const node = parseComparison(s);
  if (peek(s).type !== 'EOF') {
    throw new Error(`Unexpected token: ${peek(s).value}`);
  }
  return node;
}

function parseComparison(s: ParserState): Node {
  if (++s.depth > MAX_PARSE_DEPTH) throw new Error('Formula is nested too deeply');
  try {
    let left = parseConcat(s);
    while (isOp(s, '=', '<>', '<', '>', '<=', '>=')) {
      const op = s.tokens[s.pos++].value;
      left = { t: 'bin', op, l: left, r: parseConcat(s) };
    }
    return left;
  } finally {
    s.depth--;
  }
}

function parseConcat(s: ParserState): Node {
  let left = parseAdditive(s);
  while (isOp(s, '&')) {
    s.pos++;
    left = { t: 'bin', op: '&', l: left, r: parseAdditive(s) };
  }
  return left;
}

function parseAdditive(s: ParserState): Node {
  let left = parseTerm(s);
  while (isOp(s, '+', '-')) {
    const op = s.tokens[s.pos++].value;
    left = { t: 'bin', op, l: left, r: parseTerm(s) };
  }
  return left;
}

function parseTerm(s: ParserState): Node {
  let left = parsePower(s);
  while (isOp(s, '*', '/')) {
    const op = s.tokens[s.pos++].value;
    left = { t: 'bin', op, l: left, r: parsePower(s) };
  }
  return left;
}

function parsePower(s: ParserState): Node {
  let left = parseUnary(s);
  while (isOp(s, '^')) {
    s.pos++;
    left = { t: 'bin', op: '^', l: left, r: parseUnary(s) };
  }
  return left;
}

function parseUnary(s: ParserState): Node {
  if (isOp(s, '-', '+', '@')) {
    const op = s.tokens[s.pos++].value as '-' | '+' | '@';
    if (++s.depth > MAX_PARSE_DEPTH) throw new Error('Formula is nested too deeply');
    try {
      return { t: 'unary', op, x: parseUnary(s) };
    } finally {
      s.depth--;
    }
  }
  return parsePostfix(s);
}

function parsePostfix(s: ParserState): Node {
  let node = parseRange(s);
  while (isOp(s, '%')) {
    s.pos++;
    node = { t: 'unary', op: '%', x: node };
  }
  return node;
}

function parseRange(s: ParserState): Node {
  let left = parsePrimary(s);
  while (peek(s).type === 'COLON') {
    s.pos++;
    left = { t: 'bin', op: ':', l: left, r: parsePrimary(s) };
  }
  return left;
}

function parseArgs(s: ParserState): Array<Node | null> {
  expect(s, 'LPAREN');
  const args: Array<Node | null> = [];
  if (peek(s).type !== 'RPAREN') {
    for (;;) {
      // An omitted argument (e.g. XLOOKUP(x, A:A, B:B, , -1)) is null
      const next = peek(s).type;
      args.push(next === 'COMMA' || next === 'RPAREN' ? null : parseComparison(s));
      if (peek(s).type !== 'COMMA') break;
      s.pos++;
    }
  }
  expect(s, 'RPAREN');
  return args;
}

function parsePrimary(s: ParserState): Node {
  const token = peek(s);
  let node: Node;
  let callable = false;

  switch (token.type) {
    case 'NUMBER':
      s.pos++;
      node = { t: 'num', v: parseFloat(token.value) };
      break;
    case 'STRING':
      s.pos++;
      node = { t: 'str', v: token.value };
      break;
    case 'BOOLEAN':
      s.pos++;
      node = { t: 'bool', v: token.value === 'TRUE' };
      break;
    case 'ERROR':
      s.pos++;
      node = { t: 'err', v: token.value };
      // #REF!# is what a deleted spill reference (A1#) is rewritten to
      if (peek(s).type === 'HASH') s.pos++;
      break;
    case 'REF':
      s.pos++;
      node = { t: 'ref', ref: token.ref! };
      if (peek(s).type === 'HASH') {
        s.pos++;
        node = { t: 'unary', op: '#', x: node };
      }
      break;
    case 'NAME':
      s.pos++;
      node = { t: 'name', name: token.value };
      break;
    case 'FUNC':
      s.pos++;
      node = { t: 'call', name: token.value, args: parseArgs(s) };
      callable = true;
      break;
    case 'LPAREN':
      s.pos++;
      node = parseComparison(s);
      expect(s, 'RPAREN');
      callable = true;
      break;
    case 'LBRACE':
      node = parseArray(s);
      break;
    default:
      throw new Error(`Unexpected token: ${token.type} (${token.value})`);
  }

  // LAMBDA(x, x + 1)(2): call whatever the expression returned
  while (callable && peek(s).type === 'LPAREN') {
    node = { t: 'invoke', callee: node, args: parseArgs(s) };
  }
  return node;
}

function parseArray(s: ParserState): Node {
  expect(s, 'LBRACE');
  const rows: Node[][] = [[]];
  if (peek(s).type === 'RBRACE') throw new Error('Empty array');
  for (;;) {
    rows[rows.length - 1].push(parseComparison(s));
    const next = peek(s).type;
    if (next === 'COMMA') {
      s.pos++;
    } else if (next === 'SEMICOLON') {
      s.pos++;
      rows.push([]);
    } else {
      break;
    }
  }
  expect(s, 'RBRACE');
  const width = rows[0].length;
  if (rows.some((r) => r.length !== width)) throw new Error('Array rows differ in length');
  return { t: 'array', rows };
}

/** Whether a call to one of `names` (upper case) appears anywhere in the tree. */
export function callsAny(node: Node, names: ReadonlySet<string>): boolean {
  switch (node.t) {
    case 'call':
      return names.has(node.name) || node.args.some((a) => a !== null && callsAny(a, names));
    case 'invoke':
      return callsAny(node.callee, names) || node.args.some((a) => a !== null && callsAny(a, names));
    case 'unary':
      return callsAny(node.x, names);
    case 'bin':
      return callsAny(node.l, names) || callsAny(node.r, names);
    case 'array':
      return node.rows.some((row) => row.some((n) => callsAny(n, names)));
    default:
      return false;
  }
}

/** Whether a sheet name must be quoted in a reference (`'My sheet'!A1`). */
export function sheetNameNeedsQuotes(name: string): boolean {
  return !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) || CELL_RE.test(name);
}

/** A sheet name as written in front of a reference, quoted when needed, including the `!`. */
export function sheetPrefix(name: string): string {
  return (sheetNameNeedsQuotes(name) ? `'${name.replace(/'/g, "''")}'` : name) + '!';
}
