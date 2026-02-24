import {
  type GridData,
  type FormulaContext,
  type FormulaFunction,
  cellKey,
  refToCoord,
} from '../types.js';

/** Token types for the lexer */
type TokenType =
  | 'NUMBER'
  | 'STRING'
  | 'BOOLEAN'
  | 'REF'
  | 'RANGE'
  | 'FUNC'
  | 'OPERATOR'
  | 'LPAREN'
  | 'RPAREN'
  | 'COMMA'
  | 'EOF';

interface Token {
  type: TokenType;
  value: string;
}

/** Mutable parser state passed through the recursive descent chain */
interface ParserState {
  tokens: Token[];
  pos: number;
}

/**
 * FormulaEngine - A recursive descent parser for Excel-like formula syntax.
 *
 * Supports:
 * - Cell references: A1, B2, etc.
 * - Range references: A1:B5
 * - Functions: SUM(A1:A5), AVERAGE(...), etc.
 * - Arithmetic: +, -, *, /
 * - String literals: "hello"
 * - Numeric literals: 123, 3.14
 * - Boolean literals: TRUE, FALSE
 * - Comparison: =, <>, <, >, <=, >=
 * - String concatenation: &
 */
/** Maximum nesting depth for formula evaluation to prevent stack overflow from deeply nested formulas */
const MAX_EVAL_DEPTH = 64;

/** Set of aggregate function names that should have RangeValue flattened into individual args */
const AGGREGATE_FUNCTIONS = new Set([
  'SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'CONCAT',
]);

/**
 * Represents a 2D range of values with shape information.
 * Used by lookup functions (VLOOKUP, INDEX, etc.) that need row/col structure.
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
}

/**
 * Helper to check if a value matches a criteria string.
 * Supports operator prefixes: <>, >=, <=, >, <, =
 * Without operator prefix, does exact match (case-insensitive for strings).
 */
function matchesCriteria(value: unknown, criteria: string): boolean {
  // Parse operator from criteria
  let op = '=';
  let target = criteria;

  if (criteria.startsWith('<>')) {
    op = '<>';
    target = criteria.substring(2);
  } else if (criteria.startsWith('>=')) {
    op = '>=';
    target = criteria.substring(2);
  } else if (criteria.startsWith('<=')) {
    op = '<=';
    target = criteria.substring(2);
  } else if (criteria.startsWith('>')) {
    op = '>';
    target = criteria.substring(1);
  } else if (criteria.startsWith('<')) {
    op = '<';
    target = criteria.substring(1);
  } else if (criteria.startsWith('=')) {
    op = '=';
    target = criteria.substring(1);
  }

  const numTarget = Number(target);
  const numValue = Number(value);
  const bothNumeric = !isNaN(numTarget) && !isNaN(numValue)
    && target.trim() !== '' && String(value).trim() !== '';

  switch (op) {
    case '=':
      if (bothNumeric) return numValue === numTarget;
      return String(value).toLowerCase() === target.toLowerCase();
    case '<>':
      if (bothNumeric) return numValue !== numTarget;
      return String(value).toLowerCase() !== target.toLowerCase();
    case '>':
      return bothNumeric ? numValue > numTarget : String(value) > target;
    case '<':
      return bothNumeric ? numValue < numTarget : String(value) < target;
    case '>=':
      return bothNumeric ? numValue >= numTarget : String(value) >= target;
    case '<=':
      return bothNumeric ? numValue <= numTarget : String(value) <= target;
    default:
      return false;
  }
}

export class FormulaEngine {
  private functions: Map<string, FormulaFunction> = new Map();
  private data: GridData = new Map();
  private evaluating: Set<string> = new Set(); // circular reference detection
  private _evalDepth = 0;

  // Dependency tracking for targeted recalculation
  private _deps: Map<string, Set<string>> = new Map();
  private _reverseDeps: Map<string, Set<string>> = new Map();
  private _trackingCellKey: string | null = null;

  // Formula result cache for performance
  private _cache: Map<string, { displayValue: string; type: 'text' | 'number' | 'boolean' | 'error' }> = new Map();

  constructor() {
    this.registerBuiltins();
  }

  /** Register a user-defined function */
  registerFunction(name: string, fn: FormulaFunction): void {
    this.functions.set(name.toUpperCase(), fn);
  }

  /** Update the data reference for evaluation */
  setData(data: GridData): void {
    this.data = data;
    this._deps.clear();
    this._reverseDeps.clear();
    this._cache.clear();
  }

  /**
   * Evaluate a raw value. If it starts with '=', parse and execute the formula.
   * Otherwise, return the value as-is (possibly coerced).
   *
   * @param forCellKey - If provided, tracks formula dependencies for this cell key.
   */
  evaluate(
    rawValue: string,
    forCellKey?: string
  ): { displayValue: string; type: 'text' | 'number' | 'boolean' | 'error' } {
    if (!rawValue || rawValue.trim() === '') {
      if (forCellKey) {
        this._clearDepsFor(forCellKey);
        this._cache.delete(forCellKey);
      }
      return { displayValue: '', type: 'text' };
    }

    if (!rawValue.startsWith('=')) {
      if (forCellKey) {
        this._clearDepsFor(forCellKey);
        this._cache.delete(forCellKey);
      }
      return this.coerceValue(rawValue);
    }

    // Check cache for formula cells
    if (forCellKey) {
      const cached = this._cache.get(forCellKey);
      if (cached) {
        return { displayValue: cached.displayValue, type: cached.type };
      }
    }

    try {
      if (forCellKey) {
        this._clearDepsFor(forCellKey);
        this._trackingCellKey = forCellKey;
      }
      const formula = rawValue.substring(1);
      const result = this.parseExpression(formula);
      const evaluated = this.coerceValue(String(result));

      // Store in cache for formula cells
      if (forCellKey) {
        this._cache.set(forCellKey, { displayValue: evaluated.displayValue, type: evaluated.type });
      }

      return evaluated;
    } catch (e) {
      // Preserve specific error codes (#DIV/0!, #NAME?, #CIRC!)
      const msg = e instanceof Error ? e.message : '';
      const errorResult = msg.startsWith('#')
        ? { displayValue: msg, type: 'error' as const }
        : { displayValue: '#ERROR!', type: 'error' as const };

      if (forCellKey) {
        this._cache.set(forCellKey, errorResult);
      }

      return errorResult;
    } finally {
      this._trackingCellKey = null;
    }
  }

  /**
   * Re-evaluate all formula cells in the grid.
   * Rebuilds the dependency graph. Returns a set of cell keys that changed.
   */
  recalculate(): Set<string> {
    this._deps.clear();
    this._reverseDeps.clear();
    this._cache.clear();

    const changed = new Set<string>();

    for (const [key, cell] of this.data) {
      if (cell.rawValue.startsWith('=')) {
        const result = this.evaluate(cell.rawValue, key);
        if (cell.displayValue !== result.displayValue || cell.type !== result.type) {
          cell.displayValue = result.displayValue;
          cell.type = result.type;
          changed.add(key);
        }
      }
    }

    return changed;
  }

  /**
   * Recalculate only formulas affected by the given changed cell keys.
   * Uses the dependency graph for targeted recalculation (BFS order).
   */
  recalculateAffected(changedKeys: string[]): Set<string> {
    // If the dependency graph is empty but data exists, fall back to a full
    // recalculate so we never silently skip dependents after a setData() call.
    if (this._reverseDeps.size === 0 && this.data.size > 0) {
      return this.recalculate();
    }

    const changed = new Set<string>();
    const changedSet = new Set(changedKeys);

    // Clear cache entries for directly changed keys before re-evaluating
    for (const key of changedKeys) {
      this._cache.delete(key);
    }

    // Re-evaluate changed cells that are formulas (they may have been
    // evaluated with stale sibling values during batch application)
    for (const key of changedKeys) {
      const cell = this.data.get(key);
      if (cell?.rawValue.startsWith('=')) {
        const result = this.evaluate(cell.rawValue, key);
        if (cell.displayValue !== result.displayValue || cell.type !== result.type) {
          cell.displayValue = result.displayValue;
          cell.type = result.type;
          changed.add(key);
        }
      }
    }

    // BFS to find all transitive dependents (respects evaluation order)
    const toRecalc: string[] = [];
    const visited = new Set<string>();
    const queue = [...changedKeys];

    while (queue.length > 0) {
      const key = queue.shift()!;
      const dependents = this._reverseDeps.get(key);
      if (dependents) {
        for (const dep of dependents) {
          if (!visited.has(dep) && !changedSet.has(dep)) {
            visited.add(dep);
            toRecalc.push(dep);
            queue.push(dep);
          }
        }
      }
    }

    // Clear cache entries for BFS dependents before re-evaluating
    for (const key of toRecalc) {
      this._cache.delete(key);
    }

    // Recalculate each dependent formula in BFS order
    for (const key of toRecalc) {
      const cell = this.data.get(key);
      if (cell?.rawValue.startsWith('=')) {
        const result = this.evaluate(cell.rawValue, key);
        if (cell.displayValue !== result.displayValue || cell.type !== result.type) {
          cell.displayValue = result.displayValue;
          cell.type = result.type;
          changed.add(key);
        }
      }
    }

    return changed;
  }

  // ─── Dependency Tracking ─────────────────────────────

  private _clearDepsFor(targetKey: string): void {
    const oldDeps = this._deps.get(targetKey);
    if (oldDeps) {
      for (const dep of oldDeps) {
        this._reverseDeps.get(dep)?.delete(targetKey);
      }
      this._deps.delete(targetKey);
    }
    this._cache.delete(targetKey);
  }

  private _trackDep(referencedKey: string): void {
    if (!this._trackingCellKey) return;

    let deps = this._deps.get(this._trackingCellKey);
    if (!deps) {
      deps = new Set();
      this._deps.set(this._trackingCellKey, deps);
    }
    deps.add(referencedKey);

    let rev = this._reverseDeps.get(referencedKey);
    if (!rev) {
      rev = new Set();
      this._reverseDeps.set(referencedKey, rev);
    }
    rev.add(this._trackingCellKey);
  }

  // ─── Value Coercion ──────────────────────────────────

  private coerceValue(val: string): { displayValue: string; type: 'text' | 'number' | 'boolean' | 'error' } {
    if (val === '#ERROR!' || val === '#REF!' || val === '#DIV/0!' || val === '#NAME?' || val === '#CIRC!' || val === '#VALUE!') {
      return { displayValue: val, type: 'error' };
    }

    const upper = val.toUpperCase();
    if (upper === 'TRUE' || upper === 'FALSE') {
      return { displayValue: upper, type: 'boolean' };
    }

    const num = Number(val);
    if (val.trim() !== '' && !isNaN(num)) {
      return { displayValue: String(num), type: 'number' };
    }

    return { displayValue: val, type: 'text' };
  }

  // ─── Lexer ───────────────────────────────────────────

  private tokenize(input: string): Token[] {
    const tokens: Token[] = [];
    let i = 0;

    while (i < input.length) {
      const ch = input[i];

      // Skip whitespace
      if (/\s/.test(ch)) {
        i++;
        continue;
      }

      // String literal
      if (ch === '"') {
        let str = '';
        i++; // skip opening quote
        while (i < input.length && input[i] !== '"') {
          str += input[i];
          i++;
        }
        i++; // skip closing quote
        tokens.push({ type: 'STRING', value: str });
        continue;
      }

      // Number literal
      if (/\d/.test(ch) || (ch === '.' && i + 1 < input.length && /\d/.test(input[i + 1]))) {
        let num = '';
        while (i < input.length && (/\d/.test(input[i]) || input[i] === '.')) {
          num += input[i];
          i++;
        }
        tokens.push({ type: 'NUMBER', value: num });
        continue;
      }

      // Operators
      if (ch === '+' || ch === '-' || ch === '*' || ch === '/' || ch === '&') {
        tokens.push({ type: 'OPERATOR', value: ch });
        i++;
        continue;
      }

      // Comparison operators
      if (ch === '<' || ch === '>') {
        if (i + 1 < input.length && input[i + 1] === '=') {
          tokens.push({ type: 'OPERATOR', value: ch + '=' });
          i += 2;
        } else if (ch === '<' && i + 1 < input.length && input[i + 1] === '>') {
          tokens.push({ type: 'OPERATOR', value: '<>' });
          i += 2;
        } else {
          tokens.push({ type: 'OPERATOR', value: ch });
          i++;
        }
        continue;
      }

      if (ch === '=') {
        tokens.push({ type: 'OPERATOR', value: '=' });
        i++;
        continue;
      }

      // Parentheses
      if (ch === '(') {
        tokens.push({ type: 'LPAREN', value: '(' });
        i++;
        continue;
      }
      if (ch === ')') {
        tokens.push({ type: 'RPAREN', value: ')' });
        i++;
        continue;
      }

      // Comma
      if (ch === ',') {
        tokens.push({ type: 'COMMA', value: ',' });
        i++;
        continue;
      }

      // Identifiers: cell references, function names, booleans
      // Also match $ for absolute/mixed references like $A$1, $A1, A$1
      if (/[A-Za-z_$]/.test(ch)) {
        let ident = '';
        while (i < input.length && /[A-Za-z0-9_$]/.test(input[i])) {
          ident += input[i];
          i++;
        }

        // Strip $ for structure checks, but keep original for token value
        const stripped = ident.replace(/\$/g, '');
        const hasDollar = stripped !== ident;
        const upper = stripped.toUpperCase();

        // Check if boolean (only when no $ is present)
        if (!hasDollar && (upper === 'TRUE' || upper === 'FALSE')) {
          tokens.push({ type: 'BOOLEAN', value: upper });
          continue;
        }

        // Check for range (e.g. A1:B2 or $A$1:$B$2) - look ahead for colon
        if (i < input.length && input[i] === ':' && /^[A-Z]+\d+$/i.test(stripped)) {
          i++; // skip colon
          let end = '';
          while (i < input.length && /[A-Za-z0-9$]/.test(input[i])) {
            end += input[i];
            i++;
          }
          const endStripped = end.replace(/\$/g, '');
          if (/^[A-Z]+\d+$/i.test(endStripped)) {
            tokens.push({ type: 'RANGE', value: `${ident.toUpperCase()}:${end.toUpperCase()}` });
            continue;
          }
          // If the part after colon isn't a valid ref, treat as error
          throw new Error(`Invalid range: ${ident}:${end}`);
        }

        // Check if function call (next non-space is '(') - only when no $ is present
        if (!hasDollar) {
          let lookAhead = i;
          while (lookAhead < input.length && /\s/.test(input[lookAhead])) lookAhead++;
          if (lookAhead < input.length && input[lookAhead] === '(') {
            tokens.push({ type: 'FUNC', value: upper });
            continue;
          }
        }

        // Cell reference
        if (/^[A-Z]+\d+$/i.test(stripped)) {
          tokens.push({ type: 'REF', value: ident.toUpperCase() });
          continue;
        }

        // Unknown identifier - treat as function name or error (only without $)
        if (!hasDollar) {
          tokens.push({ type: 'FUNC', value: upper });
          continue;
        }

        throw new Error(`Unexpected identifier: ${ident}`);
      }

      throw new Error(`Unexpected character: ${ch}`);
    }

    tokens.push({ type: 'EOF', value: '' });
    return tokens;
  }

  // ─── Parser ──────────────────────────────────────────
  //
  // Parser state is passed as a mutable object through the recursive
  // descent chain so that nested evaluations (resolveRef / resolveRange)
  // each get their own independent state without save/restore.

  private _peek(s: ParserState): Token {
    return s.tokens[s.pos];
  }

  private _consume(s: ParserState, expectedType?: TokenType): Token {
    const token = s.tokens[s.pos];
    if (expectedType && token.type !== expectedType) {
      throw new Error(`Expected ${expectedType} but got ${token.type} (${token.value})`);
    }
    s.pos++;
    return token;
  }

  private parseExpression(input: string): unknown {
    if (++this._evalDepth > MAX_EVAL_DEPTH) {
      this._evalDepth--;
      throw new Error('#ERROR!');
    }
    try {
      const s: ParserState = { tokens: this.tokenize(input), pos: 0 };
      return this._parseComparison(s);
    } finally {
      this._evalDepth--;
    }
  }

  private _parseComparison(s: ParserState): unknown {
    let left = this._parseConcatenation(s);

    while (
      this._peek(s).type === 'OPERATOR' &&
      ['=', '<>', '<', '>', '<=', '>='].includes(this._peek(s).value)
    ) {
      const op = this._consume(s).value;
      const right = this._parseConcatenation(s);
      left = this._compareValues(left, right, op);
    }

    return left;
  }

  private _parseConcatenation(s: ParserState): unknown {
    let left = this._parseAddSub(s);

    while (this._peek(s).type === 'OPERATOR' && this._peek(s).value === '&') {
      this._consume(s); // &
      const right = this._parseAddSub(s);
      left = String(left) + String(right);
    }

    return left;
  }

  private _parseAddSub(s: ParserState): unknown {
    let left = this._parseMulDiv(s);

    while (
      this._peek(s).type === 'OPERATOR' &&
      (this._peek(s).value === '+' || this._peek(s).value === '-')
    ) {
      const op = this._consume(s).value;
      const right = this._parseMulDiv(s);
      if (op === '+') left = Number(left) + Number(right);
      else left = Number(left) - Number(right);
    }

    return left;
  }

  private _parseMulDiv(s: ParserState): unknown {
    let left = this._parseUnary(s);

    while (
      this._peek(s).type === 'OPERATOR' &&
      (this._peek(s).value === '*' || this._peek(s).value === '/')
    ) {
      const op = this._consume(s).value;
      const right = this._parseUnary(s);
      if (op === '*') left = Number(left) * Number(right);
      else {
        const divisor = Number(right);
        if (divisor === 0) throw new Error('#DIV/0!');
        left = Number(left) / divisor;
      }
    }

    return left;
  }

  private _parseUnary(s: ParserState): unknown {
    if (this._peek(s).type === 'OPERATOR' && this._peek(s).value === '-') {
      this._consume(s);
      return -Number(this._parsePrimary(s));
    }
    if (this._peek(s).type === 'OPERATOR' && this._peek(s).value === '+') {
      this._consume(s);
      return Number(this._parsePrimary(s));
    }
    return this._parsePrimary(s);
  }

  private _parsePrimary(s: ParserState): unknown {
    const token = this._peek(s);

    switch (token.type) {
      case 'NUMBER':
        this._consume(s);
        return parseFloat(token.value);

      case 'STRING':
        this._consume(s);
        return token.value;

      case 'BOOLEAN':
        this._consume(s);
        return token.value === 'TRUE';

      case 'REF':
        this._consume(s);
        return this._resolveRef(token.value);

      case 'RANGE':
        this._consume(s);
        return this._resolveRange(token.value);

      case 'FUNC':
        return this._parseFunction(s);

      case 'LPAREN':
        this._consume(s);
        const expr = this._parseComparison(s);
        this._consume(s, 'RPAREN');
        return expr;

      default:
        throw new Error(`Unexpected token: ${token.type} (${token.value})`);
    }
  }

  private _parseFunction(s: ParserState): unknown {
    const name = this._consume(s, 'FUNC').value;
    this._consume(s, 'LPAREN');

    const args: unknown[] = [];
    if (this._peek(s).type !== 'RPAREN') {
      args.push(this._parseComparison(s));
      while (this._peek(s).type === 'COMMA') {
        this._consume(s);
        args.push(this._parseComparison(s));
      }
    }

    this._consume(s, 'RPAREN');

    const fn = this.functions.get(name);
    if (!fn) throw new Error(`#NAME?`);

    const ctx = this._createContext();

    // For aggregate functions, flatten RangeValue into individual args
    // For other functions, pass RangeValue directly so they can access shape
    if (AGGREGATE_FUNCTIONS.has(name)) {
      const flatArgs: unknown[] = [];
      for (const arg of args) {
        if (arg instanceof RangeValue) {
          // Filter out undefined (empty cells) for aggregates
          for (const v of arg.values) {
            if (v !== undefined) flatArgs.push(v);
          }
        } else if (Array.isArray(arg)) {
          flatArgs.push(...arg);
        } else {
          flatArgs.push(arg);
        }
      }
      return fn(ctx, ...flatArgs);
    }

    return fn(ctx, ...args);
  }

  // ─── Comparison helper ────────────────────────────────

  private _compareValues(left: unknown, right: unknown, op: string): boolean {
    // Determine if both sides can be compared numerically.
    // Booleans are numeric (TRUE=1, FALSE=0). Strings that look like numbers
    // are numeric. Empty strings and non-numeric strings are NOT numeric.
    const isNumeric = (v: unknown): boolean => {
      if (typeof v === 'boolean') return true;
      if (typeof v === 'number') return true;
      const s = String(v).trim();
      return s !== '' && !isNaN(Number(s));
    };

    const bothNumeric = isNumeric(left) && isNumeric(right);

    if (bothNumeric) {
      const l = Number(left);
      const r = Number(right);
      switch (op) {
        case '=':  return l === r;
        case '<>': return l !== r;
        case '<':  return l < r;
        case '>':  return l > r;
        case '<=': return l <= r;
        case '>=': return l >= r;
        default:   return false;
      }
    }

    // String comparison (case-insensitive for = and <>, lexicographic for ordering)
    const ls = String(left ?? '');
    const rs = String(right ?? '');
    switch (op) {
      case '=':  return ls === rs;
      case '<>': return ls !== rs;
      case '<':  return ls < rs;
      case '>':  return ls > rs;
      case '<=': return ls <= rs;
      case '>=': return ls >= rs;
      default:   return false;
    }
  }

  // ─── Reference resolution ───────────────────────────
  //
  // Each nested evaluation calls parseExpression which creates its
  // own ParserState, so no save/restore is needed.

  private _resolveRef(ref: string): unknown {
    const coord = refToCoord(ref.replace(/\$/g, ''));
    const key = cellKey(coord.row, coord.col);

    this._trackDep(key);

    // Circular reference detection
    if (this.evaluating.has(key)) {
      throw new Error('#CIRC!');
    }

    const cell = this.data.get(key);
    if (!cell) return 0; // empty cells are 0

    // If this cell also has a formula, evaluate it in a fresh parser context.
    // Keep _trackingCellKey so transitive dependencies are recorded
    // (e.g. C1→B1→A1 means C1 depends on A1 too).
    if (cell.rawValue.startsWith('=')) {
      this.evaluating.add(key);
      try {
        return this.parseExpression(cell.rawValue.substring(1));
      } catch {
        throw new Error('#ERROR!');
      } finally {
        this.evaluating.delete(key);
      }
    }

    // Return the value, coerced to number if possible
    const num = Number(cell.rawValue);
    if (!isNaN(num) && cell.rawValue.trim() !== '') return num;
    if (cell.rawValue.toUpperCase() === 'TRUE') return true;
    if (cell.rawValue.toUpperCase() === 'FALSE') return false;
    return cell.rawValue;
  }

  private _resolveRange(rangeStr: string): RangeValue {
    const [startRef, endRef] = rangeStr.split(':');
    const start = refToCoord(startRef.replace(/\$/g, ''));
    const end = refToCoord(endRef.replace(/\$/g, ''));

    const minRow = Math.min(start.row, end.row);
    const maxRow = Math.max(start.row, end.row);
    const minCol = Math.min(start.col, end.col);
    const maxCol = Math.max(start.col, end.col);

    const numRows = maxRow - minRow + 1;
    const numCols = maxCol - minCol + 1;

    const values: unknown[] = [];
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const key = cellKey(r, c);

        this._trackDep(key);

        const cell = this.data.get(key);
        if (cell) {
          if (cell.rawValue.startsWith('=')) {
            // Keep _trackingCellKey so transitive deps through ranges are tracked
            if (this.evaluating.has(key)) {
              values.push('#CIRC!');
            } else {
              this.evaluating.add(key);
              try {
                values.push(this.parseExpression(cell.rawValue.substring(1)));
              } catch (e) {
                const msg = e instanceof Error ? e.message : '';
                values.push(msg.startsWith('#') ? msg : '#ERROR!');
              } finally {
                this.evaluating.delete(key);
              }
            }
          } else {
            // Match _resolveRef's coercion logic for consistency
            const num = Number(cell.rawValue);
            if (!isNaN(num) && cell.rawValue.trim() !== '') {
              values.push(num);
            } else if (cell.rawValue.toUpperCase() === 'TRUE') {
              values.push(true);
            } else if (cell.rawValue.toUpperCase() === 'FALSE') {
              values.push(false);
            } else {
              values.push(cell.rawValue);
            }
          }
        } else {
          // Empty cells contribute to shape
          values.push(undefined);
        }
      }
    }

    return new RangeValue(values, numRows, numCols);
  }

  private _createContext(): FormulaContext {
    return {
      getCellValue: (ref: string) => {
        if (ref.includes(':') && !/[A-Z]/i.test(ref.charAt(0))) {
          // It's a key like "0:0"
          const cell = this.data.get(ref);
          if (!cell) return 0;
          const num = Number(cell.rawValue);
          return !isNaN(num) && cell.rawValue.trim() !== '' ? num : cell.rawValue;
        }
        return this._resolveRef(ref);
      },
      getRangeValues: (startRef: string, endRef: string) => {
        const start = refToCoord(startRef);
        const end = refToCoord(endRef);

        const minRow = Math.min(start.row, end.row);
        const maxRow = Math.max(start.row, end.row);
        const minCol = Math.min(start.col, end.col);
        const maxCol = Math.max(start.col, end.col);

        const values: unknown[] = [];
        for (let r = minRow; r <= maxRow; r++) {
          for (let c = minCol; c <= maxCol; c++) {
            const key = cellKey(r, c);
            const cell = this.data.get(key);
            if (cell) {
              const num = Number(cell.rawValue);
              values.push(!isNaN(num) && cell.rawValue.trim() !== '' ? num : cell.rawValue);
            }
          }
        }
        return values;
      },
    };
  }

  // ─── Built-in Functions ─────────────────────────────

  private registerBuiltins(): void {
    this.registerFunction('SUM', (_ctx, ...args) => {
      return args.reduce((sum: number, v) => sum + (Number(v) || 0), 0);
    });

    this.registerFunction('AVERAGE', (_ctx, ...args) => {
      const nums = args.filter((v) => typeof v === 'number' || !isNaN(Number(v)));
      if (nums.length === 0) return 0;
      const sum = nums.reduce((s: number, v) => s + Number(v), 0);
      return sum / nums.length;
    });

    this.registerFunction('MIN', (_ctx, ...args) => {
      const nums = args.map(Number).filter((n) => !isNaN(n));
      return nums.length ? Math.min(...nums) : 0;
    });

    this.registerFunction('MAX', (_ctx, ...args) => {
      const nums = args.map(Number).filter((n) => !isNaN(n));
      return nums.length ? Math.max(...nums) : 0;
    });

    this.registerFunction('COUNT', (_ctx, ...args) => {
      return args.filter((v) => typeof v === 'number' || (!isNaN(Number(v)) && String(v).trim() !== '')).length;
    });

    this.registerFunction('COUNTA', (_ctx, ...args) => {
      return args.filter((v) => v !== '' && v !== null && v !== undefined).length;
    });

    this.registerFunction('IF', (_ctx, condition, trueVal, falseVal) => {
      return condition ? trueVal : falseVal;
    });

    this.registerFunction('CONCAT', (_ctx, ...args) => {
      return args.map(String).join('');
    });

    this.registerFunction('ABS', (_ctx, val) => {
      return Math.abs(Number(val));
    });

    this.registerFunction('ROUND', (_ctx, val, digits) => {
      const d = Number(digits) || 0;
      const factor = Math.pow(10, d);
      return Math.round(Number(val) * factor) / factor;
    });

    this.registerFunction('UPPER', (_ctx, val) => {
      return String(val).toUpperCase();
    });

    this.registerFunction('LOWER', (_ctx, val) => {
      return String(val).toLowerCase();
    });

    this.registerFunction('LEN', (_ctx, val) => {
      return String(val).length;
    });

    this.registerFunction('TRIM', (_ctx, val) => {
      return String(val).trim();
    });

    // ─── Logic/Conditional ────────────────────────────────

    this.registerFunction('IFERROR', (_ctx, value, fallback) => {
      const s = String(value);
      if (
        s === '#ERROR!' || s === '#REF!' || s === '#DIV/0!' ||
        s === '#NAME?' || s === '#CIRC!' || s === '#VALUE!'
      ) {
        return fallback;
      }
      return value;
    });

    this.registerFunction('AND', (_ctx, ...args) => {
      for (const arg of args) {
        if (arg instanceof RangeValue) {
          for (const v of arg.values) {
            if (v !== undefined && !v) return false;
          }
        } else {
          if (!arg) return false;
        }
      }
      return true;
    });

    this.registerFunction('OR', (_ctx, ...args) => {
      for (const arg of args) {
        if (arg instanceof RangeValue) {
          for (const v of arg.values) {
            if (v !== undefined && v) return true;
          }
        } else {
          if (arg) return true;
        }
      }
      return false;
    });

    this.registerFunction('NOT', (_ctx, val) => {
      return !val;
    });

    // ─── Conditional Aggregation ─────────────────────────

    this.registerFunction('SUMIF', (_ctx, range, criteria, sumRange?) => {
      const criteriaStr = String(criteria);
      const rangeArr: unknown[] = range instanceof RangeValue ? range.values : (Array.isArray(range) ? range : [range]);
      const sumArr: unknown[] | undefined = sumRange instanceof RangeValue ? sumRange.values : (Array.isArray(sumRange) ? sumRange : undefined);

      let total = 0;
      for (let i = 0; i < rangeArr.length; i++) {
        const val = rangeArr[i];
        if (val !== undefined && matchesCriteria(val, criteriaStr)) {
          if (sumArr) {
            total += Number(sumArr[i]) || 0;
          } else {
            total += Number(val) || 0;
          }
        }
      }
      return total;
    });

    this.registerFunction('COUNTIF', (_ctx, range, criteria) => {
      const criteriaStr = String(criteria);
      const rangeArr: unknown[] = range instanceof RangeValue ? range.values : (Array.isArray(range) ? range : [range]);

      let count = 0;
      for (const val of rangeArr) {
        if (val !== undefined && matchesCriteria(val, criteriaStr)) {
          count++;
        }
      }
      return count;
    });

    // ─── Lookup ─────────────────────────────────────────

    this.registerFunction('VLOOKUP', (_ctx, lookupValue, tableRange, colIndex, exactMatch?) => {
      if (!(tableRange instanceof RangeValue)) {
        throw new Error('#VALUE!');
      }
      const colIdx = Number(colIndex);
      if (colIdx < 1 || colIdx > tableRange.cols) {
        throw new Error('#REF!');
      }
      const isExact = exactMatch === undefined || exactMatch === true || exactMatch === 0;

      // Search first column
      const firstCol = tableRange.getCol(0);
      for (let r = 0; r < firstCol.length; r++) {
        if (isExact) {
          const numLookup = Number(lookupValue);
          const numCell = Number(firstCol[r]);
          const bothNum = !isNaN(numLookup) && !isNaN(numCell)
            && String(lookupValue).trim() !== '' && String(firstCol[r]).trim() !== '';
          if (bothNum ? numLookup === numCell : String(firstCol[r]).toLowerCase() === String(lookupValue).toLowerCase()) {
            return tableRange.get(r, colIdx - 1);
          }
        } else {
          // Approximate match: find largest value <= lookupValue
          // Data assumed sorted ascending
          if (Number(firstCol[r]) > Number(lookupValue)) {
            if (r === 0) throw new Error('#N/A');
            return tableRange.get(r - 1, colIdx - 1);
          }
        }
      }
      if (!isExact && firstCol.length > 0) {
        return tableRange.get(firstCol.length - 1, colIdx - 1);
      }
      throw new Error('#N/A');
    });

    this.registerFunction('HLOOKUP', (_ctx, lookupValue, tableRange, rowIndex, exactMatch?) => {
      if (!(tableRange instanceof RangeValue)) {
        throw new Error('#VALUE!');
      }
      const rowIdx = Number(rowIndex);
      if (rowIdx < 1 || rowIdx > tableRange.rows) {
        throw new Error('#REF!');
      }
      const isExact = exactMatch === undefined || exactMatch === true || exactMatch === 0;

      // Search first row
      const firstRow = tableRange.getRow(0);
      for (let c = 0; c < firstRow.length; c++) {
        if (isExact) {
          const numLookup = Number(lookupValue);
          const numCell = Number(firstRow[c]);
          const bothNum = !isNaN(numLookup) && !isNaN(numCell)
            && String(lookupValue).trim() !== '' && String(firstRow[c]).trim() !== '';
          if (bothNum ? numLookup === numCell : String(firstRow[c]).toLowerCase() === String(lookupValue).toLowerCase()) {
            return tableRange.get(rowIdx - 1, c);
          }
        } else {
          if (Number(firstRow[c]) > Number(lookupValue)) {
            if (c === 0) throw new Error('#N/A');
            return tableRange.get(rowIdx - 1, c - 1);
          }
        }
      }
      if (!isExact && firstRow.length > 0) {
        return tableRange.get(rowIdx - 1, firstRow.length - 1);
      }
      throw new Error('#N/A');
    });

    this.registerFunction('INDEX', (_ctx, rangeArg, rowNum, colNum?) => {
      if (rangeArg instanceof RangeValue) {
        const r = Number(rowNum) - 1;
        const c = colNum !== undefined ? Number(colNum) - 1 : 0;
        if (r < 0 || r >= rangeArg.rows || c < 0 || c >= rangeArg.cols) {
          throw new Error('#REF!');
        }
        const val = rangeArg.get(r, c);
        return val !== undefined ? val : 0;
      }
      throw new Error('#VALUE!');
    });

    this.registerFunction('MATCH', (_ctx, lookupValue, rangeArg, matchType?) => {
      let arr: unknown[];
      if (rangeArg instanceof RangeValue) {
        // Use flat values array for 1D lookup
        arr = rangeArg.values;
      } else if (Array.isArray(rangeArg)) {
        arr = rangeArg;
      } else {
        throw new Error('#VALUE!');
      }

      const mt = matchType !== undefined ? Number(matchType) : 0;

      if (mt === 0) {
        // Exact match
        for (let i = 0; i < arr.length; i++) {
          const numLookup = Number(lookupValue);
          const numCell = Number(arr[i]);
          const bothNum = !isNaN(numLookup) && !isNaN(numCell)
            && String(lookupValue).trim() !== '' && String(arr[i]).trim() !== '';
          if (bothNum ? numLookup === numCell : String(arr[i]).toLowerCase() === String(lookupValue).toLowerCase()) {
            return i + 1; // 1-indexed
          }
        }
        throw new Error('#N/A');
      } else if (mt === 1) {
        // Largest value <= lookupValue (data assumed sorted ascending)
        let lastMatch = -1;
        for (let i = 0; i < arr.length; i++) {
          if (Number(arr[i]) <= Number(lookupValue)) {
            lastMatch = i;
          }
        }
        if (lastMatch === -1) throw new Error('#N/A');
        return lastMatch + 1;
      } else {
        // mt === -1: Smallest value >= lookupValue (data assumed sorted descending)
        let lastMatch = -1;
        for (let i = 0; i < arr.length; i++) {
          if (Number(arr[i]) >= Number(lookupValue)) {
            lastMatch = i;
          }
        }
        if (lastMatch === -1) throw new Error('#N/A');
        return lastMatch + 1;
      }
    });

    // ─── Math ───────────────────────────────────────────

    this.registerFunction('MOD', (_ctx, num, divisor) => {
      const d = Number(divisor);
      if (d === 0) throw new Error('#DIV/0!');
      return Number(num) % d;
    });

    this.registerFunction('POWER', (_ctx, base, exp) => {
      return Math.pow(Number(base), Number(exp));
    });

    this.registerFunction('CEILING', (_ctx, num, significance?) => {
      const n = Number(num);
      const sig = significance !== undefined ? Number(significance) : 1;
      if (sig === 0) return 0;
      return Math.ceil(n / sig) * sig;
    });

    this.registerFunction('FLOOR', (_ctx, num, significance?) => {
      const n = Number(num);
      const sig = significance !== undefined ? Number(significance) : 1;
      if (sig === 0) return 0;
      return Math.floor(n / sig) * sig;
    });

    // ─── String ─────────────────────────────────────────

    this.registerFunction('LEFT', (_ctx, text, n?) => {
      const count = n !== undefined ? Number(n) : 1;
      return String(text).substring(0, count);
    });

    this.registerFunction('RIGHT', (_ctx, text, n?) => {
      const s = String(text);
      const count = n !== undefined ? Number(n) : 1;
      return s.substring(s.length - count);
    });

    this.registerFunction('MID', (_ctx, text, start, n) => {
      const s = String(text);
      const startIdx = Number(start) - 1; // 1-indexed to 0-indexed
      const count = Number(n);
      return s.substring(startIdx, startIdx + count);
    });

    this.registerFunction('SUBSTITUTE', (_ctx, text, oldStr, newStr, instance?) => {
      const s = String(text);
      const old = String(oldStr);
      const replacement = String(newStr);

      if (instance !== undefined) {
        const nth = Number(instance);
        let count = 0;
        let idx = -1;
        let searchFrom = 0;
        while (searchFrom < s.length) {
          idx = s.indexOf(old, searchFrom);
          if (idx === -1) break;
          count++;
          if (count === nth) {
            return s.substring(0, idx) + replacement + s.substring(idx + old.length);
          }
          searchFrom = idx + 1;
        }
        return s; // nth occurrence not found, return unchanged
      }

      // Replace all occurrences
      return s.split(old).join(replacement);
    });

    this.registerFunction('FIND', (_ctx, search, text, start?) => {
      const s = String(text);
      const searchStr = String(search);
      const startIdx = start !== undefined ? Number(start) - 1 : 0;
      const idx = s.indexOf(searchStr, startIdx);
      if (idx === -1) throw new Error('#VALUE!');
      return idx + 1; // 1-indexed
    });

    // ─── Conversion ─────────────────────────────────────

    this.registerFunction('TEXT', (_ctx, value, format) => {
      const num = Number(value);
      const fmt = String(format);

      if (isNaN(num)) return String(value);

      // Support common number formats
      if (fmt === '0') {
        return Math.round(num).toString();
      }
      if (fmt === '0.00' || fmt === '0.0') {
        const decimals = (fmt.split('.')[1] || '').length;
        return num.toFixed(decimals);
      }
      if (fmt === '#,##0' || fmt === '#,##0.00') {
        const decimals = fmt.includes('.') ? (fmt.split('.')[1] || '').length : 0;
        return num.toLocaleString('en-US', {
          minimumFractionDigits: decimals,
          maximumFractionDigits: decimals,
        });
      }

      return num.toString();
    });

    this.registerFunction('VALUE', (_ctx, text) => {
      const num = Number(String(text));
      if (isNaN(num)) throw new Error('#VALUE!');
      return num;
    });

    // ─── Date ───────────────────────────────────────────

    this.registerFunction('DATE', (_ctx, year, month, day) => {
      // Excel serial number: days since 1899-12-30
      const y = Number(year);
      const m = Number(month);
      const d = Number(day);
      const date = new Date(y, m - 1, d);
      const epoch = new Date(1899, 11, 30); // 1899-12-30
      const diff = date.getTime() - epoch.getTime();
      return Math.round(diff / (1000 * 60 * 60 * 24));
    });

    this.registerFunction('NOW', () => {
      const now = new Date();
      const epoch = new Date(1899, 11, 30); // 1899-12-30
      const diff = now.getTime() - epoch.getTime();
      return Math.round(diff / (1000 * 60 * 60 * 24));
    });
  }
}
