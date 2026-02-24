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
    if (val === '#ERROR!' || val === '#REF!' || val === '#DIV/0!' || val === '#NAME?' || val === '#CIRC!') {
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
      if (/[A-Za-z_]/.test(ch)) {
        let ident = '';
        while (i < input.length && /[A-Za-z0-9_]/.test(input[i])) {
          ident += input[i];
          i++;
        }

        const upper = ident.toUpperCase();

        // Check if boolean
        if (upper === 'TRUE' || upper === 'FALSE') {
          tokens.push({ type: 'BOOLEAN', value: upper });
          continue;
        }

        // Check for range (e.g. A1:B2) - look ahead for colon
        if (i < input.length && input[i] === ':' && /^[A-Z]+\d+$/i.test(ident)) {
          i++; // skip colon
          let end = '';
          while (i < input.length && /[A-Za-z0-9]/.test(input[i])) {
            end += input[i];
            i++;
          }
          if (/^[A-Z]+\d+$/i.test(end)) {
            tokens.push({ type: 'RANGE', value: `${ident.toUpperCase()}:${end.toUpperCase()}` });
            continue;
          }
          // If the part after colon isn't a valid ref, treat as error
          throw new Error(`Invalid range: ${ident}:${end}`);
        }

        // Check if function call (next non-space is '(')
        let lookAhead = i;
        while (lookAhead < input.length && /\s/.test(input[lookAhead])) lookAhead++;
        if (lookAhead < input.length && input[lookAhead] === '(') {
          tokens.push({ type: 'FUNC', value: upper });
          continue;
        }

        // Cell reference
        if (/^[A-Z]+\d+$/i.test(ident)) {
          tokens.push({ type: 'REF', value: ident.toUpperCase() });
          continue;
        }

        // Unknown identifier - treat as function name or error
        tokens.push({ type: 'FUNC', value: upper });
        continue;
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

    // Flatten range arrays into args for aggregate functions
    const flatArgs: unknown[] = [];
    for (const arg of args) {
      if (Array.isArray(arg)) {
        flatArgs.push(...arg);
      } else {
        flatArgs.push(arg);
      }
    }

    return fn(ctx, ...flatArgs);
  }

  // ─── Comparison helper ────────────────────────────────

  private _compareValues(left: unknown, right: unknown, op: string): boolean {
    const l = Number(left);
    const r = Number(right);
    const bothNumeric = !isNaN(l) && !isNaN(r)
      && String(left).trim() !== '' && String(right).trim() !== '';

    switch (op) {
      case '=':
        return left === right || (bothNumeric && l === r);
      case '<>':
        return left !== right && (!bothNumeric || l !== r);
      case '<':
        return bothNumeric ? l < r : String(left) < String(right);
      case '>':
        return bothNumeric ? l > r : String(left) > String(right);
      case '<=':
        return bothNumeric ? l <= r : String(left) <= String(right);
      case '>=':
        return bothNumeric ? l >= r : String(left) >= String(right);
      default:
        return false;
    }
  }

  // ─── Reference resolution ───────────────────────────
  //
  // Each nested evaluation calls parseExpression which creates its
  // own ParserState, so no save/restore is needed.

  private _resolveRef(ref: string): unknown {
    const coord = refToCoord(ref);
    const key = cellKey(coord.row, coord.col);

    this._trackDep(key);

    // Circular reference detection
    if (this.evaluating.has(key)) {
      throw new Error('#CIRC!');
    }

    const cell = this.data.get(key);
    if (!cell) return 0; // empty cells are 0

    // If this cell also has a formula, evaluate it in a fresh parser context
    if (cell.rawValue.startsWith('=')) {
      const savedTracking = this._trackingCellKey;
      this._trackingCellKey = null; // only track direct dependencies
      this.evaluating.add(key);
      try {
        return this.parseExpression(cell.rawValue.substring(1));
      } catch {
        throw new Error('#ERROR!');
      } finally {
        this._trackingCellKey = savedTracking;
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

  private _resolveRange(rangeStr: string): unknown[] {
    const [startRef, endRef] = rangeStr.split(':');
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

        this._trackDep(key);

        const cell = this.data.get(key);
        if (cell) {
          if (cell.rawValue.startsWith('=')) {
            const savedTracking = this._trackingCellKey;
            this._trackingCellKey = null;
            this.evaluating.add(key);
            try {
              values.push(this.parseExpression(cell.rawValue.substring(1)));
            } catch {
              values.push(0);
            } finally {
              this._trackingCellKey = savedTracking;
              this.evaluating.delete(key);
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
        }
      }
    }

    return values;
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
  }
}
