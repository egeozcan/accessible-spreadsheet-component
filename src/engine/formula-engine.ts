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
export class FormulaEngine {
  private functions: Map<string, FormulaFunction> = new Map();
  private data: GridData = new Map();
  private evaluating: Set<string> = new Set(); // circular reference detection

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
  }

  /**
   * Evaluate a raw value. If it starts with '=', parse and execute the formula.
   * Otherwise, return the value as-is (possibly coerced).
   */
  evaluate(rawValue: string): { displayValue: string; type: 'text' | 'number' | 'boolean' | 'error' } {
    if (!rawValue || rawValue.trim() === '') {
      return { displayValue: '', type: 'text' };
    }

    if (!rawValue.startsWith('=')) {
      return this.coerceValue(rawValue);
    }

    try {
      this.evaluating.clear();
      const formula = rawValue.substring(1);
      const result = this.parseExpression(formula);
      return this.coerceValue(String(result));
    } catch {
      return { displayValue: '#ERROR!', type: 'error' };
    }
  }

  /**
   * Re-evaluate all formula cells in the grid.
   * Returns a set of cell keys that changed.
   */
  recalculate(): Set<string> {
    const changed = new Set<string>();

    for (const [key, cell] of this.data) {
      if (cell.rawValue.startsWith('=')) {
        const result = this.evaluate(cell.rawValue);
        if (cell.displayValue !== result.displayValue || cell.type !== result.type) {
          cell.displayValue = result.displayValue;
          cell.type = result.type;
          changed.add(key);
        }
      }
    }

    return changed;
  }

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

  private tokens: Token[] = [];
  private pos = 0;

  private peek(): Token {
    return this.tokens[this.pos];
  }

  private consume(expectedType?: TokenType): Token {
    const token = this.tokens[this.pos];
    if (expectedType && token.type !== expectedType) {
      throw new Error(`Expected ${expectedType} but got ${token.type} (${token.value})`);
    }
    this.pos++;
    return token;
  }

  private parseExpression(input: string): unknown {
    this.tokens = this.tokenize(input);
    this.pos = 0;
    const result = this.parseComparison();
    return result;
  }

  private parseComparison(): unknown {
    let left = this.parseConcatenation();

    while (
      this.peek().type === 'OPERATOR' &&
      ['=', '<>', '<', '>', '<=', '>='].includes(this.peek().value)
    ) {
      const op = this.consume().value;
      const right = this.parseConcatenation();
      const l = Number(left);
      const r = Number(right);

      switch (op) {
        case '=':
          left = left === right || (!isNaN(l) && !isNaN(r) && l === r);
          break;
        case '<>':
          left = left !== right && (isNaN(l) || isNaN(r) || l !== r);
          break;
        case '<':
          left = l < r;
          break;
        case '>':
          left = l > r;
          break;
        case '<=':
          left = l <= r;
          break;
        case '>=':
          left = l >= r;
          break;
      }
    }

    return left;
  }

  private parseConcatenation(): unknown {
    let left = this.parseAddSub();

    while (this.peek().type === 'OPERATOR' && this.peek().value === '&') {
      this.consume(); // &
      const right = this.parseAddSub();
      left = String(left) + String(right);
    }

    return left;
  }

  private parseAddSub(): unknown {
    let left = this.parseMulDiv();

    while (
      this.peek().type === 'OPERATOR' &&
      (this.peek().value === '+' || this.peek().value === '-')
    ) {
      const op = this.consume().value;
      const right = this.parseMulDiv();
      if (op === '+') left = Number(left) + Number(right);
      else left = Number(left) - Number(right);
    }

    return left;
  }

  private parseMulDiv(): unknown {
    let left = this.parseUnary();

    while (
      this.peek().type === 'OPERATOR' &&
      (this.peek().value === '*' || this.peek().value === '/')
    ) {
      const op = this.consume().value;
      const right = this.parseUnary();
      if (op === '*') left = Number(left) * Number(right);
      else {
        const divisor = Number(right);
        if (divisor === 0) throw new Error('#DIV/0!');
        left = Number(left) / divisor;
      }
    }

    return left;
  }

  private parseUnary(): unknown {
    if (this.peek().type === 'OPERATOR' && this.peek().value === '-') {
      this.consume();
      return -Number(this.parsePrimary());
    }
    if (this.peek().type === 'OPERATOR' && this.peek().value === '+') {
      this.consume();
      return Number(this.parsePrimary());
    }
    return this.parsePrimary();
  }

  private parsePrimary(): unknown {
    const token = this.peek();

    switch (token.type) {
      case 'NUMBER':
        this.consume();
        return parseFloat(token.value);

      case 'STRING':
        this.consume();
        return token.value;

      case 'BOOLEAN':
        this.consume();
        return token.value === 'TRUE';

      case 'REF':
        this.consume();
        return this.resolveRef(token.value);

      case 'RANGE':
        this.consume();
        return this.resolveRange(token.value);

      case 'FUNC':
        return this.parseFunction();

      case 'LPAREN':
        this.consume();
        const expr = this.parseComparison();
        this.consume('RPAREN');
        return expr;

      default:
        throw new Error(`Unexpected token: ${token.type} (${token.value})`);
    }
  }

  private parseFunction(): unknown {
    const name = this.consume('FUNC').value;
    this.consume('LPAREN');

    const args: unknown[] = [];
    if (this.peek().type !== 'RPAREN') {
      args.push(this.parseComparison());
      while (this.peek().type === 'COMMA') {
        this.consume();
        args.push(this.parseComparison());
      }
    }

    this.consume('RPAREN');

    const fn = this.functions.get(name);
    if (!fn) throw new Error(`#NAME?`);

    const ctx = this.createContext();

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

  // ─── Reference resolution ───────────────────────────

  private resolveRef(ref: string): unknown {
    const coord = refToCoord(ref);
    const key = cellKey(coord.row, coord.col);

    // Circular reference detection
    if (this.evaluating.has(key)) {
      throw new Error('#CIRC!');
    }

    const cell = this.data.get(key);
    if (!cell) return 0; // empty cells are 0

    // If this cell also has a formula, evaluate it
    if (cell.rawValue.startsWith('=')) {
      this.evaluating.add(key);
      try {
        // Save and restore parser state so nested evaluation doesn't corrupt
        // the token stream of the calling formula (same pattern as resolveRange)
        const saved = { tokens: [...this.tokens], pos: this.pos };
        const result = this.parseExpression(cell.rawValue.substring(1));
        this.tokens = saved.tokens;
        this.pos = saved.pos;
        this.evaluating.delete(key);
        return result;
      } catch {
        this.evaluating.delete(key);
        throw new Error('#ERROR!');
      }
    }

    // Return the value, coerced to number if possible
    const num = Number(cell.rawValue);
    if (!isNaN(num) && cell.rawValue.trim() !== '') return num;
    if (cell.rawValue.toUpperCase() === 'TRUE') return true;
    if (cell.rawValue.toUpperCase() === 'FALSE') return false;
    return cell.rawValue;
  }

  private resolveRange(rangeStr: string): unknown[] {
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
        const cell = this.data.get(key);
        if (cell) {
          if (cell.rawValue.startsWith('=')) {
            this.evaluating.add(key);
            try {
              const saved = { tokens: [...this.tokens], pos: this.pos };
              values.push(this.parseExpression(cell.rawValue.substring(1)));
              this.tokens = saved.tokens;
              this.pos = saved.pos;
            } catch {
              values.push(0);
            }
            this.evaluating.delete(key);
          } else {
            const num = Number(cell.rawValue);
            if (!isNaN(num) && cell.rawValue.trim() !== '') {
              values.push(num);
            } else {
              values.push(cell.rawValue);
            }
          }
        }
      }
    }

    return values;
  }

  private createContext(): FormulaContext {
    return {
      getCellValue: (ref: string) => {
        if (ref.includes(':') && !/[A-Z]/i.test(ref.charAt(0))) {
          // It's a key like "0:0"
          const cell = this.data.get(ref);
          if (!cell) return 0;
          const num = Number(cell.rawValue);
          return !isNaN(num) && cell.rawValue.trim() !== '' ? num : cell.rawValue;
        }
        return this.resolveRef(ref);
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
