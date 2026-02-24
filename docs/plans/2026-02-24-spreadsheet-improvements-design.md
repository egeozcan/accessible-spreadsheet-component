# Spreadsheet Improvements Design

**Date**: 2026-02-24
**Status**: Approved

## Goal

Improve the y11n-spreadsheet component across three areas simultaneously using 3 parallel agents: comprehensive unit test coverage, formula engine feature expansion, and clipboard/performance/stories improvements.

## Approach

3 agents working in isolated git worktrees to minimize merge conflicts. Each agent owns specific files.

---

## Agent 1: Unit Test Coverage

**Files owned**: `src/__tests__/` (new test files only)

### New test files

#### `src/__tests__/formula-engine.test.ts`
- **Tokenizer**: Numbers, strings, booleans, cell refs (A1, AA1), ranges (A1:B2), operators, function names, EOF, edge cases (empty input, multi-letter columns)
- **Parser**: Operator precedence (1+2*3=7), parentheses, unary negation, comparisons, string concatenation (`&`), nested functions
- **Evaluator**: Cell ref resolution, range resolution, all 13 built-in functions with normal and edge-case inputs (empty ranges, non-numeric values, division by zero)
- **Dependency tracking**: Forward/reverse graph construction, `recalculateAffected()` BFS traversal, transitive deps (A1->B1->C1), `_clearDepsFor()` cleanup
- **Error handling**: `#CIRC!` (circular refs), `#DIV/0!`, `#REF!` (invalid refs), `#NAME?` (unknown functions), `#ERROR!` (generic), max eval depth (64)

#### `src/__tests__/selection-manager.test.ts`
- Move by delta (arrow keys), move to absolute position, extend selection with Shift
- Bounds clamping at grid edges
- `selectAll()`, `clearSelection()`
- Drag start/extend/end
- `isCellSelected()` / `isCellActive()` queries
- Range normalization (start < end always)

#### `src/__tests__/clipboard-manager.test.ts`
- **TSV serialization**: Normal values, values with tabs/newlines/quotes (quoted fields), empty cells
- **TSV parsing**: Windows (`\r\n`), Unix (`\n`), Mac (`\r`) line endings, quoted fields with escaped quotes (`""`), empty cells, single cell
- **Copy/cut/paste workflows**: Range serialization, clipboard API interaction

#### `src/__tests__/formula-bar.test.ts`
- Mode toggle (raw <-> formatted)
- Draft value management (persists across mode changes)
- Commit on Enter, commit on blur
- Cancel on Escape with blur suppression
- Read-only mode (editing disabled)

---

## Agent 2: Formula Engine Features

**Files owned**: `src/engine/formula-engine.ts`, `e2e/formulas.spec.ts`

### Absolute/Mixed References

**Syntax**: `$A$1` (both absolute), `$A1` (absolute column), `A$1` (absolute row)

**Approach**: Keep `$` in the string representation within REF tokens. The tokenizer already extracts the ref string; we just need to not strip `$` and handle it in `_resolveRef()`. Absolute vs relative only matters during copy/paste offset adjustment.

- Extend tokenizer to recognize `$` before column letters and row numbers in REF tokens
- In ClipboardManager paste, adjust relative references by paste offset but leave absolute references unchanged
- Add F4 key toggle in main component's `_handleRefArrow()` for cycling through reference modes

### New Formula Functions (~22)

| Function | Category | Signature |
|----------|----------|-----------|
| VLOOKUP | Lookup | `VLOOKUP(value, range, col_idx, [exact])` |
| HLOOKUP | Lookup | `HLOOKUP(value, range, row_idx, [exact])` |
| INDEX | Lookup | `INDEX(range, row, [col])` |
| MATCH | Lookup | `MATCH(value, range, [type])` |
| SUMIF | Conditional | `SUMIF(range, criteria, [sum_range])` |
| COUNTIF | Conditional | `COUNTIF(range, criteria)` |
| IFERROR | Logic | `IFERROR(value, fallback)` |
| AND | Logic | `AND(val1, val2, ...)` |
| OR | Logic | `OR(val1, val2, ...)` |
| NOT | Logic | `NOT(val)` |
| MOD | Math | `MOD(num, divisor)` |
| POWER | Math | `POWER(base, exp)` |
| CEILING | Math | `CEILING(num, [significance])` |
| FLOOR | Math | `FLOOR(num, [significance])` |
| LEFT | String | `LEFT(text, [n])` |
| RIGHT | String | `RIGHT(text, [n])` |
| MID | String | `MID(text, start, n)` |
| SUBSTITUTE | String | `SUBSTITUTE(text, old, new, [instance])` |
| FIND | String | `FIND(search, text, [start])` |
| TEXT | Conversion | `TEXT(value, format)` |
| VALUE | Conversion | `VALUE(text)` |
| DATE | Date | `DATE(year, month, day)` |
| NOW | Date | `NOW()` |

Each function registered in the engine's `_functions` map with argument validation and proper error codes.

### E2E Tests
- Tests for each new function with normal inputs and edge cases
- Tests for absolute/mixed references in formulas
- Tests for reference behavior during paste operations

---

## Agent 3: Clipboard + Performance + Stories

**Files owned**: `src/controllers/clipboard-manager.ts`, `stories/`, performance changes in engine

### HTML Table Clipboard Support

- Add `parseHTMLTable(html: string): string[][]` to ClipboardManager using `DOMParser`
- In `paste()`, try `text/html` MIME type first, parse `<table>` elements, fall back to TSV
- In `copy()`, write both `text/html` (as `<table>`) and `text/plain` (as TSV) for interoperability

### Formula Result Caching

- Add `_cache: Map<string, CellData>` to FormulaEngine
- Check cache before parsing/evaluating in `evaluate()`
- Invalidate cache entries during BFS traversal in `recalculateAffected()`
- Clear cache in `_clearDepsFor()` and on full `recalculate()`

### Smarter `setData()`

- Diff new data against old `_internalData`
- Collect changed keys into a set
- Use `recalculateAffected(changedKeys)` instead of full `recalculate()`
- Only rebuild deps for cells whose formulas changed

### New Storybook Stories

- Absolute References Demo
- Lookup Functions Demo (VLOOKUP/INDEX/MATCH)
- Conditional Aggregation Demo (SUMIF/COUNTIF)
- Logic Functions Demo (AND/OR/NOT/IFERROR)
- String Functions Demo (LEFT/RIGHT/MID/SUBSTITUTE/FIND)
- HTML Paste Demo (instructions to paste from Excel/Google Sheets)
