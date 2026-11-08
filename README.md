# y11n-spreadsheet

An accessible spreadsheet web component built with [Lit 3.0](https://lit.dev). Implements the [WAI-ARIA grid pattern](https://www.w3.org/WAI/ARIA/apg/patterns/grid/) with virtual rendering, a formula engine, undo/redo, and clipboard support.

## Features

- Full keyboard navigation (arrows, Ctrl+Arrow data-edge jumps, Tab, Enter, Escape, Ctrl+A)
- Row/column selection via headers (click, Shift+click, drag) or Ctrl+Space / Shift+Space
- Fill down/right (Ctrl+D / Ctrl+R), a draggable fill handle that continues numeric series, and multi-cell entry (Ctrl+Enter)
- Insert and delete rows/columns, with formula references rewritten to follow their cells
- Find (Ctrl+F) with match count and next/previous navigation
- Per-cell read-only flag, value validation, and cancelable `cell-edit-start` / `before-change` events
- CSV/TSV import and export
- Copy/cut/paste with HTML and TSV clipboard formats
- Undo/redo with Ctrl+Z / Ctrl+Shift+Z (or Ctrl+Y)
- Formula engine with 50+ built-in functions and custom function support
- Arrow key cell reference insertion during formula editing
- Virtual rendering for large datasets (1000+ rows)
- Dual-mode formula bar (raw / formatted)
- Read-only mode
- CSS custom properties for theming
- Shadow DOM encapsulation
- Screen reader support with ARIA live regions

## Install

```bash
npm install y11n-spreadsheet
```

Lit is a peer dependency:

```bash
npm install lit
```

## Quick start

```html
<script type="module">
  import 'y11n-spreadsheet';
</script>

<y11n-spreadsheet rows="50" cols="10"></y11n-spreadsheet>
```

## Usage

### Basic grid with data

```js
import 'y11n-spreadsheet';
import { cellKey } from 'y11n-spreadsheet';

const sheet = document.querySelector('y11n-spreadsheet');

const data = new Map();
data.set(cellKey(0, 0), { rawValue: 'Name',  displayValue: 'Name',  type: 'text' });
data.set(cellKey(0, 1), { rawValue: 'Score', displayValue: 'Score', type: 'text' });
data.set(cellKey(1, 0), { rawValue: 'Alice', displayValue: 'Alice', type: 'text' });
data.set(cellKey(1, 1), { rawValue: '95',    displayValue: '95',    type: 'number' });
data.set(cellKey(2, 1), { rawValue: '=SUM(B1:B2)', displayValue: '95', type: 'number' });

sheet.setData(data);
```

### Properties

| Property | Attribute | Type | Default | Description |
|----------|-----------|------|---------|-------------|
| `rows` | `rows` | `number` | `100` | Number of rows |
| `cols` | `cols` | `number` | `26` | Number of columns |
| `data` | - | `GridData` | `new Map()` | Sparse cell data (`Map<string, CellData>`) |
| `readOnly` | `read-only` | `boolean` | `false` | Disables editing |
| `functions` | - | `Record<string, FormulaFunction>` | `{}` | Custom formula functions |
| `validateCell` | - | `(cellId, value) => string \| null \| undefined` | - | Validates user-entered values; return a message to reject one (see [Validation and edit control](#validation-and-edit-control)) |

### Methods

```js
sheet.getData()          // Returns a copy of the grid data
sheet.setData(data)      // Replaces grid data (clears undo/redo)
sheet.registerFunction('DOUBLE', (ctx, val) => Number(val) * 2)

// Formatting
sheet.setCellFormat('0:0', { bold: true, verticalAlign: 'top', wrap: true })
sheet.setRangeFormat({ start: { row: 1, col: 1 }, end: { row: 9, col: 1 } },
  { numberFormat: { type: 'date', datePattern: 'DD/MM/YYYY' } })

// CSV / TSV
sheet.toCSV()                                   // display values, A1 → last used cell
sheet.toCSV({ values: 'raw', delimiter: '\t' }) // formulas, as TSV
sheet.toCSV({ range: { start: { row: 0, col: 0 }, end: { row: 9, col: 3 } } })
sheet.fromCSV(text)                             // write at A1 (undoable)
sheet.fromCSV(text, { start: { row: 5, col: 0 }, delimiter: ';' })

// Rows and columns (undoable; the grid grows/shrinks)
sheet.insertRows(2)        // insert one row before row 3
sheet.insertRows(0, 5)     // insert five rows at the top
sheet.deleteRows(4, 2)     // delete rows 5–6
sheet.insertColumns(1)     // insert a column before B
sheet.deleteColumns(3)     // delete column D

// Find
sheet.findCells('total')                      // ['5:0', ...] in reading order
sheet.findCells('SUM', { formulas: true, matchCase: true })
sheet.openFind()                              // open the find bar (Ctrl/Cmd+F)
sheet.closeFind()
```

Inserting or deleting rows/columns moves cells with their formats and read-only flags, and rewrites formula references the way Excel does: references after the change move, ranges that span an insertion grow, ranges that lose rows/columns shrink, and references to deleted cells become `#REF!`. Absolute (`$A$1`) references move too. Out-of-range arguments throw a `RangeError`, and at least one row and column must remain.

`fromCSV` drops values outside `rows` × `cols` (and throws a `RangeError` for a negative or non-integer `start`); use the exported `parseDelimited(text)` to size the grid first if needed.

### Cell formats

| Property | Type | Description |
|----------|------|-------------|
| `bold`, `italic`, `underline`, `strikethrough` | `boolean` | Text styling |
| `textColor`, `backgroundColor` | `string` | CSS colors |
| `textAlign` | `'left' \| 'center' \| 'right'` | Horizontal alignment |
| `verticalAlign` | `'top' \| 'middle' \| 'bottom'` | Vertical alignment (default middle) |
| `wrap` | `boolean` | Wrap text instead of truncating. Rows keep their fixed height, so only the lines that fit are shown |
| `fontSize` | `number` | Font size in px |
| `numberFormat` | `NumberFormatOptions` | `{ type: 'number' \| 'currency' \| 'percent' \| 'scientific' \| 'date', decimals?, currencySymbol?, thousandsSep?, datePattern? }` |

The `date` type renders date serial numbers (as returned by `DATE`, `TODAY`, `NOW`) using `datePattern` (default `YYYY-MM-DD`; tokens `YYYY YY MM M DD D HH H mm ss`).

### Events

```js
// Single cell edited
sheet.addEventListener('cell-change', (e) => {
  console.log(e.detail.cellId, e.detail.value, e.detail.oldValue);
});

// Selection moved
sheet.addEventListener('selection-change', (e) => {
  console.log(e.detail.range); // { start: { row, col }, end: { row, col } }
});

// Bulk operation (paste, cut, clear, fill, import, row/column insert/delete, undo, redo)
sheet.addEventListener('data-change', (e) => {
  console.log(e.detail.updates, e.detail.source, e.detail.operation);
});

// Rows/columns were inserted or deleted (also on undo/redo)
sheet.addEventListener('structure-change', (e) => {
  const { axis, kind, index, count, source } = e.detail; // e.g. 'row', 'insert', 2, 1, 'user'
});
```

### Validation and edit control

```js
// Lock individual cells: the user can't edit, clear, paste into, fill or format them
sheet.setData(new Map([
  ['0:0', { rawValue: 'Total', displayValue: 'Total', type: 'text', readOnly: true }],
]));

// Reject invalid input. Clearing a cell is never validated.
sheet.validateCell = (cellId, value) =>
  cellId.endsWith(':1') && isNaN(Number(value)) ? 'Quantity must be a number' : null;

sheet.addEventListener('validation-error', (e) => {
  console.log(e.detail.errors); // [{ cellId, value, message }]
});

// Veto or redirect editing
sheet.addEventListener('cell-edit-start', (e) => {
  if (e.detail.cellId === '0:3') e.preventDefault();
});

// Inspect, rewrite or cancel user changes before they are applied
sheet.addEventListener('before-change', (e) => {
  for (const change of e.detail.changes) change.value = change.value.trim();
  // e.detail.changes.splice(i, 1) skips a cell; e.preventDefault() cancels everything
});

// Veto keyboard row/column insertion or deletion
sheet.addEventListener('before-structure-change', (e) => {
  if (e.detail.kind === 'delete') e.preventDefault();
});
```

When Enter or Tab commits a rejected value in the inline editor, the editor stays open, shows the message below the cell (`part="validation-message"`, linked with `aria-describedby`) and announces it. If focus moves away instead, the edit is discarded. A rejected formula-bar entry is not applied and the message is announced. For paste, fill and Ctrl+Enter, rejected and read-only cells are skipped and the rest is applied; the count is announced.

`before-change` fires for user edits, pastes, cuts, clears and fills (`detail.operation` says which), after read-only and validation filtering. It does not fire for undo/redo or API calls such as `fromCSV`. Rewritten values are not validated again.

### Formula engine

Formulas start with `=`. Supported syntax:

**Operators**: `+`, `-`, `*`, `/`, `&` (concat), `=`, `<>`, `<`, `>`, `<=`, `>=`

**Cell references**: `A1`, `B2:D10` (ranges)

**Built-in functions**:

| Category | Functions |
|----------|-----------|
| Aggregates | `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `PRODUCT`, `MEDIAN` |
| Conditional aggregates | `SUMIF`, `COUNTIF`, `AVERAGEIF`, `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` |
| Logic | `IF`, `IFS`, `IFERROR`, `AND`, `OR`, `NOT` |
| Lookup | `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH` |
| Math | `ABS`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `MOD`, `POWER`, `SQRT`, `CEILING`, `FLOOR` |
| Text | `CONCAT`, `TEXTJOIN`, `UPPER`, `LOWER`, `LEN`, `TRIM`, `LEFT`, `RIGHT`, `MID`, `SUBSTITUTE`, `FIND`, `TEXT`, `VALUE` |
| Date | `DATE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `WEEKDAY` |

Omitted arguments use their defaults, e.g. `=XLOOKUP(75, A1:A9, B1:B9, , -1)`. Criteria in `SUMIF`/`COUNTIF`/`AVERAGEIF` and the `*IFS` functions accept operators (`">5"`, `"<>x"`) and wildcards (`*`, `?`, `~` to escape). `XLOOKUP` supports match modes `0`, `-1`, `1` (not wildcard mode `2`) and search modes `1`/`-1`; it returns a single value (no spilling).

**Custom functions**:

```js
sheet.registerFunction('TAX', (ctx, amount, rate) => {
  return Number(amount) * Number(rate);
});
// Use in cells: =TAX(B1, 0.2)
```

The `ctx` argument provides `getCellValue(ref)` and `getRangeValues(start, end)` for reading other cells.

### Theming

Style the component using CSS custom properties:

```css
y11n-spreadsheet {
  --ls-font-family: 'Inter', system-ui;
  --ls-font-size: 14px;
  --ls-cell-width: 120px;
  --ls-cell-height: 32px;
  --ls-header-bg: #f8f9fa;
  --ls-border-color: #dee2e6;
  --ls-selection-border: 2px solid #0d6efd;
  --ls-selection-bg: rgba(13, 110, 253, 0.08);
  --ls-text-color: #212529;
}
```

<details>
<summary>All CSS custom properties</summary>

| Property | Default | Description |
|----------|---------|-------------|
| `--ls-font-family` | `system-ui` | Font family |
| `--ls-font-size` | `13px` | Base font size |
| `--ls-text-color` | `#333` | Cell text color |
| `--ls-cell-width` | `100px` | Column width (any CSS length, e.g. `7rem`) |
| `--ls-cell-height` | `28px` | Row height (any CSS length) |
| `--ls-header-bg` | `#f3f3f3` | Header background |
| `--ls-border-color` | `#e0e0e0` | Cell border color |
| `--ls-selection-border` | `2px solid #1a73e8` | Selection outline |
| `--ls-selection-bg` | `rgba(26, 115, 232, 0.1)` | Selection fill |
| `--ls-focus-ring` | `2px solid #1a73e8` | Focus indicator |
| `--ls-editor-bg` | `#fff` | Inline editor background |
| `--ls-editor-shadow` | `0 2px 6px rgba(0,0,0,0.2)` | Editor drop shadow |
| `--ls-formula-bar-bg` | `#f8fafc` | Formula bar background |
| `--ls-formula-ref-bg` | `#fff` | Formula bar input background |
| `--ls-formula-mode-active-bg` | `#dbeafe` | Active mode toggle background |
| `--ls-ref-highlight-border` | `#1a73e8` | Reference highlight border |
| `--ls-ref-highlight-bg` | `rgba(26, 115, 232, 0.15)` | Reference highlight fill |
| `--ls-header-selected-bg` | `#dbe7fb` | Row/column header background for selected rows/columns |
| `--ls-readonly-cell-bg` | `transparent` | Background of read-only cells |
| `--ls-fill-handle-color` | `#1a73e8` | Fill handle square |
| `--ls-fill-preview-bg` | `rgba(26, 115, 232, 0.08)` | Cells a fill-handle drag will fill |
| `--ls-fill-preview-border` | `#1a73e8` | Dashed outline of those cells |
| `--ls-validation-bg` / `--ls-validation-color` / `--ls-validation-border` | `#fdecea` / `#8a1c12` / `#d93025` | Validation message |
| `--ls-find-bg` | `#fff` | Find bar background |
| `--ls-find-status-color` | `#555` | Find bar match count |

</details>

The inline editor is exposed as a CSS part:

```css
y11n-spreadsheet::part(editor) {
  font-weight: bold;
}
```

### Keyboard shortcuts

| Key | Action |
|-----|--------|
| Arrow keys | Move active cell |
| Shift + Arrow | Extend selection |
| Ctrl/Cmd + Arrow | Jump to the edge of the data region |
| Ctrl/Cmd + Shift + Arrow | Extend selection to the edge of the data region |
| Ctrl + Space | Select entire column(s) |
| Shift + Space | Select entire row(s) |
| Ctrl + Shift + Space, Ctrl+A | Select all |
| Tab / Shift+Tab | Move right / left |
| Enter / Shift+Enter (editing) | Commit edit and move down / up |
| Ctrl/Cmd + Enter (editing) | Commit the value into every selected cell |
| Enter, F2 | Start editing |
| Escape | Cancel edit / clear selection |
| Delete, Backspace | Clear selected cells |
| Ctrl/Cmd + D | Fill down (copy the top row of the selection into the rest) |
| Ctrl/Cmd + R | Fill right (copy the left column of the selection into the rest) |
| Ctrl + Alt + D / Ctrl + Alt + R | Fill series down / right: continue the filled leading rows/columns of the selection into its empty cells, like dragging the fill handle (on macOS use Ctrl+Option) |
| Ctrl/Cmd + Shift + = (or Ctrl+Alt+=, Ctrl+Numpad +) | Insert as many rows/columns as are selected, before the selection. Needs whole rows (Shift+Space) or columns (Ctrl+Space) selected |
| Ctrl/Cmd + - (or Ctrl+Alt+-, Ctrl+Numpad -) | Delete the selected whole rows/columns |
| Ctrl/Cmd + F | Find. In the find bar: Enter / Shift+Enter (or F3 / Shift+F3) next / previous match, Escape closes |
| Ctrl+C / X / V | Copy / Cut / Paste |
| Ctrl+Z | Undo |
| Ctrl+Shift+Z / Ctrl+Y | Redo |
| Any printable character | Start editing with that character |

During formula editing, arrow keys insert cell references instead of navigating.

With the mouse, drag the square at the bottom-right corner of the selection (the fill handle) to fill adjacent cells. A run of two or more numbers continues its linear trend (`1, 2` → `3, 4, 5`; `10, 20` → `30, 40`); text, formulas and single numbers repeat, with relative references shifted. From the keyboard, select the cells to continue plus the empty cells after them and press Ctrl+Alt+D (down) or Ctrl+Alt+R (right); Ctrl+D / Ctrl+R copy instead of continuing a series.

## Types

All types are exported for TypeScript consumers:

```ts
import type {
  CellData,          // { rawValue, displayValue, type, style? }
  CellCoord,         // { row, col }
  GridData,          // Map<string, CellData>
  SelectionRange,    // { start: CellCoord, end: CellCoord }
  FormulaContext,    // { getCellValue, getRangeValues }
  FormulaFunction,   // (ctx: FormulaContext, ...args: unknown[]) => unknown
  CellChangeDetail,
  SelectionChangeDetail,
  DataChangeDetail,
  BeforeChangeDetail,    // { changes: ProposedChange[], operation }
  CellEditStartDetail,   // { cellId, value }
  ValidationErrorDetail, // { errors: [{ cellId, value, message }], operation }
  StructureChangeDetail, // { axis, kind, index, count, source }
  CellValidator,
} from 'y11n-spreadsheet';
```

Utility functions for coordinate conversion, dates, and delimited text:

```ts
import { cellKey, parseKey, colToLetter, letterToCol, refToCoord, coordToRef } from 'y11n-spreadsheet';

cellKey(0, 0)         // "0:0"
parseKey("0:0")       // { row: 0, col: 0 }
colToLetter(0)        // "A"
letterToCol("AA")     // 26
refToCoord("B3")      // { row: 2, col: 1 }
coordToRef({row:0, col:0}) // "A1"

import { dateToSerial, dateSerialToDate, parseDelimited, serializeDelimited } from 'y11n-spreadsheet';

dateToSerial(2024, 0, 1)          // 45292 (month is 0-indexed)
parseDelimited('a,"b,c"\n1,2')    // [['a', 'b,c'], ['1', '2']]
serializeDelimited([['a', 'b,c']]) // 'a,"b,c"'

import { adjustFormulaForStructureChange } from 'y11n-spreadsheet';

adjustFormulaForStructureChange('=SUM(A1:A3)', { axis: 'row', kind: 'insert', index: 1, count: 1 })
// '=SUM(A1:A4)'
```

## Exports

The package also exports internal modules if you need them standalone:

```ts
import { FormulaEngine } from 'y11n-spreadsheet';
import { SelectionManager } from 'y11n-spreadsheet';
import { ClipboardManager } from 'y11n-spreadsheet';
import { Y11nFormulaBar } from 'y11n-spreadsheet';
```

---

## Contributing

### Prerequisites

- Node.js 22+
- npm

### Setup

```bash
git clone <repo-url>
cd accessible-spreadsheet-component
npm install
npx playwright install  # for E2E tests
```

### Development

```bash
npm run dev        # Vite dev server at http://localhost:5173
npm run storybook  # Storybook at http://localhost:6006
```

The `index.html` at the project root has a test harness. Storybook has detailed interactive examples covering all features.

### Testing

```bash
npm test                 # Unit tests (Vitest)
npm run test:watch       # Unit tests in watch mode
npm run test:e2e         # E2E tests headless (Playwright)
npm run test:e2e:headed  # E2E tests in visible browser
npm run test:e2e:ui      # E2E tests with Playwright UI
npm run typecheck        # TypeScript type checking
```

### Project structure

```
src/
  index.ts                          # Barrel exports
  types.ts                          # Types & coordinate utilities
  y11n-spreadsheet.ts               # Main component
  components/y11n-formula-bar.ts    # Formula bar sub-component
  controllers/
    selection-manager.ts            # Selection state (Lit ReactiveController)
    clipboard-manager.ts            # Clipboard operations (TSV + HTML formats)
  engine/
    formula-engine.ts               # Recursive descent formula parser
  utils/
    delimited.ts                    # CSV/TSV parsing and serialization
    references.ts                   # Formula reference rewriting for row/column insert/delete
  __tests__/                        # Unit tests colocated with source
e2e/                                # Playwright E2E tests
stories/                            # Storybook stories
```

### Build

```bash
npm run build  # Outputs ES module to dist/
```

The build externalizes `lit` as a peer dependency and produces `dist/index.js` with TypeScript declarations and source maps.

## License

MIT
