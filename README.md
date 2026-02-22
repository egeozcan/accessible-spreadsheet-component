# y11n-spreadsheet

An accessible spreadsheet web component built with [Lit 3.0](https://lit.dev). Implements the [WAI-ARIA grid pattern](https://www.w3.org/WAI/ARIA/apg/patterns/grid/) with virtual rendering, a formula engine, undo/redo, and clipboard support.

## Features

- Full keyboard navigation (arrows, Tab, Enter, Escape, Ctrl+A)
- Copy/cut/paste with TSV clipboard format
- Undo/redo with Ctrl+Z / Ctrl+Shift+Z
- Formula engine with 20+ built-in functions and custom function support
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

### Methods

```js
sheet.getData()          // Returns a copy of the grid data
sheet.setData(data)      // Replaces grid data (clears undo/redo)
sheet.registerFunction('DOUBLE', (ctx, val) => Number(val) * 2)
```

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

// Bulk operation (paste, cut, clear, undo, redo)
sheet.addEventListener('data-change', (e) => {
  console.log(e.detail.updates, e.detail.source, e.detail.operation);
});
```

### Formula engine

Formulas start with `=`. Supported syntax:

**Operators**: `+`, `-`, `*`, `/`, `&` (concat), `=`, `<>`, `<`, `>`, `<=`, `>=`

**Cell references**: `A1`, `B2:D10` (ranges)

**Built-in functions**:

| Category | Functions |
|----------|-----------|
| Aggregates | `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA` |
| Conditional | `IF(condition, trueVal, falseVal)` |
| Math | `ABS`, `ROUND` |
| String | `CONCAT`, `UPPER`, `LOWER`, `LEN`, `TRIM` |

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
| `--ls-cell-width` | `100px` | Column width |
| `--ls-cell-height` | `28px` | Row height |
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
| Tab / Shift+Tab | Move right / left |
| Enter | Start editing / commit edit |
| Escape | Cancel edit / clear selection |
| Delete, Backspace | Clear selected cells |
| Ctrl+A | Select all |
| Ctrl+C / X / V | Copy / Cut / Paste |
| Ctrl+Z | Undo |
| Ctrl+Shift+Z | Redo |
| Any printable character | Start editing with that character |

During formula editing, arrow keys insert cell references instead of navigating.

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
} from 'y11n-spreadsheet';
```

Utility functions for coordinate conversion:

```ts
import { cellKey, parseKey, colToLetter, letterToCol, refToCoord, coordToRef } from 'y11n-spreadsheet';

cellKey(0, 0)         // "0:0"
parseKey("0:0")       // { row: 0, col: 0 }
colToLetter(0)        // "A"
letterToCol("AA")     // 26
refToCoord("B3")      // { row: 2, col: 1 }
coordToRef({row:0, col:0}) // "A1"
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
    clipboard-manager.ts            # Clipboard operations (TSV format)
  engine/
    formula-engine.ts               # Recursive descent formula parser
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
