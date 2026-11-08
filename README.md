# y11n-spreadsheet

An accessible spreadsheet web component built with [Lit 3.0](https://lit.dev). Implements the [WAI-ARIA grid pattern](https://www.w3.org/WAI/ARIA/apg/patterns/grid/) with virtual rendering, a formula engine, undo/redo, and clipboard support.

## Features

- Full keyboard navigation (arrows, Ctrl+Arrow data-edge jumps, Tab, Enter, Escape, Ctrl+A)
- Row/column selection via headers (click, Shift+click, drag) or Ctrl+Space / Shift+Space
- Fill down/right (Ctrl+D / Ctrl+R), a draggable fill handle that continues numeric series, and multi-cell entry (Ctrl+Enter)
- Insert and delete rows/columns, with formula references rewritten to follow their cells
- Frozen panes: keep header rows and columns in view while scrolling (`frozen-rows` / `frozen-cols`, Alt+Shift+F)
- Resizable columns: drag a header border, double-click it to fit the content, or use Alt+Shift+Left / Right and Alt+Shift+0
- Sort a range or the table around the active cell (Alt+Shift+Up / Down), by one or more columns, with header detection
- Filter a table by the values in its columns, from an accessible header menu (Alt+Down) or the API; edits, fills and pastes skip the hidden rows
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
| `frozenRows` | `frozen-rows` | `number` | `0` | Rows at the top that stay in view while scrolling (see [Frozen panes](#frozen-panes)) |
| `frozenCols` | `frozen-cols` | `number` | `0` | Columns at the left that stay in view while scrolling |
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

// Sorting (undoable)
const table = { start: { row: 0, col: 0 }, end: { row: 9, col: 3 } };
sheet.sort(table, { col: 1 }, { hasHeader: true })              // by column B, A to Z
sheet.sort(table, [{ col: 2, direction: 'descending' }, { col: 0 }]) // C high to low, then A
sheet.sortSelection('descending')  // what Alt+Shift+Down does: the selection, or the table around the active cell

// Column widths (undoable; px, clamped to 20–2000)
sheet.setColumnWidth(0, 180)          // column A
sheet.setColumnWidths({ 1: 60, 2: null }) // B to 60px, C back to the default (--ls-cell-width)
sheet.getColumnWidth(1)               // 60
sheet.getColumnWidths()               // { 0: 180, 1: 60 }: custom widths only, to save and restore
sheet.autoFitColumns([0, 1])          // fit to the widest value, as a border double-click does

// Frozen panes (a view setting: not undoable)
sheet.freezePanes(1, 1)   // freeze row 1 and column A
sheet.freezePanes(1)      // freeze row 1 only
sheet.freezePanes()       // freeze the rows above and columns left of the active cell (at A1: unfreeze)
sheet.unfreezePanes()

// Filters (undoable; work in read-only mode)
const people = { start: { row: 0, col: 0 }, end: { row: 60, col: 3 } }; // row 1 holds the headers
sheet.setFilter(people)                                     // filter buttons, nothing hidden yet
sheet.setFilter(people, [{ col: 1, values: ['Sales', 'Support'] }]) // keep rows whose B shows Sales or Support
sheet.setColumnFilter(2, ['Oslo', ''])                      // and whose C is Oslo or blank
sheet.setColumnFilter(2, null)                              // clear column C's criteria
sheet.reapplyFilter()                                       // re-filter after edits
sheet.getFilter()        // { range, columns: [{ col, values }], hiddenRows: [...] } or null
sheet.isRowHidden(5)
sheet.removeFilter()
sheet.toggleFilter()     // what Ctrl+Shift+L does: remove the filter, or add one to the table around the active cell
sheet.openFilterMenu(1)  // open column B's filter menu (Alt+Down on its header)

// Find
sheet.findCells('total')                      // ['5:0', ...] in reading order
sheet.findCells('SUM', { formulas: true, matchCase: true })
sheet.openFind()                              // open the find bar (Ctrl/Cmd+F)
sheet.closeFind()
```

Inserting or deleting rows/columns moves cells with their formats and read-only flags, and rewrites formula references the way Excel does: references after the change move, ranges that span an insertion grow, ranges that lose rows/columns shrink, and references to deleted cells become `#REF!`. Absolute (`$A$1`) references move too. Out-of-range arguments throw a `RangeError`, and at least one row and column must remain.

Sorting reorders the rows of a range; cells outside it don't move. Ascending order is numbers (smallest first), text (A to Z, ignoring case), `FALSE`, `TRUE`, then errors; descending reverses that, and blank cells go last either way. The sort is stable, and later keys break ties. Cells move with their formats and read-only flags. Formulas in moved cells shift their relative references by the rows they moved, as if copied (`=B2*C2` moved to row 5 becomes `=B5*C5`); references from outside the range, such as `=SUM(A2:A10)`, are left alone, as in Excel. `sort()` throws a `RangeError` for a range outside the grid or a key column outside the range, and returns whether anything moved.

`sortSelection()` sorts the selection, trimmed to its last filled row and column, by the active cell's column. When a single cell is selected, or a single column of a wider table (which would split its records apart), it sorts the block of data around the active cell instead, like Excel's current region. It keeps the first row in place when it looks like a header: all text, above columns with numbers or other non-text values, or bold above a row that isn't. It refuses to move read-only cells, fires `before-sort`, and announces the result.

Column widths are presentation, not data: they can be changed in read-only mode, `setData()` keeps them, and they aren't part of `getData()`. Save them with `getColumnWidths()` (for example on `column-resize`) and restore them with `setColumnWidths()`. Widths move with their columns when columns are inserted or deleted, and `column-resize` reports the columns whose width changed, so saving on `column-resize` keeps the layout in step. Auto-fit measures the displayed values of every row, including rows that aren't rendered; an empty column goes back to the default width.

### Frozen panes

`<y11n-spreadsheet frozen-rows="1" frozen-cols="1">` keeps the first row and column in view while the rest scrolls. Frozen cells stay in the grid in their usual order, so screen readers report the same row and column numbers as without freezing. At least one row and one column always scroll. Freezing is a view setting: it isn't undoable and survives `setData()`. It follows the content like Excel: inserting rows or columns inside the frozen area grows it, and deleting frozen rows or columns shrinks it (and undo restores it).

`freeze-change` (`{ rows, cols, source }`) fires when the frozen area changes through `freezePanes()`, `unfreezePanes()`, Alt+Shift+F or an insert/delete, not when you set the properties yourself. Frozen cells need an opaque background: set `--ls-cell-bg` to your grid's background color if it isn't white. Keep the frozen area smaller than the visible grid: cells below or right of a frozen area that fills the view can't be scrolled into sight (Alt+Shift+F refuses to freeze that much; the API and attributes don't check).

`fromCSV` drops values outside `rows` × `cols` (and throws a `RangeError` for a negative or non-integer `start`); use the exported `parseDelimited(text)` to size the grid first if needed.

### Filters

A filter hides the rows of a table whose cells don't match. The table's first row is its header row: it is never hidden, and each of its cells gets a filter button. For each filtered column, a row stays visible only if the column displays one of the column's values, compared as shown on screen (number formats included; `''` matches blank cells). Filtered columns show a funnel icon and the table's row numbers turn blue.

On a header cell, Alt+Down (or a click on its button) opens the filter menu, a dialog with Sort A to Z / Z to A, Clear filter, a search box and a checklist of the column's values. The list holds the values of the rows that pass the other columns' filters, sorted like `sort()` with (Blanks) last. Arrow keys move through it, Space checks a value, and Enter applies (OK). Searching checks every match, and OK then keeps only the checked matches, as in Excel. Escape or a click outside closes the menu without changes. Ctrl+Shift+L adds a filter to the selection, or to the table around the active cell (removing it if there is one), and Ctrl+Alt+L reapplies it.

Filters are a view, like column widths: they work in read-only mode, survive `setData()`, and aren't part of `getData()`. Every change is undoable, though. Rows don't hide or reappear as cells are edited (that would move rows out from under the cursor): the filter is evaluated when it is set, changed or reapplied. When the user changes or reapplies a filter, it first grows over rows of data added right below the table. The filter moves with inserted and deleted rows and columns; deleting its header row removes it.

While rows are hidden, the user's actions skip them: arrow keys, Ctrl+Arrow, Page Up/Down and Ctrl+End step over them, and copy, cut, paste, Delete, fill (Ctrl+D, Ctrl+Alt+D, the fill handle), Ctrl+Enter and Ctrl+B/I/U only touch the visible rows. A paste goes to the visible rows from the anchor down, and formulas shift by how far each row actually moved. Sorting (from the menu, Alt+Shift+Up/Down or `sort()`) reorders only the visible rows; the hidden ones stay where they are. Deleting selected rows is refused while some of them are hidden. The find bar skips hidden cells, while `findCells()`, `toCSV()` and the other API methods see every row. Screen readers hear the real row numbers, since hidden rows keep their place in `aria-rowindex`.

`filter-change` fires after every change (`{ range, columns, hiddenRows, source }`, with `range: null` once the filter is removed), including undo/redo and inserts or deletes that move the filter. User changes fire the cancelable `before-filter` first, with the same detail describing the result.

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

// Bulk operation (paste, cut, clear, fill, import, row/column insert/delete, sort, undo, redo)
sheet.addEventListener('data-change', (e) => {
  console.log(e.detail.updates, e.detail.source, e.detail.operation);
});

// Rows/columns were inserted or deleted (also on undo/redo)
sheet.addEventListener('structure-change', (e) => {
  const { axis, kind, index, count, source } = e.detail; // e.g. 'row', 'insert', 2, 1, 'user'
});

// Column widths changed (also on undo/redo, and when widths move with inserted/deleted columns)
sheet.addEventListener('column-resize', (e) => {
  const { columns, source } = e.detail; // columns: [{ col, width }]
  localStorage.setItem('widths', JSON.stringify(sheet.getColumnWidths()));
});

// The filter changed (also on undo/redo; range is null once it was removed)
sheet.addEventListener('filter-change', (e) => {
  const { range, columns, hiddenRows, source } = e.detail; // columns: [{ col, values }]
});

// A range was sorted (also on undo/redo; on undo the detail describes the sort that was undone)
sheet.addEventListener('sort-change', (e) => {
  const { range, keys, hasHeader, source } = e.detail; // keys: [{ col, direction }]
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

// Veto a user resize (border drag, Alt+Shift+Left/Right, auto-fit); same detail as column-resize
sheet.addEventListener('before-column-resize', (e) => {
  if (e.detail.columns.some(({ width }) => width < 40)) e.preventDefault();
});

// Veto a user sort (Alt+Shift+Up/Down or sortSelection()); same detail as sort-change
sheet.addEventListener('before-sort', (e) => {
  if (e.detail.range.start.col === 0) e.preventDefault();
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
| `--ls-cell-bg` | `#fff` | Background of frozen cells (set it to your grid background) |
| `--ls-frozen-line-color` | `#9aa0a6` | Line below frozen rows and right of frozen columns |
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
| `--ls-filter-button-color` / `--ls-filter-button-bg` | `#555` / `#fff` | Filter button in a header cell |
| `--ls-filter-active-color` | `#1a73e8` | Filter button of a filtered column, and the menu's OK button |
| `--ls-filtered-row-header-color` | `#1a73e8` | Row numbers of a filtered table |
| `--ls-filter-menu-bg` | `#fff` | Filter menu background |
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
| Alt + Shift + F | Freeze the rows above and the columns left of the active cell; press again to unfreeze |
| Ctrl/Cmd + Home | Go to A1, or with frozen panes to the first unfrozen cell |
| Alt + Shift + Left / Right | Narrow / widen the selected columns by 8px; the new width is announced (on macOS use Option) |
| Alt + Shift + 0 | Fit the selected columns to their content |
| Alt + Shift + Up / Down | Sort ascending / descending by the active cell's column: the selection, or the table around the active cell (on macOS use Option) |
| Ctrl/Cmd + Shift + L | Add a filter to the selection or the table around the active cell; remove the filter |
| Alt + Down | On a filter header cell: open its filter menu (on macOS use Option) |
| Ctrl + Alt + L | Reapply the filter, after edits (on macOS use Ctrl+Option) |
| Ctrl/Cmd + F | Find. In the find bar: Enter / Shift+Enter (or F3 / Shift+F3) next / previous match, Escape closes |
| Ctrl+C / X / V | Copy / Cut / Paste |
| Ctrl+Z | Undo |
| Ctrl+Shift+Z / Ctrl+Y | Redo |
| Any printable character | Start editing with that character |

During formula editing, arrow keys insert cell references instead of navigating.

Drag the right border of a column header to resize the column (Escape during the drag cancels it); with whole columns selected, all of them get the new width. Double-click the border to fit the column to its content.

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
  SortDetail,            // { range, keys: [{ col, direction }], hasHeader, source }
  SortKey,               // { col, direction? }
  SortOptions,           // { hasHeader? }
  SortDirection,         // 'ascending' | 'descending'
  ColumnResizeDetail,    // { columns: [{ col, width }], source }
  FreezeChangeDetail,    // { rows, cols, source }
  FilterColumn,          // { col, values }
  FilterInfo,            // { range, columns, hiddenRows }
  FilterChangeDetail,    // { range | null, columns, hiddenRows, source }
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
