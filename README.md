# y11n-spreadsheet

An accessible spreadsheet web component built with [Lit 3.0](https://lit.dev). Implements the [WAI-ARIA grid pattern](https://www.w3.org/WAI/ARIA/apg/patterns/grid/) with virtual rendering, a formula engine, undo/redo, and clipboard support.

## Features

- Full keyboard navigation (arrows, Ctrl+Arrow data-edge jumps, Tab, Enter, Escape, Ctrl+A)
- Row/column selection via headers (click, Shift+click, drag) or Ctrl+Space / Shift+Space
- Fill down/right (Ctrl+D / Ctrl+R), a draggable fill handle that continues numbers, dates, day and month names and numbered text (double-click it to fill down the table), and multi-cell entry (Ctrl+Enter)
- Typed input like `15%`, `$1,200`, `1,234.5`, `2026-01-05` or `14:30` is stored as a number with the matching format
- Status bar with the Sum, Average and Count of the selection (Alt+Shift+S reads them out)
- AutoSum (Alt+=), today's date / the time (Ctrl+; / Ctrl+Shift+;), quick number formats (Ctrl+Shift+1 / 3 / 4 / 5 / `)
- Paste values (Ctrl+Shift+V) and an accessible Paste Special dialog (Ctrl+Alt+V) with Values, Formats and Transpose
- Insert and delete rows/columns, with formula references rewritten to follow their cells
- Frozen panes: keep header rows and columns in view while scrolling (`frozen-rows` / `frozen-cols`, Alt+Shift+F)
- Resizable columns: drag a header border, double-click it to fit the content, or use Alt+Shift+Left / Right and Alt+Shift+0
- Sort a range or the table around the active cell (Alt+Shift+Up / Down), by one or more columns, with header detection
- Filter a table by the values in its columns, from an accessible header menu (Alt+Down) or the API; edits, fills and pastes skip the hidden rows
- Conditional formatting: highlight rules (values, text, dates, duplicates, top/bottom, averages, formulas), color scales, data bars and icon sets, with a rule manager dialog (Alt+Shift+C). Every rule has a label that is announced with its cells, so color never carries the meaning alone
- Find (Ctrl+F) with match count and next/previous navigation
- Per-cell read-only flag, value validation, and cancelable `cell-edit-start` / `before-change` events
- CSV/TSV import and export
- Copy/cut/paste with HTML and TSV clipboard formats
- Undo/redo with Ctrl+Z / Ctrl+Shift+Z (or Ctrl+Y)
- Formula engine with 80+ built-in functions and custom function support
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
| `rawInput` | `raw-input` | `boolean` | `false` | Store typed values exactly as entered, without recognising percentages, currency, dates and times (see [Typed input](#typed-input)) |
| `hideStatusBar` | `hide-status-bar` | `boolean` | `false` | Hide the status bar under the grid |

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

// Conditional formatting (undoable; rules are kept in priority order, first wins)
const scores = { start: { row: 1, col: 3 }, end: { row: 60, col: 3 } };
const id = sheet.addConditionalFormat({                    // returns the rule's id
  range: scores,
  condition: { type: 'aboveAverage' },
  format: { backgroundColor: '#c6efce', textColor: '#006100' },
})
sheet.addConditionalFormat({ range: scores, condition: { type: 'dataBar' } })
sheet.updateConditionalFormat(id, { stopIfTrue: true, label: 'strong' })
sheet.moveConditionalFormat(id, 0)                         // to the top priority
sheet.removeConditionalFormat(id)
sheet.getConditionalFormats()                              // all rules, to save
sheet.setConditionalFormats(savedRules)                    // replace all rules
sheet.clearConditionalFormats(scores)                      // clear a range; no argument clears everything
sheet.getConditionalFormatResult('4:3')  // { format, labels: ['above average'], ruleIds, dataBar?, icon?, scale? } or null
sheet.applyConditionalFormatPreset('top10')                // a one-step rule on the selection, as the user
sheet.clearSelectionConditionalFormats()                   // clear the selection's rules, as the user
sheet.openConditionalFormatManager()                       // the rule manager dialog (Alt+Shift+C)

// Find
sheet.findCells('total')                      // ['5:0', ...] in reading order
sheet.findCells('SUM', { formulas: true, matchCase: true })
sheet.openFind()                              // open the find bar (Ctrl/Cmd+F)
sheet.closeFind()

// Paste (reads the system clipboard; undoable)
sheet.paste()                                 // what Ctrl+V does
sheet.paste({ mode: 'values' })               // values only, no formulas or formats (Ctrl+Shift+V)
sheet.paste({ mode: 'formats' })              // formats only
sheet.paste({ transpose: true })              // rows become columns
sheet.openPasteSpecial()                      // the Paste Special dialog (Ctrl+Alt+V)
```

Inserting or deleting rows/columns moves cells with their formats and read-only flags, and rewrites formula references the way Excel does: references after the change move, ranges that span an insertion grow, ranges that lose rows/columns shrink, and references to deleted cells become `#REF!`. Absolute (`$A$1`) references move too. Out-of-range arguments throw a `RangeError`, and at least one row and column must remain.

Sorting reorders the rows of a range; cells outside it don't move. Ascending order is numbers (smallest first), text (A to Z, ignoring case), `FALSE`, `TRUE`, then errors; descending reverses that, and blank cells go last either way. The sort is stable, and later keys break ties. Cells move with their formats and read-only flags. Formulas in moved cells shift their relative references by the rows they moved, as if copied (`=B2*C2` moved to row 5 becomes `=B5*C5`); references from outside the range, such as `=SUM(A2:A10)`, are left alone, as in Excel. `sort()` throws a `RangeError` for a range outside the grid or a key column outside the range, and returns whether anything moved.

`sortSelection()` sorts the selection, trimmed to its last filled row and column, by the active cell's column. When a single cell is selected, or a single column of a wider table (which would split its records apart), it sorts the block of data around the active cell instead, like Excel's current region. It keeps the first row in place when it looks like a header: all text, above columns with numbers or other non-text values, or bold above a row that isn't. It refuses to move read-only cells, fires `before-sort`, and announces the result.

Column widths are presentation, not data: they can be changed in read-only mode, `setData()` keeps them, and they aren't part of `getData()`. Save them with `getColumnWidths()` (for example on `column-resize`) and restore them with `setColumnWidths()`. Widths move with their columns when columns are inserted or deleted, and `column-resize` reports the columns whose width changed, so saving on `column-resize` keeps the layout in step. Auto-fit measures the displayed values of every row, including rows that aren't rendered; an empty column goes back to the default width.

### Typed input

As in Excel, what the user types (in a cell, the formula bar or with Ctrl+Enter) is stored as a number when it is clearly one with a format:

| Typed | Stored | Format |
|-------|--------|--------|
| `15%`, `-2.5%` | `0.15`, `-0.025` | percent, with the decimals typed |
| `$1,200`, `-€3.50`, `£-3` (also `¥`) | `1200`, `-3.5`, `-3` | currency with that symbol, 0 or 2 decimals |
| `1,234,567.8` | `1234567.8` | number with thousands separators |
| `2026-01-05`, `2026/1/5` | the date serial | date, `YYYY-MM-DD` |
| `14:30`, `9:05:30` | the fraction of a day | date, `HH:mm` / `HH:mm:ss` |
| `2026-01-05 14:30` | serial plus time | date, `YYYY-MM-DD HH:mm` |

Only year-first dates are recognised, since `1/5/2026` could be January or May. A plain number typed into a percent cell is a percentage, as with Excel's automatic percent entry: `13` and `0.13` both give 13% (numbers of 1 or more are divided by 100). A cell that already has a number format of the same type keeps it, so a custom `DD/MM/YYYY` pattern survives typing a new date. Percent and date cells are edited as they are displayed (`15%`, `2026-01-05`), and the formula bar shows them that way too. `validateCell`, `before-change` and `cell-change` see the stored number. Pastes, fills and API calls are not converted. Set `raw-input` to store everything as typed; `parseInput(text)` is exported if you want the same rules elsewhere.

### Status bar

The bar under the grid shows the Average, Count and Sum of the selection when two or more of its cells are filled. Count counts non-empty cells; Sum and Average cover the number cells, formula results included, and use the number format of the first one. Rows hidden by a filter are left out. The bar isn't a live region, so it doesn't interrupt screen reader users as the selection moves: Alt+Shift+S announces the same figures. Hide it with `hide-status-bar`, or style it with `::part(status-bar)`, `--ls-status-bg` and `--ls-status-color`.

### Paste Special

Ctrl+Shift+V pastes values: formulas copied in this grid paste as the values they had when copied, and the target cells keep their formats. Ctrl+Alt+V opens the Paste Special dialog, a modal dialog with All / Values / Formats radio buttons and a Transpose checkbox (arrow keys pick the mode, Space toggles Transpose, Enter pastes, Escape closes). Formats pastes only the copied cells' formats; the targets keep their values. Transpose turns copied rows into columns; formulas shift their relative references by how far each cell moved. All of this is also available through `paste(options)`.

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

### Conditional formatting

A conditional-format rule is a range, a condition and a format: `{ range, condition, format }`. Rules live beside the grid data. They are evaluated after every recalculation, and the winning format is merged over the cell's own format when the cell is drawn. Nothing is written into the cells, so `getData()`, copy, sort and CSV export keep seeing the real formats.

```js
const sales = { start: { row: 1, col: 1 }, end: { row: 50, col: 1 } };

sheet.addConditionalFormat({
  range: sales,
  condition: { type: 'cellValue', operator: 'greaterThan', value: 100 },
  format: { backgroundColor: '#c6efce', textColor: '#006100' },
});

// A formula rule: written for the range's first cell and shifted for the others, like a filled formula
sheet.addConditionalFormat({
  range: { start: { row: 1, col: 0 }, end: { row: 50, col: 5 } },
  condition: { type: 'formula', formula: '=$F2="overdue"' },
  format: { bold: true, textColor: '#9c0006' },
  label: 'overdue',
});
```

| `condition.type` | Other properties | Matches |
|------------------|------------------|---------|
| `cellValue` | `operator`: `greaterThan`, `greaterThanOrEqual`, `lessThan`, `lessThanOrEqual`, `equal`, `notEqual`, `between`, `notBetween`; `value`; `value2` for the between operators | Cells whose value compares that way. A value is a number, text (compared ignoring case), or a formula such as `'=$B$1'` |
| `text` | `operator`: `contains`, `notContains`, `beginsWith`, `endsWith`; `text` | Cells whose text does, ignoring case |
| `date` | `period`: `yesterday`, `today`, `tomorrow`, `last7Days`, `lastWeek`, `thisWeek`, `nextWeek`, `lastMonth`, `thisMonth`, `nextMonth` | Date serial numbers in that period (weeks start on Sunday) |
| `blanks`, `notBlanks`, `errors`, `notErrors` | | Empty (or whitespace-only) cells, filled cells, cells showing an error, filled cells without one |
| `duplicates`, `unique` | | Values that occur more than once, or exactly once, in the range (ignoring case) |
| `top`, `bottom` | `rank`; `percent?` | The highest or lowest `rank` numbers of the range, or that percentage of them. Ties with the last one are included |
| `aboveAverage`, `belowAverage` | `orEqual?` | Numbers above or below the average of the range |
| `formula` | `formula` | Cells for which the formula is `TRUE` or a non-zero number |
| `colorScale` | `min`, `max`, optional `mid`: `{ type, value?, color }` | Draws: fills each number cell with a color between the stops |
| `dataBar` | `color?`, `negativeColor?`, `min?`, `max?`, `showValue?` | Draws: a bar from zero to the cell's number; negative bars grow left from the zero axis |
| `iconSet` | `icons`: `arrows3`, `arrows5`, `shapes3`, `symbols3`, `bars5`; `thresholds?`, `reverse?`, `showValue?`, `labels?` | Draws: one icon of the set, by where the number lies in the range |

Scale points (`min`, `max`, `mid`, `thresholds`) are `{ type: 'min' | 'max' | 'number' | 'percent' | 'percentile', value? }`, worked out from the numbers in the rule's range. Color scale colors must be `#rgb` or `#rrggbb`, since they are blended; other colors can be any CSS color. A rule's `format` takes `bold`, `italic`, `underline`, `strikethrough`, `textColor` and `backgroundColor`.

A few deliberate differences from Excel: empty cells only match `blanks` and `formula` rules (Excel highlights them for "less than 5"), and a number is never greater or less than text. Statistics (`top`, averages, duplicates, scales) cover the whole range, including rows hidden by a filter.

**Precedence.** Rules are kept in priority order, and new rules go first, as in Excel (`addConditionalFormat(rule, index)` puts one elsewhere). Every matching rule applies; where two set the same property, the earlier one wins. A matching rule with `stopIfTrue: true` ends the evaluation for that cell. A cell shows at most one data bar, one icon and one color scale: the first of each.

**Accessibility.** Color alone never carries the meaning. Every rule has a short label, generated from its condition ("greater than 100", "above average", "duplicate") or set with `label`. It is part of each matching cell's accessible name and of the announcement when the cell becomes active ("B6: 150, above average"). Color scales and data bars say where the cell is on the scale ("color scale 40%", "data bar 64%"), and icons have text alternatives ("up arrow"; replace them with `labels`). The icons of a set differ in shape, not only in color. A color scale picks black or white text, whichever contrasts more with the fill. With `showValue: false` the value is hidden on screen but stays in the accessible name. In forced-colors mode, where fills are dropped, highlighted cells get a marker on their left edge, color scales and data bars become a meter line under the value (dashed for negative bars), and icons take the system text color.

**Rule manager.** Alt+Shift+C (or `openConditionalFormatManager()`) opens a modal dialog that lists the rules in priority order with Edit, Move up, Move down, Delete and a Stop if true checkbox, plus a form for a new rule (applies to, rule type, values, format, label). Changes apply at once, each as one undo step, and are announced inside the dialog. Escape goes back from the form to the list and then closes the dialog. In read-only mode it shows the rules but changes nothing.

**Toolbar.** `<y11n-format-toolbar>` has a Conditional formatting menu button with one-step rules for the selection. It fires `conditional-format-action`:

```js
toolbar.addEventListener('conditional-format-action', (e) => {
  const { action } = e.detail; // 'duplicates', 'unique', 'aboveAverage', 'belowAverage', 'top10',
                               // 'bottom10', 'colorScale', 'dataBar', 'iconSet', 'clear' or 'manage'
  if (action === 'manage') sheet.openConditionalFormatManager();
  else if (action === 'clear') sheet.clearSelectionConditionalFormats();
  else sheet.applyConditionalFormatPreset(action);
});
```

Rules move with their cells when rows or columns are inserted or deleted: ranges shift, grow and shrink like range references, formulas in rules are rewritten, and a rule whose cells are all deleted is dropped (undo brings it back). Sorting moves cells, not rules. Like filters, rules survive `setData()` and aren't part of `getData()`: save them with `getConditionalFormats()` (for example on `conditional-format-change`) and restore them with `setConditionalFormats()`. Every rule change is undoable. Changes the user makes (the rule manager, `applyConditionalFormatPreset()`, `clearSelectionConditionalFormats()`) are refused in read-only mode and fire the cancelable `before-conditional-format-change` first; `conditional-format-change` (`{ rules, source }`) fires after every change. Invalid rules throw a `TypeError`, and ranges outside the grid a `RangeError`.

Each rule covers one rectangle. Clearing rules from a range removes the rules inside it and trims those it cuts across a whole edge; a rule that would no longer be one rectangle is left as it is.

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

// The conditional-format rules changed (also on undo/redo, and when rows/columns move them)
sheet.addEventListener('conditional-format-change', (e) => {
  const { rules, source } = e.detail; // all rules, in priority order
  localStorage.setItem('rules', JSON.stringify(rules));
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
| Aggregates | `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `PRODUCT`, `MEDIAN`, `SUMPRODUCT` |
| Statistics | `STDEV`, `STDEV.S`, `STDEV.P`, `VAR`, `VAR.S`, `VAR.P` |
| Conditional aggregates | `SUMIF`, `COUNTIF`, `AVERAGEIF`, `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` |
| Logic | `IF`, `IFS`, `IFERROR`, `AND`, `OR`, `NOT` |
| Lookup | `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH` |
| Math | `ABS`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `MOD`, `POWER`, `SQRT`, `CEILING`, `FLOOR` |
| Text | `CONCAT`, `TEXTJOIN`, `UPPER`, `LOWER`, `LEN`, `TRIM`, `LEFT`, `RIGHT`, `MID`, `SUBSTITUTE`, `FIND`, `TEXT`, `VALUE` |
| Date | `DATE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `WEEKDAY`, `EDATE`, `EOMONTH` |
| Financial | `PMT`, `FV`, `PV`, `NPV` |
| Information | `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISLOGICAL`, `ISERROR`, `ISERR`, `ISNA` |

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
| `--ls-data-bar-color` / `--ls-data-bar-negative-color` | `#8fb8ea` / `#f0a1a1` | Conditional-format data bars without a color of their own, for positive / negative numbers |
| `--ls-find-bg` | `#fff` | Find bar background |
| `--ls-find-status-color` | `#555` | Find bar match count |
| `--ls-status-bg` / `--ls-status-color` | `#f8fafc` / `#555` | Status bar under the grid |

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
| Alt + = | AutoSum: on a cell, start a SUM of the numbers above it (or to its left); on a range, write a SUM under each column (or right of a single row) |
| Ctrl + ; / Ctrl + Shift + ; | Enter today's date / the current time (inserted at the caret while editing) |
| Ctrl + ' / Ctrl + Shift + ' | Edit with the formula / the value of the cell above |
| Ctrl + Shift + 1 / 3 / 4 / 5 / ` | Number (2 decimals, separators) / date / currency / percent / general format |
| Alt + Shift + S | Announce the Sum, Average and Count of the selection |
| Ctrl/Cmd + D | Fill down (copy the top row of the selection into the rest) |
| Ctrl/Cmd + R | Fill right (copy the left column of the selection into the rest) |
| Ctrl + Alt + D / Ctrl + Alt + R | Fill series down / right: continue the filled leading rows/columns of the selection into its empty cells, like dragging the fill handle. With no empty cells selected, Ctrl+Alt+D fills down as far as the data next to the selection goes, like double-clicking the fill handle (on macOS use Ctrl+Option) |
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
| Alt + Shift + C | Open the conditional-format rule manager (on macOS use Option) |
| Ctrl/Cmd + F | Find. In the find bar: Enter / Shift+Enter (or F3 / Shift+F3) next / previous match, Escape closes |
| Ctrl+C / X / V | Copy / Cut / Paste |
| Ctrl + Shift + V | Paste values only |
| Ctrl + Alt + V | Paste Special: values, formats, transpose (on macOS use Ctrl+Option) |
| Ctrl+Z | Undo |
| Ctrl+Shift+Z / Ctrl+Y | Redo |
| Any printable character | Start editing with that character |

During formula editing, arrow keys insert cell references instead of navigating.

Drag the right border of a column header to resize the column (Escape during the drag cancels it); with whole columns selected, all of them get the new width. Double-click the border to fit the column to its content.

With the mouse, drag the square at the bottom-right corner of the selection (the fill handle) to fill adjacent cells. A run of two or more numbers continues its linear trend (`1, 2` → `3, 4, 5`; `10, 20` → `30, 40`), a single date steps by a day, day and month names continue and wrap (`Mon` → `Tue`, `Nov, Dec` → `Jan`, keeping their case), and text ending in a number counts up (`Item 1` → `Item 2`, `Q09` → `Q10`; `Row 1, Row 3` → `Row 5`). Other text, formulas and single numbers repeat, with relative references shifted. Double-click the fill handle to fill down as far as the column to the left (or right) has data, stopping above any filled cell. From the keyboard, select the cells to continue plus the empty cells after them and press Ctrl+Alt+D (down) or Ctrl+Alt+R (right); Ctrl+D / Ctrl+R copy instead of continuing a series.

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
  ConditionalFormatRule,        // { id, range, condition, format?, stopIfTrue?, label? }
  ConditionalFormatRuleInput,   // the same, with an optional id
  ConditionalFormatCondition,   // { type: 'cellValue' | 'text' | ... | 'iconSet', ... }
  ConditionalFormatStyle,       // { bold?, italic?, underline?, strikethrough?, textColor?, backgroundColor? }
  ConditionalFormatResult,      // { format, labels, ruleIds, dataBar?, icon?, scale? }
  ConditionalFormatPreset,      // 'duplicates' | 'aboveAverage' | 'top10' | 'colorScale' | ...
  ConditionalFormatChangeDetail, // { rules, source }
  ConditionalThreshold,         // { type: 'min' | 'max' | 'number' | 'percent' | 'percentile', value? }
  ConditionalColorStop,         // ConditionalThreshold & { color }
  ConditionalIconSetName,       // 'arrows3' | 'arrows5' | 'shapes3' | 'symbols3' | 'bars5'
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
import { Y11nFormatToolbar, Y11nConditionalFormatDialog } from 'y11n-spreadsheet';

// Conditional formatting without the component: validate rules and evaluate them over your own data
import { ConditionalFormatEvaluator, normalizeConditionalRule, describeCondition } from 'y11n-spreadsheet';
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
  components/
    y11n-formula-bar.ts             # Formula bar sub-component
    y11n-format-toolbar.ts          # Formatting toolbar (with the Conditional formatting menu)
    y11n-conditional-format-dialog.ts # Rule manager dialog
    conditional-icons.ts            # Icon-set icons (SVG)
  controllers/
    selection-manager.ts            # Selection state (Lit ReactiveController)
    clipboard-manager.ts            # Clipboard operations (TSV + HTML formats)
  engine/
    formula-engine.ts               # Recursive descent formula parser
  utils/
    delimited.ts                    # CSV/TSV parsing and serialization
    references.ts                   # Formula reference rewriting for row/column insert/delete
    conditional-format.ts           # Conditional-format rule validation and evaluation
    input-parse.ts                  # Recognises typed percentages, currency, dates and times
    series.ts                       # Fill-handle series (numbers, dates, day/month names, numbered text)
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
