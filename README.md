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
- An optional toolbar element with buttons for all of the above, which wires itself to a grid and shows only the controls you choose
- Per-cell read-only flag, value validation, and cancelable `cell-edit-start` / `before-change` events
- CSV/TSV import and export
- Copy/cut/paste with HTML and TSV clipboard formats
- Undo/redo with Ctrl+Z / Ctrl+Shift+Z (or Ctrl+Y)
- Formula engine with more than 100 built-in functions and custom function support
- Dynamic arrays: a formula whose result is an array spills into the cells below and to the right of it (`=A1:A3*2`, `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`), with `#SPILL!`, `A1#` and `@`
- Lists and arrays in cells: several values in one cell (Ctrl+J), a popup that shows them (Alt+Down), filtering by item, and `HAS` / `HASANY` / `HASALL` / `FLATTEN`
- Multiple sheets with accessible sheet tabs and cross-sheet references (`=Sheet2!A1`, `=SUM(Data!A:A)`)
- Defined names for ranges, constants and functions (`=SUM(sales)`, `=double(4)`), with a Name Box to go to a cell, range or name and to name the selection (Ctrl+G), and a name manager dialog (Ctrl+F3)
- `LET` and `LAMBDA`, with `MAP`, `REDUCE`, `BYROW` and `BYCOL`
- References picked with the arrow keys or the mouse while typing a formula, on other sheets too
- Virtual rendering for large datasets (1000+ rows)
- Dual-mode formula bar (raw / formatted) with a Name Box
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
| `sheetTabs` | `sheet-tabs` | `boolean` | `false` | Show the sheet tabs even with a single sheet, so users can add sheets. With more than one sheet the tabs always show (see [Multiple sheets](#multiple-sheets)) |
| `sheets` | - | `string[]` (read-only) | `['Sheet1']` | Names of the sheets, in tab order |
| `activeSheet` | - | `string` | `'Sheet1'` | Name of the sheet being shown. Assign a sheet's name to show it |

With several sheets, `rows`, `cols`, `data`, `frozenRows` and `frozenCols` are those of the sheet being shown.

### Methods

```js
sheet.getData()          // Returns a copy of the grid data (of the sheet being shown)
sheet.setData(data)      // Replaces grid data (clears undo/redo)
sheet.registerFunction('DOUBLE', (ctx, val) => Number(val) * 2)
sheet.registerFunction('NET', (ctx, gross) => gross / 1.19, { args: ['number'] }) // called once per element of a range

// Formatting
sheet.setCellFormat('0:0', { bold: true, verticalAlign: 'top', wrap: true })
sheet.setRangeFormat({ start: { row: 1, col: 1 }, end: { row: 9, col: 1 } },
  { numberFormat: { type: 'date', datePattern: 'DD/MM/YYYY' } })

// The selection, and what a user does to it. These back your own buttons (the toolbar uses them).
// Each does what its keyboard shortcut does: those that change cells are refused in read-only mode,
// leave read-only cells and rows hidden by the filter alone, go through the cancelable before-…
// events, are one undo step and are announced. An edit in progress is committed first
sheet.getSelection()                  // { range: { start, end }, activeCell: { row, col } }
sheet.getSelectionFormat()            // the format all selected cells share, e.g. { bold: true }
sheet.formatSelection({ textAlign: 'center' })
sheet.toggleFormat('bold')            // what Ctrl+B does; also 'italic', 'underline', 'strikethrough'
sheet.clearSelectionFormat()
sheet.changeDecimals(1)               // one more decimal place in each number's own format; -1 for one fewer
sheet.insertRowsAtSelection('above')  // as many rows as the selection spans; or 'below'
sheet.insertColumnsAtSelection('left') // or 'right'
sheet.deleteSelectedRows()
sheet.deleteSelectedColumns()
sheet.autoSum()                       // what Alt+= does
sheet.toggleLists()                   // what Ctrl+J does
sheet.cut()                           // what Ctrl+X does
sheet.copy()                          // what Ctrl+C does (also in read-only mode, like the three below)
sheet.autoFitSelectedColumns()        // what Alt+Shift+0 does
sheet.freezePanesAsUser()             // what Alt+Shift+F freezes: up to the active cell. (1, 0): the first row; (0, 0): unfreeze.
                                      // Unlike freezePanes() it refuses what would not leave room to scroll, and announces
sheet.reapplyFilterAsUser()           // what Ctrl+Alt+L does: unlike reapplyFilter() it takes in rows added under the table,
                                      // fires before-filter and announces the rows shown

// Undo and redo
sheet.undo()                          // true if a change was undone
sheet.redo()
sheet.canUndo                         // history-change fires when canUndo or canRedo change, and when setData() starts a fresh history
sheet.canRedo

// A message for screen readers, through the grid's live region
sheet.announce('Saved')

// Move the keyboard focus to the active cell (or to the editor while a cell is being edited)
sheet.focus()

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

// Sheets (not undoable)
sheet.sheets                                  // ['Sheet1']
sheet.addSheet()                              // 'Sheet2': the next free name is the default
sheet.addSheet('Orders', { data: orders, rows: 200, cols: 8, index: 0, activate: true }) // every option is optional
sheet.activeSheet                             // 'Orders'
sheet.activateSheet('Sheet1')                 // or: sheet.activeSheet = 'Sheet1'
sheet.renameSheet('Sheet1', 'Summary')        // formulas that refer to it are rewritten
sheet.moveSheet('Summary', 0)                 // make it the first tab
sheet.getSheetData('Orders')                  // like getData(), for any sheet
sheet.setSheetData('Orders', data)            // like setData(), for any sheet
sheet.removeSheet('Sheet2')

// Defined names (workbook-wide; not undoable)
sheet.defineName('tax', '=0.19')                                  // a constant: =B2*(1+tax)
sheet.defineName('sales', '=Orders!C2:C60')                       // a range: =SUM(sales)
sheet.defineName('withTax', '=LAMBDA(amount, amount * (1 + tax))') // a function: =withTax(B5)
sheet.defineName('prices', '=B2:B60', 'Orders')                   // third argument: the sheet B2:B60 is on
sheet.getNames()                              // [{ name, refersTo, sheet }], to save
sheet.setNames(savedNames)                    // replace all names
sheet.removeName('tax')                       // true if it existed; formulas that use it show #NAME?
sheet.goTo('sales')                           // select the cells of a name or reference, on their sheet: 'B5', 'A1:C3', 'Orders!A2'
sheet.focusNameBox()                          // move the focus to the Name Box (Ctrl+G)
sheet.openNameManager()                       // the name manager dialog (Ctrl+F3)
sheet.closeNameManager()

// Lists and arrays in cells
const owners = { start: { row: 1, col: 1 }, end: { row: 9, col: 1 } };
sheet.convertToList(owners)                   // "Ann, Bob" becomes a list of two values; returns the number of cells converted (undoable)
sheet.convertListToText(owners)               // and back
sheet.convertToList()                         // no argument: the selection
sheet.getCellValues('1:1')                    // { rows: 2, cols: 1, values: ['Ann', 'Bob'] }, or null for a cell with a single value
sheet.showCellValues('1:1')                   // open the popup with the cell's values (Alt+Down or Alt+Shift+V); no argument: the active cell
sheet.closeCellValues()

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

A filter hides the rows of a table whose cells don't match. The table's first row is its header row: it is never hidden, and each of its cells gets a filter button. For each filtered column, a row stays visible only if the column displays one of the column's values, compared as shown on screen (number formats included; `''` matches blank cells). A cell that holds [several values](#lists-and-arrays-in-cells) matches when one of its items is among them. Filtered columns show a funnel icon and the table's row numbers turn blue.

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

**Toolbar.** The [toolbar](#toolbar)'s Conditional formatting menu has one-step rules for the selection (`applyConditionalFormatPreset()`), Clear rules from selection (`clearSelectionConditionalFormats()`) and Manage rules (`openConditionalFormatManager()`).

Rules move with their cells when rows or columns are inserted or deleted: ranges shift, grow and shrink like range references, formulas in rules are rewritten, and a rule whose cells are all deleted is dropped (undo brings it back). Sorting moves cells, not rules. Like filters, rules survive `setData()` and aren't part of `getData()`: save them with `getConditionalFormats()` (for example on `conditional-format-change`) and restore them with `setConditionalFormats()`. Every rule change is undoable. Changes the user makes (the rule manager, `applyConditionalFormatPreset()`, `clearSelectionConditionalFormats()`) are refused in read-only mode and fire the cancelable `before-conditional-format-change` first; `conditional-format-change` (`{ rules, source }`) fires after every change. Invalid rules throw a `TypeError`, and ranges outside the grid a `RangeError`.

Each rule covers one rectangle. Clearing rules from a range removes the rules inside it and trims those it cuts across a whole edge; a rule that would no longer be one rectangle is left as it is.

### Dynamic arrays

A formula whose result is an array spills: its values fill the cells below and to the right of the formula's cell. Operators and single-value functions work element by element on ranges and arrays, so `=A1:A3*2`, `=UPPER(A1:A3)` and `=IF(A1:A3>1, "yes", "no")` each give three values. The array functions (`FILTER`, `SORT`, `UNIQUE`, `SEQUENCE` and the others listed under [Formula engine](#formula-engine)) return arrays of any size.

```js
import { cellKey } from 'y11n-spreadsheet';

const cell = (rawValue) => ({ rawValue, displayValue: rawValue, type: 'text' });
const data = new Map();
['North', 'South', 'North', 'East'].forEach((region, row) => data.set(cellKey(row, 0), cell(region)));
data.set('0:2', cell('=UNIQUE(A1:A4)')); // C1 spills North, South, East into C1:C3
data.set('0:4', cell('=ROWS(C1#)'));     // E1: 3
sheet.setData(data);

sheet.getData().get('0:2') // { rawValue: '=UNIQUE(A1:A4)', displayValue: 'North', type: 'text', spill: { rows: 3, cols: 1 } }
sheet.getData().get('1:2') // { rawValue: '', displayValue: 'South', type: 'text', spilledFrom: '0:2' }
```

A formula shows `#SPILL!` and spills nothing when a cell in the way holds a value, or when the result would leave the grid. Clearing the cell lets the spill through. Typing into a spilled cell blocks the spill in the same way.

`C1#` refers to everything `C1` spills, whatever its size (`=SUM(C1#)`). `@` picks one value from a range: the one on the formula's own row or column (`=@A1:A9` in `C4` is the value of `A4`).

Spilled values are computed, not stored. In `getData()` the formula's cell has `spill` (the size of the range it fills), and the cells that show its values have an empty `rawValue` and `spilledFrom` (the key of the formula's cell). `setData()` ignores spilled values in the data it is given, so the result of `getData()` can be passed back as it is. Sorting and inserting or deleting rows and columns move the stored cells only, and the formula spills again where it lands. Copying a spilled cell copies its value. CSV export of display values, filters, the status bar and number formats treat spilled values like any others.

**Accessibility.** On a spilled value the active cell is announced with "spilled from C1", on the formula's cell with "spills to C3", and on a formula showing `#SPILL!` with the reason: "the cells the result would fill are not empty" or "the result does not fit on the sheet". While the active cell is in a spill range, the range is outlined (`--ls-spill-border`), and on a spilled value the formula bar shows the formula it comes from greyed out.

### Lists and arrays in cells

A cell can hold several values, as in [Excel](https://techcommunity.microsoft.com/blog/excelblog/excel-now-supports-multiple-values-in-a-single-cell/4549756). There are two kinds: a list, whose values the user types, and an array, which a formula keeps in its cell instead of spilling.

**Lists.** Ctrl+J turns the comma-separated text of the selected cells into lists: `Ann, Bob` becomes two values. When the selection has lists and no text left to convert, it turns the lists back into text. On an empty cell it opens the editor for a new list, and in the editor it commits what was typed as a list. Formula cells are left alone. Like other edits, the conversion skips read-only cells and rows hidden by a filter, and is one undo step.

A list is edited as its items: Enter or F2 on a list shows `Ann, Bob, "Smith, J"`. Items are separated by commas; double quotes keep a comma inside an item, or keep a number as text (`"007"`). Numbers and `TRUE` / `FALSE` become values, and empty items are dropped. Comma-separated text typed right below a list becomes a list too (not with `raw-input`).

A list is stored as a formula, `={{"Ann";"Bob"}}`, in the cell's `rawValue`, so undo, copy, export and `setData()` need nothing new. `listToFormula()` and `listItemsOf()` convert between items and that formula:

```js
import { listToFormula } from 'y11n-spreadsheet';

const cell = (rawValue) => ({ rawValue, displayValue: '', type: 'text' });
sheet.setData(new Map([
  ['0:0', cell('Project')],  ['0:1', cell('Owners')],                         ['0:2', cell('Owned by Bob?')],
  ['1:0', cell('Website')],  ['1:1', cell(listToFormula(['Ann', 'Bob']))],    ['1:2', cell('=HAS(B2, "Bob")')],
  ['2:0', cell('App')],      ['2:1', cell('={{"Bob";"Cy";"Di"}}')],           ['2:2', cell('=HAS(B3, "Bob")')],
]));

sheet.getCellValues('1:1')   // { rows: 2, cols: 1, values: ['Ann', 'Bob'] }
sheet.getData().get('1:1')   // { rawValue: '={{"Ann";"Bob"}}', displayValue: 'Ann, Bob', type: 'text', array: { ... } }
```

**Arrays in cells.** Wrapping a formula in braces keeps a result that would spill in one cell: `={FILTER(A2:A9, B2:B9="x")}`. Arrays nest: `={{1,2,3};{4,5,6}}` fills two cells with an array of three values each, and `=TEXTSPLIT(A1:A3, ",")` gives one array per text.

**In formulas.** A reference to a cell with several values gives all of them: `=B2` spills them, and `=SUM(B2)`, `=COUNTA(B2)` and `=B2*2` work on the items. Four functions are made for these cells:

| Function | Result |
|----------|--------|
| `HAS(array, value)` | Whether the value is one of the items, at any depth of nesting (text is compared ignoring case) |
| `HASANY(array, values)` | Whether at least one of the values is among the items |
| `HASALL(array, values)` | Whether every one of the values is among the items |
| `FLATTEN(array, [pad_value], [levels])` | Lays nested arrays out as a table, e.g. `=UNIQUE(FLATTEN(B2:B4))` for everyone in three lists. Shorter arrays are padded with `pad_value` (default `#N/A`); `levels` is how many levels of nesting to remove (default: all) |

**Filters** list and match the individual items instead of the whole text: the filter menu of a column of lists shows each item once, and a row stays visible when one of its items is kept. `sheet.setColumnFilter(1, ['Bob'])` keeps the rows whose list has Bob in it.

**Display and the values popup.** The cell shows its values as text (`Ann, Bob`; the rows of a table are separated by `; `, and nested arrays are in braces) next to an icon. `CellData.array` has them as data: `{ rows, cols, values }`, with `values` row by row, `null` for a blank and a nested `{ rows, cols, values }` for an array inside the array. Alt+Down on the cell, a click on its icon or `showCellValues()` opens a popup with the values. Alt+Shift+V opens it too, and is the key to use on a filter header cell, where Alt+Down opens the filter menu. The popup shows a numbered list, or a table for an array with several rows and columns (up to 200 rows and 20 columns). A nested array is a button that opens it; Back or Backspace returns, and Escape or a click outside closes the popup.

**Accessibility.** The cell's accessible name and the active-cell announcement say what the cell holds, "list of 3 items" or "array of 2 rows and 3 columns", and the announcement names the key that shows the values: "Alt+Down to show the values", or "Alt+Shift+V to show the values" on a filter header cell. The popup is a dialog labelled with the cell ("Values in B2").

`list-change` (`{ cellIds, toList, source }`) fires when cells are converted with Ctrl+J, `convertToList()` or `convertListToText()`. Typing or editing a list is an edit like any other (`before-change`, `cell-change`).

### Multiple sheets

A workbook has one or more sheets, and starts with one named `Sheet1`. Formulas refer to cells on another sheet with the sheet's name and `!`: `=Sheet2!A1`, `=SUM(Data!A:A)`, and with quotes when the name needs them, `='My sheet'!A1:B2`. A reference to a sheet that doesn't exist is `#REF!`.

```js
const cell = (rawValue) => ({ rawValue, displayValue: rawValue, type: 'text' });

sheet.setData(new Map([
  ['0:0', cell('North')],
  ['0:1', cell('=SUMIF(Orders!A:A, A1, Orders!C:C)')],
]));
sheet.renameSheet('Sheet1', 'Summary');
sheet.addSheet('Orders', {
  data: new Map([
    ['0:0', cell('North')], ['0:2', cell('120')],
    ['1:0', cell('South')], ['1:2', cell('80')],
    ['2:0', cell('North')], ['2:2', cell('45')],
  ]),
});

sheet.sheets                    // ['Summary', 'Orders']
sheet.activeSheet               // 'Summary'; B1 shows 165
sheet.activateSheet('Orders');  // show the other sheet
```

**What is per sheet.** Each sheet has its own data, size, selection, scroll position, column widths, frozen panes, filter and conditional formats. `data`, `rows`, `cols`, the selection and every method that takes cell keys or ranges work on the sheet being shown; `getSheetData(name)` and `setSheetData(name, data)` reach the other sheets. Defined names, custom functions and the undo history belong to the whole workbook. `addSheet(name, options)` takes the new sheet's `data`, `rows` and `cols` (default: the size of the sheet being shown), its `index` among the tabs (default: last) and `activate` to show it.

**Sheet tabs.** The tabs (`<y11n-sheet-tabs>`) appear under the grid when there is more than one sheet, or always with the `sheet-tabs` attribute, which lets users add the second sheet. They are a tab list: Left / Right and Home / End move between the tabs and show each sheet as they go. F2 or a double-click renames a sheet in place (Enter commits, Escape cancels, and a name that isn't allowed shows why), Delete asks for confirmation and then deletes the sheet, and Ctrl+Shift+Left / Right moves it. A tab can also be dragged to another position: with the mouse once it has moved a little sideways, with a finger or pen after resting on it briefly. A line shows where the tab will land, and Escape during the drag puts it back. The `+` button adds a sheet, and the Sheet options menu has Rename, Delete, Move left and Move right for the active sheet. From the grid, Ctrl+PageDown / PageUp show the next / previous sheet; Alt+PageDown / PageUp do the same, since browsers keep Ctrl+PageDown / PageUp for their own tabs. In read-only mode the tabs switch sheets but don't add, rename, delete or move them.

**References to another sheet while typing.** With a formula open, a click on another sheet's tab shows that sheet without ending the edit, and so do Ctrl+PageDown / PageUp and Alt+PageDown / PageUp in the editor. Cells picked there, with a click, a drag or the arrow keys, are written with the sheet's name: `=A1+Orders!B3`, `=A1+'My data'!B2`. Meanwhile the formula stays in view in a strip at the bottom of the grid, and the announcement says which cell is being edited. Enter commits and Escape cancels, and both show the edited cell's sheet again. With anything other than a formula in the editor, a click on a tab commits the edit first, and the sheet keys do nothing.

**Sheet names** have 1 to 31 characters, none of `: \ / ? * [ ]`, no space or apostrophe at either end, and are unique ignoring case. `addSheet()` and `renameSheet()` throw a `RangeError` for any other name, as every sheet method does for a sheet that doesn't exist. `validateSheetName(name, otherNames)` is exported and returns what is wrong with a name, or `null`.

**Undo.** Undo and redo cover all sheets: undoing a change made on another sheet shows that sheet first. Adding, removing, renaming and moving sheets are not undoable. Removing a sheet, or replacing its data with `setSheetData()`, takes the changes made on it out of the undo history.

**Keeping references right.** Inserting or deleting rows and columns rewrites the references to that sheet on every sheet, and in defined names. Renaming a sheet rewrites the formulas, defined names and conditional formats that refer to it, and the undo history. Removing a sheet keeps the text of the formulas that refer to it: they show `#REF!` until a sheet of that name exists again. The last sheet can't be removed.

**Events.** `sheet-change` (`{ sheet, previous, source }`) fires after another sheet is shown, including when undo or redo shows the sheet it changes. `sheets-change` (`{ sheets, kind, sheet, previous?, source }`, with `kind` one of `'add'`, `'remove'`, `'rename'`, `'move'`) fires after the list of sheets changes. Changes users make from the tabs fire the cancelable `before-sheets-change` first, with the same detail describing the result:

```js
sheet.addEventListener('before-sheets-change', (e) => {
  if (e.detail.kind === 'remove' && e.detail.sheet === 'Summary') e.preventDefault();
});
```

To save a workbook, store `getSheetData(name)` for each of `sheet.sheets`, plus `getNames()`. `getColumnWidths()`, `getFilter()` and `getConditionalFormats()` return those of the sheet being shown.

### Defined names

A defined name is a name that formulas can use for a range, a constant or a function.

```js
sheet.defineName('sales', '=B2:B60');
sheet.defineName('tax', '=0.19');
sheet.defineName('withTax', '=LAMBDA(amount, amount * (1 + tax))');
// In cells: =SUM(sales), =B2*(1+tax), =withTax(B5)
```

A name is made of letters, digits, `_` and `.`, doesn't start with a digit, and can't look like a cell reference (`A1`) or be `TRUE` or `FALSE`; `defineName()` throws a `RangeError` otherwise. Names are not case-sensitive and belong to the whole workbook; defining an existing name replaces it. The leading `=` of `refersTo` is optional. References without a sheet name point to the sheet given as the third argument (default: the sheet being shown). A name that is not defined is `#NAME?`.

The references a name holds move when rows or columns are inserted or deleted (and move back on undo), and follow a renamed sheet. Defining, changing and removing names is not undoable, whether it is done from code or by the user, and names survive `setData()`. `names-change` (`{ names, source }`) fires after every change, including when rows or columns move a name's references and when a sheet a name is based on or refers to is renamed or removed: save the names with `getNames()` and restore them with `setNames()`. A name whose sheet was removed has `sheet: ''`: it keeps working as a constant or with sheet-qualified references, its references without a sheet name are `#REF!`, and `setNames()` restores it like that instead of basing it on another sheet.

**Name Box.** The box at the left of the formula bar shows the active cell's reference, or a name when the selection is exactly that name's cells. It is an input. Type a cell (`C5`), a range (`A1:C3`), a reference to another sheet (`Orders!A2`) or a name that stands for cells, and press Enter: those cells are selected, on their sheet, and the focus returns to the grid. Type a name that doesn't exist yet to name the selection: it is defined as an absolute reference with the sheet's name (`=Sheet1!$A$1:$A$2`), and the new name is announced. Other text changes nothing and is announced as not valid; a name for a constant or a LAMBDA has no cells to go to. The box suggests the defined names, Escape returns to the grid, and Ctrl+G or `focusNameBox()` moves the focus to the box.

`goTo(reference)` does from code what typing in the Name Box and pressing Enter does for a reference or a name. It returns `false` when the text is neither, when the name doesn't stand for cells, or when the cells are outside the grid.

```js
sheet.goTo('Orders!A2');  // shows the Orders sheet and selects A2
sheet.goTo('sales');      // selects B2:B60
sheet.goTo('tax');        // false: a constant has no cells
```

**Name manager.** Ctrl+F3 or `openNameManager()` opens a dialog that lists the names in alphabetical order, each with what it refers to, its sheet, and Go to, Edit and Delete buttons. New name opens a form with three fields: Name, Refers to (it starts as the selection, such as `=Sheet1!$A$1:$A$2`) and Sheet (the sheet that references without a sheet name point to). Enter saves, and a name or formula that isn't valid is explained after its field (and announced). Renaming a name rewrites the formulas that use it on every sheet, so `=SUM(sales)` becomes `=SUM(revenue)`. A rename that would change what a formula means is refused as a whole, with the reason shown in the dialog: when a formula that uses the name has a `LET` or `LAMBDA` variable of the old or the new name, or calls the name and the new name is a function (`double` to `SUM`). Delete asks for confirmation, and formulas that used the name then show `#NAME?`. Go to closes the dialog and selects the name's cells. Every change is announced inside the dialog ("Name added: sales"). Escape goes back from the form to the list and then closes the dialog; a click outside or `closeNameManager()` closes it too.

**Changes by users.** Naming the selection in the Name Box and every change in the name manager fire the cancelable `before-names-change` first (`{ names, source: 'user' }`, with the names as they will be). They are refused in read-only mode, where the name manager still lists the names and goes to them. `defineName()`, `removeName()` and `setNames()` are not restricted and don't fire it.

```js
// Keep users from removing or renaming `tax`
sheet.addEventListener('before-names-change', (e) => {
  if (!e.detail.names.some((n) => n.name === 'tax')) e.preventDefault();
});
```

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

### Toolbar

`<y11n-toolbar>` is an optional element with buttons for what the grid can do. It is separate from the grid, so it goes wherever your layout wants it. Point it at a grid and it wires itself: its controls show the format of the selected cells, follow undo, the filter, frozen panes and read-only mode, and act on the selection.

```html
<y11n-toolbar for="sheet"></y11n-toolbar>
<y11n-spreadsheet id="sheet"></y11n-spreadsheet>
```

```js
toolbar.target = sheet; // from script, instead of the for attribute
```

**Choosing the controls.** `items` lists the controls to show, in order: names of items, names of groups, and `|` for a separator. Without it the toolbar shows everything.

```html
<y11n-toolbar for="sheet" items="history | bold italic textColor | sortAscending filter"></y11n-toolbar>
```

```js
toolbar.items = 'history | bold italic textColor | sortAscending filter';
toolbar.items = ['history', '|', 'bold', 'italic'];              // an array works too, and is type-checked
toolbar.items = TOOLBAR_ITEMS.filter((item) => item !== 'paste'); // everything but
```

| Group | Items |
|-------|-------|
| `history` | `undo`, `redo` |
| `clipboard` | `cut`, `copy`, `paste`, `pasteSpecial` |
| `font` | `bold`, `italic`, `underline`, `strikethrough`, `fontSize`, `textColor`, `backgroundColor` |
| `alignment` | `textAlign` (left, center, right), `verticalAlign` (top, middle, bottom), `wrap` |
| `number` | `numberFormat` (General, Number, Currency, Percent, Scientific, Date), `decreaseDecimals`, `increaseDecimals` |
| `styles` | `conditionalFormat` (a menu), `clearFormat` |
| `cells` | `insert` (a menu: row above / below, column left / right), `delete` (a menu: row, column), `autoFit` |
| `data` | `sortAscending`, `sortDescending`, `filter`, `reapplyFilter`, `list`, `autoSum` |
| `view` | `freeze` (a menu: up to the active cell, first row, first column, unfreeze), `find`, `names` |

A group name is a group of its own, with a separator on each side. An unknown name is skipped with a console warning. `TOOLBAR_GROUPS` and `TOOLBAR_ITEMS` hold these names.

| Property | Attribute | Type | Default | Description |
|----------|-----------|------|---------|-------------|
| `target` | - | `Y11nSpreadsheet \| null` | `null` | The grid to act on and to follow |
| `htmlFor` | `for` | `string` | `''` | The id of that grid, looked up in the toolbar's document or shadow root. The grid may be added later or replaced by another with the same id. `target` wins |
| `items` | `items` | `string \| ToolbarItemSpec[] \| null` | `null` | The controls to show; null for all of them |
| `label` | `label` | `string` | `'Spreadsheet toolbar'` | The toolbar's accessible name. Give each toolbar its own when a page has several |
| `readOnly` | `read-only` | `boolean` | `false` | Disables the controls that change cells |

With a target, the toolbar sets the properties below itself. Without one it shows what they say, so it can be wired by hand: `bold`, `italic`, `underline`, `strikethrough`, `textColor`, `backgroundColor`, `textAlign`, `verticalAlign`, `wrap`, `fontSize`, `numberFormat` (as in a [cell format](#cell-formats)), `canUndo`, `canRedo` (both default to true), `filtered` and `frozen`.

**Events.** Every control fires an event before it acts, and `preventDefault()` stops it, so your code can step in. Without a target the events are all the toolbar does.

| Event | Detail | Fired by |
|-------|--------|----------|
| `format-action` | `{ action, value }` | The formatting controls (the color controls when a color is chosen, not while the picker is dragged). `action` is `bold`, `italic`, `underline`, `strikethrough` (`value`: the new state), `textColor`, `backgroundColor`, `textAlign`, `verticalAlign`, `wrap`, `fontSize`, `numberFormat` (`value`: the new value, so `{ [action]: value }` is a cell format; no `numberFormat` value means General) or `clearFormat` |
| `conditional-format-action` | `{ action }` | The Conditional formatting menu: `duplicates`, `unique`, `aboveAverage`, `belowAverage`, `top10`, `bottom10`, `colorScale`, `dataBar`, `iconSet`, `clear` or `manage` |
| `toolbar-action` | `{ action }` | Everything else: `undo`, `redo`, `cut`, `copy`, `paste`, `pasteSpecial`, `decreaseDecimals`, `increaseDecimals`, `insertRowAbove`, `insertRowBelow`, `insertColumnLeft`, `insertColumnRight`, `deleteRow`, `deleteColumn`, `autoFit`, `sortAscending`, `sortDescending`, `filter`, `reapplyFilter`, `list`, `autoSum`, `freeze`, `freezeFirstRow`, `freezeFirstColumn`, `unfreeze`, `find` or `names` |

```js
// Paste values only, whatever was copied
toolbar.addEventListener('toolbar-action', (e) => {
  if (e.detail.action !== 'paste') return;
  e.preventDefault();
  sheet.paste({ mode: 'values' });
});
```

What the toolbar does with a target is all public API: `toggleFormat()`, `formatSelection()` and `clearSelectionFormat()` for `format-action`; `applyConditionalFormatPreset()`, `clearSelectionConditionalFormats()` and `openConditionalFormatManager()` for `conditional-format-action`; and for `toolbar-action` the methods listed under "The selection, and what a user does to it" in [Methods](#methods), plus `sortSelection()`, `toggleFilter()`, `paste()`, `openPasteSpecial()`, `openFind()` and `openNameManager()`. So a toolbar button does exactly what the keyboard shortcut in its tooltip does, with the same checks, cancelable events and announcements. It reads the grid's state with `getSelectionFormat()`, `canUndo`, `canRedo`, `getFilter()`, `frozenRows`, `frozenCols` and `readOnly`.

**Read-only.** In read-only mode the controls that change cells are disabled. Copy, Fit column width, Filter, Reapply filter, Freeze panes, Find and the name manager stay available, as they do in a read-only grid.

**Keyboard.** The toolbar is one Tab stop. Left and Right move between its controls (wrapping around), Home and End go to the first and last. Tab leaves it, and Shift+Tab comes back to the control it was left on. Down or Up on a menu button opens its menu at the first or last item; in a menu, Up and Down move, Enter chooses, Escape closes it and returns to the button. The buttons' tooltips name the grid's keyboard shortcut for the same action. A control used from the keyboard (or by a screen reader) keeps the focus after it acts, so Bold and then Italic, or Undo several times, work; a control that becomes disabled while it has the focus (Undo after the last step) hands it to its neighbor, and a dialog opened this way returns the focus to its button when it closes.

**Mouse.** A click on a button, on a disabled control or on the toolbar between the controls does not move the keyboard focus: it stays in the grid, so you can click Bold and go on typing. A select or color control needs the focus to open and gives it back to the grid once a value is chosen. A disabled control still shows its tooltip, so a greyed-out icon can be identified. A menu opened with the mouse takes the focus only when Down or Up is pressed; Escape closes it, and so does any other key, which goes to the grid as usual. A click while a cell is being edited commits the edit first; if the edit is a formula kept open on another sheet, the click only ends the edit.

**Theming.** The toolbar is outside the grid's shadow DOM, so it takes its own custom properties: set the ones it shares with the grid (`--ls-font-family`, `--ls-text-color`, `--ls-border-color`, `--ls-focus-ring`) on both elements or on a common ancestor, and see `--ls-toolbar-*` under [Theming](#theming).

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

// Undo or redo became available or stopped being available, or new data started a fresh history
sheet.addEventListener('history-change', (e) => {
  undoButton.disabled = !e.detail.canUndo;
  redoButton.disabled = !e.detail.canRedo;
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

// Cells were turned into lists, or lists into text (Ctrl+J, convertToList(), convertListToText())
sheet.addEventListener('list-change', (e) => {
  const { cellIds, toList, source } = e.detail; // toList: false when lists became text
});

// Another sheet is shown (also when undo/redo shows the sheet it changes)
sheet.addEventListener('sheet-change', (e) => {
  const { sheet: name, previous, source } = e.detail; // sheet names
});

// A sheet was added, removed, renamed or moved
sheet.addEventListener('sheets-change', (e) => {
  const { sheets, kind, sheet: name, previous, source } = e.detail; // kind: 'add' | 'remove' | 'rename' | 'move'; previous: the old name of a renamed sheet
});

// The defined names changed: from code, by the user (Name Box, name manager), or when inserted or deleted rows/columns move their references
sheet.addEventListener('names-change', (e) => {
  const { names, source } = e.detail; // names: [{ name, refersTo, sheet }]
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

// Veto a sheet being added, deleted, renamed or moved from the sheet tabs; same detail as sheets-change, describing the result
sheet.addEventListener('before-sheets-change', (e) => {
  if (e.detail.kind === 'remove') e.preventDefault();
});

// Veto a change to the defined names made in the Name Box or the name manager; detail: { names, source: 'user' }, names as they will be
sheet.addEventListener('before-names-change', (e) => {
  if (e.detail.names.length > 50) e.preventDefault();
});
```

When Enter or Tab commits a rejected value in the inline editor, the editor stays open, shows the message below the cell (`part="validation-message"`, linked with `aria-describedby`) and announces it. If focus moves away instead, the edit is discarded. A rejected formula-bar entry is not applied and the message is announced. For paste, fill and Ctrl+Enter, rejected and read-only cells are skipped and the rest is applied; the count is announced.

`before-change` fires for user edits, pastes, cuts, clears and fills (`detail.operation` says which), after read-only and validation filtering. It does not fire for undo/redo or API calls such as `fromCSV`. Rewritten values are not validated again.

### Formula engine

Formulas start with `=`. A formula is parsed once into a syntax tree and evaluated from it. The engine records the cells and ranges each formula reads, so a change recalculates only the formulas it can reach; a large range counts as one dependency, and its values are read once per recalculation. Supported syntax:

**Operators**: `+`, `-`, `*`, `/`, `^`, `%` (after a value: `50%`), `&` (concat), `=`, `<>`, `<`, `>`, `<=`, `>=`, `:` (range), `@` (one value of a range), `#` (what a cell spills: `A1#`). As in Excel, `^` is left-associative and `-2^2` is 4.

**Values**: numbers (`1.5`, `1e3`), text in double quotes, `TRUE` / `FALSE`, and array literals (`{1,2;3,4}`: commas between columns, semicolons between rows; the elements can be expressions).

**References**: cells (`A1`, `$A$1`), ranges (`B2:D10`), whole columns and rows (`A:A`, `$A:$C`, `1:1`), references to other sheets (`Sheet2!A1`, `'My sheet'!A1:B2`, `Data!A:A`), and [defined names](#defined-names). A whole-column reference is as tall as the grid (`=ROWS(A:A)` is the number of rows) and is read up to the last row that holds anything on any sheet, values and spilled results alike, so `A:A` and `Other!A:A` pair up in `SUMIFS` or `XLOOKUP`. A formula's own spilled results don't count for what it reads: `=A:A*2` gives one result per row of data, wherever it is entered. Two such formulas that each read the other's results can't settle and one of them shows `#SPILL!`.

**Built-in functions**:

| Category | Functions |
|----------|-----------|
| Aggregates | `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `PRODUCT`, `MEDIAN`, `SUMPRODUCT` |
| Statistics | `STDEV`, `STDEV.S`, `STDEV.P`, `VAR`, `VAR.S`, `VAR.P` |
| Conditional aggregates | `SUMIF`, `COUNTIF`, `AVERAGEIF`, `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` |
| Logic | `IF`, `IFS`, `IFERROR`, `IFNA`, `AND`, `OR`, `NOT`, `TRUE`, `FALSE` |
| Lookup | `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`, `XMATCH` |
| References | `ROW`, `COLUMN`, `ROWS`, `COLUMNS`, `OFFSET`, `INDIRECT`, `ADDRESS` |
| Dynamic arrays | `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `TRANSPOSE`, `VSTACK`, `HSTACK`, `TAKE`, `DROP`, `CHOOSECOLS`, `CHOOSEROWS`, `TOCOL`, `TOROW` |
| Several values in a cell | `HAS`, `HASANY`, `HASALL`, `FLATTEN` |
| Variables and functions | `LET`, `LAMBDA`, `MAP`, `REDUCE`, `BYROW`, `BYCOL` |
| Math | `ABS`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `MOD`, `POWER`, `SQRT`, `CEILING`, `FLOOR` |
| Text | `CONCAT`, `TEXTJOIN`, `TEXTSPLIT`, `UPPER`, `LOWER`, `LEN`, `TRIM`, `LEFT`, `RIGHT`, `MID`, `SUBSTITUTE`, `FIND`, `TEXT`, `VALUE` |
| Date | `DATE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `WEEKDAY`, `EDATE`, `EOMONTH` |
| Financial | `PMT`, `FV`, `PV`, `NPV` |
| Information | `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISLOGICAL`, `ISERROR`, `ISERR`, `ISNA`, `NA` |

The functions take their arguments in Excel's order. Omitted arguments use their defaults, e.g. `=XLOOKUP(75, A1:A9, B1:B9, , -1)`. Criteria in `SUMIF`/`COUNTIF`/`AVERAGEIF` and the `*IFS` functions accept operators (`">5"`, `"<>x"`) and wildcards (`*`, `?`, `~` to escape). `XLOOKUP` supports match modes `0`, `-1`, `1` (not wildcard mode `2`) and search modes `1`/`-1`, as does `XMATCH`. `XLOOKUP` returns a single value, or the whole matching row (or column) when the return range has several columns (or rows).

**Arrays.** Operators and single-value functions work element by element on ranges and arrays, and a result that is an array spills (see [Dynamic arrays](#dynamic-arrays)): `=A1:A3*2`, `=IF(A1:A3>1, "yes", "no")`, `=SUMPRODUCT((A1:A3>1)*A1:A3)`. `HAS`, `HASANY`, `HASALL` and `FLATTEN` are described under [Lists and arrays in cells](#lists-and-arrays-in-cells).

**References as values.** `INDEX` returns a reference when it is given one, so `=SUM(A1:INDEX(A1:A9, 3))` adds up `A1:A3`, and a row or column index of `0` gives the whole column or row. `OFFSET` and `INDIRECT` return references too. `ROW()` and `COLUMN()` without an argument are the row and column of the formula's own cell; `ROW`, `COLUMN`, `ROWS` and `COLUMNS` also take a range. `ADDRESS` writes and `INDIRECT` reads `A1`-style references only.

**LET and LAMBDA.** `LET` names intermediate values: `=LET(net, B2-C2, net*net)`. `LAMBDA` makes a function, which can be called on the spot (`=LAMBDA(x, x*2)(4)`), bound by `LET`, stored as a defined name (`=double(4)`), or passed to `MAP`, `REDUCE`, `BYROW` and `BYCOL`:

```
=MAP(A1:A3, LAMBDA(v, v*2))
=REDUCE(0, A1:A3, LAMBDA(total, v, total+v))
=BYROW(A1:C3, LAMBDA(row, SUM(row)))
```

**Errors and blanks.** A name or function that is not defined is `#NAME?`. A formula that can't be parsed is `#ERROR!`. `#SPILL!` and `#CALC!` (for example a `FILTER` that keeps nothing) come with arrays. A blank cell is `""` to `&` and equal to both `0` and `""`. Aggregates skip a blank cell, whether it is in a range or given on its own: `=COUNT(A1)` and `=COUNTA(A1)` are 0 for an empty `A1`. Error text typed into a cell (`#N/A`) is an error to the formulas that read it, `AND` and `OR` over a range that holds an error give that error, and a result too large for a number is `#NUM!` (`=2^1024`).

**Custom functions**:

```js
sheet.registerFunction('TAX', (ctx, amount, rate) => {
  return Number(amount) * Number(rate);
});
// Use in cells: =TAX(B1, 0.2)
```

The `ctx` argument provides `getCellValue(ref)` and `getRangeValues(start, end)` for reading other cells, `cell` (the `{ row, col }` of the formula's cell, 0-based) and `sheet` (the name of the formula's sheet).

`registerFunction(name, fn, options)` takes two options. `args` says how each argument is passed, and `volatile: true` recalculates the function on every change, like `NOW`.

```js
// 'number' arguments make the function work element by element:
// =NET(B2, 0.19) is one value, =NET(B2:B9, 0.19) spills eight
sheet.registerFunction('NET', (ctx, gross, rate) => gross / (1 + rate), { args: ['number', 'number'] });

sheet.registerFunction('ROLL', () => 1 + Math.floor(Math.random() * 6), { volatile: true });
```

| `args` entry | The function receives |
|--------------|-----------------------|
| `'number'`, `'text'`, `'value'` | A single value. Given an array or a range of several cells, the function is called once per element and the results form an array. A blank cell arrives as `0`, `''` or `undefined` respectively. `'value'` also receives an error in the argument as its code (`'#N/A'`) instead of failing the call |
| `'array'` (default) | The argument as it is: a `RangeValue` (`{ values, rows, cols }`, row by row) for ranges and arrays |
| `'reference'` | The reference itself, a `RefValue` (`top`, `left`, `bottom`, `right`, 0-based; `rows`, `cols`; `sheet.name`), not its values |

Arguments beyond the listed kinds use the last one.

A function can return a JavaScript array, which spills like any other array result: an array of arrays is a table of rows, and a flat array is a column.

```js
// =MULTIPLES(3, 4) spills 3, 6, 9, 12 down a column
sheet.registerFunction('MULTIPLES', (ctx, n, count) =>
  Array.from({ length: Number(count) }, (_, i) => Number(n) * (i + 1)));
```

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
| `--ls-formula-ref-bg` | `#fff` | Name Box background |
| `--ls-formula-input-bg` | `#fff` | Background of the formula input and the Raw / Formatted toggle |
| `--ls-name-box-width` | `9em` | Width of the Name Box |
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
| `--ls-spill-border` | `#1a73e8` | Outline of the spill range the active cell is in |
| `--ls-array-icon-color` / `--ls-array-icon-hover-bg` | `#5f6368` / `rgba(0, 0, 0, 0.08)` | Icon of a cell that holds several values, and its background when hovered |
| `--ls-sheet-bar-bg` / `--ls-sheet-bar-height` | `--ls-header-bg` / `30px` | Sheet tab bar |
| `--ls-sheet-tab-color` / `--ls-sheet-tab-active-color` | `--ls-text-color` | Text of a sheet tab / of the active one |
| `--ls-sheet-tab-active-bg` | `--ls-cell-bg` | Background of the active sheet tab |
| `--ls-sheet-tab-accent` | `#1a73e8` | Line above the active sheet tab, and the border of the rename field |
| `--ls-sheet-tab-hover-bg` | 10% of the text color | Hovered sheet tabs, buttons and menu items |
| `--ls-sheet-tab-max-width` | `160px` | Widest a sheet tab gets; longer names are cut with an ellipsis |
| `--ls-sheet-delete-bg` / `--ls-sheet-delete-color` | `#c5221f` / `#fff` | Delete button of the sheet delete confirmation |
| `--ls-name-delete-bg` / `--ls-name-delete-color` | `--ls-sheet-delete-bg` / `--ls-sheet-delete-color` | Delete button of the name manager's delete confirmation |
| `--ls-find-bg` | `#fff` | Find bar background |
| `--ls-find-status-color` | `#555` | Find bar match count |
| `--ls-status-bg` / `--ls-status-color` | `#f8fafc` / `#555` | Status bar under the grid |
| `--ls-toolbar-bg` | `#f8fafc` | Toolbar background (set the `--ls-toolbar-*` properties on `<y11n-toolbar>`) |
| `--ls-toolbar-hover-bg` | `#e8eaed` | Hovered toolbar button |
| `--ls-toolbar-active-bg` / `--ls-toolbar-active-border` | `#d3e3fd` / `#a8c7fa` | Pressed toolbar button (Bold on a bold cell) |
| `--ls-toolbar-input-bg` | `#fff` | Background of the toolbar's Font size and Number format selects |

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
| Alt + Down | On a filter header cell: open its filter menu. On a cell that holds several values: show them (on macOS use Option) |
| Alt + Shift + V | Show the values of a cell that holds several, also on a filter header cell (on macOS use Option) |
| Ctrl/Cmd + J | Turn the comma-separated text of the selected cells into lists, or lists back into text; on an empty cell, start a new list. While editing: commit what was typed as a list |
| Ctrl/Cmd + PageDown / PageUp (or Alt + PageDown / PageUp) | Show the next / previous sheet. While typing a formula: show it without ending the edit, to pick a reference there |
| Ctrl/Cmd + G | Go to the Name Box. There: type a cell, a range or a name and press Enter to select it, or a new name to name the selection; Escape returns to the grid |
| Ctrl/Cmd + F3 | Open the name manager. In it: Enter saves the form, Escape goes back from the form to the list and then closes the dialog |
| Ctrl + Alt + L | Reapply the filter, after edits (on macOS use Ctrl+Option) |
| Alt + Shift + C | Open the conditional-format rule manager (on macOS use Option) |
| Ctrl/Cmd + F | Find. In the find bar: Enter / Shift+Enter (or F3 / Shift+F3) next / previous match, Escape closes |
| Ctrl+C / X / V | Copy / Cut / Paste |
| Ctrl + Shift + V | Paste values only |
| Ctrl + Alt + V | Paste Special: values, formats, transpose (on macOS use Ctrl+Option) |
| Ctrl+Z | Undo |
| Ctrl+Shift+Z / Ctrl+Y | Redo |
| Any printable character | Start editing with that character |

During formula editing, arrow keys insert cell references instead of navigating. The mouse does the same where a reference fits, which is right after `=`, an operator, `(`, `,` or `:`: a click on a cell puts its reference into the formula, and a drag puts a range (`=SUM(` and a drag from A1 to B2 gives `=SUM(A1:B2`). Typing then goes on after the reference. Anywhere else in a formula, and in any other edit, a click on a cell commits the edit. To pick cells on another sheet, see [Multiple sheets](#multiple-sheets).

On a sheet tab:

| Key | Action |
|-----|--------|
| Left / Right | Go to the previous / next tab and show its sheet (wraps around) |
| Home / End | Go to the first / last tab and show its sheet |
| F2 | Rename the sheet in place (Enter commits, Escape cancels) |
| Delete, Backspace | Delete the sheet, after a confirmation |
| Ctrl/Cmd + Shift + Left / Right | Move the sheet one place left / right |

A sheet tab can also be dragged to another position; Escape during the drag puts it back.

Drag the right border of a column header to resize the column (Escape during the drag cancels it); with whole columns selected, all of them get the new width. Double-click the border to fit the column to its content.

With the mouse, drag the square at the bottom-right corner of the selection (the fill handle) to fill adjacent cells. A run of two or more numbers continues its linear trend (`1, 2` → `3, 4, 5`; `10, 20` → `30, 40`), a single date steps by a day, day and month names continue and wrap (`Mon` → `Tue`, `Nov, Dec` → `Jan`, keeping their case), and text ending in a number counts up (`Item 1` → `Item 2`, `Q09` → `Q10`; `Row 1, Row 3` → `Row 5`). Other text, formulas and single numbers repeat, with relative references shifted. Double-click the fill handle to fill down as far as the column to the left (or right) has data, stopping above any filled cell. From the keyboard, select the cells to continue plus the empty cells after them and press Ctrl+Alt+D (down) or Ctrl+Alt+R (right); Ctrl+D / Ctrl+R copy instead of continuing a series.

## Types

All types are exported for TypeScript consumers:

```ts
import type {
  CellData,          // { rawValue, displayValue, type, format?, readOnly?, spill?, spilledFrom?, array? }
  CellArray,         // { rows, cols, values: CellArrayItem[] }: the values of a cell that holds several, row by row
  CellArrayItem,     // string | number | boolean | null | CellArray
  CellCoord,         // { row, col }
  GridData,          // Map<string, CellData>
  SelectionRange,    // { start: CellCoord, end: CellCoord }
  FormulaContext,    // { getCellValue, getRangeValues, cell?, sheet? }
  FormulaFunction,   // (ctx: FormulaContext, ...args: unknown[]) => unknown
  FormulaFunctionOptions, // { args?: FormulaArgumentKind[], volatile? }
  FormulaArgumentKind,    // 'number' | 'text' | 'value' | 'array' | 'reference'
  CellChangeDetail,
  SelectionChangeDetail,
  SelectionInfo,         // { range, activeCell }: what getSelection() returns
  HistoryChangeDetail,   // { canUndo, canRedo }
  ToolbarItem,           // 'undo' | 'bold' | 'filter' | ...: a control of <y11n-toolbar>
  ToolbarGroup,          // 'history' | 'clipboard' | 'font' | 'alignment' | 'number' | 'styles' | 'cells' | 'data' | 'view'
  ToolbarItemSpec,       // ToolbarItem | ToolbarGroup | '|'
  FormatActionDetail,    // { action, value? }
  ConditionalFormatActionDetail, // { action }
  ToolbarActionDetail,   // { action: ToolbarAction }
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
  AddSheetOptions,       // { data?, rows?, cols?, index?, activate? }
  SheetChangeDetail,     // { sheet, previous, source }
  SheetsChangeDetail,    // { sheets, kind: 'add' | 'remove' | 'rename' | 'move', sheet, previous?, source }
  DefinedName,           // { name, refersTo, sheet }
  NamesChangeDetail,     // { names: DefinedName[], source }
  ListConversionDetail,  // { cellIds, toList, source }
  ListItem,              // string | number | boolean
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

// With several sheets, the third argument names the formula's sheet and the sheet that changed
adjustFormulaForStructureChange('=SUM(Data!A1:A3)+A2', { axis: 'row', kind: 'insert', index: 1, count: 1 },
  { formulaSheet: 'Summary', changedSheet: 'Data' })
// '=SUM(Data!A1:A4)+A2'
```

Helpers for lists, sheet names and formula text:

```ts
import { listToFormula, listItemsOf, listToText, parseListText } from 'y11n-spreadsheet';

listToFormula(['Ann', 'Bob', 3])          // '={{"Ann";"Bob";3}}', the rawValue of a list
listItemsOf('={{"Ann";"Bob";3}}')         // ['Ann', 'Bob', 3]; null for a value that isn't a list
listToText(['Ann', 'Smith, J', '007'])    // 'Ann, "Smith, J", "007"', the text a list is edited as
parseListText('Ann, Bob, "Smith, J", 3')  // ['Ann', 'Bob', 'Smith, J', 3]

import { validateSheetName, isValidName, renameSheetInFormula, renameNameInFormula } from 'y11n-spreadsheet';

validateSheetName('a/b', [])              // 'A sheet name cannot contain : \ / ? * [ or ]'
validateSheetName('Data', ['Summary'])    // null: the name is fine
isValidName('tax')                        // true: usable as a defined name ('A1' is not)
renameSheetInFormula('=Data!A1', 'Data', 'My data') // "='My data'!A1"
renameNameInFormula('=SUM(sales)&"sales"', 'sales', 'revenue') // '=SUM(revenue)&"sales"': text in quotes is left alone

import { parseFormula, tokenizeFormula, mapFormulaReferences, formatReference } from 'y11n-spreadsheet';

parseFormula('A1+2')                      // the syntax tree (FormulaNode) of a formula, given without its '='
tokenizeFormula('Data!A1+2')              // its tokens (FormulaToken[]), with offsets and parsed references
mapFormulaReferences('=Data!A1+B2', (ref, text) => text.sheet + '$' + formatReference(ref))
// '=Data!$A1+$B2': rewrite every reference, leaving the rest of the formula as written
```

## Exports

The package also exports internal modules if you need them standalone:

```ts
import { FormulaEngine } from 'y11n-spreadsheet';
import { SelectionManager } from 'y11n-spreadsheet';
import { ClipboardManager } from 'y11n-spreadsheet';
import { Y11nFormulaBar } from 'y11n-spreadsheet';
import { Y11nToolbar, TOOLBAR_GROUPS, TOOLBAR_ITEMS } from 'y11n-spreadsheet';
import { Y11nConditionalFormatDialog } from 'y11n-spreadsheet';
import { Y11nSheetTabs, Y11nNameManager } from 'y11n-spreadsheet';

// Conditional formatting without the component: validate rules and evaluate them over your own data
import { ConditionalFormatEvaluator, normalizeConditionalRule, describeCondition } from 'y11n-spreadsheet';

// The values formulas compute with: arrays, references and LAMBDA functions
import { RangeValue, RefValue, LambdaValue } from 'y11n-spreadsheet';
```

`FormulaEngine` works without the component, on one or more `GridData` maps:

```ts
const engine = new FormulaEngine();        // starts with one sheet, 'Sheet1'
engine.setData(data);                      // the cells of the active sheet
engine.setSheetSize(100, 26);              // bounds spills and whole-column references
engine.addSheet('Data', otherData);
engine.defineName('tax', '=0.19');
engine.recalculate();                      // computes displayValue and type, and spills arrays into the maps
engine.changes                             // Map of sheet name → keys of the cells the last recalculation changed
engine.spillRange('0:2')                   // { rows: 3, cols: 1 }, or null when the cell doesn't spill
engine.blockedSpill('0:6')                 // for a formula showing #SPILL!: the size its result needs; null otherwise
engine.referenceOf('Data!A1:C3')           // { sheet: 'Data', start: { row: 0, col: 0 }, end: { row: 2, col: 2 } }
engine.referenceOf('tax')                  // null: not a reference, or a name that doesn't stand for cells

engine.evaluate('=SUM(A1:A3)')             // { displayValue: '6', type: 'number' }
engine.evaluateValue('=SUM(A1:A3)')        // 6: the value instead of its text
engine.evaluateValue('=A1:A3*2')           // a RangeValue: { values: [2, 4, 6], rows: 3, cols: 1 }
engine.evaluateValue('=1/0')               // '#DIV/0!'
engine.evaluateCommitted('=A1>0')          // against the results of the last recalculation: precedents are not computed again
```

`setData()`, `updateData()`, `recalculateAffected()`, `evaluate()` and `evaluateValue()` take a sheet name as an optional argument (default: the active sheet, set with `setActiveSheet()`). `<y11n-sheet-tabs>` holds no sheets itself: give it `sheets` (`[{ id, name }]`) and `activeId`, and apply the `sheet-select`, `sheet-add`, `sheet-rename`, `sheet-remove` and `sheet-move` events it fires (their detail types are exported as `SheetSelectDetail`, `SheetRenameDetail`, `SheetRemoveDetail` and `SheetMoveDetail`).

`<y11n-name-manager>` works the same way: give it `names` (`DefinedName[]`), `sheets` (the sheet names), `activeSheet`, `selectionRef` (what a new name refers to at first) and `status` (the outcome to announce), and apply its `name-save` (`NameSaveDetail`: `{ name, refersTo, sheet, previous? }`), `name-remove` and `name-go` (`{ name }`) and `name-close` events.

The Name Box is part of `<y11n-formula-bar>`: set `selectionLabel` to show a name instead of `cellRef`, and `names` for its suggestions. It fires `name-box-commit` (`NameBoxCommitDetail`: `{ value }`) on Enter and `name-box-cancel` on Escape; `resetNameBox()` shows the selection again and `focusNameBox()` focuses the box. In the formula bar's shadow root the Name Box is `input.name-box` and the formula input is `input.formula-input`.

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

The `index.html` at the project root has a test harness.

Storybook is the getting-started site. It has three parts: **Getting Started** (installation and a five-step tutorial), **Examples** (one runnable file per feature, shown in the Code panel next to the result) and **Reference** (the sections of this README, plus the list of every formula function).

To add an example, write `stories/examples/<name>.ts`: a default-exported function that returns the element to show, with a comment at the top saying what it demonstrates. Import the package as `'y11n-spreadsheet'` and use only its public API, since the file is displayed as it is. Then add a story for it in the group's `stories/*.stories.ts` with `parameters: example('<name>')`, and its id to `e2e/storybook-a11y.spec.ts`.

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
    y11n-formula-bar.ts             # Formula bar sub-component (Name Box, raw / formatted input)
    y11n-toolbar.ts                 # Toolbar: configurable items, wires itself to a grid (target / for)
    y11n-conditional-format-dialog.ts # Rule manager dialog
    conditional-icons.ts            # Icon-set icons (SVG)
    y11n-sheet-tabs.ts              # Sheet tab bar (switch, add, rename, delete, reorder)
    y11n-name-manager.ts            # Name manager dialog; holds no names, fires name-save/name-remove/name-go/name-close
  controllers/
    selection-manager.ts            # Selection state (Lit ReactiveController)
    clipboard-manager.ts            # Clipboard operations (TSV + HTML formats)
  engine/
    formula-engine.ts               # Formula evaluator: arrays and spilling, sheets, names, dependency tracking
    parser.ts                       # Formula tokenizer and parser (text to syntax tree)
    values.ts                       # Value types (RangeValue, RefValue, LambdaValue) and coercions
    functions.ts                    # Built-in functions
  utils/
    delimited.ts                    # CSV/TSV parsing and serialization
    references.ts                   # Formula reference rewriting for row/column insert/delete
    list.ts                         # Lists in cells: conversion between items, list text and the stored formula
    sheet-names.ts                  # Sheet name rules
    conditional-format.ts           # Conditional-format rule validation and evaluation
    input-parse.ts                  # Recognises typed percentages, currency, dates and times
    series.ts                       # Fill-handle series (numbers, dates, day/month names, numbered text)
  __tests__/                        # Unit tests colocated with source
e2e/                                # Playwright E2E tests
stories/
  guides/                           # Getting Started pages (MDX)
  examples/                         # One file per example: what a story runs and what its Code panel shows
  reference/                        # Reference pages (MDX), made of sections of this README
  support/                          # example(): story parameters for an example file; readmeSection()
  *.stories.ts                      # One per example group
```

### Build

```bash
npm run build  # Outputs ES module to dist/
```

The build externalizes `lit` as a peer dependency and produces `dist/index.js` with TypeScript declarations and source maps.

## License

MIT
