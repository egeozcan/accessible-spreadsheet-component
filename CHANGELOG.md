# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- The toolbar, now `<y11n-toolbar>`, covers what the grid can do, wires itself to a grid, and shows only the controls you choose.
  - `target` (or the `for` attribute with the grid's id) connects it: the controls show the state of the selection, follow undo, the filter, frozen panes and read-only mode, and act on the selection. Without a target it works as before: it shows its properties and fires events.
  - New controls: undo and redo; cut, copy, paste and Paste Special; vertical alignment and wrap; a Number format select with decrease / increase decimals; Insert and Delete menus for rows and columns; fit column width; sort ascending / descending; filter and reapply filter; list in cell; AutoSum; a Freeze panes menu; find; the name manager.
  - `items` chooses the controls and their order, as text or an array: names of items, names of groups (`history`, `clipboard`, `font`, `alignment`, `number`, `styles`, `cells`, `data`, `view`) and `|` for a separator. `TOOLBAR_GROUPS` and `TOOLBAR_ITEMS` are exported.
  - Events: a new `toolbar-action` for the controls that are not formats, and `format-action` now also carries `verticalAlign`, `wrap` and `numberFormat`. With a target, all three events are fired before the toolbar acts, and `preventDefault()` stops it.
  - New properties `verticalAlign`, `wrap`, `numberFormat`, `canUndo`, `canRedo`, `filtered`, `frozen` and `label` (its accessible name); new CSS custom property `--ls-toolbar-input-bg`.
  - The toolbar is one Tab stop that remembers the control it was left on, color inputs and selects included.
- Methods that do to the selection what a user does, to back your own buttons: `getSelection()`, `getSelectionFormat()`, `formatSelection()`, `clearSelectionFormat()`, `changeDecimals()`, `insertRowsAtSelection()`, `insertColumnsAtSelection()`, `deleteSelectedRows()`, `deleteSelectedColumns()`, `autoFitSelectedColumns()`, `autoSum()`, `toggleLists()`, `copy()` and `cut()`.
- `freezePanesAsUser()` and `reapplyFilterAsUser()`: what Alt+Shift+F and Ctrl+Alt+L do, with their checks, events and announcements, to back buttons.
- `undo()`, `redo()`, `canUndo`, `canRedo` and the `history-change` event (also fired when `setData()` starts a fresh history).
- `announce(message)`: a message for screen readers through the grid's live region.
- Storybook: a Toolbar example group and reference page, and a toolbar with the fitting controls in most other examples.
- A new formula engine: formulas are parsed once into an AST and then evaluated, instead of being evaluated while they are parsed. Everything below builds on it; existing formulas keep their results.
  - Grammar: `^` (left-associative, and `-2^2` is 4 as in Excel), postfix `%`, scientific literals (`1e3`), array literals (`{1,2;3,4}`, elements can be expressions), `@` and the spill operator `A1#`.
  - `LET` and `LAMBDA` (`=LAMBDA(x, x*2)(4)`, lambdas bound by `LET`, by a defined name, or passed to `MAP`, `REDUCE`, `BYROW`, `BYCOL`), `ROW()` / `COLUMN()` of the formula's own cell, `IFNA`, `NA`, `TRUE` / `FALSE`.
  - A name that is not defined is `#NAME?` (it used to be `#ERROR!`). A blank cell is `""` to `&` and equal to both `0` and `""`.
- Array values and dynamic arrays. Operators and single-value functions work element by element on ranges and arrays (`=A1:A3*2`, `=UPPER(A1:A3)`, `=IF(A1:A3>1, "yes", "no")`), and a formula whose result is an array spills into the cells below and to the right of it.
  - `#SPILL!` when a value is in the way or the result would leave the grid; clearing the obstruction lets it through. `=C1#` refers to everything `C1` spills, `@` picks the value on the formula's own row or column.
  - Functions: `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `TRANSPOSE`, `XMATCH`, `VSTACK`, `HSTACK`, `TAKE`, `DROP`, `CHOOSECOLS`, `CHOOSEROWS`, `TOCOL`, `TOROW`, `TEXTSPLIT`, `MAP`, `REDUCE`, `BYROW`, `BYCOL`. `XLOOKUP` returns the whole row (or column) when the return range has several. `SUMPRODUCT((A1:A3>1)*A1:A3)` works.
  - Copy, paste values and fill keep what cells computed: a spilled value on its own is written as a value, a spill range copied with its formula as the formula (also into another spreadsheet, through the HTML clipboard), text that reads like a number (`00456`) stays text (`="00456"`), and a cell that holds several values is written as an array constant (`arrayToFormula()`).
  - Spilled values are computed, not stored: the cells that show them have an empty `rawValue` and `CellData.spilledFrom` (the formula cell's key), and the formula cell has `CellData.spill` (the range's size). Typing into one blocks the spill; sorting, inserting rows and `setData()` leave them to their formula; copy, CSV export, filters, the status bar and number formats use their values.
  - Accessibility: the active cell announces "spilled from C1" / "spills to C3" and why a formula is `#SPILL!`; the spill range is outlined while the active cell is in it (`--ls-spill-border`); the formula bar shows the formula behind a spilled value greyed out.
- Several values in one cell: lists and arrays in cells, as in [Excel](https://techcommunity.microsoft.com/blog/excelblog/excel-now-supports-multiple-values-in-a-single-cell/4549756).
  - Lists: Ctrl+J turns the selected cells' comma-separated text into lists (`Ann, Bob` becomes two values) and lists back into text; in the editor it commits what was typed as a list, and on an empty cell it starts one. A list is edited as its items (`Ann, Bob, "Smith, J"`); comma-separated text typed right below a list becomes a list too. API: `convertToList()`, `convertListToText()`, `getCellValues()`; fires `list-change`. A list is stored as the formula `={{"Ann";"Bob"}}`, so undo, copy and export need nothing new.
  - Arrays in cells: wrapping a formula in braces keeps a result that would spill in one cell (`={FILTER(A2:A9, B2:B9="x")}`), and arrays nest (`={{1,2,3};{4,5,6}}` spills one array per row; `=TEXTSPLIT(A1:A3, ",")` gives one array per text).
  - A reference to such a cell gives all its values: `=B2` spills them, `=SUM(B2)`, `=COUNTA(B2)` and `=B2*2` work on the items. New functions `HAS`, `HASANY` and `HASALL` test for items, and `FLATTEN(array, [pad_value], [levels])` lays nested arrays out as a table.
  - Filters list and match the individual items of lists instead of the whole text.
  - The cell shows the values as text with an icon; `CellData.array` has them as data. Alt+Down (or the icon) opens a popup with the values, where nested arrays can be opened (Alt+Shift+V does the same, also on a filter header cell, where Alt+Down opens the filter menu); `showCellValues()` / `closeCellValues()`. The cell's accessible name and the announcement say "list of 3 items" or "array of 2 rows and 3 columns".
- References as values: whole columns and rows (`A:A`, `$A:$C`, `1:1`), `OFFSET`, `INDIRECT`, `INDEX` that returns a reference (`=SUM(A1:INDEX(A1:A9, 3))`, and whole rows or columns with a `0` index), `ROWS`, `COLUMNS`, `ROW` / `COLUMN` of a range, `ADDRESS`.
  - Dependencies on larger ranges are tracked as one range node instead of one edge per cell, and the values of a range are read once per recalculation, so lookups over big tables recalculate much faster (3,000 `VLOOKUP`s over 3,000 rows: from about 8 s to 0.1 s).
  - Whole-column references are as tall as the grid (`=ROWS(A:A)`), and are read up to the last row that holds anything on any sheet (values and spilled results), so `A:A` and `Other!A:A` pair up in `SUMIFS`. A formula's own spill doesn't count for what it reads: `=A:A*2` gives one result per row of data and gets shorter again when data is deleted. Formulas that read whole columns are recomputed when that last row moves.
- Multiple sheets and sheet-qualified references (`=Sheet2!A1`, `='My sheet'!A1:B2`, `=SUM(Data!A:A)`).
  - API: `sheets`, `activeSheet`, `addSheet()`, `removeSheet()`, `renameSheet()`, `moveSheet()`, `activateSheet()`, `getSheetData()`, `setSheetData()`. `data`, `rows`, `cols`, the selection and every method that takes cell keys work on the sheet being shown; each sheet keeps its own size, selection, scroll position, column widths, frozen panes, filter and conditional formats.
  - Sheet tabs (`<y11n-sheet-tabs>`) appear under the grid when there is more than one sheet, or always with the `sheet-tabs` attribute: a tab list with arrow-key navigation, rename in place (F2 or double-click), delete with confirmation, reorder (Ctrl+Shift+Left/Right, or by dragging a tab) and a menu with the same actions. Ctrl+PageDown / Ctrl+PageUp (or Alt+PageDown / Alt+PageUp) switch sheets from the grid.
  - Undo and redo cover all sheets and show the sheet they change. Inserting or deleting rows and columns rewrites the references to that sheet on every sheet; renaming a sheet rewrites formulas, defined names, conditional formats and the undo history. Deleting a sheet keeps the text of formulas that refer to it and shows `#REF!`.
  - Events: `sheet-change`, `sheets-change`, and the cancelable `before-sheets-change` for user changes.
- Defined names: `defineName()`, `removeName()`, `getNames()`, `setNames()` for ranges (`=SUM(sales)`), constants and LAMBDAs (`=double(4)`). Names are workbook-wide, move with inserted and deleted rows and columns, and fire `names-change`.
  - Name Box: the cell reference in the formula bar is now an input. Type a cell, a range, a sheet-qualified reference or a name and press Enter to go there (`goTo()` does the same); type a new name to name the selection. It shows the name when the selection is a named range, and suggests the defined names. Ctrl+G moves to it.
  - Name manager (Ctrl+F3, `openNameManager()` / `closeNameManager()`, `<y11n-name-manager>`): a dialog that lists the names and adds, edits, deletes and goes to them. Renaming a name rewrites the formulas that use it, on every sheet; a rename that would change what a formula means (the formula has a variable of that name, or calls the name and the new name is a function) is refused.
  - A name whose sheet is removed keeps working as a constant or with sheet-qualified references (`getNames()` reports `sheet: ''`), and `setNames(getNames())` restores it unchanged. `names-change` also fires when a sheet rename or removal changes the names.
  - Changes users make to names fire the cancelable `before-names-change` and are refused in read-only mode.
- Picking references while typing a formula: where a reference fits (after `=`, an operator, `(` or `,`), a click on a cell puts its reference into the formula and a drag puts a range. With a formula open, clicking another sheet's tab (or Ctrl+PageDown / PageUp, Alt+PageDown / PageUp in the editor) shows that sheet without ending the edit, and cells picked there with the mouse or the arrow keys are written with the sheet's name (`=A1+Orders!B3`). Enter and Escape return to the edited cell's sheet.
- `registerFunction(name, fn, options)`: `options.args` says how each argument is passed (`'number'`, `'text'` or `'value'` arguments make the function work element by element on arrays, `'reference'` passes a `RefValue`), `options.volatile` recalculates the function on every change. Functions receive `ctx.cell` and `ctx.sheet`. A function can return an array (a JS array of rows, or a flat array for a column), which spills.
- `FormulaEngine`: `evaluateValue()` (the value instead of its text, a `RangeValue` for arrays), `evaluateCommitted()` (evaluates against the results of the last recalculation instead of computing precedents again; conditional-format formulas use it, so they test what the grid shows and `ROW()` / `COLUMN()` in them are the tested cell's), `referenceOf()` (the cells a reference or name stands for), sheets (`addSheet()`, `setActiveSheet()`, a `sheet` argument on the data and recalculation methods, `changes` by sheet), names, `spillRange()`, `setSheetSize()`. New exports: `RefValue`, `LambdaValue`, `parseFormula()`, `tokenizeFormula()`, `mapFormulaReferences()`, `formatReference()`, `isValidName()`, `renameSheetInFormula()`, `renameNameInFormula()`, the list helpers (`listToFormula()`, `listItemsOf()`, `listToText()`, `parseListText()`), `validateSheetName()`, `Y11nSheetTabs`, `Y11nNameManager`.

- Conditional formatting. A rule is a range, a condition and a format; rules live beside the data, are evaluated after recalculation and are merged over the cells' own formats when they are drawn, so `getData()`, copy, sort and CSV export still see the real formats.
  - Highlight rules: cell value (greater, less, between, equal, with numbers, text or `=formula` operands), text contains / begins / ends, dates (today, last 7 days, this week, next month, ...), blanks, errors, duplicates and unique values, top or bottom N (or percent), above or below average, and formula rules written for the top-left cell and shifted like a filled formula.
  - In-cell graphics: 2- and 3-color scales, data bars (negative bars grow left from a zero axis) and five icon sets (`arrows3`, `arrows5`, `shapes3`, `symbols3`, `bars5`).
  - Excel's precedence: rules are kept in priority order, formats merge with the earlier rule winning, and `stopIfTrue` ends the evaluation for a cell.
  - Accessibility: every rule has a short label (generated, or set with `label`) that is part of each matching cell's accessible name and of the active-cell announcement ("B6: 150, above average"). Color scales and data bars announce the cell's position, icons have text alternatives and differ in shape, color scales pick a contrasting text color, and forced-colors mode shows an edge marker for highlights, a meter line for scales and bars, and system-colored icons.
  - API: `addConditionalFormat()`, `updateConditionalFormat()`, `removeConditionalFormat()`, `moveConditionalFormat()`, `getConditionalFormats()`, `setConditionalFormats()`, `clearConditionalFormats()`, `getConditionalFormatResult()`, and for user actions `applyConditionalFormatPreset()`, `clearSelectionConditionalFormats()` and `openConditionalFormatManager()` / `closeConditionalFormatManager()`.
  - A rule manager dialog (Alt+Shift+C) to add, edit, reorder and delete rules and set stop-if-true, and a Conditional formatting menu on the toolbar with one-step presets (fires `conditional-format-action`).
  - Undoable; fires `conditional-format-change` and, for user changes, the cancelable `before-conditional-format-change`. Rules move with inserted and deleted rows/columns (ranges and rule formulas are rewritten), survive `setData()`, and are refused for users in read-only mode.
  - Range statistics (averages, top/bottom cut-offs, duplicate counts, scale bounds) are cached per rule and recomputed only when a cell inside the rule's range changes.
  - New CSS custom properties `--ls-data-bar-color` and `--ls-data-bar-negative-color`. New exports: `ConditionalFormatEvaluator`, `normalizeConditionalRule()`, `describeCondition()`, `conditionalFormatPreset()`, `adjustRulesForStructureChange()`, `clearRulesInRange()`, `rangeToA1()`, `parseA1Range()`, `Y11nConditionalFormatDialog` and the `ConditionalFormat*` types.

### Changed

- `<y11n-format-toolbar>` is now `<y11n-toolbar>` and its class `Y11nToolbar` (was `Y11nFormatToolbar`), since it is no longer only about formats. Rename the element in your markup and the import, if you use the class.
- `<y11n-toolbar>` shows all its controls by default, the new ones included. A toolbar that is wired by hand needs `items="font textAlign styles"` (the earlier set) or handlers for `toolbar-action`; or give it a `target` and drop the handlers.
- In read-only mode the toolbar keeps the controls that a read-only grid still supports (copy, fit column width, filter, freeze panes, find, name manager) instead of disabling everything.
- The toolbar's accessible name is "Spreadsheet toolbar" (was "Cell formatting"); set `label` to change it. Its Clear formatting button has an icon instead of ✘.
- The toolbar's color controls fire `format-action` when a color is chosen (`change`), not on every step of a drag in the color picker, so a color is one undo step.
- A change made while the keyboard focus is outside the grid (a toolbar button, a button of the page that calls `undo()` or `goTo()`) no longer moves the focus to the active cell. The rule manager, name manager and Paste Special return the focus to the outside control that opened them.
- `computeSelectionFormat()` also reports `verticalAlign`, `wrap` and `numberFormat`.
- The formula bar's cell reference is an `<input>` (the Name Box) instead of an `<output>`, and the formula input has the class `formula-input`: code that reached into the formula bar with `querySelector('input')` now finds the Name Box first.
- A click on a cell while a formula is being typed, with the caret where a reference fits, picks that cell instead of committing the edit.
- `adjustFormulaForStructureChange()` and `shiftFormulaReferences()` use the formula tokenizer instead of a regular expression, so they handle whole-column, whole-row and sheet-qualified references. `adjustFormulaForStructureChange()` takes a third argument naming the formula's sheet and the sheet that changed.
- Formulas that used to fail now work, which can change what a sheet shows: `=A1:A3` and other array results spill instead of showing `#VALUE!`, and `=IF(A1:A3>1, 1, 0)` returns an array instead of testing only the range's first cell.
- Errors propagate in more places, as in Excel: `AND` / `OR` over a range that contains an error give that error, and error text typed into a cell (`#N/A`) is an error to the formulas that read it. Text always compares as less than `TRUE` / `FALSE`. `=2^1024` is `#NUM!`.
- Aggregates skip a blank cell given as a single reference, as they do in ranges: `=COUNT(A1)` and `=COUNTA(A1)` are 0 for an empty `A1`, and `=AVERAGE(A1, B1)` ignores an empty cell.

### Fixed

- Selected cells with a background color of their own (from a format, and now from a conditional format) show the selection tint over the fill. The inline background used to hide it, leaving only the active cell's outline.

## [1.4.0] - 2026-11-10

### Added

- Typed input is recognised as in Excel: `15%`, `$1,200` (also `€`, `£`, `¥`), `1,234.5`, `2026-01-05`, `14:30` and `2026-01-05 14:30` are stored as numbers and the cell gets the matching number format (unless it already has one of that type). A plain number typed into a percent cell is a percentage (`13` → 13%, like Excel's automatic percent entry). Percent and date cells are edited as displayed. The `raw-input` attribute turns this off; `parseInput()` is exported.
- Status bar under the grid with the Average, Count and Sum of the selection (skipping filtered rows), and Alt+Shift+S to announce them. `hide-status-bar` attribute, `status-bar` CSS part, `--ls-status-bg` / `--ls-status-color`.
- Paste values (Ctrl+Shift+V) and a Paste Special dialog (Ctrl+Alt+V) with All / Values / Formats and Transpose. API: `paste({ mode, transpose })`, `openPasteSpecial()`. `ClipboardManager.paste()` takes `{ values, transpose }` options, and `copy()` / `cut()` an optional function giving formula cells' computed values.
- Shortcuts: Alt+= (AutoSum), Ctrl+; / Ctrl+Shift+; (today's date / the current time), Ctrl+' / Ctrl+Shift+' (formula / value from the cell above), Ctrl+Shift+1 / 3 / 4 / 5 / ` (number, date, currency, percent, general format).
- The fill handle continues single dates, day and month names, and text ending in a number (`Item 1` → `Item 2`). Double-clicking it fills down as far as the adjacent column has data; Ctrl+Alt+D does the same when the selection has no empty cells. `seriesFor()` is exported.
- Formula functions: `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISLOGICAL`, `ISERROR`, `ISERR`, `ISNA`, `STDEV`, `STDEV.S`, `STDEV.P`, `VAR`, `VAR.S`, `VAR.P`, `SUMPRODUCT`, `EDATE`, `EOMONTH`, `PMT`, `FV`, `PV`, `NPV`. Function names may contain dots.

## [1.3.0] - 2026-11-08

### Added

- Filters: Ctrl+Shift+L adds a filter to a table (its first row is the header row) and Alt+Down on a header cell, or its button, opens an accessible filter menu: a dialog with sort, clear, a search box and a value checklist. API: `setFilter()`, `setColumnFilter()`, `reapplyFilter()` (also Ctrl+Alt+L), `removeFilter()`, `getFilter()`, `isRowHidden()`, `toggleFilter()`, `openFilterMenu()`. Values are compared as displayed. Rows re-filter only when the filter changes or is reapplied, not on every edit. Hidden rows take no space and keep their `aria-rowindex`. Navigation, copy, cut, paste, clear, fill, Ctrl+Enter, formatting shortcuts and sorting skip them, and deleting rows that include hidden ones is refused. Works in read-only mode and survives `setData()`. Undoable; fires `filter-change` and, for user changes, the cancelable `before-filter`. The filter moves with inserted and deleted rows/columns. New CSS custom properties `--ls-filter-button-color`, `--ls-filter-button-bg`, `--ls-filter-active-color`, `--ls-filtered-row-header-color` and `--ls-filter-menu-bg`.
- `ClipboardManager.copy()` / `cut()` take an optional list of rows to copy, and `paste()` / `parseTSV()` an optional row mapper, so copies and pastes can skip rows
- Sorting: `sort(range, keys, { hasHeader })` sorts the rows of a range by one or more columns, and `sortSelection(direction)` / Alt+Shift+Up / Alt+Shift+Down sort the selection or the table around the active cell by its column, keeping a detected header row in place. Numbers, text, `FALSE`, `TRUE`, errors, then blanks (always last); stable. Formulas in moved rows keep pointing at their own row. Undoable; fires `data-change` (`operation: 'sort'`), `sort-change` and, for user sorts, the cancelable `before-sort`. User sorts refuse to move read-only cells.
- `compareSortValues()` and `sortValueOf()` utility exports
- Frozen panes: `frozen-rows` / `frozen-cols` attributes, `freezePanes()` / `unfreezePanes()`, and Alt+Shift+F to freeze above and left of the active cell (or unfreeze), announced. Frozen cells are sticky within the grid, so ARIA row/column indexes are unchanged. The frozen area follows inserted and deleted rows/columns like Excel; not undoable (a view setting); fires `freeze-change`. Ctrl+Home goes to the first unfrozen cell. New CSS custom properties `--ls-cell-bg` and `--ls-frozen-line-color`.
- Resizable columns: drag a column header's right border (Escape cancels), double-click it to fit the content, or press Alt+Shift+Left / Right (8px steps, announced) and Alt+Shift+0 (fit). Selected whole columns resize together. Works in read-only mode. API: `getColumnWidth()`, `getColumnWidths()`, `setColumnWidth()`, `setColumnWidths()`, `autoFitColumns()`. Undoable; fires `column-resize` and, for user resizes, the cancelable `before-column-resize`. Widths move with their columns on insert/delete.

### Fixed

- The grid sets `scroll-padding` for its sticky header row and row header (and frozen panes), so the browser's own focus scrolling and scroll snapping no longer leave cells under the headers.
- Number formatting no longer re-evaluates formulas from scratch after a recalculation; it reuses the engine's cached results (new `FormulaEngine.resultFor()`). Formatting long formula chains was quadratic: sorting or undoing 2,000 rows of chained currency formulas took about 4 s and now takes about 20 ms.
- Recalculating after very large changes (hundreds of thousands of cells) no longer slows down quadratically.
- Undo no longer reverts changes in the wrong order, scrambling data, when an event listener makes a change while another is applied (for example re-sorting on every `data-change`). Each change is now recorded for undo before its events fire; a change made during undo clears the redo stack.
- Pasting an HTML table with merged cells (`colspan`/`rowspan`), as Excel and Google Sheets put on the clipboard, no longer shifts the following values into the wrong columns. The value lands in the merge's top-left cell and the cells it covered are pasted as empty.
- `--ls-cell-width` and `--ls-cell-height` now accept any CSS length (`7rem`, `2em`, `calc(...)`). Non-px values were read as px (`7rem` as 7px), which broke scrolling, virtual rendering and editor placement. Changing either variable at runtime now re-renders the grid.
- The grid now sets `aria-multiselectable="true"`, since a rectangular selection selects many cells at once

## [1.2.0] - 2026-11-07

### Added

- Insert and delete rows/columns: `insertRows()`, `deleteRows()`, `insertColumns()`, `deleteColumns()` and Ctrl+Shift+= / Ctrl+- on whole-row/column selections. Formula references are rewritten to follow their cells; references to deleted cells become `#REF!`. Undoable; fires `structure-change` and the cancelable `before-structure-change`.
- Find bar (Ctrl/Cmd+F) with match count and next/previous navigation, plus `findCells()`, `openFind()` and `closeFind()`
- Fill handle: drag the selection's corner to fill; numeric runs continue their linear trend. Keyboard equivalent: Ctrl+Alt+D / Ctrl+Alt+R
- Per-cell `readOnly` flag on `CellData`
- `validateCell` property and `validation-error` event
- Cancelable `cell-edit-start` and `before-change` events
- `#REF!` and other error codes can be written as literals in formulas
- `adjustFormulaForStructureChange()` and `shiftFormulaReferences()` utility exports

### Fixed

- `IFERROR` no longer evaluates its fallback when the value is not an error, so an error in an unused fallback can't replace a valid result
- Copying, pasting or filling a range formula so that one end moves off the grid now yields a parseable `#REF!` instead of `#ERROR!`
- The formula bar no longer keeps a rejected draft, which could be committed into the next cell

## [1.0.0] - 2026-02-26

### Added

- Accessible spreadsheet grid implementing the WAI-ARIA grid pattern
- Virtual rendering for large datasets (only visible cells + buffer are rendered)
- Formula engine with recursive descent parser supporting 30+ functions
  - Arithmetic: SUM, AVERAGE, MIN, MAX, COUNT, ABS, ROUND, MOD, POWER, CEILING, FLOOR
  - String: UPPER, LOWER, LEN, TRIM, CONCAT, LEFT, RIGHT, MID, SUBSTITUTE, FIND, TEXT, VALUE
  - Logical: IF, AND, OR, NOT, IFERROR
  - Lookup: VLOOKUP, INDEX, MATCH, SUMIF, COUNTIF
- Cell references (relative, absolute, mixed) with F4 cycling
- Range references (e.g. A1:B10) in formulas
- Dependency tracking with BFS-based targeted recalculation
- Custom function registration via `registerFunction()` API
- Keyboard navigation (arrow keys, Tab, Shift+Tab, Ctrl+A, Escape)
- Arrow-key reference insertion while editing formulas
- Cell editing via Enter, double-click, or direct character input
- Formula bar component (`<y11n-formula-bar>`) with raw/formatted mode toggle
- Format toolbar component (`<y11n-format-toolbar>`) for cell styling
- Cell-level formatting: bold, italic, underline, strikethrough, text color, background color, font size, text alignment
- Undo/redo with up to 100 history entries
- Clipboard support (copy, cut, paste) with TSV and HTML formats
- Format-aware clipboard (preserves cell formatting on copy/paste)
- Range selection via Shift+Click, Shift+Arrow, mouse drag, Ctrl+A
- Read-only mode
- Custom events: `cell-change`, `selection-change`, `data-change`, `format-change`
- Public API: `getData()`, `setData()`, `setCellFormat()`, `getCellFormat()`, `setRangeFormat()`, `clearCellFormat()`
- Configurable rows and columns via properties
- CSS custom properties for theming (prefixed with `--ls-`)
- Shadow DOM encapsulation with `part="editor"` for external styling

[1.0.0]: https://github.com/egeozcan/accessible-spreadsheet-component/releases/tag/v1.0.0
