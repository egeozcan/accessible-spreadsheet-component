# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
