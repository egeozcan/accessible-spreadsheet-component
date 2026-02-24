# Future Work

Items identified during the 2026-02-24 codebase review that were deferred from the current round of improvements.

## Spreadsheet Features

### Column/Row Freezing
Freeze panes so headers stay visible during scroll. Would require splitting the virtual rendering into frozen and scrollable regions, with synchronized scroll positions.

### Sorting & Filtering
Sort columns ascending/descending, filter rows by criteria. Needs a data view layer between the raw `GridData` and the rendered output, plus UI for sort indicators and filter dropdowns.

### Data Validation
Cell-level validation rules (number ranges, dropdown lists, date formats, custom regex). Needs a validation rule model, visual indicators for invalid cells, and integration with the editing flow.

### Conditional Formatting
Rules-based cell styling (color scales, data bars, icon sets, custom rules). Needs a rule evaluation engine that runs after recalculation and applies CSS overrides to affected cells.

### Cell Merging
Merge multiple cells into a single visual cell spanning rows/columns. Requires changes to the grid layout, selection handling, and virtual rendering to account for merged regions.

### Cell Comments/Notes
Attach notes to cells with visual indicators (corner triangle). Needs a comment data model, popover UI, and keyboard accessibility.

## Code Structure

### Main Component Decomposition
The main `y11n-spreadsheet.ts` is 1,215 lines. Could be improved by:
- Extracting keyboard handling into a `KeyboardController`
- Extracting editing logic into an `EditingController`
- Breaking rendering into sub-components: header row, cell renderer, editor overlay

This was deferred to avoid a large diff conflicting with feature work.

## Clipboard

### Mouse Range Selection in Reference Mode
Currently, arrow keys insert cell references during formula editing. Mouse click/drag to select a range reference is not yet supported.

## Formula Engine

### Array Formulas
Support for formulas that return arrays and spill into adjacent cells (e.g., `=SORT(A1:A10)`).

### More Functions
Beyond the ~22 functions being added now, candidates for future rounds:
- Statistical: STDEV, VAR, MEDIAN, PERCENTILE, RANK
- Financial: NPV, IRR, PMT, FV, PV
- Advanced lookup: SUMPRODUCT, OFFSET, INDIRECT
- Date/time: YEAR, MONTH, DAY, HOUR, MINUTE, DATEVALUE, TIMEVALUE, EDATE, EOMONTH
- Information: ISBLANK, ISNUMBER, ISTEXT, ISERROR, TYPE

### Error Propagation in Ranges
Currently errors in ranges are included as values in the array. Consider short-circuiting or propagating errors more aggressively (matching Excel behavior).
