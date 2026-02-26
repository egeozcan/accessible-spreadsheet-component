# Six Codebase Improvements - Design Document

Date: 2026-02-26

## 1. Async Batch Paste

**Problem**: `_handlePaste()` applies all updates synchronously. Large pastes (10K+ cells) freeze the UI.

**Solution**: Chunk updates into batches of ~500 cells, yielding to the event loop between chunks via `requestAnimationFrame`.

- The full paste remains a single undo entry (one `CommandBatch`)
- Add `_applyBatchProgressive()` that splits deltas into chunks
- First chunk applies immediately; subsequent chunks yield via `requestAnimationFrame`
- `_pasteInProgress` flag prevents concurrent pastes
- Threshold: only use async path for >500 cells; small pastes stay synchronous

**Files**: `src/y11n-spreadsheet.ts`

## 2. JSDoc Documentation

Add `@param` / `@returns` / `@throws` JSDoc to:

- All public API methods on `Y11nSpreadsheet` (`getData`, `setData`, `registerFunction`, `getCellFormat`, `setCellFormat`, `setRangeFormat`, `clearCellFormat`, `clearRangeFormat`, `toggleFormat`)
- All public properties on `Y11nSpreadsheet` (`rows`, `cols`, `data`, `readOnly`, `functions`)
- All exported functions in `types.ts` (`cellKey`, `parseKey`, `colToLetter`, `letterToCol`, `refToCoord`, `coordToRef`, `formatsEqual`)
- All exported interfaces/types in `types.ts`
- `FormulaEngine` class: public methods + parser grammar comment
- `SelectionManager` class: public methods
- `ClipboardManager` class: public methods
- Formula bar and format toolbar components

Private methods with clear names get lighter treatment (1-line summary only where missing).

**Files**: All source files in `src/`

## 3. Semantic Headers (Accessibility)

**Current**: Column/row headers are `<div role="columnheader">` and `<div role="rowheader">`. Valid ARIA but missing descriptive labels.

**Solution**: Keep `<div>` elements (changing to `<th>` breaks CSS Grid layout in shadow DOM). Add:

- `aria-label="Column A"` (etc.) to column headers
- `aria-label="Row 1"` (etc.) to row headers
- Corner header gets `aria-label="Select all"`

**Files**: `src/y11n-spreadsheet.ts` (render methods)

## 4. Integration Tests

New file `e2e/integration.spec.ts` covering cross-subsystem workflows:

1. **Copy formula + paste with ref adjustment**: Copy cell with `=A1+B1`, paste 2 rows down, verify formula becomes `=A3+B3`
2. **Edit + format + undo + redo**: Edit cell value, apply bold, undo twice (restores format then value), redo twice (restores both)
3. **Cut + paste + undo**: Cut cells, paste elsewhere, verify source cleared and target populated, undo restores both
4. **Paste external HTML with format + undo**: Write styled HTML to clipboard, paste, verify format applied, undo removes it
5. **Format + clear cells + undo**: Format cells, clear them, undo restores values with format intact

**Files**: `e2e/integration.spec.ts`

## 5. Extract Bounds-Checking Logic

**Current**: `SelectionManager.clamp()` is private. `_handleRefArrow()` in the main component duplicates bounds checking with inline `Math.max(0, Math.min(...))`.

**Solution**: Make `clamp()` public on `SelectionManager`. Replace inline bounds checking in `_handleRefArrow()` with `this._selection.clamp(...)`.

**Files**: `src/controllers/selection-manager.ts`, `src/y11n-spreadsheet.ts`

## 6. Preset Number Formats

Add number formatting to `CellFormat`:

```typescript
type NumberFormatType = 'number' | 'currency' | 'percent' | 'scientific';

interface NumberFormatOptions {
  type: NumberFormatType;
  decimals?: number;       // default: 2
  currencySymbol?: string; // default: '$'
  thousandsSep?: boolean;  // default: true for currency/number
}
```

- Add `numberFormat?: NumberFormatOptions` to `CellFormat`
- Add `formatNumber(value: number, opts: NumberFormatOptions): string` pure function in `types.ts`
- Format applied at display time in `FormulaEngine.evaluate()` — stored in `displayValue`
- Raw numeric value preserved in `rawValue`
- Clipboard copies display value; existing format round-trip via `data-format` handles `numberFormat`
- Undo/redo works via existing `formatBefore`/`formatAfter` delta tracking
- `formatsEqual()` updated to deep-compare `numberFormat`

**Files**: `src/types.ts`, `src/engine/formula-engine.ts`, `src/y11n-spreadsheet.ts`
