# Six Codebase Improvements Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Implement 6 improvements: async batch paste, JSDoc documentation, semantic a11y headers, integration tests, extracted bounds-checking, and preset number formats.

**Architecture:** Each improvement is independent and touches different concerns. Tasks 5 (bounds-checking) and 8 (number formats) modify shared types, so they go first. Task 1 (async paste) modifies the main component. Task 3 (headers) is a render-only change. Task 2 (JSDoc) goes last since it touches all files. Task 4 (integration tests) runs last to verify everything together.

**Tech Stack:** Lit 3.0, TypeScript strict, Vitest (unit), Playwright (E2E)

---

### Task 1: Extract bounds-checking from SelectionManager

**Files:**
- Modify: `src/controllers/selection-manager.ts:154-159`
- Modify: `src/y11n-spreadsheet.ts:681-682,696-697`
- Test: `src/controllers/__tests__/selection-manager.test.ts`

**Step 1: Write failing test for public clamp**

Add to `src/controllers/__tests__/selection-manager.test.ts`:

```typescript
describe('clamp (public)', () => {
  it('clamps negative coordinates to zero', () => {
    const manager = new SelectionManager(mockHost(), 10, 5);
    expect(manager.clamp({ row: -1, col: -3 })).toEqual({ row: 0, col: 0 });
  });

  it('clamps coordinates exceeding max bounds', () => {
    const manager = new SelectionManager(mockHost(), 10, 5);
    expect(manager.clamp({ row: 15, col: 8 })).toEqual({ row: 9, col: 4 });
  });

  it('returns coordinates within bounds unchanged', () => {
    const manager = new SelectionManager(mockHost(), 10, 5);
    expect(manager.clamp({ row: 5, col: 3 })).toEqual({ row: 5, col: 3 });
  });
});
```

**Step 2: Run test to verify it fails**

Run: `npm test -- --run src/controllers/__tests__/selection-manager.test.ts`
Expected: FAIL — `clamp` is private / not accessible

**Step 3: Make clamp public**

In `src/controllers/selection-manager.ts`, change line 154:

```typescript
// FROM:
private clamp(coord: CellCoord): CellCoord {
// TO:
clamp(coord: CellCoord): CellCoord {
```

**Step 4: Run test to verify it passes**

Run: `npm test -- --run src/controllers/__tests__/selection-manager.test.ts`
Expected: PASS

**Step 5: Replace inline bounds-checking in _handleRefArrow**

In `src/y11n-spreadsheet.ts`, replace the inline `Math.max/Math.min` in `_handleRefArrow()`:

```typescript
// FROM (lines ~681-682):
this._refCursorRow = Math.max(0, Math.min(row + dRow, this.rows - 1));
this._refCursorCol = Math.max(0, Math.min(col + dCol, this.cols - 1));

// TO:
const clamped = this._selection.clamp({ row: row + dRow, col: col + dCol });
this._refCursorRow = clamped.row;
this._refCursorCol = clamped.col;
```

And similarly for lines ~696-697:

```typescript
// FROM:
this._refCursorRow = Math.max(0, Math.min(this._refCursorRow + dRow, this.rows - 1));
this._refCursorCol = Math.max(0, Math.min(this._refCursorCol + dCol, this.cols - 1));

// TO:
const moved = this._selection.clamp({ row: this._refCursorRow + dRow, col: this._refCursorCol + dCol });
this._refCursorRow = moved.row;
this._refCursorCol = moved.col;
```

**Step 6: Run full test suite**

Run: `npm test -- --run`
Expected: All 511+ tests pass

**Step 7: Commit**

```bash
git add src/controllers/selection-manager.ts src/y11n-spreadsheet.ts src/controllers/__tests__/selection-manager.test.ts
git commit -m "refactor: make SelectionManager.clamp() public, remove inline bounds-checking"
```

---

### Task 2: Add preset number formats

**Files:**
- Modify: `src/types.ts`
- Modify: `src/engine/formula-engine.ts:394-411`
- Modify: `src/y11n-spreadsheet.ts:409-418` (display pipeline)
- Modify: `src/index.ts` (export new types)
- Test: `src/__tests__/types.test.ts`
- Test: `src/engine/__tests__/formula-engine.test.ts`

**Step 1: Add types and formatNumber to types.ts**

Add after the `CellFormat` interface in `src/types.ts`:

```typescript
/** Available preset number format types */
export type NumberFormatType = 'number' | 'currency' | 'percent' | 'scientific';

/** Options for number formatting */
export interface NumberFormatOptions {
  type: NumberFormatType;
  /** Number of decimal places (default: 2) */
  decimals?: number;
  /** Currency symbol (default: '$'), only used when type is 'currency' */
  currencySymbol?: string;
  /** Whether to use thousands separators (default: true for 'number' and 'currency') */
  thousandsSep?: boolean;
}
```

Add `numberFormat?: NumberFormatOptions;` to `CellFormat` interface.

Add the pure formatting function:

```typescript
/**
 * Format a number according to the given options.
 * Returns the formatted string representation.
 */
export function formatNumber(value: number, opts: NumberFormatOptions): string {
  const decimals = opts.decimals ?? 2;

  switch (opts.type) {
    case 'percent':
      return (value * 100).toFixed(decimals) + '%';

    case 'scientific':
      return value.toExponential(decimals);

    case 'currency': {
      const symbol = opts.currencySymbol ?? '$';
      const useSep = opts.thousandsSep !== false;
      const abs = Math.abs(value);
      const formatted = useSep
        ? addThousandsSep(abs.toFixed(decimals))
        : abs.toFixed(decimals);
      return value < 0 ? `-${symbol}${formatted}` : `${symbol}${formatted}`;
    }

    case 'number': {
      const useSep = opts.thousandsSep !== false;
      const fixed = value.toFixed(decimals);
      return useSep ? addThousandsSep(fixed) : fixed;
    }
  }
}

function addThousandsSep(numStr: string): string {
  const [intPart, decPart] = numStr.split('.');
  const withCommas = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return decPart !== undefined ? `${withCommas}.${decPart}` : withCommas;
}
```

Update `formatsEqual` to deep-compare `numberFormat`:

```typescript
// In formatsEqual, replace the simple comparison with:
export function formatsEqual(a: CellFormat | undefined, b: CellFormat | undefined): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  const keysA = Object.keys(a) as (keyof CellFormat)[];
  const keysB = Object.keys(b) as (keyof CellFormat)[];
  if (keysA.length !== keysB.length) return false;
  return keysA.every((k) => {
    if (k === 'numberFormat') {
      return numberFormatEqual(a.numberFormat, b.numberFormat);
    }
    return a[k] === b[k];
  });
}

function numberFormatEqual(a: NumberFormatOptions | undefined, b: NumberFormatOptions | undefined): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return a.type === b.type
    && a.decimals === b.decimals
    && a.currencySymbol === b.currencySymbol
    && a.thousandsSep === b.thousandsSep;
}
```

Export new types from `src/index.ts`:

```typescript
export type { NumberFormatType, NumberFormatOptions } from './types.js';
export { formatNumber } from './types.js';
```

**Step 2: Write failing tests for formatNumber**

Add to `src/__tests__/types.test.ts`:

```typescript
import { formatNumber } from '../types.js';

describe('formatNumber', () => {
  it('formats number with default options', () => {
    expect(formatNumber(1234.5, { type: 'number' })).toBe('1,234.50');
  });

  it('formats number without thousands separator', () => {
    expect(formatNumber(1234.5, { type: 'number', thousandsSep: false })).toBe('1234.50');
  });

  it('formats number with custom decimals', () => {
    expect(formatNumber(3.14159, { type: 'number', decimals: 3 })).toBe('3.142');
  });

  it('formats currency with default symbol', () => {
    expect(formatNumber(1234.5, { type: 'currency' })).toBe('$1,234.50');
  });

  it('formats currency with custom symbol', () => {
    expect(formatNumber(1234.5, { type: 'currency', currencySymbol: '€' })).toBe('€1,234.50');
  });

  it('formats negative currency', () => {
    expect(formatNumber(-50, { type: 'currency' })).toBe('-$50.00');
  });

  it('formats percent', () => {
    expect(formatNumber(0.1234, { type: 'percent' })).toBe('12.34%');
  });

  it('formats percent with custom decimals', () => {
    expect(formatNumber(0.5, { type: 'percent', decimals: 0 })).toBe('50%');
  });

  it('formats scientific notation', () => {
    expect(formatNumber(1234.5, { type: 'scientific' })).toBe('1.23e+3');
  });

  it('formats zero', () => {
    expect(formatNumber(0, { type: 'number' })).toBe('0.00');
  });
});

describe('formatsEqual with numberFormat', () => {
  it('returns true when both have identical numberFormat', () => {
    const a = { bold: true, numberFormat: { type: 'currency' as const, decimals: 2 } };
    const b = { bold: true, numberFormat: { type: 'currency' as const, decimals: 2 } };
    expect(formatsEqual(a, b)).toBe(true);
  });

  it('returns false when numberFormat differs', () => {
    const a = { numberFormat: { type: 'currency' as const } };
    const b = { numberFormat: { type: 'percent' as const } };
    expect(formatsEqual(a, b)).toBe(false);
  });

  it('returns false when one has numberFormat and other does not', () => {
    const a = { numberFormat: { type: 'number' as const } };
    const b = {};
    expect(formatsEqual(a, b)).toBe(false);
  });
});
```

**Step 3: Run tests to verify they fail**

Run: `npm test -- --run src/__tests__/types.test.ts`
Expected: FAIL — `formatNumber` not exported yet

**Step 4: Implement the types and formatNumber**

Apply the changes from Step 1 to `src/types.ts` and `src/index.ts`.

**Step 5: Run tests to verify they pass**

Run: `npm test -- --run src/__tests__/types.test.ts`
Expected: PASS

**Step 6: Hook number formatting into the display pipeline**

In `src/y11n-spreadsheet.ts`, modify `_setCellRaw` (around line 409) to apply number formatting after evaluation:

```typescript
private _setCellRaw(key: string, rawValue: string): void {
  const existing = this._internalData.get(key);
  const evaluated = this._formulaEngine.evaluate(rawValue, key);

  // Apply number format to display value if the cell has one and the value is numeric
  let displayValue = evaluated.displayValue;
  const numFmt = existing?.format?.numberFormat;
  if (numFmt && evaluated.type === 'number') {
    displayValue = formatNumber(Number(evaluated.displayValue), numFmt);
  }

  this._internalData.set(key, {
    rawValue,
    displayValue,
    type: evaluated.type,
    format: existing?.format,
  });
}
```

Add `formatNumber` to the imports from `./types.js` at the top of the file.

Also update `_applyRawValueByString` — when format changes (via format application), we need display values to reflect the new format. The existing `_applyCellFormat` method sets the format, and then `_finalizeBatch` calls `_recalcAffected` which re-evaluates. But for cells whose raw value hasn't changed, we need to re-apply number format. Modify `_applyCellFormat`:

```typescript
private _applyCellFormat(cellId: string, format: CellFormat | undefined): void {
  let cell = this._internalData.get(cellId);
  if (!cell) {
    cell = { rawValue: '', displayValue: '', type: 'text' };
    this._internalData.set(cellId, cell);
  }
  if (format && Object.keys(format).length > 0) {
    cell.format = format;
  } else {
    delete cell.format;
    if (cell.rawValue === '') {
      this._internalData.delete(cellId);
      return;
    }
  }
  // Re-apply number format to display value
  if (cell.type === 'number' && cell.format?.numberFormat) {
    cell.displayValue = formatNumber(Number(cell.rawValue.startsWith('=')
      ? cell.displayValue : cell.rawValue), cell.format.numberFormat);
  } else if (cell.type === 'number' && !cell.format?.numberFormat) {
    // Remove number formatting — re-evaluate to get plain display
    const evaluated = this._formulaEngine.evaluate(cell.rawValue, cellId);
    cell.displayValue = evaluated.displayValue;
  }
}
```

**Step 7: Write unit test for number-formatted display**

Add to `src/engine/__tests__/formula-engine.test.ts` or create a focused test:

Actually, the display formatting happens in the spreadsheet component, not the engine. We'll test it via E2E in Task 5 (integration tests). For now, validate the pure function works (Step 5 above).

**Step 8: Run full test suite**

Run: `npm test -- --run`
Expected: All tests pass

**Step 9: Commit**

```bash
git add src/types.ts src/index.ts src/__tests__/types.test.ts src/y11n-spreadsheet.ts
git commit -m "feat: add preset number formats (number, currency, percent, scientific)"
```

---

### Task 3: Async batch paste

**Files:**
- Modify: `src/y11n-spreadsheet.ts:1072-1113`
- Test: `src/__tests__/cell-format.test.ts` (or new unit test)

**Step 1: Add _pasteInProgress state and PASTE_CHUNK_SIZE constant**

At the top of `Y11nSpreadsheet` class (near the other state declarations):

```typescript
/** Whether an async paste operation is currently in progress */
private _pasteInProgress = false;

/** Number of cells to process per animation frame during large pastes */
private static readonly PASTE_CHUNK_SIZE = 500;
```

**Step 2: Rewrite _handlePaste to use chunked async application**

Replace the existing `_handlePaste` method:

```typescript
private async _handlePaste(): Promise<void> {
  if (this.readOnly || this._pasteInProgress) return;

  const { row, col } = this._selection.activeCell;
  const updates = await this._clipboardManager.paste(row, col, this.rows, this.cols);

  if (!updates || updates.length === 0) return;

  const selection = this._snapshotSelection();

  // Build the full command batch (for undo) before applying anything
  const valueUpdates = updates.map((u) => ({ id: u.id, value: u.value }));
  const batch = this._buildCommandBatch(valueUpdates, 'paste', selection, selection);
  if (!batch) return;

  // Inject format deltas
  const updateMap = new Map(updates.map((u) => [u.id, u]));
  for (const delta of batch.deltas) {
    const pasteUpdate = updateMap.get(delta.id);
    if (pasteUpdate?.format) {
      const existing = this._internalData.get(delta.id);
      delta.formatBefore = existing?.format ? { ...existing.format } : undefined;
      delta.formatAfter = pasteUpdate.format;
    }
  }
  const deltaIds = new Set(batch.deltas.map((d) => d.id));
  for (const u of updates) {
    if (u.format && !deltaIds.has(u.id)) {
      const existing = this._internalData.get(u.id);
      batch.deltas.push({
        id: u.id,
        before: existing?.rawValue ?? '',
        after: u.value,
        formatBefore: existing?.format ? { ...existing.format } : undefined,
        formatAfter: u.format,
      });
    }
  }

  // Small paste: apply synchronously (no overhead)
  if (batch.deltas.length <= Y11nSpreadsheet.PASTE_CHUNK_SIZE) {
    this._executeUserBatch(batch);
    return;
  }

  // Large paste: apply in chunks with rAF yields
  this._pasteInProgress = true;
  try {
    await this._applyBatchProgressive(batch);
    this._pushHistory(batch);
  } finally {
    this._pasteInProgress = false;
  }
}

/**
 * Apply a command batch progressively in chunks, yielding to the event loop
 * between chunks via requestAnimationFrame to keep the UI responsive.
 */
private _applyBatchProgressive(batch: CommandBatch): Promise<void> {
  return new Promise((resolve) => {
    const deltas = batch.deltas;
    let offset = 0;

    const applyChunk = () => {
      const end = Math.min(offset + Y11nSpreadsheet.PASTE_CHUNK_SIZE, deltas.length);

      for (let i = offset; i < end; i++) {
        const delta = deltas[i];
        this._applyRawValueByString(delta.id, delta.after);
        if ('formatAfter' in delta) {
          this._applyCellFormat(delta.id, delta.formatAfter);
        }
      }

      offset = end;

      if (offset < deltas.length) {
        this.requestUpdate();
        requestAnimationFrame(applyChunk);
      } else {
        // Final: restore selection, recalc, dispatch events
        this._restoreSelection(batch.selectionAfter);
        this._finalizeBatch(batch, 'user', 'after');
        resolve();
      }
    };

    applyChunk();
  });
}
```

**Step 3: Run full test suite**

Run: `npm test -- --run`
Expected: All tests pass (small pastes still go through sync path)

**Step 4: Commit**

```bash
git add src/y11n-spreadsheet.ts
git commit -m "perf: async chunked paste for large operations (>500 cells)"
```

---

### Task 4: Semantic a11y headers

**Files:**
- Modify: `src/y11n-spreadsheet.ts:1522-1543,1593-1595`
- Test: E2E test `e2e/rendering-a11y.spec.ts` (add assertions)

**Step 1: Add aria-labels to column headers**

In `_renderHeaderRow` (around line 1522-1543), update the corner header and column headers:

```typescript
return html`
  <div class="ls-header-row" role="row" aria-rowindex="1">
    <div class="ls-corner-header" role="columnheader" aria-label="Select all"></div>
    ${startCol > 0
      ? html`<div style="grid-column: span ${startCol};"></div>`
      : nothing}
    ${headers.map(
      (h) => html`
        <div
          class="ls-col-header"
          role="columnheader"
          aria-colindex="${h.index + 2}"
          aria-label="Column ${h.letter}"
        >
          ${h.letter}
        </div>
      `
    )}
    ${endCol < this.cols
      ? html`<div style="grid-column: span ${this.cols - endCol};"></div>`
      : nothing}
  </div>
`;
```

**Step 2: Add aria-label to row headers**

In `_renderRow` (around line 1594-1595):

```typescript
// FROM:
<div class="ls-row-header" role="rowheader">${row + 1}</div>

// TO:
<div class="ls-row-header" role="rowheader" aria-label="Row ${row + 1}">${row + 1}</div>
```

**Step 3: Add E2E assertions**

Add to `e2e/rendering-a11y.spec.ts`:

```typescript
test('column headers have aria-label', async ({ spreadsheet }) => {
  const firstHeader = spreadsheet.shadow('.ls-col-header').first();
  await expect(firstHeader).toHaveAttribute('aria-label', 'Column A');
});

test('row headers have aria-label', async ({ spreadsheet }) => {
  const firstRow = spreadsheet.shadow('.ls-row-header').first();
  await expect(firstRow).toHaveAttribute('aria-label', 'Row 1');
});

test('corner header has aria-label', async ({ spreadsheet }) => {
  const corner = spreadsheet.shadow('.ls-corner-header');
  await expect(corner).toHaveAttribute('aria-label', 'Select all');
});
```

**Step 4: Run E2E tests**

Run: `npm run test:e2e -- rendering-a11y.spec.ts`
Expected: PASS

**Step 5: Commit**

```bash
git add src/y11n-spreadsheet.ts e2e/rendering-a11y.spec.ts
git commit -m "a11y: add descriptive aria-labels to column and row headers"
```

---

### Task 5: Integration tests

**Files:**
- Create: `e2e/integration.spec.ts`

**Step 1: Create integration test file**

```typescript
import { test, expect } from './fixtures.js';

test.describe('Integration: cross-subsystem workflows', () => {

  test('copy formula cell and paste with reference adjustment', async ({ spreadsheet, page }) => {
    // Set up: A1=10, B1=20, C1=A1+B1
    await spreadsheet.setData({
      '0:0': { rawValue: '10', displayValue: '10', type: 'number' },
      '0:1': { rawValue: '20', displayValue: '20', type: 'number' },
      '0:2': { rawValue: '=A1+B1', displayValue: '30', type: 'number' },
    });

    // Select C1 and copy
    await spreadsheet.clickCell(0, 2);
    await page.keyboard.press('Control+c');

    // Move to C3 and paste
    await spreadsheet.clickCell(2, 2);
    await page.keyboard.press('Control+v');

    // Wait for paste to complete
    await spreadsheet.waitForCellText(2, 2, '0');

    // Verify the pasted formula references were adjusted
    const data = await spreadsheet.getData();
    expect(data['2:2']?.rawValue).toBe('=A3+B3');
  });

  test('edit cell, apply bold, undo twice, redo twice', async ({ spreadsheet, page }) => {
    // Type a value into A1
    await spreadsheet.dblClickCell(0, 0);
    await spreadsheet.typeInEditor('Hello');
    await spreadsheet.commitWithEnter();
    await spreadsheet.waitForCellText(0, 0, 'Hello');

    // Move back to A1 and apply bold
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.press('Control+b');

    // Verify bold is applied
    let format = await spreadsheet.getCellFormat('0:0');
    expect(format?.bold).toBe(true);

    // Undo (should remove bold)
    await page.keyboard.press('Control+z');
    format = await spreadsheet.getCellFormat('0:0');
    expect(format?.bold).toBeFalsy();

    // Undo again (should remove value)
    await page.keyboard.press('Control+z');
    await spreadsheet.waitForCellText(0, 0, '');

    // Redo (should restore value)
    await page.keyboard.press('Control+Shift+z');
    await spreadsheet.waitForCellText(0, 0, 'Hello');

    // Redo again (should restore bold)
    await page.keyboard.press('Control+Shift+z');
    format = await spreadsheet.getCellFormat('0:0');
    expect(format?.bold).toBe(true);
  });

  test('cut cells, paste elsewhere, undo restores both', async ({ spreadsheet, page }) => {
    // Set up: A1=100, A2=200
    await spreadsheet.setData({
      '0:0': { rawValue: '100', displayValue: '100', type: 'number' },
      '1:0': { rawValue: '200', displayValue: '200', type: 'number' },
    });

    // Select A1:A2
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.down('Shift');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.up('Shift');

    // Cut
    await page.keyboard.press('Control+x');

    // Move to C1 and paste
    await spreadsheet.clickCell(0, 2);
    await page.keyboard.press('Control+v');

    // Verify source cleared and target populated
    await spreadsheet.waitForCellText(0, 2, '100');
    await spreadsheet.waitForCellText(1, 2, '200');
    const textA1 = await spreadsheet.getCellText(0, 0);
    expect(textA1).toBe('');

    // Undo paste
    await page.keyboard.press('Control+z');
    // Undo cut
    await page.keyboard.press('Control+z');

    // Verify source restored
    await spreadsheet.waitForCellText(0, 0, '100');
    await spreadsheet.waitForCellText(1, 0, '200');
  });

  test('format cells, clear them, undo restores values with format', async ({ spreadsheet, page }) => {
    // Set up: A1=42
    await spreadsheet.setData({
      '0:0': { rawValue: '42', displayValue: '42', type: 'number' },
    });

    // Apply bold to A1
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.press('Control+b');

    let format = await spreadsheet.getCellFormat('0:0');
    expect(format?.bold).toBe(true);

    // Clear A1
    await page.keyboard.press('Delete');
    await spreadsheet.waitForCellText(0, 0, '');

    // Undo clear (should restore value with bold)
    await page.keyboard.press('Control+z');
    await spreadsheet.waitForCellText(0, 0, '42');
    format = await spreadsheet.getCellFormat('0:0');
    expect(format?.bold).toBe(true);
  });
});
```

**Step 2: Run integration tests**

Run: `npm run test:e2e -- integration.spec.ts`
Expected: All PASS

**Step 3: Commit**

```bash
git add e2e/integration.spec.ts
git commit -m "test: add cross-subsystem integration E2E tests"
```

---

### Task 6: JSDoc documentation

**Files:**
- Modify: `src/types.ts`
- Modify: `src/y11n-spreadsheet.ts`
- Modify: `src/engine/formula-engine.ts`
- Modify: `src/controllers/selection-manager.ts`
- Modify: `src/controllers/clipboard-manager.ts`
- Modify: `src/components/y11n-formula-bar.ts`
- Modify: `src/components/y11n-format-toolbar.ts`

This task is documentation-only. No tests needed — just run typecheck to confirm nothing broke.

**Step 1: Document types.ts**

Add/update JSDoc for all exported interfaces, types, and functions. Key additions:

- `CellCoord`: add `@property` descriptions for `row` and `col`
- `CellData`: document each field
- `CellFormat`: document each property including new `numberFormat`
- `GridData`: document the key format
- `SelectionRange`: document normalized bounds semantics
- `FormulaContext`: document the callback signatures
- `FormulaFunction`: document the context + variadic args pattern
- All event detail interfaces: document fields
- All utility functions: `@param`, `@returns`, `@throws` where applicable
- `formatNumber`: already documented from Task 2

**Step 2: Document y11n-spreadsheet.ts**

- Class-level JSDoc: describe the component, its ARIA pattern, and key features
- Public properties: `@property` JSDoc with types and defaults
- All public API methods: `@param`, `@returns`
- Key private methods that implement complex logic: 1-line summaries where missing

**Step 3: Document formula-engine.ts**

- Class-level: describe the recursive descent parser, supported syntax
- Add grammar comment block near the top:

```typescript
/**
 * Grammar:
 *   expression   = comparison
 *   comparison   = concat (("=" | "<>" | "<" | ">" | "<=" | ">=") concat)*
 *   concat       = additive ("&" additive)*
 *   additive     = multiplicative (("+" | "-") multiplicative)*
 *   multiplicative = unary (("*" | "/") unary)*
 *   unary        = ("-" unary) | primary
 *   primary      = NUMBER | STRING | BOOLEAN | REF | RANGE | FUNC "(" args ")" | "(" expression ")"
 */
```

- `evaluate()`: document `@param forCellKey` and dependency tracking side-effect
- `recalculate()` / `recalculateAffected()`: document return values
- `registerFunction()`: document name normalization

**Step 4: Document selection-manager.ts**

- Class-level: describe anchor/head model
- `clamp()`: document the public API (newly public)
- `move()`, `moveTo()`: document extend parameter
- `startSelection()`, `extendSelection()`, `endSelection()`: document mouse interaction lifecycle

**Step 5: Document clipboard-manager.ts**

- `copy()`: document formats written (HTML + TSV)
- `cut()`: document return value (keys to clear)
- `paste()`: document the priority chain (internal → HTML → TSV)
- `adjustFormulaReferences()`: document absolute/mixed ref handling
- `parseTSV()`: document RFC 4180 compliance

**Step 6: Document formula-bar and format-toolbar**

- Light-touch: ensure class-level JSDoc is clear on wiring pattern
- Document the custom events each component dispatches

**Step 7: Run typecheck**

Run: `npm run typecheck`
Expected: No errors

**Step 8: Commit**

```bash
git add src/
git commit -m "docs: add comprehensive JSDoc to all public APIs and key internals"
```

---

### Task 7: Final verification

**Step 1: Run full unit test suite**

Run: `npm test -- --run`
Expected: All tests pass

**Step 2: Run full E2E test suite**

Run: `npm run test:e2e`
Expected: All tests pass (including new integration tests)

**Step 3: Run typecheck**

Run: `npm run typecheck`
Expected: No errors

**Step 4: Run build**

Run: `npm run build`
Expected: Clean build, no warnings
