# CLAUDE.md

## Project overview

`y11n-spreadsheet` is an accessible spreadsheet web component built with Lit 3.0 and TypeScript. It implements the WAI-ARIA grid pattern with virtual rendering, a formula engine, undo/redo, clipboard support, and keyboard navigation.

## Commands

```bash
npm run dev              # Start Vite dev server (port 5173)
npm run build            # Typecheck + build library to dist/
npm run typecheck        # TypeScript check only (tsc --noEmit)
npm test                 # Run unit tests (vitest)
npm run test:watch       # Unit tests in watch mode
npm run test:e2e         # Playwright E2E tests (headless)
npm run test:e2e:headed  # E2E tests in visible browser
npm run test:e2e:ui      # E2E tests with Playwright UI
npm run storybook        # Dev Storybook on port 6006
npm run build-storybook  # Build static Storybook
```

## Architecture

```
src/
  index.ts                          # Barrel exports
  types.ts                          # Shared types & coordinate utilities
  y11n-spreadsheet.ts               # Main component (<y11n-spreadsheet>)
  components/y11n-formula-bar.ts    # Formula bar sub-component
  components/y11n-format-toolbar.ts # Stateless formatting toolbar (fires format-action / conditional-format-action)
  components/y11n-conditional-format-dialog.ts # Rule manager dialog; holds no rules, fires cf-save/cf-remove/cf-move/cf-close
  components/conditional-icons.ts   # Icon-set SVGs + the styles any component rendering them must include
  controllers/selection-manager.ts  # Selection state (Lit ReactiveController)
  controllers/clipboard-manager.ts  # Copy/cut/paste via Clipboard API + TSV
  engine/formula-engine.ts          # Recursive descent formula parser/evaluator
  utils/delimited.ts                # CSV/TSV parse + serialize (shared by clipboard and toCSV/fromCSV)
  utils/references.ts               # Rewrites formula refs when rows/columns are inserted/deleted
  utils/sort.ts                     # Sort value ordering (numbers, text, FALSE, TRUE, errors, blanks last)
  utils/input-parse.ts              # Recognises typed 15% / $1,200 / 2026-01-05 / 14:30 as numbers + formats
  utils/series.ts                   # Fill-handle series: numbers, dates, day/month names, "Item 1"
  utils/conditional-format.ts       # Conditional formats: rule validation, ConditionalFormatEvaluator, structure rewriting
```

- **Data model**: `GridData = Map<"row:col", CellData>` (sparse)
- **Cell keys**: `"row:col"` format (0-indexed), e.g. `"0:0"` = A1
- **Coordinate utils**: `cellKey()`, `parseKey()`, `colToLetter()`, `letterToCol()`, `refToCoord()`, `coordToRef()` in `types.ts`
- **Formula engine**: Tokenizer + recursive descent parser. Each `parseExpression()` call creates its own `ParserState`, so nested evaluation is safe without save/restore.
- **Dependency tracking**: Forward/reverse dep graph in the formula engine; `recalculateAffected()` uses BFS for targeted recalc.
- **Virtual rendering**: Only visible cells + 5-cell buffer are rendered. Scroll position drives which rows/cols appear.
- **Column positions**: columns can have custom widths (`_colWidths`, sparse; others use the measured `--ls-cell-width`). Use `_colLeft()`, `_colWidth()` and `_colAt()` (cached prefix sums) for any horizontal math, never `col * width`. Rows still share one height (`_cellSize.height`).
- **Frozen panes**: frozen cells/headers are `position: sticky` in the same grid (DOM order and ARIA indexes unchanged), always rendered, with top/left set inline (`_stickyTop`/`_stickyLeft`). Read the effective counts via `_frozenRowCount`/`_frozenColCount` (clamped), never the raw properties. Scroll math must treat the frozen area (`_frozenHeight`/`_frozenWidth`) as covering the start of the viewport. The grid's `scroll-padding` (`_scrollPadding()`) covers the header and frozen area, so a row's scroll-snap position is exactly where `_ensureCellVisible` scrolls it to.
- **Filters / hidden rows**: `_filter` (`FilterState`: range, criteria, `hidden` rows) is immutable and replaced on every change, so batches hold before/after snapshots (`CommandBatch.filter`). Hidden rows take no height: use `_rowTop()` / `_visibleIndex()` for any vertical math, never `row * height`, and `_isRowHidden()` / `_visibleRowsBetween()` for anything a user does over a range (writes, copy, sort, navigation all skip hidden rows; `SelectionManager.stepRows()` for keyboard moves). The active cell must never be on a hidden row. Edits don't re-filter; only setting/changing/reapplying the filter does.
- **Conditional formats**: `_cfRules` (priority order, first wins) is immutable and replaced on every change like `_filter`, so batches hold before/after arrays (`CommandBatch.conditional`; structure batches carry one when rules move). Change rules only through `_commitConditional()`. Rules never write into `CellData`: `_renderCell` asks `_cf.resultFor(row, col)` and merges the style over the cell's format (`_getCellStyles`), so copy/sort/export see the real formats. `_cf` caches results per cell (dropped on any change) and range statistics per rule (dropped only when a cell in that rule's range changes): every data write must reach `_recalcAffected()` or `_recalcAll()`, which invalidate it. Every rule needs a label (`describeCondition()` default) rendered in the cell's `.ls-cf-label` and in `_getAnnouncement`; bars and icons are `aria-hidden`, and each rule type needs a forced-colors cue (see the `forced-colors` block in the styles).
- **Undo/redo**: `CommandBatch[]` stacks tracking cell deltas + selection state. Max 100 entries. Row/column insert/delete batches carry a `structure` field with whole-cell before/after snapshots and grid sizes instead of deltas; sort batches carry a `sort` field with whole-cell snapshots and the sort spec; column-width batches carry a `widths` field; filter batches carry only a `filter` field (structure batches carry one too when they move the filter); conditional-format batches carry only a `conditional` field. Apply batches with `_commitBatch()`, which records history *before* applying, so changes made by event listeners during the apply are recorded after it.
- **User writes**: every user batch (edit, paste, cut, clear, fill) goes through `_vetUserBatch()`, which drops `CellData.readOnly` cells and values rejected by `validateCell`, then fires the cancelable `before-change`. Route new user write paths through it. Values the user *types* (inline editor, formula bar, Ctrl+Enter) go through `_interpretInput()` first, which turns `15%`-style input into a number plus a number format; pastes, fills and API calls don't. The editor and formula bar start from `_editTextFor(cell)`, not `rawValue`, so percent/date cells edit as displayed.
- **Selection**: anchor/head range plus an optional active-cell override (`SelectionManager._active`) so whole-row/column selections keep the active cell in place. Read the active cell via `activeCell`, never `head`.

## Code style

- TypeScript strict mode (`noUnusedLocals`, `noUnusedParameters`, `noFallthroughCasesInSwitch`)
- ES2021 target, ESNext modules
- Lit decorators (`@property`, `@state`, `@customElement`)
- `useDefineForClassFields: false` (required for Lit)
- Private members use `_` prefix convention
- CSS custom properties prefixed with `--ls-`
- CSS parts: `part="editor"` on the inline editor

## Testing

- **Unit tests**: Vitest, files in `src/**/__tests__/*.test.ts`
- **E2E tests**: Playwright, files in `e2e/*.spec.ts` plus `e2e/fixtures.ts` for helpers
- Playwright runs against the Vite dev server on port 5173
- E2E tests use shadow DOM piercing via Lit's locator patterns
- Run `npx playwright install` if browsers aren't installed

## Build output

Library build (`dist/`) uses ES module format only. `lit` is externalized as a peer dependency. Declaration files and source maps are generated.
