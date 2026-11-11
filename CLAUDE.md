# CLAUDE.md

## Project overview

`y11n-spreadsheet` is an accessible spreadsheet web component built with Lit 3.0 and TypeScript. It implements the WAI-ARIA grid pattern with virtual rendering, a formula engine (dynamic arrays, several values per cell, multiple sheets, defined names), undo/redo, clipboard support, and keyboard navigation.

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
npm run storybook        # Dev Storybook on port 6006 (getting-started site: guides, examples, reference)
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
  components/y11n-sheet-tabs.ts     # Sheet tab list; holds no sheets, fires sheet-select/add/rename/remove/move
  components/y11n-name-manager.ts   # Defined-names dialog; holds no names, fires name-save/remove/go/close
  controllers/selection-manager.ts  # Selection state (Lit ReactiveController)
  controllers/clipboard-manager.ts  # Copy/cut/paste via Clipboard API + TSV
  engine/parser.ts                  # Formula tokenizer + recursive descent parser (text -> AST)
  engine/values.ts                  # Value types: RangeValue (arrays), RefValue (references), LambdaValue; coercions
  engine/functions.ts               # Built-in functions that only need their argument values (core, lookups, registrar); calls the domain files below
  engine/functions-{math,stats,distributions,text,datetime,financial,logical-info}.ts  # The rest of the built-ins by domain, each `register…Builtins(reg)`
  engine/formula-engine.ts          # AST evaluator, special forms (IF, LET, LAMBDA, INDEX, OFFSET, ...), spill, dep graph, sheets, names
  utils/delimited.ts                # CSV/TSV parse + serialize (shared by clipboard and toCSV/fromCSV)
  utils/references.ts               # Rewrites formula refs (copy/fill shift, row/column insert/delete), on the tokenizer
  utils/list.ts                     # Lists: `Ann, Bob` text <-> the `={{"Ann";"Bob"}}` formula that stores it
  utils/sheet-names.ts              # Sheet name rules (validateSheetName, nextSheetName)
  utils/sort.ts                     # Sort value ordering (numbers, text, FALSE, TRUE, errors, blanks last)
  utils/input-parse.ts              # Recognises typed 15% / $1,200 / 2026-01-05 / 14:30 as numbers + formats
  utils/series.ts                   # Fill-handle series: numbers, dates, day/month names, "Item 1"
  utils/conditional-format.ts       # Conditional formats: rule validation, ConditionalFormatEvaluator, structure rewriting
```

- **Data model**: `GridData = Map<"row:col", CellData>` (sparse)
- **Cell keys**: `"row:col"` format (0-indexed), e.g. `"0:0"` = A1
- **Coordinate utils**: `cellKey()`, `parseKey()`, `colToLetter()`, `letterToCol()`, `refToCoord()`, `coordToRef()` in `types.ts`
- **Formula engine**: `parser.ts` turns formula text into an AST once (cached by text); `formula-engine.ts` evaluates it. A scalar error is thrown (`new Error('#N/A')`); inside an array it is the code string. `_eval` may return a `RefValue` or `LambdaValue`; `_value` resolves references to values. Functions are registered with argument kinds (`n`/`t`/`v` single values that make the call go element by element over arrays, `a` array as is, `r` reference, `x` error-tolerant); functions that are lazy or need references live in the engine's `_special` map.
- **Evaluation passes**: every evaluation runs in a pass with a memo. `pure` (standalone `evaluate(raw)`: stores nothing, recomputes what it reaches), `root` (`evaluate(raw, key)`: caches that one cell with transitive deps), `commit` (`recalculate()` / `recalculateAffected()`: each cell records its own deps, results are cached, spills are written). Long reference chains are resolved through `_deferring` (EvalDepthError), never by deep recursion.
- **Dependency tracking**: recorded while evaluating. Cells and small ranges are cell edges; ranges over 64 cells are one range node. `_closure()` is the BFS over dependents (including readers of a formula's spilled cells). Range values are cached per pass (`pass.ranges`), so never mutate a `RangeValue`.
- **Whole columns and rows**: `A:A` is as tall as the grid but is read up to the workbook's used extent (`_usedExtent`: stored content of all sheets plus how far spills reach, `_spillReach`), leaving out the reading formula's own spill (`_current`), so a spill can't lengthen what it reads. Readers are remembered with the extent they read (`_wholeReaders`) and recomputed when it moves (`_movedWholeReaders`). Formulas that lengthen each other are caught by `_growsWithoutEnd` (evaluated once with `pass.untrimmed`) and become `#SPILL!`.
- **Outside a recalculation**: `pure` / `root` passes write nothing, so spills computed in the pass are kept in `pass.pureSpills` and read from there, and are checked for being blocked (`_spillBlocked`) like committed ones. `evaluateCommitted()` is a pure pass that reads formula cells from `_results` (volatile ones too): use it for anything that tests what the grid shows (conditional formats do).
- **Computed values as input**: when a computed value is written somewhere as a value (copy of a spilled cell, paste values, fill, HTML clipboard `data-value`), go through `computedLiteral()` in `clipboard-manager.ts`: arrays become constants (`arrayToFormula`), text that would read back as something else becomes `="..."`.
- **Spill**: a formula whose result is an array larger than 1x1 spills. The engine writes the spilled values into the sheet data as cells with `rawValue: ''` and `spilledFrom` (the anchor's key), sets `spill` on the anchor, and removes them again (`_materialize`); `#SPILL!` when a value is in the way or the range leaves the grid (`setSheetSize`, kept in sync by `_syncEngineSize()`). In the component, use `hasContent(cell)` for "does this cell show something" and `storedCell(cell)` wherever cells are moved, snapshotted or imported (sort, structure changes, `_syncData`), so spilled values are never carried around; anything that replaces or deletes cell objects must not drop a cell with `spilledFrom` unless a recalc follows. `_plainNumber()` gives the number behind any number cell, spilled ones included.
- **Several values in one cell**: a 1x1 array whose element is an array stays in the cell (`={{1;2;3}}`, `={FILTER(...)}`): `displayValue` is its text, `CellData.array` the values, `type` is `'text'`. A list is exactly such a formula of constants (`utils/list.ts`); `_editTextFor` shows it as its items and `_interpretInput(key, value, asList)` turns typed text back into one (`_editListMode`). Filters match such cells by item (`_rowMatches`, `_filterMenuValues`).
- **Sheets**: `_sheets` (`SheetState[]`) holds every sheet; for the sheet being shown only `id`/`name` are current and everything else lives in the component's fields (`_internalData`, `rows`, `_filter`, `_cfRules`, `_colWidths`, selection, scroll) until `_stashSheet()` / `_loadSheet()` swap them. The engine has the same sheets by name (`setActiveSheet`), holds each sheet's data map by reference, and reports changed cells of all sheets in `engine.changes` (`_formatOtherSheets()` applies number formats there). Batches carry `sheetId`; undo/redo call `_showSheetOf()` first. Structure changes also rewrite formulas on other sheets and defined names (`_addCrossSheetStructure`, `StructureBatch.others` / `.names`); renaming a sheet rewrites formulas, names, conditional-format formulas and the undo history. Sheet add/remove/rename/move and name changes are not undoable.
- **Virtual rendering**: Only visible cells + 5-cell buffer are rendered. Scroll position drives which rows/cols appear.
- **Column positions**: columns can have custom widths (`_colWidths`, sparse; others use the measured `--ls-cell-width`). Use `_colLeft()`, `_colWidth()` and `_colAt()` (cached prefix sums) for any horizontal math, never `col * width`. Rows still share one height (`_cellSize.height`).
- **Frozen panes**: frozen cells/headers are `position: sticky` in the same grid (DOM order and ARIA indexes unchanged), always rendered, with top/left set inline (`_stickyTop`/`_stickyLeft`). Read the effective counts via `_frozenRowCount`/`_frozenColCount` (clamped), never the raw properties. Scroll math must treat the frozen area (`_frozenHeight`/`_frozenWidth`) as covering the start of the viewport. The grid's `scroll-padding` (`_scrollPadding()`) covers the header and frozen area, so a row's scroll-snap position is exactly where `_ensureCellVisible` scrolls it to.
- **Filters / hidden rows**: `_filter` (`FilterState`: range, criteria, `hidden` rows) is immutable and replaced on every change, so batches hold before/after snapshots (`CommandBatch.filter`). Hidden rows take no height: use `_rowTop()` / `_visibleIndex()` for any vertical math, never `row * height`, and `_isRowHidden()` / `_visibleRowsBetween()` for anything a user does over a range (writes, copy, sort, navigation all skip hidden rows; `SelectionManager.stepRows()` for keyboard moves). The active cell must never be on a hidden row. Edits don't re-filter; only setting/changing/reapplying the filter does.
- **Conditional formats**: `_cfRules` (priority order, first wins) is immutable and replaced on every change like `_filter`, so batches hold before/after arrays (`CommandBatch.conditional`; structure batches carry one when rules move). Change rules only through `_commitConditional()`. Rules never write into `CellData`: `_renderCell` asks `_cf.resultFor(row, col)` and merges the style over the cell's format (`_getCellStyles`), so copy/sort/export see the real formats. `_cf` caches results per cell (dropped on any change) and range statistics per rule (dropped only when a cell in that rule's range changes): every data write must reach `_recalcAffected()` or `_recalcAll()`, which invalidate it. Every rule needs a label (`describeCondition()` default) rendered in the cell's `.ls-cf-label` and in `_getAnnouncement`; bars and icons are `aria-hidden`, and each rule type needs a forced-colors cue (see the `forced-colors` block in the styles).
- **Names UI**: the formula bar's Name Box fires `name-box-commit`; `_handleNameBoxCommit` goes to the cells (`engine.referenceOf()` -> `_goToRange()`) or names the selection. The name manager's events are applied by `_handleNameSave` / `_handleNameRemove` / `_handleNameGo`. Every user change to names goes through `_changeNamesByUser()` (read-only check, cancelable `before-names-change`); a rename rewrites the name everywhere (`_renameNameEverywhere`, which uses `_rewriteStoredFormulas` like a sheet rename does).
- **Picking references**: while a formula is edited at a position where a reference fits (`_isRefPositionInFormula`), arrows (`_handleRefArrow`) and the pointer (`_pickReferenceWithPointer`) write through `_insertReference(_refText(...))`. The edit can outlive a sheet switch: `_editSheetId` is the edited cell's sheet, `_editingElsewhere` is true while another sheet is shown, `_refText` then prefixes the sheet name, and `_commitEdit` / `_cancelEdit` call `_returnToEditSheet()` first. Anything that ends or moves an edit must do the same. `_activateSheet()` keeps a formula edit open only for user switches; `_switchSheetView()` is the raw view swap.
- **Undo/redo**: `CommandBatch[]` stacks tracking cell deltas + selection state, for all sheets (each batch has a `sheetId`). Max 100 entries. Row/column insert/delete batches carry a `structure` field with whole-cell before/after snapshots and grid sizes instead of deltas; sort batches carry a `sort` field with whole-cell snapshots and the sort spec; column-width batches carry a `widths` field; filter batches carry only a `filter` field (structure batches carry one too when they move the filter); conditional-format batches carry only a `conditional` field. Apply batches with `_commitBatch()`, which records history *before* applying, so changes made by event listeners during the apply are recorded after it.
- **User writes**: every user batch (edit, paste, cut, clear, fill) goes through `_vetUserBatch()`, which drops `CellData.readOnly` cells and values rejected by `validateCell`, then fires the cancelable `before-change`. Route new user write paths through it. Values the user *types* (inline editor, formula bar, Ctrl+Enter) go through `_interpretInput()` first, which turns `15%`-style input into a number plus a number format; pastes, fills and API calls don't. The editor and formula bar start from `_editTextFor(cell)`, not `rawValue`, so percent/date cells edit as displayed.
- **Selection**: anchor/head range plus an optional active-cell override (`SelectionManager._active`) so whole-row/column selections keep the active cell in place. Read the active cell via `activeCell`, never `head`.

- **Modals**: every dialog or popup that takes over interaction (filter menu, Paste Special, rule manager, name manager, values popup, delete confirmations) has `aria-modal="true"`, traps Tab and Shift+Tab (wrapping at its ends), closes on Escape and returns focus to what opened it. New ones must too.

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

## Storybook

Storybook is the getting-started site for users of the component, so everything in it is user-facing documentation.

- **Examples** (`stories/examples/<name>.ts`): one file per example, a default-exported function that returns the element to show. The file is displayed verbatim in the Code panel and under "Show code" (`example('<name>')` in `stories/support/example.ts` reads it with `?raw`), and its leading `/** */` comment becomes the story description. So: import the package as `'y11n-spreadsheet'` (aliased to `src/index.ts` in `.storybook/main.ts` and `tsconfig.json`), use public API only (no `_private` members, no `as any`), and keep code lines short enough to read in a side panel. `gridFromRows` (`examples/grid-from-rows.ts`) is appended to every listing that imports it.
- **Setup order in examples**: `setData()` starts a fresh undo history for the sheet, so setup that should not be undoable by the user (column widths, conditional-format rules, formats inside the data) goes before it. A filter needs the data first and stays an undo step.
- **Stories** (`stories/*.stories.ts`): one file per group under `Examples/`, each story `{ name, render, parameters: example('<name>', { height }) }`. Group Docs pages use `.storybook/ExamplesPage.mdx`. New stories must be added to the list in `e2e/storybook-a11y.spec.ts`, and new groups to `storySort` in `.storybook/preview.ts`.
- **Guides** (`stories/guides/*.mdx`) are hand-written. **Reference** pages (`stories/reference/*.mdx`) render README sections through `readmeSection('<heading>')`: change the README, not the page, and keep README section headings in step with those calls.
- `.storybook/manager.ts` is only read at startup: restart `npm run storybook` after changing it.

## Conformance harness

`src/engine/__tests__/conformance.test.ts` runs every `conformance/*.json` case `{f, e, src, grid?, skip?}` against a fresh engine on a default grid (documented in the test). `src` is `excel` (computed by Excel), `docs` (a Microsoft documentation example) or `expected` (author's knowledge, not yet checked against Excel). New functions get cases there; a `skip` records a known gap.

## Build output

Library build (`dist/`) uses ES module format only. `lit` is externalized as a peer dependency. Declaration files and source maps are generated.
