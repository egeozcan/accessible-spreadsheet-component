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
  controllers/selection-manager.ts  # Selection state (Lit ReactiveController)
  controllers/clipboard-manager.ts  # Copy/cut/paste via Clipboard API + TSV
  engine/formula-engine.ts          # Recursive descent formula parser/evaluator
```

- **Data model**: `GridData = Map<"row:col", CellData>` (sparse)
- **Cell keys**: `"row:col"` format (0-indexed), e.g. `"0:0"` = A1
- **Coordinate utils**: `cellKey()`, `parseKey()`, `colToLetter()`, `letterToCol()`, `refToCoord()`, `coordToRef()` in `types.ts`
- **Formula engine**: Tokenizer + recursive descent parser. Each `parseExpression()` call creates its own `ParserState`, so nested evaluation is safe without save/restore.
- **Dependency tracking**: Forward/reverse dep graph in the formula engine; `recalculateAffected()` uses BFS for targeted recalc.
- **Virtual rendering**: Only visible cells + 5-cell buffer are rendered. Scroll position drives which rows/cols appear.
- **Undo/redo**: `CommandBatch[]` stacks tracking cell deltas + selection state. Max 100 entries.

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
