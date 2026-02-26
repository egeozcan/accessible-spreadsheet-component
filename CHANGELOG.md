# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
