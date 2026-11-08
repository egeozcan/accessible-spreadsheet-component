import { LitElement, html, css, nothing } from 'lit';
import { customElement, property, state, query } from 'lit/decorators.js';
import { repeat } from 'lit/directives/repeat.js';
import { styleMap } from 'lit/directives/style-map.js';
import {
  type CellCoord,
  type CellData,
  type CellFormat,
  type GridData,
  type SelectionRange,
  type FormulaFunction,
  type CellChangeDetail,
  type SelectionChangeDetail,
  type DataChangeDetail,
  type FormatChangeDetail,
  type DataChangeOperation,
  type CellEditStartDetail,
  type ProposedChange,
  type BeforeChangeDetail,
  type CellValidator,
  type ValidationErrorDetail,
  type StructureChangeDetail,
  type SortDirection,
  type SortKey,
  type SortOptions,
  type SortDetail,
  type ColumnResizeDetail,
  type FreezeChangeDetail,
  cellKey,
  colToLetter,
  coordToRef,
  formatNumber,
  formatsEqual,
  parseKey,
  refToCoord,
} from './types.js';
import { SelectionManager, findDataEdge } from './controllers/selection-manager.js';
import { parseDelimited, serializeDelimited } from './utils/delimited.js';
import { adjustFormulaForStructureChange, type StructureChange } from './utils/references.js';
import { compareSortValues, sortValueOf, type SortValue } from './utils/sort.js';
import { FormulaEngine } from './engine/formula-engine.js';
import { ClipboardManager } from './controllers/clipboard-manager.js';
import './components/y11n-formula-bar.js';
import type { FormulaBarMode, Y11nFormulaBar } from './components/y11n-formula-bar.js';

type ChangeSource = NonNullable<DataChangeDetail['source']>;

function cloneCell(cell: CellData): CellData {
  const copy: CellData = { ...cell };
  if (cell.format) {
    copy.format = { ...cell.format };
    if (cell.format.numberFormat) {
      copy.format.numberFormat = { ...cell.format.numberFormat };
    }
  }
  return copy;
}

/**
 * Lay out a sorted list of rendered indices across `total` slots, emitting
 * a spacer for every run of unrendered slots so the CSS grid keeps its geometry.
 */
function layoutWithGaps<T>(
  indices: number[],
  total: number,
  item: (index: number) => T,
  gap: (start: number, size: number) => T
): T[] {
  const out: T[] = [];
  let next = 0;
  for (const i of indices) {
    if (i > next) out.push(gap(next, i - next));
    out.push(item(i));
    next = i + 1;
  }
  if (total > next) out.push(gap(next, total - next));
  return out;
}

/** Contiguous `[start, end)` indices plus any `extras` outside that window, sorted. */
function windowWithExtras(start: number, end: number, extras: number[]): number[] {
  const out: number[] = [];
  for (let i = start; i < end; i++) out.push(i);
  for (const e of extras) {
    if ((e < start || e >= end) && !out.includes(e)) out.push(e);
  }
  return out.sort((a, b) => a - b);
}

const ARROW_DELTAS: Record<string, [number, number]> = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
};

/**
 * A cell write for `_buildCellBatch`. `format: undefined` leaves the cell's
 * format alone, `null` clears it, and an object replaces it.
 */
interface CellUpdate {
  id: string;
  value: string;
  format?: CellFormat | null;
}

/** Options for `toCSV()`. */
export interface CSVExportOptions {
  /** Range to export. Default: from A1 to the last row/column that has content. */
  range?: SelectionRange;
  /** Field separator. Default: `,` */
  delimiter?: string;
  /** Export computed display values (default) or raw values (formulas). */
  values?: 'display' | 'raw';
}

/** Options for `findCells()`. */
export interface FindOptions {
  /** Match upper/lower case exactly. Default: false */
  matchCase?: boolean;
  /** Search raw values (formula text) instead of displayed values. Default: false */
  formulas?: boolean;
}

/** Options for `fromCSV()`. */
export interface CSVImportOptions {
  /** Field separator. Default: `,` */
  delimiter?: string;
  /** Top-left cell to write to. Default: A1 */
  start?: CellCoord;
}

interface CellDelta {
  id: string;
  before: string;
  after: string;
  formatBefore?: CellFormat;
  formatAfter?: CellFormat;
}

interface SelectionSnapshot {
  anchor: CellCoord;
  head: CellCoord;
  /** Active-cell override (whole-row/column selections); null when it follows the head */
  active: CellCoord | null;
}

/** Whole-cell before/after state for a row/column insertion or deletion. */
interface StructureCellDelta {
  id: string;
  before?: CellData;
  after?: CellData;
}

interface StructureBatch {
  change: StructureChange;
  cells: StructureCellDelta[];
  rowsBefore: number;
  colsBefore: number;
  rowsAfter: number;
  colsAfter: number;
  /** Custom column widths before/after, when columns with custom widths move */
  colWidths?: { before: Map<number, number>; after: Map<number, number> };
}

/** A column width change; `null` is the default width (`--ls-cell-width`). */
interface WidthDelta {
  col: number;
  before: number | null;
  after: number | null;
}

/** Whole-cell before/after state for a sort, plus what was sorted. */
interface SortBatch {
  detail: Omit<SortDetail, 'source'>;
  cells: StructureCellDelta[];
}

interface CommandBatch {
  deltas: CellDelta[];
  selectionBefore: SelectionSnapshot;
  selectionAfter: SelectionSnapshot;
  operation: DataChangeOperation;
  /** Set for row/column insertions and deletions, which replace whole cells */
  structure?: StructureBatch;
  /** Set for sorts, which move whole cells */
  sort?: SortBatch;
  /** Set for column width changes, which touch no cells */
  widths?: WidthDelta[];
}

const STRUCTURE_OPERATIONS: Record<string, DataChangeOperation> = {
  'insert:row': 'insert-rows',
  'delete:row': 'delete-rows',
  'insert:col': 'insert-columns',
  'delete:col': 'delete-columns',
};

/**
 * Least-squares linear trend of a run of numbers (for fill-handle series),
 * or null if the run has fewer than two values or any non-number (null).
 */
function linearTrend(values: Array<number | null>): { intercept: number; slope: number } | null {
  if (values.length < 2 || values.some((v) => v === null)) return null;
  const ys = values as number[];
  const n = ys.length;
  const meanX = (n - 1) / 2;
  const meanY = ys.reduce((a, b) => a + b, 0) / n;
  let num = 0;
  let den = 0;
  ys.forEach((y, x) => {
    num += (x - meanX) * (y - meanY);
    den += (x - meanX) ** 2;
  });
  const slope = num / den;
  return { intercept: meanY - slope * meanX, slope };
}

/** Render a computed series value without floating-point noise (0.1+0.2 → 0.3). */
function formatSeriesNumber(value: number): string {
  return String(parseFloat(value.toPrecision(15)));
}

/**
 * `<y11n-spreadsheet>` -- An accessible spreadsheet web component.
 *
 * Implements the WAI-ARIA Grid pattern with:
 * - Roving tabindex and full keyboard navigation (arrows, Tab, Home/End, Ctrl+arrows)
 * - Recursive-descent formula engine (`=SUM(A1:B5)`, `=IF(...)`, etc.)
 * - Copy/cut/paste via the Clipboard API (HTML table + TSV formats)
 * - Undo/redo command history (up to 100 entries)
 * - Virtual rendering for large grids (only visible cells + buffer are in the DOM)
 * - Cell formatting (bold, italic, colors, number formats, etc.)
 *
 * @fires cell-change - When a single cell value is edited
 * @fires selection-change - When the selected range changes
 * @fires data-change - On bulk updates (paste, cut, clear, undo/redo)
 * @fires format-change - When cell formatting is modified
 * @fires cell-edit-start - Before the inline editor opens (cancelable)
 * @fires before-change - Before a user edit, paste, cut, clear or fill is applied (cancelable)
 * @fires validation-error - When `validateCell` rejects one or more values
 * @fires before-structure-change - Before a user inserts or deletes rows/columns (cancelable)
 * @fires structure-change - After rows/columns are inserted or deleted
 * @fires before-sort - Before a user sorts a range (cancelable)
 * @fires sort-change - After a range is sorted (also on undo/redo)
 * @fires before-column-resize - Before a user resizes columns (cancelable)
 * @fires column-resize - After column widths change (also on undo/redo)
 * @fires freeze-change - After the frozen rows/columns change (not when the properties are set directly)
 *
 * @csspart editor - The inline cell editor input element
 * @csspart validation-message - The message shown under the editor when a value is rejected
 */
@customElement('y11n-spreadsheet')
export class Y11nSpreadsheet extends LitElement {
  // ─── Public Properties ──────────────────────────────

  /** Number of rows in the grid. Default: 100 */
  @property({ type: Number }) rows = 100;
  /** Number of columns in the grid. Default: 26 */
  @property({ type: Number }) cols = 26;
  /** The sparse grid data map. Set externally or via `setData()`. */
  @property({ attribute: false }) data: GridData = new Map();
  /** When true, all editing is disabled. Default: false */
  @property({ type: Boolean, attribute: 'read-only', reflect: true }) readOnly = false;
  /** Accessible label for the grid. Default: 'Spreadsheet' */
  @property({ type: String, attribute: 'grid-label' }) gridLabel = 'Spreadsheet';
  /** Record of custom formula functions to register, keyed by name. */
  @property({ attribute: false }) functions: Record<string, FormulaFunction> = {};
  /**
   * Optional validator for values the user enters (typing, formula bar,
   * paste, fill). Return an error message to reject a value, or null/undefined
   * to accept it. Clearing a cell is never validated.
   */
  @property({ attribute: false }) validateCell?: CellValidator;
  /**
   * Number of rows at the top that stay in view while scrolling. At least
   * one row always scrolls. Default: 0
   */
  @property({ type: Number, attribute: 'frozen-rows', reflect: true }) frozenRows = 0;
  /**
   * Number of columns at the left that stay in view while scrolling. At
   * least one column always scrolls. Default: 0
   */
  @property({ type: Number, attribute: 'frozen-cols', reflect: true }) frozenCols = 0;

  // ─── Internal State ─────────────────────────────────

  @state() private _isEditing = false;
  @state() private _editValue = '';
  @state() private _formulaBarMode: FormulaBarMode = 'raw';
  /** When true, the next render should reposition the editor cursor (set on edit start / ref mode) */
  private _editorNeedsCursorReset = false;

  /** Whether arrow keys are currently selecting a cell reference in a formula */
  @state() private _refMode = false;

  /** Screen reader announcement text for mode changes */
  @state() private _modeAnnouncement = '';

  /** Message from `validateCell` for the value in the editor, shown below it */
  @state() private _validationMessage = '';

  /** Find bar state */
  @state() private _findOpen = false;
  @state() private _findQuery = '';
  @state() private _findStatus = '';

  /** Cells a fill-handle drag would fill, for the drag preview */
  @state() private _fillPreview: SelectionRange | null = null;

  /** Track scroll position for virtual rendering */
  @state() private _scrollTop = 0;
  @state() private _scrollLeft = 0;

  // ─── DOM Refs ───────────────────────────────────────

  @query('#editor') private _editor!: HTMLInputElement;
  @query('#find-input') private _findInput?: HTMLInputElement;
  @query('.ls-grid') private _grid!: HTMLDivElement;

  // ─── Controllers / Managers ─────────────────────────

  private _selection = new SelectionManager(this, this.rows, this.cols);
  private _formulaEngine = new FormulaEngine();
  private _clipboardManager = new ClipboardManager();

  /** Internal data store - we keep a separate mutable copy */
  private _internalData: GridData = new Map();
  private _lastSyncedDataRef: GridData | null = null;

  /** Track which cell started editing for commit logic */
  private _editingCellKey: string | null = null;

  private _undoStack: CommandBatch[] = [];
  private _redoStack: CommandBatch[] = [];
  private readonly _maxHistory = 100;

  /** Reference mode cursor position */
  private _refCursorRow = 0;
  private _refCursorCol = 0;

  /** Start/end positions of the reference text within the formula string */
  private _refInsertStart = 0;
  private _refInsertEnd = 0;

  /** Whether an async paste operation is currently in progress */
  private _pasteInProgress = false;

  /**
   * Bumped whenever rows/columns are inserted or deleted (including undo/redo),
   * so async clipboard work can tell that its cell coordinates went stale.
   */
  private _structureRevision = 0;

  /**
   * Bumped when the data is replaced (`setData()` / `data`), which clears the
   * undo history. Operations that started before then must not push history
   * entries or keep writing (e.g. a progressive fill between frames).
   */
  private _dataGeneration = 0;

  /**
   * Bumped by every data change. Cancelable events compare it before and after
   * dispatch to detect a listener that changed the grid reentrantly.
   */
  private _mutationCount = 0;

  /** Number of cells to process per animation frame during large pastes */
  private static readonly PASTE_CHUNK_SIZE = 500;
  private static readonly INITIAL_VIEW_ROWS = 15;
  private static readonly INITIAL_VIEW_COLS = 8;
  private static readonly ROW_HEADER_WIDTH = 50;

  /**
   * Rendered cell size in px, measured from `--ls-cell-width` /
   * `--ls-cell-height` (see `_measureCellSize()`). Defaults match the CSS
   * fallbacks until the first measurement.
   */
  private _cellSize = { width: 100, height: 28 };

  /** Column widths set by the user or API, in px; other columns use `_cellSize.width` */
  private _colWidths = new Map<number, number>();
  /** Live widths while a header border is dragged, before the resize is committed */
  private _resizePreview: { cols: number[]; width: number } | null = null;
  /** Bumped whenever a column width (or preview) changes, to invalidate `_colLayout` */
  private _colWidthsRevision = 0;
  /** Cached left offsets of every column (`cols + 1` entries), see `_colOffsets()` */
  private _colLayout: { key: string; offsets: Float64Array } | null = null;
  private _measureCanvas: CanvasRenderingContext2D | null | undefined;

  private static readonly MIN_COL_WIDTH = 20;
  private static readonly MAX_COL_WIDTH = 2000;
  /** Alt+Shift+Left/Right step, in px */
  private static readonly COL_RESIZE_STEP = 8;
  /** The `_cellSize` object the last render used */
  private _renderedCellSize = this._cellSize;
  private _cellSizeObserver: ResizeObserver | null = null;

  // ─── Lifecycle ──────────────────────────────────────

  connectedCallback(): void {
    super.connectedCallback();
    // Re-attach after a move in the DOM; the first attach happens in firstUpdated
    if (this.hasUpdated) this._observeCellSize();
  }

  disconnectedCallback(): void {
    super.disconnectedCallback();
    this._cellSizeObserver?.disconnect();
  }

  protected firstUpdated(): void {
    // Measure synchronously so API calls right after the first render
    // (scrolling, editor placement) already use the real size. The re-render
    // with it is left to the ResizeObserver's initial callback: requesting an
    // update from firstUpdated trips Lit's change-in-update warning.
    this._measureCellSize();
    if (!this._observeCellSize() && this._cellSize !== this._renderedCellSize) this.requestUpdate();
  }

  willUpdate(changedProps: Map<string, unknown>): void {
    if (changedProps.has('data') && this.data !== this._lastSyncedDataRef) {
      this._syncData();
    }
    if (changedProps.has('readOnly') && this.readOnly && this._isEditing) {
      // The grid became read-only mid-edit: discard the edit
      this._exitRefMode();
      this._cancelEdit();
    }
    if (changedProps.has('rows') || changedProps.has('cols')) {
      this.rows = Math.max(1, Math.round(this.rows));
      this.cols = Math.max(1, Math.round(this.cols));
      this._selection.setBounds(this.rows, this.cols);
    }
    if (changedProps.has('functions')) {
      // On first render the old value is undefined and this.functions is the
      // default empty object.  Skip the setFunctions call in that case so we
      // don't wipe out functions added imperatively via registerFunction().
      const prev = changedProps.get('functions') as Record<string, FormulaFunction> | undefined;
      if (prev !== undefined || Object.keys(this.functions).length > 0) {
        this._formulaEngine.setFunctions(this.functions);
        this._recalcAll();
      }
    }
  }

  // ─── Public API Methods ─────────────────────────────

  /**
   * Returns a shallow copy of the current grid data.
   *
   * @returns A new Map containing all cell data
   */
  getData(): GridData {
    const copy: GridData = new Map();
    for (const [key, cell] of this._internalData) {
      copy.set(key, cloneCell(cell));
    }
    return copy;
  }

  /**
   * Replaces the entire grid data and resets undo/redo history.
   *
   * @param data - The new sparse grid data map
   */
  setData(data: GridData): void {
    this.data = data;
    this._syncData();
    this.requestUpdate();
  }

  /**
   * Register a custom formula function at runtime.
   * Triggers a full recalculation of all formulas.
   *
   * @param name - Function name (will be uppercased for lookup)
   * @param fn - The function implementation
   */
  registerFunction(name: string, fn: FormulaFunction): void {
    this._formulaEngine.registerFunction(name, fn);
    this._recalcAll();
  }

  /**
   * Serialize grid contents as CSV (or TSV etc. via `delimiter`).
   * By default exports display values from A1 to the last row and column
   * that contain a value.
   *
   * @param options - Range, delimiter, and whether to export display or raw values
   * @returns The delimited text (rows joined with `\n`); empty string for an empty grid
   */
  toCSV(options: CSVExportOptions = {}): string {
    const { delimiter = ',', values = 'display' } = options;
    let range = options.range;
    if (!range) {
      let maxRow = -1;
      let maxCol = -1;
      for (const [key, cell] of this._internalData) {
        if (cell.rawValue === '') continue;
        const { row, col } = parseKey(key);
        if (row > maxRow) maxRow = row;
        if (col > maxCol) maxCol = col;
      }
      if (maxRow < 0) return '';
      range = { start: { row: 0, col: 0 }, end: { row: maxRow, col: maxCol } };
    }

    const rows: string[][] = [];
    for (let r = range.start.row; r <= range.end.row; r++) {
      const fields: string[] = [];
      for (let c = range.start.col; c <= range.end.col; c++) {
        const cell = this._internalData.get(cellKey(r, c));
        fields.push((values === 'raw' ? cell?.rawValue : cell?.displayValue) ?? '');
      }
      rows.push(fields);
    }
    return serializeDelimited(rows, delimiter);
  }

  /**
   * Write CSV (or TSV etc. via `delimiter`) text into the grid starting at
   * `start` (default A1). Values beginning with `=` become formulas. Values
   * outside the grid's `rows`/`cols` are dropped, so grow the grid first if
   * needed. The import is a single undoable step and fires `data-change`
   * with `operation: 'import'`.
   *
   * @param text - Delimited text
   * @param options - Delimiter and top-left target cell
   */
  fromCSV(text: string, options: CSVImportOptions = {}): void {
    if (this._pasteInProgress) return;
    this._syncPendingData();
    const { delimiter = ',', start = { row: 0, col: 0 } } = options;
    if (!Number.isInteger(start.row) || !Number.isInteger(start.col) || start.row < 0 || start.col < 0) {
      throw new RangeError('fromCSV: start must be a cell with non-negative integer row and col');
    }
    const updates: CellUpdate[] = [];
    parseDelimited(text, delimiter).forEach((fields, r) => {
      fields.forEach((value, c) => {
        const row = start.row + r;
        const col = start.col + c;
        if (row < this.rows && col < this.cols) {
          updates.push({ id: cellKey(row, col), value });
        }
      });
    });

    const selection = this._snapshotSelection();
    const batch = this._buildCellBatch(updates, 'import', selection);
    if (!batch) return;
    this._commitBatch(batch, 'programmatic');
  }

  // ─── Rows / Columns API ──────────────────────────────

  /**
   * Insert `count` empty rows before row `index` (0-based; `index === rows`
   * appends). Cells below move down, the grid grows by `count` rows, and
   * formula references are rewritten to follow the cells they point to.
   * Undoable; fires `data-change` and `structure-change`.
   *
   * @param index - Row to insert before
   * @param count - Number of rows to insert (default 1)
   */
  insertRows(index: number, count = 1): void {
    this._changeStructure({ axis: 'row', kind: 'insert', index, count }, 'programmatic');
  }

  /**
   * Delete `count` rows starting at row `index` (0-based). Cells below move
   * up and the grid shrinks; references to deleted cells become `#REF!`.
   * At least one row must remain. Undoable; fires `data-change` and
   * `structure-change`.
   *
   * @param index - First row to delete
   * @param count - Number of rows to delete (default 1)
   */
  deleteRows(index: number, count = 1): void {
    this._changeStructure({ axis: 'row', kind: 'delete', index, count }, 'programmatic');
  }

  /**
   * Insert `count` empty columns before column `index` (0-based;
   * `index === cols` appends). See `insertRows()`.
   *
   * @param index - Column to insert before
   * @param count - Number of columns to insert (default 1)
   */
  insertColumns(index: number, count = 1): void {
    this._changeStructure({ axis: 'col', kind: 'insert', index, count }, 'programmatic');
  }

  /**
   * Delete `count` columns starting at column `index` (0-based). See
   * `deleteRows()`.
   *
   * @param index - First column to delete
   * @param count - Number of columns to delete (default 1)
   */
  deleteColumns(index: number, count = 1): void {
    this._changeStructure({ axis: 'col', kind: 'delete', index, count }, 'programmatic');
  }

  // ─── Sort API ────────────────────────────────────────

  /**
   * Sort the rows of `range` by one or more key columns. Only cells inside
   * the range move; with `hasHeader`, its first row stays in place.
   *
   * Ascending order is numbers (smallest first), text (A to Z, ignoring
   * case), FALSE, TRUE, then errors; descending reverses it. Blank cells go
   * last either way. The sort is stable, and later keys break ties of
   * earlier ones. Cells move with their formats and read-only flags.
   * Formulas in moved cells shift their relative references by the rows they
   * moved, as if copied; references from outside the range are not changed.
   *
   * One undoable step; fires `data-change` and `sort-change` (not
   * `before-sort`). Throws a RangeError for a range outside the grid or a key
   * column outside the range.
   *
   * @param range - The cells to sort
   * @param keys - Sort key(s) in priority order
   * @param options - Whether the first row is a header
   * @returns Whether any cell moved
   */
  sort(range: SelectionRange, keys: SortKey | SortKey[], options: SortOptions = {}): boolean {
    const keyList = Array.isArray(keys) ? keys : [keys];
    const start = range?.start;
    const end = range?.end;
    const isIndex = (n: unknown): n is number => Number.isInteger(n) && (n as number) >= 0;
    if (!start || !end || ![start.row, start.col, end.row, end.col].every(isIndex)) {
      throw new RangeError('sort: range must have non-negative integer start and end coordinates');
    }
    const normalized: SelectionRange = {
      start: { row: Math.min(start.row, end.row), col: Math.min(start.col, end.col) },
      end: { row: Math.max(start.row, end.row), col: Math.max(start.col, end.col) },
    };
    if (normalized.end.row >= this.rows || normalized.end.col >= this.cols) {
      throw new RangeError(`sort: the range is outside the grid (${this.rows} rows, ${this.cols} columns)`);
    }
    if (keyList.length === 0) throw new RangeError('sort: at least one key is required');
    const resolved = keyList.map((key) => {
      if (!key || !Number.isInteger(key.col) || key.col < normalized.start.col || key.col > normalized.end.col) {
        throw new RangeError(`sort: key column ${key?.col} is outside the range`);
      }
      const direction = key.direction ?? 'ascending';
      if (direction !== 'ascending' && direction !== 'descending') {
        throw new RangeError(`sort: direction must be 'ascending' or 'descending'`);
      }
      return { col: key.col, direction };
    });
    return this._sortRange(
      { range: normalized, keys: resolved, hasHeader: options.hasHeader === true },
      'programmatic'
    );
  }

  /**
   * Sort as the user would with Alt+Shift+Up / Alt+Shift+Down: by the
   * active cell's column, over the selection, or over the block of data
   * around the active cell when a single cell or a single column of a wider
   * table is selected. A first row of
   * text headers over other data is detected and kept in place. Respects
   * `readOnly`, refuses to move read-only cells, and fires the cancelable
   * `before-sort`. The outcome is announced to screen readers. Use this to
   * back a "Sort A to Z" button.
   *
   * @param direction - Default: `'ascending'`
   * @returns Whether any cell moved
   */
  sortSelection(direction: SortDirection = 'ascending'): boolean {
    if (this.readOnly || this._pasteInProgress) return false;
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();

    const active = this._selection.activeCell;
    const selected = this._selection.range;
    const single = selected.start.row === selected.end.row && selected.start.col === selected.end.col;
    let range = single ? this._currentRegion(active.row, active.col) : this._clipToData(selected, active);
    // Sorting one column of a wider table would split its records apart, so
    // sort the whole table instead (Excel's default when it asks)
    if (range && range.start.col === range.end.col) {
      // Grow from the column's first filled cell: the active cell may be below the data
      let row = range.start.row;
      while (row < range.end.row && !this._isCellFilled(row, range.start.col)) row++;
      const region = this._currentRegion(row, range.start.col);
      if (region && region.start.col < region.end.col) range = region;
    }
    if (!range) {
      this._modeAnnouncement = 'Nothing to sort';
      return false;
    }
    return this._sortRange(
      { range, keys: [{ col: active.col, direction }], hasHeader: this._looksLikeHeader(range) },
      'user'
    );
  }

  // ─── Column Widths API ───────────────────────────────

  /**
   * The width of a column in px: its custom width, or the default from
   * `--ls-cell-width`.
   *
   * @param col - 0-based column index
   */
  getColumnWidth(col: number): number {
    return this._colWidth(col);
  }

  /**
   * The custom column widths, keyed by 0-based column index. Columns at the
   * default width are left out. Pass the result to `setColumnWidths()` to
   * restore a saved layout.
   */
  getColumnWidths(): Record<number, number> {
    const out: Record<number, number> = {};
    for (const [col, width] of [...this._colWidths].sort((a, b) => a[0] - b[0])) {
      if (col < this.cols) out[col] = width;
    }
    return out;
  }

  /**
   * Set the width of a column in px, or `null` for the default width.
   * Widths are rounded and clamped to 20–2000px. Undoable; fires
   * `column-resize`.
   *
   * @param col - 0-based column index
   * @param width - Width in px, or null to reset
   */
  setColumnWidth(col: number, width: number | null): void {
    this.setColumnWidths({ [col]: width });
  }

  /**
   * Set several column widths at once (one undo step). Keys are 0-based
   * column indices; `null` resets a column to the default width. Throws a
   * RangeError for a column outside the grid or a width that isn't a
   * positive number.
   *
   * @param widths - Widths in px keyed by column index
   */
  setColumnWidths(widths: Record<number, number | null>): void {
    const changes = Object.entries(widths).map(([key, width]) => {
      const col = Number(key);
      if (!Number.isInteger(col) || col < 0 || col >= this.cols) {
        throw new RangeError(`setColumnWidths: column ${key} is outside the grid (${this.cols} columns)`);
      }
      if (width !== null && !(typeof width === 'number' && Number.isFinite(width) && width > 0)) {
        throw new RangeError(`setColumnWidths: the width of column ${key} must be a positive number or null`);
      }
      return { col, width };
    });
    this._resizeColumns(changes, 'programmatic');
  }

  /**
   * Fit columns to their widest value, as a double-click on a header border
   * does. Empty columns go back to the default width. Undoable; fires
   * `column-resize`.
   *
   * @param cols - 0-based column indices
   */
  autoFitColumns(cols: number[]): void {
    for (const col of cols) {
      if (!Number.isInteger(col) || col < 0 || col >= this.cols) {
        throw new RangeError(`autoFitColumns: column ${col} is outside the grid (${this.cols} columns)`);
      }
    }
    this._resizeColumns(cols.map((col) => ({ col, width: this._fitWidth(col) })), 'programmatic');
  }

  // ─── Frozen Panes API ────────────────────────────────

  /**
   * Freeze the top `rows` rows and left `cols` columns (default 0) so they
   * stay in view while the rest scrolls. Without arguments, freezes the rows
   * above and the columns left of the active cell, as Alt+Shift+F does (at
   * A1 that unfreezes). Values are
   * clamped so at least one row and column scroll. Not undoable (a view
   * setting); fires `freeze-change`.
   *
   * @param rows - Number of rows to freeze
   * @param cols - Number of columns to freeze
   */
  freezePanes(rows?: number, cols?: number): void {
    const active = this._selection.activeCell;
    const r = rows ?? active.row;
    // freezePanes(1) freezes one row, not also the columns left of the active cell
    const c = cols ?? (rows === undefined ? active.col : 0);
    if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0) {
      throw new RangeError('freezePanes: rows and cols must be non-negative integers');
    }
    this._setFrozen(r, c, 'programmatic');
  }

  /** Unfreeze all rows and columns. Fires `freeze-change`. */
  unfreezePanes(): void {
    this._setFrozen(0, 0, 'programmatic');
  }

  // ─── Find API ────────────────────────────────────────

  /**
   * Find the cells whose displayed value (or formula text, with
   * `formulas: true`) contains `query`. Case-insensitive by default.
   *
   * @param query - Text to search for
   * @param options - Case sensitivity and whether to search formulas
   * @returns Matching cell keys in reading order (row by row)
   */
  findCells(query: string, options: FindOptions = {}): string[] {
    if (!query) return [];
    const { matchCase = false, formulas = false } = options;
    const needle = matchCase ? query : query.toLowerCase();
    const hits: CellCoord[] = [];
    for (const [key, cell] of this._internalData) {
      const text = formulas ? cell.rawValue : cell.displayValue;
      if (!text) continue;
      const haystack = matchCase ? text : text.toLowerCase();
      if (!haystack.includes(needle)) continue;
      const coord = parseKey(key);
      if (coord.row < this.rows && coord.col < this.cols) hits.push(coord);
    }
    hits.sort((a, b) => a.row - b.row || a.col - b.col);
    return hits.map((c) => cellKey(c.row, c.col));
  }

  /** Open the find bar (as Ctrl/Cmd+F does) and focus its input. */
  async openFind(): Promise<void> {
    if (this._isEditing) this._commitEdit();
    this._findOpen = true;
    await this.updateComplete;
    this._findInput?.focus();
    this._findInput?.select();
  }

  /** Close the find bar and return focus to the active cell. */
  closeFind(): void {
    if (!this._findOpen) return;
    this._findOpen = false;
    this._findStatus = '';
    this.updateComplete.then(() => this._focusActiveCell());
  }

  // ─── Format API ──────────────────────────────────────

  /**
   * Returns the format for a cell, or undefined if unformatted.
   *
   * @param cellId - Cell key in `"row:col"` format
   * @returns The cell's format, or undefined
   */
  getCellFormat(cellId: string): CellFormat | undefined {
    return this._internalData.get(cellId)?.format;
  }

  /**
   * Merges format properties into a cell. Creates the cell if it does not exist.
   *
   * @param cellId - Cell key in `"row:col"` format
   * @param format - Format properties to merge
   */
  setCellFormat(cellId: string, format: CellFormat): void {
    this._applyFormat([cellId], format, 'programmatic');
  }

  /**
   * Applies format to all cells in a rectangular range.
   *
   * @param range - The selection range to format
   * @param format - Format properties to apply
   */
  setRangeFormat(range: SelectionRange, format: CellFormat): void {
    const cellIds: string[] = [];
    for (let r = range.start.row; r <= range.end.row; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        cellIds.push(cellKey(r, c));
      }
    }
    this._applyFormat(cellIds, format, 'programmatic');
  }

  /**
   * Removes all formatting from a single cell.
   *
   * @param cellId - Cell key in `"row:col"` format
   */
  clearCellFormat(cellId: string): void {
    this._clearFormat([cellId], 'programmatic');
  }

  /**
   * Removes all formatting from every cell in a range.
   *
   * @param range - The selection range to clear formatting from
   */
  clearRangeFormat(range: SelectionRange): void {
    const cellIds: string[] = [];
    for (let r = range.start.row; r <= range.end.row; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        cellIds.push(cellKey(r, c));
      }
    }
    this._clearFormat(cellIds, 'programmatic');
  }

  /**
   * Toggles a boolean format property on the current selection.
   * If all selected cells have the property set, it is removed; otherwise it is applied to all.
   *
   * @param prop - The boolean format property to toggle
   */
  toggleFormat(prop: 'bold' | 'italic' | 'underline' | 'strikethrough'): void {
    if (this.readOnly || this._pasteInProgress) return;
    const range = this._selection.range;
    const cellIds: string[] = [];
    let allHaveProp = true;
    let locked = 0;

    for (let r = range.start.row; r <= range.end.row; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        const id = cellKey(r, c);
        if (this._isCellLocked(id)) {
          locked++;
          continue;
        }
        cellIds.push(id);
        const cell = this._internalData.get(id);
        if (!cell?.format?.[prop]) {
          allHaveProp = false;
        }
      }
    }

    if (locked > 0) this._announceSkipped(locked, []);
    if (cellIds.length === 0) return;
    const format: CellFormat = { [prop]: !allHaveProp };
    this._applyFormat(cellIds, format, 'user');
  }

  // ─── Format Internals ────────────────────────────────

  private _applyFormat(cellIds: string[], format: CellFormat, source: FormatChangeDetail['source']): void {
    const selection = this._snapshotSelection();
    const batch = this._buildFormatBatch(cellIds, format, selection);
    if (!batch) return;

    if (source === 'user') {
      this._executeUserBatch(batch);
    } else {
      this._commitBatch(batch, source);
    }

    this._dispatchFormatChange({ cellIds, format, source });
  }

  private _clearFormat(cellIds: string[], source: FormatChangeDetail['source']): void {
    const selection = this._snapshotSelection();
    const deltas: CellDelta[] = [];

    for (const id of cellIds) {
      const cell = this._internalData.get(id);
      if (cell?.format) {
        deltas.push({
          id,
          before: cell.rawValue,
          after: cell.rawValue,
          formatBefore: { ...cell.format },
          formatAfter: undefined,
        });
      }
    }

    if (deltas.length === 0) return;

    const batch: CommandBatch = {
      deltas,
      selectionBefore: selection,
      selectionAfter: selection,
      operation: 'format',
    };

    if (source === 'user') {
      this._executeUserBatch(batch);
    } else {
      this._commitBatch(batch, source);
    }

    this._dispatchFormatChange({ cellIds, format: {}, source });
  }

  private _buildFormatBatch(
    cellIds: string[],
    format: CellFormat,
    selection: SelectionSnapshot
  ): CommandBatch | null {
    const deltas: CellDelta[] = [];

    for (const id of cellIds) {
      const cell = this._internalData.get(id);
      const existingFormat = cell?.format;
      const merged = { ...existingFormat, ...format };
      // Clean out undefined values
      for (const k of Object.keys(merged) as (keyof CellFormat)[]) {
        if (merged[k] === undefined) delete merged[k];
      }

      const formatChanged = !formatsEqual(existingFormat, merged);
      if (formatChanged) {
        deltas.push({
          id,
          before: cell?.rawValue ?? '',
          after: cell?.rawValue ?? '',
          formatBefore: existingFormat ? { ...existingFormat } : undefined,
          formatAfter: Object.keys(merged).length > 0 ? merged : undefined,
        });
      }
    }

    if (deltas.length === 0) return null;

    return {
      deltas,
      selectionBefore: selection,
      selectionAfter: selection,
      operation: 'format',
    };
  }

  /** Convert CellFormat to inline styles for styleMap */
  private _getCellStyles(row: number, col: number): Record<string, string> {
    const cell = this._getCell(row, col);
    if (!cell?.format) return {};

    const styles: Record<string, string> = {};
    const fmt = cell.format;

    if (fmt.bold) styles['font-weight'] = 'bold';
    if (fmt.italic) styles['font-style'] = 'italic';

    const decorations: string[] = [];
    if (fmt.underline) decorations.push('underline');
    if (fmt.strikethrough) decorations.push('line-through');
    if (decorations.length > 0) styles['text-decoration'] = decorations.join(' ');

    if (fmt.textColor) styles.color = fmt.textColor;
    if (fmt.backgroundColor) styles['background-color'] = fmt.backgroundColor;
    if (fmt.fontSize) styles['font-size'] = `${fmt.fontSize}px`;

    if (fmt.textAlign) {
      styles['text-align'] = fmt.textAlign;
      const justifyMap = { left: 'flex-start', center: 'center', right: 'flex-end' };
      styles['justify-content'] = justifyMap[fmt.textAlign];
    }

    if (fmt.verticalAlign) {
      const alignMap = { top: 'flex-start', middle: 'center', bottom: 'flex-end' };
      styles['align-items'] = alignMap[fmt.verticalAlign];
    }

    if (fmt.wrap) {
      styles['white-space'] = 'normal';
      styles['overflow-wrap'] = 'anywhere';
    }

    return styles;
  }

  // ─── Data Sync ──────────────────────────────────────

  private _syncData(): void {
    const oldData = this._internalData;
    // Clone each cell so internal mutations (recalc, number formatting) never
    // leak into the caller's objects, and caller mutations never leak in.
    const newData: GridData = new Map();
    for (const [key, cell] of this.data) {
      newData.set(key, cloneCell(cell));
    }

    // Diff old vs new to collect changed keys for future targeted recalc
    const changedKeys: string[] = [];

    // Keys in new data that are different or absent in old data
    for (const [key, cell] of newData) {
      const oldCell = oldData.get(key);
      if (
        !oldCell ||
        oldCell.rawValue !== cell.rawValue ||
        !formatsEqual(oldCell.format, cell.format)
      ) {
        changedKeys.push(key);
      } else {
        // Unchanged cells keep their computed values; the caller's copy may
        // carry a placeholder displayValue (e.g. '' for formulas).
        cell.displayValue = oldCell.displayValue;
        cell.type = oldCell.type;
      }
    }

    // Keys in old data that are absent in new data (deleted)
    for (const key of oldData.keys()) {
      if (!newData.has(key)) {
        changedKeys.push(key);
      }
    }

    this._internalData = newData;
    this._lastSyncedDataRef = this.data;
    this._undoStack = [];
    this._redoStack = [];
    this._dataGeneration++;
    this._mutationCount++;

    if (changedKeys.length > 0 && oldData.size > 0) {
      // Incremental update: preserve dep graph for unchanged cells
      this._formulaEngine.updateData(this._internalData, changedKeys);
      // Literal cells aren't recalculated by the engine: derive their display
      // from the raw value so an added or removed number format takes effect.
      for (const key of changedKeys) {
        const cell = this._internalData.get(key);
        if (cell && !cell.rawValue.startsWith('=')) {
          const evaluated = this._formulaEngine.evaluate(cell.rawValue);
          cell.displayValue = evaluated.displayValue;
          cell.type = evaluated.type;
        }
      }
      this._recalcAffected(changedKeys);
      this._applyNumberFormatsToKeys(new Set(changedKeys));
    } else {
      // Full reset: first load or complete replacement
      this._formulaEngine.setData(this._internalData);
      this._recalcAll();
    }
  }

  /**
   * Whether the grid changed since `mutations` was read from `_mutationCount`,
   * counting a `data` assignment that hasn't been imported yet.
   */
  private _changedSince(mutations: number): boolean {
    return mutations !== this._mutationCount || this.data !== this._lastSyncedDataRef;
  }

  /**
   * Import a `data` assignment that Lit hasn't processed yet, so an API call
   * made right after it works on the new data instead of overwriting it.
   */
  private _syncPendingData(): void {
    if (this.data !== this._lastSyncedDataRef) this._syncData();
  }

  private _recalcAll(): void {
    this._formulaEngine.recalculate();
    this._reapplyNumberFormats();
  }

  private _recalcAffected(changedKeys: string[]): void {
    const changed = this._formulaEngine.recalculateAffected(changedKeys);
    this._applyNumberFormatsToKeys(changed);
  }

  /** Re-apply number formatting to all cells that have a numberFormat */
  private _reapplyNumberFormats(): void {
    for (const [key, cell] of this._internalData) {
      this._applyNumberFormatToCell(key, cell);
    }
  }

  /** Re-apply number formatting only to the given set of cell keys */
  private _applyNumberFormatsToKeys(keys: Set<string>): void {
    for (const key of keys) {
      const cell = this._internalData.get(key);
      if (cell) {
        this._applyNumberFormatToCell(key, cell);
      }
    }
  }

  /** Apply number formatting to a single cell using rawValue for robustness */
  private _applyNumberFormatToCell(key: string, cell: CellData): void {
    if (cell.type === 'number' && cell.format?.numberFormat) {
      const numericValue = cell.rawValue.startsWith('=')
        ? Number(this._formulaEngine.resultFor(key, cell.rawValue).displayValue)
        : Number(cell.rawValue);
      cell.displayValue = formatNumber(numericValue, cell.format.numberFormat);
    }
  }

  // ─── Cell Access ────────────────────────────────────

  private _getCell(row: number, col: number): CellData | undefined {
    return this._internalData.get(cellKey(row, col));
  }

  private _getCellDisplay(row: number, col: number): string {
    const cell = this._getCell(row, col);
    return cell?.displayValue ?? '';
  }

  private _getCellRaw(row: number, col: number): string {
    const cell = this._getCell(row, col);
    return cell?.rawValue ?? '';
  }

  private _getActiveCellKey(): string {
    const { row, col } = this._selection.activeCell;
    return cellKey(row, col);
  }

  private _getActiveCellRef(): string {
    return coordToRef(this._selection.activeCell);
  }

  private _setCellRaw(key: string, rawValue: string): void {
    const existing = this._internalData.get(key);
    const evaluated = this._formulaEngine.evaluate(rawValue, key);

    // Apply number format to display value if the cell has one and the value is numeric
    let displayValue = evaluated.displayValue;
    const numFmt = existing?.format?.numberFormat;
    if (numFmt && evaluated.type === 'number') {
      displayValue = formatNumber(Number(evaluated.displayValue), numFmt);
    }

    const cell: CellData = {
      rawValue,
      displayValue,
      type: evaluated.type,
      format: existing?.format,
    };
    if (existing?.readOnly) cell.readOnly = true;
    this._internalData.set(key, cell);
  }

  private _applyRawValueByString(key: string, rawValue: string): void {
    if (rawValue === '') {
      const existing = this._internalData.get(key);
      if (existing?.format || existing?.readOnly) {
        // Keep the entry so the format / read-only flag survives
        const cell: CellData = { rawValue: '', displayValue: '', type: 'text' };
        if (existing.format) cell.format = existing.format;
        if (existing.readOnly) cell.readOnly = true;
        this._internalData.set(key, cell);
      } else {
        this._internalData.delete(key);
      }
      return;
    }
    this._setCellRaw(key, rawValue);
  }

  private _snapshotSelection(): SelectionSnapshot {
    return {
      anchor: { ...this._selection.anchor },
      head: { ...this._selection.head },
      active: this._selection.activeOverride,
    };
  }

  private _restoreSelection(snapshot: SelectionSnapshot): void {
    this._selection.setSelection(snapshot.anchor, snapshot.head, snapshot.active);
    this._dispatchSelectionChange();
    this.updateComplete.then(() => this._focusActiveCell());
  }

  private _buildCommandBatch(
    updates: Array<{ id: string; value: string }>,
    operation: DataChangeOperation,
    selectionBefore: SelectionSnapshot,
    selectionAfter: SelectionSnapshot
  ): CommandBatch | null {
    const deltas: CellDelta[] = [];

    for (const update of updates) {
      const before = this._internalData.get(update.id)?.rawValue ?? '';
      const after = update.value;
      if (before !== after) {
        deltas.push({ id: update.id, before, after });
      }
    }

    if (deltas.length === 0) {
      return null;
    }

    return {
      deltas,
      selectionBefore,
      selectionAfter,
      operation,
    };
  }

  /**
   * Build a batch that writes values and, optionally, formats. Cells whose
   * value and format would both be unchanged are skipped.
   */
  private _buildCellBatch(
    updates: CellUpdate[],
    operation: DataChangeOperation,
    selection: SelectionSnapshot
  ): CommandBatch | null {
    const deltas: CellDelta[] = [];
    for (const update of updates) {
      const existing = this._internalData.get(update.id);
      const before = existing?.rawValue ?? '';
      const delta: CellDelta = { id: update.id, before, after: update.value };
      if (update.format !== undefined) {
        const formatAfter = update.format && Object.keys(update.format).length > 0
          ? { ...update.format }
          : undefined;
        if (!formatsEqual(existing?.format, formatAfter)) {
          delta.formatBefore = existing?.format ? { ...existing.format } : undefined;
          delta.formatAfter = formatAfter;
        }
      }
      if (before !== update.value || 'formatAfter' in delta) {
        deltas.push(delta);
      }
    }
    if (deltas.length === 0) return null;
    return { deltas, selectionBefore: selection, selectionAfter: selection, operation };
  }

  /**
   * Execute and record a user batch. Large batches are applied in chunks
   * across animation frames so the UI stays responsive.
   */
  private async _runUserBatch(batch: CommandBatch): Promise<void> {
    if (batch.deltas.length <= Y11nSpreadsheet.PASTE_CHUNK_SIZE) {
      this._executeUserBatch(batch);
      return;
    }
    this._pasteInProgress = true;
    const generation = this._dataGeneration;
    try {
      if (await this._applyBatchProgressive(batch)) {
        this._pushHistory(batch, generation);
      }
    } finally {
      this._pasteInProgress = false;
    }
  }

  /**
   * Record a batch for undo, unless the data was replaced (which clears the
   * history) since `generation` was captured, before the batch was applied.
   */
  private _pushHistory(batch: CommandBatch, generation: number): void {
    if (generation !== this._dataGeneration) return;
    this._undoStack.push(batch);
    this._redoStack = [];
    if (this._undoStack.length > this._maxHistory) {
      this._undoStack.shift();
    }
  }

  private _finalizeBatch(
    batch: CommandBatch,
    source: ChangeSource,
    valueSide: 'before' | 'after'
  ): void {
    this._recalcAffected(batch.deltas.map((d) => d.id));

    const shouldEmitDataChange = source !== 'user' || batch.operation !== 'edit';
    if (shouldEmitDataChange) {
      this._dispatchDataChange({
        updates: batch.deltas.map((delta) => ({ id: delta.id, value: delta[valueSide] })),
        source,
        operation: batch.operation,
      });
    }

    this.requestUpdate();
  }

  private _applyBatch(batch: CommandBatch, source: ChangeSource): void {
    // Column widths are layout, not data: they must not make pending writes look stale
    if (!batch.widths) this._mutationCount++;
    if (batch.structure) {
      this._applyStructureBatch(batch, 'after', source);
      return;
    }
    if (batch.sort) {
      this._applySortBatch(batch, 'after', source);
      return;
    }
    if (batch.widths) {
      this._applyWidthBatch(batch, 'after', source);
      return;
    }
    for (const delta of batch.deltas) {
      this._applyRawValueByString(delta.id, delta.after);
      if ('formatAfter' in delta) {
        this._applyCellFormat(delta.id, delta.formatAfter);
      }
    }
    this._lastSyncedDataRef = this.data;
    this._restoreSelection(batch.selectionAfter);
    this._finalizeBatch(batch, source, 'after');
  }

  private _revertBatch(batch: CommandBatch): void {
    if (!batch.widths) this._mutationCount++;
    if (batch.structure) {
      this._applyStructureBatch(batch, 'before', 'undo');
      return;
    }
    if (batch.sort) {
      this._applySortBatch(batch, 'before', 'undo');
      return;
    }
    if (batch.widths) {
      this._applyWidthBatch(batch, 'before', 'undo');
      return;
    }
    for (const delta of batch.deltas) {
      this._applyRawValueByString(delta.id, delta.before);
      if ('formatBefore' in delta) {
        this._applyCellFormat(delta.id, delta.formatBefore);
      }
    }
    this._lastSyncedDataRef = this.data;
    this._restoreSelection(batch.selectionBefore);
    this._finalizeBatch(batch, 'undo', 'before');
  }

  /** Low-level: set or remove format on a cell */
  private _applyCellFormat(cellId: string, format: CellFormat | undefined): void {
    let cell = this._internalData.get(cellId);
    if (!cell) {
      // Create an empty cell to hold the format
      cell = { rawValue: '', displayValue: '', type: 'text' };
      this._internalData.set(cellId, cell);
    }
    if (format && Object.keys(format).length > 0) {
      cell.format = format;
    } else {
      delete cell.format;
      // Remove empty cells
      if (cell.rawValue === '' && !cell.readOnly) {
        this._internalData.delete(cellId);
        return;
      }
    }
    // Re-apply number format to display value
    if (cell.type === 'number' && cell.format?.numberFormat) {
      // For formula cells, get the unformatted result from the engine
      // (displayValue may already be formatted, e.g., "$1,234.50")
      if (cell.rawValue.startsWith('=')) {
        const evaluated = this._formulaEngine.resultFor(cellId, cell.rawValue);
        cell.displayValue = formatNumber(Number(evaluated.displayValue), cell.format.numberFormat);
      } else {
        cell.displayValue = formatNumber(Number(cell.rawValue), cell.format.numberFormat);
      }
    } else if (cell.type === 'number' && !cell.format?.numberFormat) {
      // Remove number formatting — show the plain result
      const evaluated = this._formulaEngine.resultFor(cellId, cell.rawValue);
      cell.displayValue = evaluated.displayValue;
    }
  }

  private _executeUserBatch(batch: CommandBatch): void {
    this._commitBatch(batch, 'user');
  }

  /**
   * Record a batch for undo, then apply it. Recording first keeps the history
   * in order when a listener of the events fired while applying makes a
   * change of its own (say, re-sorting on every `data-change`): that change
   * is recorded after this one, so undo reverts them in the right order. If a
   * listener replaces the data instead, the history is cleared with it.
   */
  private _commitBatch(batch: CommandBatch, source: ChangeSource): void {
    this._pushHistory(batch, this._dataGeneration);
    this._applyBatch(batch, source);
  }

  private _commitRawValue(key: string, newValue: string, options: { validate?: boolean } = {}): boolean {
    const oldCell = this._internalData.get(key);
    const oldValue = oldCell?.rawValue ?? '';
    if (newValue === oldValue) return false;

    const selection = this._snapshotSelection();
    const batch = this._buildCommandBatch(
      [{ id: key, value: newValue }],
      'edit',
      selection,
      selection
    );

    const vetted = batch && this._vetUserBatch(batch, options);
    if (!vetted) return false;
    this._executeUserBatch(vetted);
    this._dispatchCellChange({ cellId: key, value: vetted.deltas[0].after, oldValue });
    return true;
  }

  // ─── Read-only cells, validation, before-change ─────

  /** Whether a cell is individually marked read-only (`CellData.readOnly`). */
  private _isCellLocked(key: string): boolean {
    return this._internalData.get(key)?.readOnly === true;
  }

  /** Run `validateCell`; returns the rejection message, or null if valid. */
  private _validate(cellId: string, value: string): string | null {
    if (!this.validateCell || value === '') return null;
    try {
      const result = this.validateCell(cellId, value);
      return typeof result === 'string' && result !== '' ? result : null;
    } catch (err) {
      console.error('y11n-spreadsheet: validateCell threw', err);
      return 'Invalid value';
    }
  }

  private _announceSkipped(locked: number, errors: ValidationErrorDetail['errors']): void {
    const parts: string[] = [];
    if (locked > 0) {
      parts.push(locked === 1 ? '1 read-only cell was not changed' : `${locked} read-only cells were not changed`);
    }
    if (errors.length === 1) {
      parts.push(`Invalid value: ${errors[0].message}`);
    } else if (errors.length > 1) {
      parts.push(`${errors.length} values were rejected: ${errors[0].message}`);
    }
    if (parts.length > 0) this._modeAnnouncement = parts.join('. ');
  }

  /**
   * Prepare a user batch for applying: drop read-only cells and values that
   * `validateCell` rejects, then fire the cancelable `before-change` event
   * and apply the host's rewrites. Returns null when nothing is left to apply.
   */
  private _vetUserBatch(batch: CommandBatch, options: { validate?: boolean } = {}): CommandBatch | null {
    const { validate = true } = options;
    // Re-checked here because async operations (paste, cut) await the
    // clipboard between their initial checks and applying.
    if (this.readOnly || this._pasteInProgress) return null;
    // Validators and event listeners may change the grid; if they do, the
    // batch's before-values are stale and it must not be applied.
    const mutations = this._mutationCount;
    const stale = () => {
      if (!this._changedSince(mutations)) return false;
      this._modeAnnouncement = 'Change not applied because the data changed';
      return true;
    };
    let locked = 0;
    const errors: ValidationErrorDetail['errors'] = [];
    let deltas = batch.deltas.filter((d) => {
      if (this._isCellLocked(d.id)) {
        locked++;
        return false;
      }
      if (validate && d.after !== d.before) {
        const message = this._validate(d.id, d.after);
        if (message) {
          errors.push({ cellId: d.id, value: d.after, message });
          return false;
        }
      }
      return true;
    });

    if (errors.length > 0) this._dispatchValidationError({ errors, operation: batch.operation });
    this._announceSkipped(locked, errors);
    if (stale() || deltas.length === 0) return null;

    if (batch.operation !== 'format') {
      const changes: ProposedChange[] = deltas.map((d) => ({
        cellId: d.id,
        value: d.after,
        oldValue: d.before,
      }));
      const detail: BeforeChangeDetail = { changes, operation: batch.operation };
      const event = new CustomEvent<BeforeChangeDetail>('before-change', {
        detail,
        bubbles: true,
        composed: true,
        cancelable: true,
      });
      if (!this.dispatchEvent(event)) {
        this._modeAnnouncement = 'Change cancelled';
        return null;
      }
      if (stale()) return null;
      // Apply the host's edits to the list: rewritten values and removed cells
      const finalValues = new Map<string, string>();
      for (const change of Array.isArray(detail.changes) ? detail.changes : []) {
        if (change && typeof change.cellId === 'string') {
          finalValues.set(change.cellId, String(change.value ?? ''));
        }
      }
      deltas = deltas.flatMap((d) => {
        const after = finalValues.get(d.id);
        if (after === undefined) return [];
        if (after === d.before && !('formatAfter' in d)) return [];
        return [{ ...d, after }];
      });
      if (deltas.length === 0) return null;
    }

    return { ...batch, deltas };
  }

  // ─── Row / Column Insertion and Deletion ────────────

  /**
   * Insert or delete rows/columns as one undoable step. Programmatic calls
   * throw a RangeError for out-of-range arguments; user-initiated ones are
   * pre-checked by the caller and can be vetoed via `before-structure-change`.
   */
  private _changeStructure(change: StructureChange, source: 'user' | 'programmatic'): boolean {
    const { axis, kind, index, count } = change;
    const size = axis === 'row' ? this.rows : this.cols;
    const noun = axis === 'row' ? 'row' : 'column';
    if (!Number.isInteger(index) || !Number.isInteger(count) || index < 0 || count < 1) {
      throw new RangeError(`${kind} ${noun}s: index must be a non-negative integer and count a positive integer`);
    }
    if (kind === 'insert' ? index > size : index + count > size) {
      throw new RangeError(`${kind} ${noun}s: ${index}+${count} is outside the grid (${size} ${noun}s)`);
    }
    if (kind === 'delete' && count >= size) {
      throw new RangeError(`delete ${noun}s: at least one ${noun} must remain`);
    }
    if (this._pasteInProgress) return false;
    // Import a pending data assignment first, so the edit commits into it
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();

    if (source === 'user') {
      if (this.readOnly) return false;
      if (kind === 'delete' && this._hasLockedCellIn(change)) {
        this._modeAnnouncement = `Cannot delete: the ${noun}s contain read-only cells`;
        return false;
      }
      const detail: StructureChangeDetail = { ...change, source };
      const event = new CustomEvent<StructureChangeDetail>('before-structure-change', {
        detail,
        bubbles: true,
        composed: true,
        cancelable: true,
      });
      const mutations = this._mutationCount;
      if (!this.dispatchEvent(event)) return false;
      // A listener changed (or replaced) the grid: the checks above no longer hold
      if (this._changedSince(mutations)) return false;
    }

    const rowsBefore = this.rows;
    const colsBefore = this.cols;
    const delta = kind === 'insert' ? count : -count;
    const rowsAfter = axis === 'row' ? rowsBefore + delta : rowsBefore;
    const colsAfter = axis === 'col' ? colsBefore + delta : colsBefore;
    const selectionBefore = this._snapshotSelection();

    const batch: CommandBatch = {
      deltas: [],
      selectionBefore,
      selectionAfter: this._remapSelection(selectionBefore, rowsBefore, colsBefore, rowsAfter, colsAfter),
      operation: STRUCTURE_OPERATIONS[`${kind}:${axis}`],
      structure: {
        change,
        cells: this._computeStructureCells(change),
        rowsBefore,
        colsBefore,
        rowsAfter,
        colsAfter,
        colWidths: this._computeStructureWidths(change),
      },
    };
    this._commitBatch(batch, source);

    if (source === 'user') {
      const label = (i: number) => (axis === 'row' ? String(i + 1) : colToLetter(i));
      const what =
        count === 1
          ? `${noun} ${label(index)}`
          : `${noun}s ${label(index)} to ${label(index + count - 1)}`;
      this._modeAnnouncement = `${kind === 'insert' ? 'Inserted' : 'Deleted'} ${what}`;
    }
    return true;
  }

  /** Custom column widths after a column insertion or deletion: they move with their columns. */
  private _computeStructureWidths(change: StructureChange): StructureBatch['colWidths'] {
    if (change.axis !== 'col' || this._colWidths.size === 0) return undefined;
    const { kind, index, count } = change;
    const after = new Map<number, number>();
    for (const [col, width] of this._colWidths) {
      if (kind === 'insert') after.set(col >= index ? col + count : col, width);
      else if (col < index) after.set(col, width);
      else if (col >= index + count) after.set(col - count, width);
    }
    return { before: new Map(this._colWidths), after };
  }

  private _hasLockedCellIn(change: StructureChange): boolean {
    for (const [key, cell] of this._internalData) {
      if (!cell.readOnly) continue;
      const { row, col } = parseKey(key);
      const p = change.axis === 'row' ? row : col;
      if (p >= change.index && p < change.index + change.count) return true;
    }
    return false;
  }

  /** Before/after state of every cell a row/column insertion or deletion changes. */
  private _computeStructureCells(change: StructureChange): StructureCellDelta[] {
    const { axis, kind, index, count } = change;
    const next: GridData = new Map();
    for (const [key, cell] of this._internalData) {
      const { row, col } = parseKey(key);
      const p = axis === 'row' ? row : col;
      let moved = p;
      if (kind === 'insert') {
        if (p >= index) moved = p + count;
      } else if (p >= index + count) {
        moved = p - count;
      } else if (p >= index) {
        continue; // deleted
      }
      const copy = cloneCell(cell);
      copy.rawValue = adjustFormulaForStructureChange(cell.rawValue, change);
      next.set(axis === 'row' ? cellKey(moved, col) : cellKey(row, moved), copy);
    }

    const cells: StructureCellDelta[] = [];
    for (const id of new Set([...this._internalData.keys(), ...next.keys()])) {
      const before = this._internalData.get(id);
      const after = next.get(id);
      if (
        before &&
        after &&
        before.rawValue === after.rawValue &&
        formatsEqual(before.format, after.format) &&
        !!before.readOnly === !!after.readOnly
      ) {
        continue;
      }
      cells.push({ id, before: before && cloneCell(before), after });
    }
    return cells;
  }

  /**
   * Selection after a structure change: the same cells by position, clamped
   * to the new size. A selection of whole rows (or columns) stays whole.
   */
  private _remapSelection(
    snapshot: SelectionSnapshot,
    rowsBefore: number,
    colsBefore: number,
    rowsAfter: number,
    colsAfter: number
  ): SelectionSnapshot {
    const minCol = Math.min(snapshot.anchor.col, snapshot.head.col);
    const maxCol = Math.max(snapshot.anchor.col, snapshot.head.col);
    const minRow = Math.min(snapshot.anchor.row, snapshot.head.row);
    const maxRow = Math.max(snapshot.anchor.row, snapshot.head.row);
    const fullWidth = minCol === 0 && maxCol === colsBefore - 1;
    const fullHeight = minRow === 0 && maxRow === rowsBefore - 1;
    const remap = (c: CellCoord): CellCoord => ({
      row: fullHeight && c.row === rowsBefore - 1 ? rowsAfter - 1 : Math.min(c.row, rowsAfter - 1),
      col: fullWidth && c.col === colsBefore - 1 ? colsAfter - 1 : Math.min(c.col, colsAfter - 1),
    });
    return {
      anchor: remap(snapshot.anchor),
      head: remap(snapshot.head),
      active: snapshot.active
        ? { row: Math.min(snapshot.active.row, rowsAfter - 1), col: Math.min(snapshot.active.col, colsAfter - 1) }
        : null,
    };
  }

  private _applyStructureBatch(batch: CommandBatch, side: 'before' | 'after', source: ChangeSource): void {
    const structure = batch.structure!;
    for (const cell of structure.cells) {
      const value = side === 'after' ? cell.after : cell.before;
      if (value) {
        this._internalData.set(cell.id, cloneCell(value));
      } else {
        this._internalData.delete(cell.id);
      }
    }
    this._structureRevision++;
    let movedWidths: ColumnResizeDetail['columns'] = [];
    if (structure.colWidths) {
      const from = side === 'after' ? structure.colWidths.before : structure.colWidths.after;
      const to = side === 'after' ? structure.colWidths.after : structure.colWidths.before;
      this._colWidths = new Map(to);
      this._colWidthsRevision++;
      // Columns whose width changed as widths moved with their columns
      const cols = [...new Set([...from.keys(), ...to.keys()])].sort((a, b) => a - b);
      movedWidths = cols
        .filter((c) => from.get(c) !== to.get(c))
        .map((c) => ({ col: c, width: to.get(c) ?? this._cellSize.width }));
    }
    // The change as applied now: undo applies the inverse
    const { change } = structure;
    const applied: StructureChange =
      side === 'after' ? change : { ...change, kind: change.kind === 'insert' ? 'delete' : 'insert' };
    const frozenBefore = { rows: this._frozenRowCount, cols: this._frozenColCount };
    const frozenNext = this._frozenAfterStructureChange(applied);
    this.rows = side === 'after' ? structure.rowsAfter : structure.rowsBefore;
    this.cols = side === 'after' ? structure.colsAfter : structure.colsBefore;
    // Store the effective counts, so a clamped value can't come back later
    if (applied.axis === 'row') this.frozenRows = Math.min(frozenNext, this.rows - 1);
    else this.frozenCols = Math.min(frozenNext, this.cols - 1);
    if (applied.axis !== 'row') this.frozenRows = this._frozenRowCount;
    else this.frozenCols = this._frozenColCount;
    const frozen =
      this._frozenRowCount !== frozenBefore.rows || this._frozenColCount !== frozenBefore.cols
        ? { rows: this._frozenRowCount, cols: this._frozenColCount }
        : null;
    this._selection.setBounds(this.rows, this.cols);
    this._lastSyncedDataRef = this.data;

    // Every reference may have moved, so rebuild the dependency graph.
    this._formulaEngine.setData(this._internalData);
    this._recalcAll();
    this._restoreSelection(side === 'after' ? batch.selectionAfter : batch.selectionBefore);

    this._dispatchDataChange({
      updates: structure.cells.map((cell) => ({
        id: cell.id,
        value: (side === 'after' ? cell.after : cell.before)?.rawValue ?? '',
      })),
      source,
      operation: batch.operation,
    });
    this._dispatchStructureChange({ ...change, kind: applied.kind, source });
    if (frozen) this._dispatchFreezeChange({ ...frozen, source });
    if (movedWidths.length > 0) {
      // So hosts that save widths on column-resize see them move
      this.dispatchEvent(
        new CustomEvent<ColumnResizeDetail>('column-resize', {
          detail: { columns: movedWidths, source },
          bubbles: true,
          composed: true,
        })
      );
    }
    this.requestUpdate();
  }

  // ─── Sorting ────────────────────────────────────────

  private _sortRange(spec: Omit<SortDetail, 'source'>, source: 'user' | 'programmatic'): boolean {
    if (this._pasteInProgress) return false;
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();
    if (source === 'user' && this.readOnly) return false;

    const { range, hasHeader } = spec;
    const first = range.start.row + (hasHeader ? 1 : 0);
    const order = this._computeSortOrder(first, range.end.row, spec.keys);
    const moved = order.filter((src, i) => src !== first + i);
    if (moved.length === 0) {
      if (source === 'user') this._modeAnnouncement = 'Already sorted';
      return false;
    }

    if (source === 'user') {
      const locked = moved.some((row) => {
        for (let c = range.start.col; c <= range.end.col; c++) {
          if (this._isCellLocked(cellKey(row, c))) return true;
        }
        return false;
      });
      if (locked) {
        this._modeAnnouncement = 'Cannot sort: the rows to move contain read-only cells';
        return false;
      }
      // A copy, so a listener that mutates the detail can't change what is sorted
      const detail: SortDetail = {
        range: { start: { ...range.start }, end: { ...range.end } },
        keys: spec.keys.map((k) => ({ ...k })),
        hasHeader,
        source,
      };
      const event = new CustomEvent<SortDetail>('before-sort', {
        detail,
        bubbles: true,
        composed: true,
        cancelable: true,
      });
      const mutations = this._mutationCount;
      if (!this.dispatchEvent(event)) {
        this._modeAnnouncement = 'Sort cancelled';
        return false;
      }
      if (this._changedSince(mutations)) {
        this._modeAnnouncement = 'Sort not applied because the data changed';
        return false;
      }
    }

    const cells = this._computeSortCells(order, first, range.start.col, range.end.col);
    const selection = this._snapshotSelection();
    const batch: CommandBatch = {
      deltas: [],
      selectionBefore: selection,
      selectionAfter: selection,
      operation: 'sort',
      sort: {
        detail: {
          range: { start: { ...range.start }, end: { ...range.end } },
          keys: spec.keys.map((k) => ({ ...k })),
          hasHeader,
        },
        cells,
      },
    };
    this._commitBatch(batch, source);

    if (source === 'user') {
      const sortedRef = `${coordToRef({ row: first, col: range.start.col })}:${coordToRef(range.end)}`;
      const by = spec.keys.map((k) => `column ${colToLetter(k.col)}, ${k.direction}`).join(', then ');
      this._modeAnnouncement =
        `Sorted ${sortedRef} by ${by}` + (hasHeader ? `. Row ${range.start.row + 1} kept as the header` : '');
    }
    return true;
  }

  /** Source row for each target row `first..last`, in sorted order (stable). */
  private _computeSortOrder(first: number, last: number, keys: SortDetail['keys']): number[] {
    const rows: Array<{ row: number; values: SortValue[] }> = [];
    for (let row = first; row <= last; row++) {
      rows.push({
        row,
        values: keys.map((k) => {
          const key = cellKey(row, k.col);
          return sortValueOf(this._internalData.get(key), (cell) => this._sortNumber(key, cell));
        }),
      });
    }
    // Array.prototype.sort is stable (ES2019)
    rows.sort((a, b) => {
      for (let i = 0; i < keys.length; i++) {
        const result = compareSortValues(a.values[i], b.values[i], keys[i].direction);
        if (result !== 0) return result;
      }
      return 0;
    });
    return rows.map((r) => r.row);
  }

  /** The number behind a number cell, whose display may be formatted. */
  private _sortNumber(key: string, cell: CellData): number {
    if (!cell.format?.numberFormat) return Number(cell.displayValue);
    return Number(this._formulaEngine.resultFor(key, cell.rawValue).displayValue);
  }

  /** Before/after state of every cell in columns `c0..c1` that the sort changes. */
  private _computeSortCells(order: number[], first: number, c0: number, c1: number): StructureCellDelta[] {
    const cells: StructureCellDelta[] = [];
    order.forEach((src, i) => {
      const target = first + i;
      if (src === target) return;
      for (let c = c0; c <= c1; c++) {
        const id = cellKey(target, c);
        const before = this._internalData.get(id);
        const source = this._internalData.get(cellKey(src, c));
        let after: CellData | undefined;
        if (source) {
          after = cloneCell(source);
          after.rawValue = this._shiftFormula(source.rawValue, target - src, 0);
        }
        if (
          before &&
          after &&
          before.rawValue === after.rawValue &&
          formatsEqual(before.format, after.format) &&
          !!before.readOnly === !!after.readOnly
        ) {
          continue;
        }
        if (!before && !after) continue;
        cells.push({ id, before: before && cloneCell(before), after });
      }
    });
    return cells;
  }

  private _applySortBatch(batch: CommandBatch, side: 'before' | 'after', source: ChangeSource): void {
    const sort = batch.sort!;
    for (const cell of sort.cells) {
      const value = side === 'after' ? cell.after : cell.before;
      if (value) {
        this._internalData.set(cell.id, cloneCell(value));
      } else {
        this._internalData.delete(cell.id);
      }
    }
    this._lastSyncedDataRef = this.data;

    const ids = sort.cells.map((cell) => cell.id);
    this._formulaEngine.updateData(this._internalData, ids);
    this._recalcAffected(ids);
    this._applyNumberFormatsToKeys(new Set(ids));
    this._restoreSelection(side === 'after' ? batch.selectionAfter : batch.selectionBefore);

    this._dispatchDataChange({
      updates: sort.cells.map((cell) => ({
        id: cell.id,
        value: (side === 'after' ? cell.after : cell.before)?.rawValue ?? '',
      })),
      source,
      operation: 'sort',
    });
    const { range, keys, hasHeader } = sort.detail;
    this._dispatchSortChange({
      range: { start: { ...range.start }, end: { ...range.end } },
      keys: keys.map((k) => ({ ...k })),
      hasHeader,
      source,
    });
    this.requestUpdate();
  }

  /**
   * The block of data around a cell, like Excel's current region: grown
   * while any cell bordering it (diagonals included) is filled. Null when the
   * block is a single empty cell.
   */
  private _currentRegion(row: number, col: number): SelectionRange | null {
    let r0 = row;
    let r1 = row;
    let c0 = col;
    let c1 = col;
    const rowHasData = (r: number) => {
      for (let c = Math.max(0, c0 - 1); c <= Math.min(this.cols - 1, c1 + 1); c++) {
        if (this._isCellFilled(r, c)) return true;
      }
      return false;
    };
    const colHasData = (c: number) => {
      for (let r = Math.max(0, r0 - 1); r <= Math.min(this.rows - 1, r1 + 1); r++) {
        if (this._isCellFilled(r, c)) return true;
      }
      return false;
    };
    let grew = true;
    while (grew) {
      const before = [r0, r1, c0, c1].join();
      if (r0 > 0 && rowHasData(r0 - 1)) r0--;
      if (r1 < this.rows - 1 && rowHasData(r1 + 1)) r1++;
      if (c0 > 0 && colHasData(c0 - 1)) c0--;
      if (c1 < this.cols - 1 && colHasData(c1 + 1)) c1++;
      grew = [r0, r1, c0, c1].join() !== before;
    }
    if (r0 === r1 && c0 === c1 && !this._isCellFilled(row, col)) return null;
    return { start: { row: r0, col: c0 }, end: { row: r1, col: c1 } };
  }

  /**
   * A selection trimmed to its last filled row and column, so sorting whole
   * columns doesn't move blank rows around. Keeps the active cell's column
   * (the sort key) inside. Null when the selection has no data.
   */
  private _clipToData(range: SelectionRange, active: CellCoord): SelectionRange | null {
    let lastRow = -1;
    let lastCol = -1;
    for (const [key, cell] of this._internalData) {
      if (cell.rawValue === '') continue;
      const { row, col } = parseKey(key);
      if (row < range.start.row || row > range.end.row || col < range.start.col || col > range.end.col) continue;
      if (row > lastRow) lastRow = row;
      if (col > lastCol) lastCol = col;
    }
    if (lastRow < 0) return null;
    return {
      start: { ...range.start },
      end: { row: lastRow, col: Math.max(lastCol, active.col) },
    };
  }

  /**
   * Whether the range's first row looks like column headers: every cell is
   * non-empty text, and either some column has non-text data below it or
   * the first row is bold while the second isn't.
   */
  private _looksLikeHeader(range: SelectionRange): boolean {
    const { start, end } = range;
    if (end.row <= start.row) return false;
    for (let c = start.col; c <= end.col; c++) {
      const cell = this._internalData.get(cellKey(start.row, c));
      if (!cell || cell.rawValue === '' || cell.type !== 'text') return false;
    }
    for (const [key, cell] of this._internalData) {
      if (cell.rawValue === '' || cell.type === 'text') continue;
      const { row, col } = parseKey(key);
      if (row > start.row && row <= end.row && col >= start.col && col <= end.col) return true;
    }
    const bold = (row: number) => {
      for (let c = start.col; c <= end.col; c++) {
        if (!this._internalData.get(cellKey(row, c))?.format?.bold) return false;
      }
      return true;
    };
    return bold(start.row) && !bold(start.row + 1);
  }

  // ─── Frozen Panes ───────────────────────────────────

  /** Frozen rows in effect: the property clamped so at least one row scrolls. */
  private get _frozenRowCount(): number {
    const n = Math.floor(Number(this.frozenRows));
    return Number.isFinite(n) ? Math.max(0, Math.min(n, this.rows - 1)) : 0;
  }

  private get _frozenColCount(): number {
    const n = Math.floor(Number(this.frozenCols));
    return Number.isFinite(n) ? Math.max(0, Math.min(n, this.cols - 1)) : 0;
  }

  /** Height in px of the frozen rows, below the column header. */
  private get _frozenHeight(): number {
    return this._frozenRowCount * this._cellSize.height;
  }

  /** Width in px of the frozen columns, right of the row header. */
  private get _frozenWidth(): number {
    return this._colLeft(this._frozenColCount);
  }

  private _setFrozen(rows: number, cols: number, source: FreezeChangeDetail['source']): boolean {
    const r = Math.max(0, Math.min(rows, this.rows - 1));
    const c = Math.max(0, Math.min(cols, this.cols - 1));
    if (r === this._frozenRowCount && c === this._frozenColCount) {
      // Normalize an out-of-range property value without an event
      this.frozenRows = r;
      this.frozenCols = c;
      return false;
    }
    this.frozenRows = r;
    this.frozenCols = c;
    // The frozen rows/columns now cover the start of the viewport, which may
    // hide the active cell: scroll it clear (without moving focus), after the
    // render applies the new scroll padding its snap position depends on
    this.updateComplete.then(() => {
      const { row, col } = this._selection.activeCell;
      this._ensureCellVisible(row, col);
    });
    this._dispatchFreezeChange({ rows: r, cols: c, source });
    return true;
  }

  private _dispatchFreezeChange(detail: FreezeChangeDetail): void {
    this.dispatchEvent(new CustomEvent<FreezeChangeDetail>('freeze-change', { detail, bubbles: true, composed: true }));
  }

  /** Alt+Shift+F: unfreeze, or freeze the rows above and columns left of the active cell. */
  private _toggleFreeze(): void {
    if (this._frozenRowCount > 0 || this._frozenColCount > 0) {
      this._setFrozen(0, 0, 'user');
      this._modeAnnouncement = 'Panes unfrozen';
      return;
    }
    const { row, col } = this._selection.activeCell;
    if (row === 0 && col === 0) {
      this._modeAnnouncement = 'Nothing to freeze: move below or right of the rows and columns to keep in view';
      return;
    }
    const gridEl = this._grid;
    const h = this._cellSize.height;
    // Leave at least two rows (one column) of scrolling area, when the grid has a size
    if (
      gridEl?.clientHeight &&
      (row * h > gridEl.clientHeight - 3 * h ||
        this._colLeft(col) > gridEl.clientWidth - Y11nSpreadsheet.ROW_HEADER_WIDTH - this._colWidth(col))
    ) {
      this._modeAnnouncement = 'Cannot freeze: the frozen area would not fit on screen';
      return;
    }
    this._setFrozen(row, col, 'user');
    const parts: string[] = [];
    if (row > 0) parts.push(row === 1 ? 'row 1' : `rows 1 to ${row}`);
    if (col > 0) parts.push(col === 1 ? 'column A' : `columns A to ${colToLetter(col - 1)}`);
    this._modeAnnouncement = `Froze ${parts.join(' and ')}`;
  }

  /**
   * Frozen count on `change.axis` after a row/column insertion or deletion
   * (the freeze line follows the content, as in Excel), from the current
   * effective count. Undo/redo pass the inverse change, so a freeze made
   * after the original change is respected.
   */
  private _frozenAfterStructureChange(change: StructureChange): number {
    const frozen = change.axis === 'row' ? this._frozenRowCount : this._frozenColCount;
    if (change.kind === 'insert') return change.index < frozen ? frozen + change.count : frozen;
    const overlap = Math.max(0, Math.min(frozen, change.index + change.count) - change.index);
    return frozen - overlap;
  }

  /**
   * The grid's scroll padding: the sticky header row and frozen rows at the
   * top, the row header and frozen columns at the left. Scroll snapping then
   * aligns a row just below the frozen area (its snap position is exactly
   * where `_ensureCellVisible` scrolls a row to), and the browser's own
   * focus scrolling keeps cells out from under the sticky parts.
   */
  private _scrollPadding(): string {
    return (
      `scroll-padding-top: calc(var(--_cell-height) * ${1 + this._frozenRowCount}); ` +
      `scroll-padding-left: ${this._stickyLeft(this._frozenColCount)};`
    );
  }

  /** CSS for a frozen row's sticky top offset: below the header and the frozen rows above it. */
  private _stickyTop(row: number): string {
    return `calc(var(--_cell-height) * ${row + 1})`;
  }

  /** Classes marking a cell or header as frozen (pass -1 for a header's other axis). */
  private _frozenClasses(row: number, col: number): string {
    const fr = this._frozenRowCount;
    const fc = this._frozenColCount;
    const classes: string[] = [];
    // Headers pass -1 for the other axis, which must never match an edge
    if (row >= 0 && row < fr) classes.push('frozen-row');
    if (row >= 0 && row === fr - 1) classes.push('frozen-row-edge');
    if (col >= 0 && col < fc) classes.push('frozen-col');
    if (col >= 0 && col === fc - 1) classes.push('frozen-col-edge');
    return classes.join(' ');
  }

  /**
   * CSS for a frozen column's sticky left offset: the row header plus the
   * widths of the columns before it, with default-width columns kept as
   * `var(--_cell-width)` so the offset follows the CSS variable exactly.
   */
  private _stickyLeft(col: number): string {
    let px = 0;
    let defaults = 0;
    for (let c = 0; c < col; c++) {
      const custom = this._resizePreview?.cols.includes(c) ? this._resizePreview.width : this._colWidths.get(c);
      if (custom === undefined) defaults++;
      else px += custom;
    }
    return `calc(var(--_row-header-width) + ${px}px + ${defaults} * var(--_cell-width))`;
  }

  // ─── Column Widths ──────────────────────────────────

  private _colWidth(col: number): number {
    if (this._resizePreview?.cols.includes(col)) return this._resizePreview.width;
    return this._colWidths.get(col) ?? this._cellSize.width;
  }

  /** Left offset of every column within the data area, plus the total width at `[cols]`. */
  private _colOffsets(): Float64Array {
    const key = `${this.cols}|${this._cellSize.width}|${this._colWidthsRevision}`;
    if (this._colLayout?.key !== key) {
      const offsets = new Float64Array(this.cols + 1);
      for (let c = 0; c < this.cols; c++) offsets[c + 1] = offsets[c] + this._colWidth(c);
      this._colLayout = { key, offsets };
    }
    return this._colLayout.offsets;
  }

  private _colLeft(col: number): number {
    return this._colOffsets()[Math.max(0, Math.min(col, this.cols))];
  }

  /** The column at `x` px into the data area (clamped to the grid). */
  private _colAt(x: number): number {
    const offsets = this._colOffsets();
    let lo = 0;
    let hi = this.cols - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (offsets[mid] <= x) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }

  /** `grid-template-columns`, with runs of default-width columns kept as `repeat()`. */
  private _gridTemplateColumns(): string {
    const parts = ['var(--_row-header-width)'];
    let run = 0;
    const flush = () => {
      if (run > 0) parts.push(run === 1 ? 'var(--_cell-width)' : `repeat(${run}, var(--_cell-width))`);
      run = 0;
    };
    for (let c = 0; c < this.cols; c++) {
      const custom = this._resizePreview?.cols.includes(c) ? this._resizePreview.width : this._colWidths.get(c);
      if (custom === undefined) {
        run++;
      } else {
        flush();
        parts.push(`${custom}px`);
      }
    }
    flush();
    return parts.join(' ');
  }

  private _clampWidth(width: number): number {
    return Math.round(Math.min(Y11nSpreadsheet.MAX_COL_WIDTH, Math.max(Y11nSpreadsheet.MIN_COL_WIDTH, width)));
  }

  /**
   * Resize columns as one undoable step. User resizes (allowed in read-only
   * mode: widths are presentation, not data) can be vetoed via
   * `before-column-resize`. Returns whether anything changed.
   */
  private _resizeColumns(
    changes: Array<{ col: number; width: number | null }>,
    source: 'user' | 'programmatic'
  ): 'applied' | 'unchanged' | 'cancelled' {
    const deltas: WidthDelta[] = [];
    const seen = new Set<number>();
    for (const { col, width } of changes) {
      if (seen.has(col)) continue;
      seen.add(col);
      const before = this._colWidths.get(col) ?? null;
      const after = width === null ? null : this._clampWidth(width);
      if (after !== before) deltas.push({ col, before, after });
    }
    if (deltas.length === 0) return 'unchanged';

    if (source === 'user') {
      const detail: ColumnResizeDetail = { columns: this._resizedColumns(deltas, 'after'), source };
      const event = new CustomEvent<ColumnResizeDetail>('before-column-resize', {
        detail,
        bubbles: true,
        composed: true,
        cancelable: true,
      });
      if (!this.dispatchEvent(event)) return 'cancelled';
    }

    const selection = this._snapshotSelection();
    this._commitBatch(
      // `operation` is never reported for width batches: they fire column-resize, not data-change
      { deltas: [], selectionBefore: selection, selectionAfter: selection, operation: 'format', widths: deltas },
      source
    );
    return 'applied';
  }

  /** Effective widths of the columns in `deltas` on one side of the change. */
  private _resizedColumns(deltas: WidthDelta[], side: 'before' | 'after'): ColumnResizeDetail['columns'] {
    return deltas.map((d) => ({ col: d.col, width: d[side] ?? this._cellSize.width }));
  }

  private _applyWidthBatch(batch: CommandBatch, side: 'before' | 'after', source: ChangeSource): void {
    const deltas = batch.widths!;
    const { row, col } = this._selection.activeCell;
    const activeInView = this._isCellInView(row, col);
    for (const delta of deltas) {
      const width = delta[side];
      if (width === null) this._colWidths.delete(delta.col);
      else this._colWidths.set(delta.col, width);
    }
    this._colWidthsRevision++;
    // The selection is unchanged, so don't restore it (which would move focus,
    // e.g. when a page restores saved widths). If the active cell was in view,
    // keep it there: widening its column can push it off-screen. Scroll after
    // the render, once the grid has its new width (requested first, so
    // updateComplete waits for that render).
    this.requestUpdate();
    this.updateComplete.then(() => {
      if (activeInView && source !== 'programmatic') this._ensureCellVisible(row, col);
      // Narrower columns can shrink the scroll range under the current position
      const gridEl = this._grid;
      if (gridEl && gridEl.scrollLeft !== this._scrollLeft) this._scrollLeft = gridEl.scrollLeft;
    });
    this.dispatchEvent(
      new CustomEvent<ColumnResizeDetail>('column-resize', {
        detail: { columns: this._resizedColumns(deltas, side), source },
        bubbles: true,
        composed: true,
      })
    );
  }

  /** Whether any part of a cell is inside the scrolled data viewport. */
  private _isCellInView(row: number, col: number): boolean {
    const gridEl = this._grid;
    if (!gridEl) return false;
    const left = this._colLeft(col);
    const right = left + this._colWidth(col);
    const top = row * this._cellSize.height;
    const viewWidth = gridEl.clientWidth - Y11nSpreadsheet.ROW_HEADER_WIDTH;
    const viewHeight = gridEl.clientHeight - this._cellSize.height;
    // Frozen cells are always in view; the others scroll beside the frozen area
    const inX =
      col < this._frozenColCount ||
      (right > this._scrollLeft + this._frozenWidth && left < this._scrollLeft + viewWidth);
    const inY =
      row < this._frozenRowCount ||
      (top + this._cellSize.height > this._scrollTop + this._frozenHeight && top < this._scrollTop + viewHeight);
    return inX && inY;
  }

  /** The columns a resize of `col` applies to: all selected whole columns, if `col` is one of them. */
  private _columnsToResize(col: number): number[] {
    const { start, end } = this._selection.range;
    const wholeColumns = start.row === 0 && end.row === this.rows - 1;
    if (wholeColumns && col >= start.col && col <= end.col) {
      return Array.from({ length: end.col - start.col + 1 }, (_, i) => start.col + i);
    }
    return [col];
  }

  /** Alt+Shift+Left/Right: resize the selected columns from the active column's width. */
  private _resizeSelectedColumns(delta: number): void {
    const { start, end } = this._selection.range;
    const active = this._selection.activeCell.col;
    const width = this._clampWidth(this._colWidth(active) + delta);
    const cols = Array.from({ length: end.col - start.col + 1 }, (_, i) => start.col + i);
    this._announceWidths(cols, this._resizeColumns(cols.map((col) => ({ col, width })), 'user'));
  }

  /** Alt+Shift+0 and header-border double-click: fit columns to their content. */
  private _autoFitUserColumns(cols: number[]): void {
    this._announceWidths(cols, this._resizeColumns(cols.map((col) => ({ col, width: this._fitWidth(col) })), 'user'));
  }

  /** Announce the widths after a user resize (also when they didn't change, e.g. at the minimum). */
  private _announceWidths(cols: number[], result: ReturnType<Y11nSpreadsheet['_resizeColumns']>): void {
    if (result === 'cancelled') {
      this._modeAnnouncement = 'Resize cancelled';
      return;
    }
    const first = colToLetter(cols[0]);
    const last = colToLetter(cols[cols.length - 1]);
    const widths = new Set(cols.map((c) => Math.round(this._colWidth(c))));
    const what = cols.length === 1 ? `Column ${first}` : `Columns ${first} to ${last}`;
    this._modeAnnouncement =
      widths.size === 1 ? `${what} width ${[...widths][0]} pixels` : `${what} resized to fit their content`;
  }

  /**
   * The width that fits a column's widest displayed value, or null (the
   * default width) for an empty column. Text is measured on a canvas, so
   * rows outside the rendered window count too.
   */
  private _fitWidth(col: number): number | null {
    const host = getComputedStyle(this);
    const family = host.fontFamily || 'sans-serif';
    const baseSize = parseFloat(host.fontSize) || 13;
    if (this._measureCanvas === undefined) {
      try {
        this._measureCanvas = document.createElement('canvas').getContext('2d');
      } catch {
        this._measureCanvas = null;
      }
    }
    const ctx = this._measureCanvas;
    let widest = 0;
    for (const [key, cell] of this._internalData) {
      if (!cell.displayValue) continue;
      const pos = parseKey(key);
      if (pos.col !== col || pos.row >= this.rows) continue;
      const fmt = cell.format;
      const size = fmt?.fontSize ?? baseSize;
      let width: number;
      if (ctx) {
        ctx.font = `${fmt?.italic ? 'italic ' : ''}${fmt?.bold ? 'bold ' : ''}${size}px ${family}`;
        width = ctx.measureText(cell.displayValue).width;
      } else {
        width = cell.displayValue.length * size * 0.6; // no canvas (e.g. a test DOM): estimate
      }
      widest = Math.max(widest, width);
    }
    if (widest === 0) return null;
    // Cell padding (2 x 4px), right border, and a little room so text isn't cut off
    return this._clampWidth(Math.ceil(widest) + 15);
  }

  /** Pointer drag on a column header's right border: resize with a live preview. */
  private _handleResizerPointerDown(e: PointerEvent): void {
    if (e.button !== 0) return;
    const handle = e.currentTarget as HTMLElement;
    const col = parseInt(handle.dataset.resizeCol!, 10);
    e.preventDefault();
    e.stopPropagation();
    if (this._isEditing) this._commitEdit();

    const cols = this._columnsToResize(col);
    const startX = e.clientX;
    const startWidth = this._colWidth(col);
    const controller = new AbortController();
    const { signal } = controller;
    const setPreview = (preview: typeof this._resizePreview) => {
      this._resizePreview = preview;
      this._colWidthsRevision++;
      this.requestUpdate();
    };

    const onMove = (ev: PointerEvent) => {
      setPreview({ cols, width: this._clampWidth(startWidth + ev.clientX - startX) });
    };
    const finish = (commit: boolean) => {
      controller.abort();
      const width = this._resizePreview?.width;
      setPreview(null);
      if (commit && width !== undefined) {
        this._announceWidths(cols, this._resizeColumns(cols.map((c) => ({ col: c, width })), 'user'));
      }
    };
    document.addEventListener('pointermove', onMove, { signal });
    document.addEventListener('pointerup', () => finish(true), { signal });
    document.addEventListener('pointercancel', () => finish(false), { signal });
    // Escape cancels the drag, restoring the width
    document.addEventListener(
      'keydown',
      (ev) => {
        if (ev.key !== 'Escape') return;
        ev.preventDefault();
        ev.stopPropagation();
        finish(false);
      },
      { signal, capture: true }
    );
    window.addEventListener('blur', () => finish(false), { signal });
  }

  private _handleResizerDblClick(e: MouseEvent): void {
    e.preventDefault();
    e.stopPropagation();
    const col = parseInt((e.currentTarget as HTMLElement).dataset.resizeCol!, 10);
    this._autoFitUserColumns(this._columnsToResize(col));
  }

  private _undo(): void {
    if (this.readOnly || this._pasteInProgress) return;

    const batch = this._undoStack.pop();
    if (!batch) return;

    // Recorded before reverting, for the same reason as in _commitBatch: a
    // change a listener makes while it reverts clears the redo stack
    this._redoStack.push(batch);
    this._revertBatch(batch);
    this._dispatchFormatChangeForBatch(batch, 'undo');
  }

  private _redo(): void {
    if (this.readOnly || this._pasteInProgress) return;

    const batch = this._redoStack.pop();
    if (!batch) return;

    // Recorded before applying (see _commitBatch)
    this._undoStack.push(batch);
    if (this._undoStack.length > this._maxHistory) {
      this._undoStack.shift();
    }
    this._applyBatch(batch, 'redo');
    this._dispatchFormatChangeForBatch(batch, 'redo');
  }

  /** Dispatch format-change for any batch that contains format deltas.
   *  NOTE: The event carries the format from the first delta only. For heterogeneous
   *  batches (e.g., paste with varied per-cell formatting), this is an approximation —
   *  the cellIds are accurate but the format may not represent every cell. Consumers
   *  that need per-cell accuracy should read cell data directly. */
  private _dispatchFormatChangeForBatch(batch: CommandBatch, source: 'undo' | 'redo'): void {
    const formatDeltas = batch.deltas.filter(
      (d) => d.formatBefore !== undefined || d.formatAfter !== undefined
    );
    if (formatDeltas.length === 0) return;

    const cellIds = formatDeltas.map((d) => d.id);
    // For undo, the applied format is formatBefore; for redo, formatAfter
    const format: CellFormat = source === 'undo'
      ? (formatDeltas[0].formatBefore ?? {})
      : (formatDeltas[0].formatAfter ?? {});
    this._dispatchFormatChange({ cellIds, format, source });
  }

  // ─── Editing ────────────────────────────────────────

  private _startEditing(initialValue?: string): void {
    if (this.readOnly || this._isEditing || this._pasteInProgress) return;

    const { row, col } = this._selection.activeCell;
    const key = cellKey(row, col);
    if (this._isCellLocked(key)) {
      this._modeAnnouncement = `${coordToRef({ row, col })} is read-only`;
      return;
    }

    const cell = this._internalData.get(key);
    const detail: CellEditStartDetail = { cellId: key, value: cell?.rawValue ?? '' };
    const event = new CustomEvent<CellEditStartDetail>('cell-edit-start', {
      detail,
      bubbles: true,
      composed: true,
      cancelable: true,
    });
    if (!this.dispatchEvent(event)) return;

    this._editingCellKey = key;
    this._editValue = initialValue ?? cell?.rawValue ?? '';
    this._validationMessage = '';
    this._isEditing = true;
    this._editorNeedsCursorReset = true;
    this._modeAnnouncement = `Editing ${coordToRef(this._selection.activeCell)}`;
  }

  /**
   * Commit the inline edit. Returns false if `validateCell` rejected the
   * value: with `onInvalid: 'keep'` (keyboard commits) the editor stays open
   * and shows the message; with `'cancel'` (focus moving away) the edit is
   * discarded and the message announced.
   */
  private _commitEdit(onInvalid: 'keep' | 'cancel' = 'cancel'): boolean {
    if (!this._isEditing || !this._editingCellKey) return true;

    const key = this._editingCellKey;
    const newValue = this._editValue;
    if (!this._checkEditValue(key, newValue, onInvalid)) return false;

    this._isEditing = false;
    this._editingCellKey = null;
    this._validationMessage = '';
    this._modeAnnouncement = 'Edit committed';

    this._commitRawValue(key, newValue, { validate: false });

    // Return focus to the grid
    this.updateComplete.then(() => this._focusActiveCell());
    return true;
  }

  /**
   * Validate the value being edited in `key`. On rejection, fires
   * `validation-error` and either keeps the editor open with the message or
   * cancels the edit. Unchanged values are not validated.
   */
  private _checkEditValue(key: string, value: string, onInvalid: 'keep' | 'cancel'): boolean {
    if (value === (this._internalData.get(key)?.rawValue ?? '')) return true;
    const message = this._validate(key, value);
    if (!message) return true;
    this._dispatchValidationError({ errors: [{ cellId: key, value, message }], operation: 'edit' });
    if (onInvalid === 'keep') {
      this._validationMessage = message;
      this._modeAnnouncement = `Invalid value: ${message}`;
    } else {
      this._cancelEdit(`Invalid value: ${message}. Edit cancelled`);
    }
    return false;
  }

  private _cancelEdit(announcement = 'Edit cancelled'): void {
    this._isEditing = false;
    this._editingCellKey = null;
    this._validationMessage = '';
    this._modeAnnouncement = announcement;
    this.updateComplete.then(() => this._focusActiveCell());
  }

  // ─── Reference Selection ────────────────────────────

  /**
   * Check if the cursor is at a position in the formula where a cell reference
   * can be inserted (after =, +, -, *, /, &, (, ,, or comparison operators).
   */
  private _isRefPositionInFormula(): boolean {
    if (!this._editValue.startsWith('=')) return false;

    const cursorPos = this._editor?.selectionStart ?? this._editValue.length;
    const before = this._editValue.substring(0, cursorPos).trimEnd();
    if (before === '=') return true;

    const lastChar = before[before.length - 1];
    return ['+', '-', '*', '/', '&', '(', ',', ':', '<', '>', '='].includes(lastChar);
  }

  /**
   * Handle an arrow key press during formula editing to insert or move
   * a cell reference.
   */
  private _handleRefArrow(key: string): void {
    const dRow = key === 'ArrowDown' ? 1 : key === 'ArrowUp' ? -1 : 0;
    const dCol = key === 'ArrowRight' ? 1 : key === 'ArrowLeft' ? -1 : 0;

    if (!this._refMode) {
      // Enter reference mode - start from the editing cell and move
      const { row, col } = this._selection.activeCell;
      const clamped = this._selection.clamp({ row: row + dRow, col: col + dCol });
      this._refCursorRow = clamped.row;
      this._refCursorCol = clamped.col;
      this._refMode = true;
      this._modeAnnouncement = `Reference selection mode. Selecting ${coordToRef({ row: this._refCursorRow, col: this._refCursorCol })}`;

      const cursorPos = this._editor?.selectionStart ?? this._editValue.length;
      this._refInsertStart = cursorPos;
      const ref = coordToRef({ row: this._refCursorRow, col: this._refCursorCol });
      this._editValue =
        this._editValue.substring(0, cursorPos) +
        ref +
        this._editValue.substring(cursorPos);
      this._refInsertEnd = cursorPos + ref.length;
    } else {
      // Already in reference mode - move the reference cursor
      const moved = this._selection.clamp({ row: this._refCursorRow + dRow, col: this._refCursorCol + dCol });
      this._refCursorRow = moved.row;
      this._refCursorCol = moved.col;

      const ref = coordToRef({ row: this._refCursorRow, col: this._refCursorCol });
      this._modeAnnouncement = `Selecting ${ref}`;
      this._editValue =
        this._editValue.substring(0, this._refInsertStart) +
        ref +
        this._editValue.substring(this._refInsertEnd);
      this._refInsertEnd = this._refInsertStart + ref.length;
    }

    this._ensureCellVisible(this._refCursorRow, this._refCursorCol);
  }

  private _exitRefMode(): void {
    if (this._refMode) {
      this._modeAnnouncement = 'Reference selection ended';
    }
    this._refMode = false;
  }

  /**
   * Cycle the cell reference at/before the cursor through reference modes:
   * A1 → $A$1 → A$1 → $A1 → A1
   */
  private _cycleReferenceMode(): void {
    const cursorPos = this._editor?.selectionStart ?? this._editValue.length;
    const text = this._editValue;

    // Find the reference at or just before the cursor
    const refPattern = /(?<![A-Z0-9_.$])(\$?)([A-Z]+)(\$?)(\d+)(?![A-Z0-9_.(])/gi;
    let match: RegExpExecArray | null;
    let bestMatch: RegExpExecArray | null = null;

    while ((match = refPattern.exec(text)) !== null) {
      const start = match.index;
      const end = start + match[0].length;
      // The ref contains or is immediately before the cursor
      if (start <= cursorPos && end >= cursorPos - 1) {
        bestMatch = match;
      }
    }

    if (!bestMatch) return;

    const [full, colDollar, colLetters, rowDollar, rowDigits] = bestMatch;
    const start = bestMatch.index;
    const hasColDollar = colDollar === '$';
    const hasRowDollar = rowDollar === '$';

    let newRef: string;
    if (!hasColDollar && !hasRowDollar) {
      // A1 → $A$1
      newRef = `$${colLetters}$${rowDigits}`;
    } else if (hasColDollar && hasRowDollar) {
      // $A$1 → A$1
      newRef = `${colLetters}$${rowDigits}`;
    } else if (!hasColDollar && hasRowDollar) {
      // A$1 → $A1
      newRef = `$${colLetters}${rowDigits}`;
    } else {
      // $A1 → A1
      newRef = `${colLetters}${rowDigits}`;
    }

    this._editValue = text.substring(0, start) + newRef + text.substring(start + full.length);

    // Adjust cursor position to stay at the end of the reference
    const newCursorPos = start + newRef.length;
    this.updateComplete.then(() => {
      if (this._editor) {
        this._editor.setSelectionRange(newCursorPos, newCursorPos);
      }
    });
  }

  // ─── Focus Management ───────────────────────────────

  private _focusActiveCell(): void {
    const { row, col } = this._selection.activeCell;
    this._ensureCellVisible(row, col);
    const cellEl = this.shadowRoot?.querySelector(
      `[data-row="${row}"][data-col="${col}"]`
    ) as HTMLElement | null;

    if (cellEl) {
      cellEl.focus();
    }
  }

  // ─── Keyboard Handling ──────────────────────────────

  private _handleGridKeydown(e: KeyboardEvent): void {
    if (this._isEditing) {
      this._handleEditorKeydown(e);
      return;
    }

    const shift = e.shiftKey;
    const ctrl = e.ctrlKey || e.metaKey;

    // Ctrl+Space selects columns, Shift+Space selects rows, both select all.
    // Checked before the printable-character branch, which would otherwise
    // treat Shift+Space as typing a space.
    if (e.key === ' ' && (ctrl || shift) && !e.altKey) {
      e.preventDefault();
      e.stopPropagation();
      this._selectWholeRowsOrColumns(ctrl, shift);
      return;
    }

    // Only stop propagation for keys we actually handle, so unhandled keys
    // (e.g. Tab) propagate normally and focus can leave the grid.
    // Start editing on printable character input. This must run before the
    // shortcut switch so plain letters like "a" or "c" aren't swallowed by the
    // Ctrl+letter cases. Count code points so emoji also start editing.
    if (!ctrl && !e.altKey && [...e.key].length === 1) {
      if (!this.readOnly) {
        e.preventDefault();
        e.stopPropagation();
        this._startEditing(e.key);
      }
      return;
    }

    // Insert/delete rows or columns: Ctrl+Shift+= / Ctrl+Alt+= / Ctrl+NumpadAdd
    // and Ctrl+- / Ctrl+Alt+- / Ctrl+NumpadSubtract.
    if (ctrl && this._handleStructureShortcut(e)) return;

    // Ctrl+Alt+D / Ctrl+Alt+R: fill series down / right (the fill handle's
    // keyboard equivalent). Matched by e.code because Option changes e.key on macOS.
    if (ctrl && e.altKey && !shift && (e.code === 'KeyD' || e.code === 'KeyR')) {
      e.preventDefault();
      e.stopPropagation();
      this._fillSeriesInSelection(e.code === 'KeyD' ? 'down' : 'right');
      return;
    }

    // Alt+Shift+F: freeze the rows above and columns left of the active cell,
    // or unfreeze. A view setting, so allowed in read-only mode.
    if (e.altKey && shift && !ctrl && e.code === 'KeyF') {
      e.preventDefault();
      e.stopPropagation();
      this._toggleFreeze();
      return;
    }

    // Alt+Shift+Left / Alt+Shift+Right: narrow / widen the selected columns.
    // Alt+Shift+0: fit them to their content. Allowed in read-only mode.
    if (e.altKey && shift && !ctrl && (e.key === 'ArrowLeft' || e.key === 'ArrowRight' || e.code === 'Digit0')) {
      e.preventDefault();
      e.stopPropagation();
      if (e.code === 'Digit0') {
        const { start, end } = this._selection.range;
        this._autoFitUserColumns(Array.from({ length: end.col - start.col + 1 }, (_, i) => start.col + i));
      } else {
        const step = Y11nSpreadsheet.COL_RESIZE_STEP;
        this._resizeSelectedColumns(e.key === 'ArrowLeft' ? -step : step);
      }
      return;
    }

    // Alt+Shift+Up / Alt+Shift+Down: sort ascending / descending by the
    // active cell's column
    if (e.altKey && shift && !ctrl && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
      e.preventDefault();
      e.stopPropagation();
      this.sortSelection(e.key === 'ArrowUp' ? 'ascending' : 'descending');
      return;
    }

    let handled = true;

    // Normalize letter keys so shortcuts work with Caps Lock / Shift held.
    const key = ctrl && e.key.length === 1 ? e.key.toLowerCase() : e.key;

    switch (key) {
      case 'ArrowUp':
      case 'ArrowDown':
      case 'ArrowLeft':
      case 'ArrowRight': {
        e.preventDefault();
        const [dRow, dCol] = ARROW_DELTAS[key];
        if (ctrl) {
          // Jump to the edge of the data region (Ctrl/Cmd+Arrow); with Shift,
          // extend the selection from its head to that edge.
          const from = shift ? this._selection.head : this._selection.activeCell;
          const target = findDataEdge(from, dRow, dCol, this.rows, this.cols, (r, c) =>
            this._isCellFilled(r, c)
          );
          this._selection.moveTo(target.row, target.col, shift);
        } else {
          this._selection.move(dRow, dCol, shift);
        }
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;
      }

      case 'Enter':
        e.preventDefault();
        if (!this.readOnly) {
          this._startEditing();
        }
        break;

      case 'F2':
        e.preventDefault();
        if (!this.readOnly) {
          this._startEditing();
        }
        break;

      case 'Escape':
        e.preventDefault();
        this._selection.clearSelection();
        this._dispatchSelectionChange();
        break;

      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        if (!this.readOnly) {
          this._clearSelectedCells();
        }
        break;

      case 'Home':
        e.preventDefault();
        if (ctrl) {
          // With frozen panes, the top-left of the scrolling area (as in Excel)
          this._selection.moveTo(this._frozenRowCount, this._frozenColCount, shift);
        } else {
          this._selection.moveTo(this._selection.activeCell.row, 0, shift);
        }
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'End':
        e.preventDefault();
        if (ctrl) {
          this._selection.moveTo(this.rows - 1, this.cols - 1, shift);
        } else {
          this._selection.moveTo(this._selection.activeCell.row, this.cols - 1, shift);
        }
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'PageDown': {
        e.preventDefault();
        const cellHeight = this._cellSize.height;
        const gridEl = this._grid;
        const viewHeight = gridEl?.clientHeight || cellHeight * Y11nSpreadsheet.INITIAL_VIEW_ROWS;
        const pageRows = Math.max(1, Math.floor((viewHeight - this._frozenHeight) / cellHeight) - 1);
        this._selection.move(pageRows, 0, shift);
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;
      }

      case 'PageUp': {
        e.preventDefault();
        const cellHeight = this._cellSize.height;
        const gridEl = this._grid;
        const viewHeight = gridEl?.clientHeight || cellHeight * Y11nSpreadsheet.INITIAL_VIEW_ROWS;
        const pageRows = Math.max(1, Math.floor((viewHeight - this._frozenHeight) / cellHeight) - 1);
        this._selection.move(-pageRows, 0, shift);
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;
      }

      case 'a':
        if (ctrl) {
          e.preventDefault();
          this._selection.selectAll();
          this._dispatchSelectionChange();
        } else {
          handled = false;
        }
        break;

      case 'c':
        if (ctrl) {
          e.preventDefault();
          this._handleCopy();
        } else {
          handled = false;
        }
        break;

      case 'x':
        if (ctrl) {
          e.preventDefault();
          this._handleCut();
        } else {
          handled = false;
        }
        break;

      case 'v':
        if (ctrl) {
          e.preventDefault();
          this._handlePaste();
        } else {
          handled = false;
        }
        break;

      case 'b':
        if (ctrl) {
          e.preventDefault();
          this.toggleFormat('bold');
        } else {
          handled = false;
        }
        break;

      case 'i':
        if (ctrl) {
          e.preventDefault();
          this.toggleFormat('italic');
        } else {
          handled = false;
        }
        break;

      case 'u':
        if (ctrl) {
          e.preventDefault();
          this.toggleFormat('underline');
        } else {
          handled = false;
        }
        break;

      case 'f':
        if (ctrl && !shift) {
          e.preventDefault();
          void this.openFind();
        } else {
          handled = false;
        }
        break;

      case 'd':
        if (ctrl && !shift && !e.altKey) {
          e.preventDefault();
          this._fillSelection('down');
        } else {
          handled = false;
        }
        break;

      case 'r':
        if (ctrl && !shift && !e.altKey) {
          e.preventDefault();
          this._fillSelection('right');
        } else {
          handled = false;
        }
        break;

      case 'y':
        if (ctrl) {
          e.preventDefault();
          this._redo();
        } else {
          handled = false;
        }
        break;

      case 'z':
        if (ctrl) {
          e.preventDefault();
          if (shift) {
            this._redo();
          } else {
            this._undo();
          }
        } else {
          handled = false;
        }
        break;

      default:
        handled = false;
        break;
    }

    if (handled) {
      e.stopPropagation();
    }
  }

  private _handleEditorKeydown(e: KeyboardEvent): void {
    e.stopPropagation();

    switch (e.key) {
      case 'Enter':
        e.preventDefault();
        this._exitRefMode();
        if (e.ctrlKey || e.metaKey) {
          // Ctrl/Cmd+Enter: write the value into every selected cell, stay put
          this._commitEditToSelection();
          break;
        }
        if (!this._commitEdit('keep')) break;
        // Enter moves down after committing, Shift+Enter moves up
        this._selection.move(e.shiftKey ? -1 : 1, 0);
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'Tab':
        e.preventDefault();
        this._exitRefMode();
        if (!this._commitEdit('keep')) break;
        if (e.shiftKey) {
          this._selection.move(0, -1);
        } else {
          this._selection.move(0, 1);
        }
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'Escape':
        e.preventDefault();
        this._exitRefMode();
        this._cancelEdit();
        break;

      case 'ArrowUp':
      case 'ArrowDown':
      case 'ArrowLeft':
      case 'ArrowRight':
        if (this._refMode || this._isRefPositionInFormula()) {
          e.preventDefault();
          this._handleRefArrow(e.key);
        }
        break;

      case 'f':
      case 'F':
        if (e.ctrlKey || e.metaKey) {
          e.preventDefault();
          this._exitRefMode();
          void this.openFind();
        }
        break;

      case 'F4':
        if (this._editValue.startsWith('=')) {
          e.preventDefault();
          this._cycleReferenceMode();
        }
        break;
    }
  }

  /**
   * Keyboard insertion/deletion of rows or columns. Acts on whole-row or
   * whole-column selections (Shift+Space / Ctrl+Space); with any other
   * selection it announces how to select rows or columns first.
   *
   * @returns true if the key was an insert/delete shortcut
   */
  private _handleStructureShortcut(e: KeyboardEvent): boolean {
    const insert =
      e.code === 'NumpadAdd' || e.key === '+' || ((e.shiftKey || e.altKey) && e.code === 'Equal');
    const remove =
      !e.shiftKey && (e.code === 'NumpadSubtract' || e.code === 'Minus' || e.key === '-');
    if (!insert && !remove) return false;
    e.preventDefault();
    e.stopPropagation();

    if (this.readOnly) {
      this._modeAnnouncement = 'The spreadsheet is read-only';
      return true;
    }
    const { start, end } = this._selection.range;
    const wholeRows = start.col === 0 && end.col === this.cols - 1;
    const wholeCols = start.row === 0 && end.row === this.rows - 1;
    if (!wholeRows && !wholeCols) {
      this._modeAnnouncement =
        'Select entire rows with Shift+Space or columns with Ctrl+Space first';
      return true;
    }
    // A whole-grid selection counts as rows
    const axis = wholeRows ? 'row' : 'col';
    const index = axis === 'row' ? start.row : start.col;
    const count = axis === 'row' ? end.row - start.row + 1 : end.col - start.col + 1;
    if (remove && count >= (axis === 'row' ? this.rows : this.cols)) {
      this._modeAnnouncement = `Cannot delete every ${axis === 'row' ? 'row' : 'column'}`;
      return true;
    }
    this._changeStructure({ axis, kind: insert ? 'insert' : 'delete', index, count }, 'user');
    this.updateComplete.then(() => this._focusActiveCell());
    return true;
  }

  private _isCellFilled(row: number, col: number): boolean {
    const cell = this._internalData.get(cellKey(row, col));
    return cell !== undefined && cell.rawValue !== '';
  }

  /**
   * Ctrl+Space: select the full columns spanned by the selection.
   * Shift+Space: select the full rows. Both: select all.
   * The active cell stays where it is.
   */
  private _selectWholeRowsOrColumns(columns: boolean, rows: boolean): void {
    const sel = this._selection;
    const active = sel.activeCell;
    if (columns && rows) {
      sel.selectAll();
    } else if (columns) {
      sel.selectColumns(sel.anchor.col, sel.head.col, active);
    } else {
      sel.selectRows(sel.anchor.row, sel.head.row, active);
    }
    this._dispatchSelectionChange();
    this.updateComplete.then(() => this._focusActiveCell());
  }

  // ─── Row / Column Headers ───────────────────────────

  /**
   * Pointer selection on row/column headers: click selects the row or
   * column, Shift+click extends from the selection's anchor, and dragging
   * across headers selects a span. The active cell moves into the clicked
   * column (or row) but keeps its row (or column), so the view doesn't jump.
   */
  private _handleHeaderPointerDown(e: PointerEvent): void {
    if (e.button !== 0) return;
    const headerEl = (e.target as HTMLElement).closest(
      '[data-header-col], [data-header-row]'
    ) as HTMLElement | null;
    if (!headerEl) return;
    e.preventDefault();

    if (this._isEditing) this._commitEdit();

    const isCol = headerEl.dataset.headerCol !== undefined;
    const attr = isCol ? 'headerCol' : 'headerRow';
    const index = parseInt(headerEl.dataset[attr]!, 10);
    const sel = this._selection;
    const from = e.shiftKey ? (isCol ? sel.anchor.col : sel.anchor.row) : index;

    const select = (to: number) => {
      const lo = Math.min(from, to);
      const hi = Math.max(from, to);
      const active = sel.activeCell;
      if (isCol) {
        const col = e.shiftKey && active.col >= lo && active.col <= hi ? active.col : from;
        sel.selectColumns(from, to, { row: active.row, col });
      } else {
        const row = e.shiftKey && active.row >= lo && active.row <= hi ? active.row : from;
        sel.selectRows(from, to, { row, col: active.col });
      }
      this._dispatchSelectionChange();
      // Keep DOM focus on the active cell, which a drag can move
      this.updateComplete.then(() => this._focusActiveCell());
    };

    select(index);

    const controller = new AbortController();
    const { signal } = controller;
    let last = index;
    const onPointerMove = (ev: PointerEvent) => {
      const el = this.shadowRoot?.elementFromPoint(ev.clientX, ev.clientY) as HTMLElement | null;
      const over = el?.closest?.(isCol ? '[data-header-col]' : '[data-header-row]') as HTMLElement | null;
      if (!over) return;
      const to = parseInt(over.dataset[attr]!, 10);
      if (to !== last) {
        last = to;
        select(to);
      }
    };
    const end = () => controller.abort();
    document.addEventListener('pointermove', onPointerMove, { signal });
    document.addEventListener('pointerup', end, { signal });
    document.addEventListener('pointercancel', end, { signal });
    window.addEventListener('blur', end, { signal });
  }

  // ─── Select All (Corner Header) ─────────────────────

  private _handleSelectAll(): void {
    this._selection.selectAll();
    this._dispatchSelectionChange();
    this.updateComplete.then(() => this._focusActiveCell());
  }

  private _handleCornerKeydown(e: KeyboardEvent): void {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      // Keep the grid handler from also treating this key as "start editing".
      e.stopPropagation();
      this._handleSelectAll();
    }
  }

  // ─── Mouse Handling ─────────────────────────────────

  private _handleCellPointerDown(e: PointerEvent): void {
    const target = e.target as HTMLElement;
    const cellEl = target.closest('[data-row]') as HTMLElement | null;
    if (!cellEl) return;

    const row = parseInt(cellEl.dataset.row!, 10);
    const col = parseInt(cellEl.dataset.col!, 10);

    // Commit any in-progress edit
    if (this._isEditing) {
      this._commitEdit();
    }

    this._selection.startSelection(row, col, e.shiftKey);
    this._dispatchSelectionChange();
    cellEl.focus();

    // Set up drag tracking on the document with guaranteed cleanup
    const controller = new AbortController();
    const { signal } = controller;

    const cleanup = () => controller.abort();

    const onPointerMove = (ev: PointerEvent) => {
      // Find the cell under the pointer
      const el = this.shadowRoot?.elementFromPoint(ev.clientX, ev.clientY) as HTMLElement | null;
      const nearest = el?.closest?.('[data-row]') as HTMLElement | null;
      if (nearest) {
        const r = parseInt(nearest.dataset.row!, 10);
        const c = parseInt(nearest.dataset.col!, 10);
        this._selection.extendSelection(r, c);
        this._dispatchSelectionChange();
      }
    };

    const onPointerUp = () => {
      this._selection.endSelection();
      cleanup();
    };

    document.addEventListener('pointermove', onPointerMove, { signal });
    document.addEventListener('pointerup', onPointerUp, { signal });
    document.addEventListener('pointercancel', onPointerUp, { signal });
    window.addEventListener('blur', onPointerUp, { signal });
  }

  private _handleCellDblClick(e: MouseEvent): void {
    if (this.readOnly) return;
    const target = e.target as HTMLElement;
    const cellEl = target.closest('[data-row]') as HTMLElement | null;
    if (!cellEl) return;

    const row = parseInt(cellEl.dataset.row!, 10);
    const col = parseInt(cellEl.dataset.col!, 10);
    this._selection.moveTo(row, col);
    this._startEditing();
  }

  // ─── Clipboard ──────────────────────────────────────

  private async _handleCopy(): Promise<void> {
    await this._clipboardManager.copy(this._internalData, this._selection.range);
  }

  private async _handleCut(): Promise<void> {
    if (this.readOnly || this._pasteInProgress) return;

    const mutations = this._mutationCount;
    const keysToDelete = await this._clipboardManager.cut(
      this._internalData,
      this._selection.range
    );
    // The grid changed while the clipboard was being written (rows moved, data
    // replaced, cells edited): clearing the saved coordinates could erase
    // something other than what was copied, so leave the data alone.
    if (this._changedSince(mutations)) {
      this._modeAnnouncement = 'Cut not applied because the data changed';
      return;
    }
    if (this._pasteInProgress) {
      this._modeAnnouncement = 'Cut not applied because another change is in progress';
      return;
    }

    const updates: Array<{ id: string; value: string }> = [];
    const selection = this._snapshotSelection();
    for (const key of keysToDelete) {
      if (this._internalData.has(key)) {
        updates.push({ id: key, value: '' });
      }
    }

    const batch = this._buildCommandBatch(updates, 'cut', selection, selection);
    const vetted = batch && this._vetUserBatch(batch);
    if (vetted) {
      this._executeUserBatch(vetted);
    }
  }

  private async _handlePaste(): Promise<void> {
    if (this.readOnly || this._pasteInProgress) return;

    // Paste anchors at the top-left of the selection (not the head, which may
    // be the bottom-right corner after a Shift+arrow or drag selection).
    const { row, col } = this._selection.range.start;
    const revision = this._structureRevision;
    const generation = this._dataGeneration;
    const dataRef = this._lastSyncedDataRef;
    const updates = await this._clipboardManager.paste(row, col, this.rows, this.cols);

    if (!updates || updates.length === 0) return;
    if (
      revision !== this._structureRevision ||
      generation !== this._dataGeneration ||
      dataRef !== this.data
    ) {
      this._modeAnnouncement = 'Paste not applied because rows, columns or data changed';
      return;
    }
    if (this._pasteInProgress) {
      this._modeAnnouncement = 'Paste not applied because another change is in progress';
      return;
    }

    const selection = this._snapshotSelection();
    // Pasted formats replace the target's; cells pasted without one keep theirs.
    const batch = this._buildCellBatch(
      updates.map((u) => ({ id: u.id, value: u.value, format: u.format })),
      'paste',
      selection
    );
    const vetted = batch && this._vetUserBatch(batch);
    if (!vetted) return;
    await this._runUserBatch(vetted);
  }

  /**
   * Apply a command batch progressively in chunks, yielding to the event loop
   * between chunks via requestAnimationFrame to keep the UI responsive.
   */
  private _applyBatchProgressive(batch: CommandBatch): Promise<boolean> {
    const generation = this._dataGeneration;
    return new Promise((resolve) => {
      const deltas = batch.deltas;
      let offset = 0;

      const applyChunk = () => {
        // The data was replaced between frames: the rest of this batch (and
        // its history entry) would apply to cells it was never built for.
        if (generation !== this._dataGeneration) {
          resolve(false);
          return;
        }
        this._mutationCount++;
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
          resolve(true);
        }
      };

      applyChunk();
    });
  }

  // ─── Cell Clearing ──────────────────────────────────

  private _clearSelectedCells(): void {
    if (this._pasteInProgress) return;
    const range = this._selection.range;
    const updates: Array<{ id: string; value: string }> = [];
    const selection = this._snapshotSelection();

    for (let r = range.start.row; r <= range.end.row; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        const key = cellKey(r, c);
        if (this._internalData.has(key)) {
          updates.push({ id: key, value: '' });
        }
      }
    }

    const batch = this._buildCommandBatch(updates, 'clear', selection, selection);
    const vetted = batch && this._vetUserBatch(batch);
    if (vetted) {
      this._executeUserBatch(vetted);
    }
  }

  // ─── Fill ───────────────────────────────────────────

  /**
   * Fill down (Ctrl+D) or right (Ctrl+R): copy the first row/column of the
   * selection into the rest of it, shifting relative formula references and
   * copying formats. With a single row/column selected, copies from the
   * row above / column to the left instead.
   */
  private _fillSelection(direction: 'down' | 'right'): void {
    if (this.readOnly || this._pasteInProgress) return;
    const { start, end } = this._selection.range;
    const down = direction === 'down';

    let source = down ? start.row : start.col;
    const last = down ? end.row : end.col;
    if (last === source) {
      if (source === 0) return;
      source -= 1;
    }

    const updates: CellUpdate[] = [];
    const crossStart = down ? start.col : start.row;
    const crossEnd = down ? end.col : end.row;
    for (let target = source + 1; target <= last; target++) {
      for (let cross = crossStart; cross <= crossEnd; cross++) {
        const srcKey = down ? cellKey(source, cross) : cellKey(cross, source);
        const dstKey = down ? cellKey(target, cross) : cellKey(cross, target);
        const src = this._internalData.get(srcKey);
        const offset = target - source;
        updates.push({
          id: dstKey,
          value: this._shiftFormula(src?.rawValue ?? '', down ? offset : 0, down ? 0 : offset),
          format: src?.format ?? null,
        });
      }
    }

    const batch = this._buildCellBatch(updates, 'fill', this._snapshotSelection());
    const vetted = batch && this._vetUserBatch(batch);
    if (vetted) void this._runUserBatch(vetted);
  }

  /**
   * Keyboard fill series (Ctrl+Alt+D / Ctrl+Alt+R): the leading rows (or
   * columns) of the selection that contain data are the source, and the rest
   * of the selection is filled from them exactly as a fill-handle drag would,
   * so `1, 2` followed by empty cells continues as `3, 4, ...`.
   */
  private _fillSeriesInSelection(direction: 'down' | 'right'): void {
    if (this.readOnly || this._pasteInProgress) return;
    const range = this._selection.range;
    const down = direction === 'down';
    const first = down ? range.start.row : range.start.col;
    const last = down ? range.end.row : range.end.col;
    const crossStart = down ? range.start.col : range.start.row;
    const crossEnd = down ? range.end.col : range.end.row;

    let seedEnd = first - 1;
    for (let p = first; p <= last; p++) {
      let filled = false;
      for (let x = crossStart; x <= crossEnd && !filled; x++) {
        filled = down ? this._isCellFilled(p, x) : this._isCellFilled(x, p);
      }
      if (!filled) break;
      seedEnd = p;
    }
    if (seedEnd < first || seedEnd >= last) {
      this._modeAnnouncement = `Select the cells to continue and the empty cells ${down ? 'below' : 'to the right of'} them`;
      return;
    }
    const source: SelectionRange = {
      start: { ...range.start },
      end: down ? { row: seedEnd, col: range.end.col } : { row: range.end.row, col: seedEnd },
    };
    this._autoFill(source, { start: { ...range.start }, end: { ...range.end } });
    this.updateComplete.then(() => this._focusActiveCell());
  }

  /**
   * Fill-handle drag: dragging the square at the selection's bottom-right
   * corner previews the cells to fill (along whichever axis the pointer has
   * moved further past the selection) and fills them on release.
   */
  private _handleFillHandlePointerDown(e: PointerEvent): void {
    if (e.button !== 0 || this.readOnly || this._pasteInProgress) return;
    // Keep the cell underneath from starting a new selection
    e.preventDefault();
    e.stopPropagation();

    const { start, end } = this._selection.range;
    const source: SelectionRange = { start: { ...start }, end: { ...end } };
    let target: SelectionRange | null = null;

    const controller = new AbortController();
    const { signal } = controller;
    const onPointerMove = (ev: PointerEvent) => {
      const el = this.shadowRoot?.elementFromPoint(ev.clientX, ev.clientY) as HTMLElement | null;
      const cellEl = el?.closest?.('[data-row]') as HTMLElement | null;
      if (!cellEl) return;
      const next = this._fillTargetFor(
        source,
        parseInt(cellEl.dataset.row!, 10),
        parseInt(cellEl.dataset.col!, 10)
      );
      const same =
        next === target ||
        (next !== null &&
          target !== null &&
          next.start.row === target.start.row &&
          next.start.col === target.start.col &&
          next.end.row === target.end.row &&
          next.end.col === target.end.col);
      if (!same) {
        target = next;
        this._fillPreview = next;
      }
    };
    const finish = (apply: boolean) => {
      controller.abort();
      this._fillPreview = null;
      if (apply && target) this._autoFill(source, target);
      this.updateComplete.then(() => this._focusActiveCell());
    };
    document.addEventListener('pointermove', onPointerMove, { signal });
    document.addEventListener('pointerup', () => finish(true), { signal });
    document.addEventListener('pointercancel', () => finish(false), { signal });
    window.addEventListener('blur', () => finish(false), { signal });
  }

  /**
   * The range a fill-handle drag from `source` to cell (row, col) covers, or
   * null while the pointer is still inside the source.
   */
  private _fillTargetFor(source: SelectionRange, row: number, col: number): SelectionRange | null {
    const below = row - source.end.row;
    const above = source.start.row - row;
    const right = col - source.end.col;
    const left = source.start.col - col;
    const vertical = Math.max(below, above, 0);
    const horizontal = Math.max(right, left, 0);
    if (vertical === 0 && horizontal === 0) return null;
    if (vertical >= horizontal) {
      return below > 0
        ? { start: { ...source.start }, end: { row, col: source.end.col } }
        : { start: { row, col: source.start.col }, end: { ...source.end } };
    }
    return right > 0
      ? { start: { ...source.start }, end: { row: source.end.row, col } }
      : { start: { row: source.start.row, col }, end: { ...source.end } };
  }

  /**
   * Fill `target` (which contains `source` and extends it along one axis)
   * from `source`, as the fill handle does. Each line of the source along
   * the fill direction is handled separately:
   *
   * - two or more plain numbers continue their linear trend (1, 2 → 3, 4);
   * - anything else repeats, with formulas' relative references shifted and
   *   formats copied.
   */
  private _autoFill(source: SelectionRange, target: SelectionRange): void {
    if (this.readOnly || this._pasteInProgress) return;
    const vertical = target.start.row < source.start.row || target.end.row > source.end.row;
    const along = vertical ? 'row' : 'col';
    const cross = vertical ? 'col' : 'row';
    const lineStart = source.start[along];
    const lineLength = source.end[along] - lineStart + 1;

    const updates: CellUpdate[] = [];
    for (let x = source.start[cross]; x <= source.end[cross]; x++) {
      const keyAt = (p: number) => (vertical ? cellKey(p, x) : cellKey(x, p));
      const line = Array.from({ length: lineLength }, (_, i) =>
        this._internalData.get(keyAt(lineStart + i))
      );
      const trend = linearTrend(line.map((cell) => this._literalNumber(cell)));
      for (let p = target.start[along]; p <= target.end[along]; p++) {
        if (p >= lineStart && p < lineStart + lineLength) continue;
        const offset = p - lineStart;
        const srcIndex = ((offset % lineLength) + lineLength) % lineLength;
        const src = line[srcIndex];
        let value: string;
        if (trend) {
          value = formatSeriesNumber(trend.intercept + trend.slope * offset);
        } else {
          const shift = offset - srcIndex;
          value = this._shiftFormula(src?.rawValue ?? '', vertical ? shift : 0, vertical ? 0 : shift);
        }
        updates.push({ id: keyAt(p), value, format: src?.format ?? null });
      }
    }

    const selectionAfter: SelectionSnapshot = {
      anchor: { ...target.start },
      head: { ...target.end },
      active: { ...this._selection.activeCell },
    };
    this._modeAnnouncement = `Filled ${updates.length} cell${updates.length === 1 ? '' : 's'}`;
    const batch = this._buildCellBatch(updates, 'fill', this._snapshotSelection());
    const vetted = batch && this._vetUserBatch(batch);
    if (vetted) {
      vetted.selectionAfter = selectionAfter;
      void this._runUserBatch(vetted);
    } else {
      this._restoreSelection(selectionAfter);
    }
  }

  /** The number a cell holds as a literal (not a formula), or null. */
  private _literalNumber(cell: CellData | undefined): number | null {
    if (!cell || cell.rawValue === '' || cell.rawValue.startsWith('=')) return null;
    const evaluated = this._formulaEngine.evaluate(cell.rawValue);
    if (evaluated.type !== 'number') return null;
    const n = Number(evaluated.displayValue);
    return Number.isFinite(n) ? n : null;
  }

  /** Shift the relative references of a formula; non-formulas are returned as-is. */
  private _shiftFormula(raw: string, rowOffset: number, colOffset: number): string {
    if (!raw.startsWith('=') || (rowOffset === 0 && colOffset === 0)) return raw;
    return this._clipboardManager.adjustFormulaReferences(raw, rowOffset, colOffset);
  }

  /**
   * Ctrl/Cmd+Enter while editing: commit the edited value into every cell of
   * the selection. Formulas are written relative to each cell, as if the
   * edited cell were copied there.
   */
  private _commitEditToSelection(): void {
    if (!this._isEditing || !this._editingCellKey) return;
    const origin = parseKey(this._editingCellKey);
    const value = this._editValue;
    const { start, end } = this._selection.range;

    if (start.row === end.row && start.col === end.col) {
      this._commitEdit('keep');
      return;
    }
    // Validate the edited cell's value up front so the editor can stay open;
    // the other cells are validated individually when the batch is vetted.
    if (!this._checkEditValue(this._editingCellKey, value, 'keep')) return;

    this._isEditing = false;
    this._editingCellKey = null;
    this._validationMessage = '';

    const updates: CellUpdate[] = [];
    for (let r = start.row; r <= end.row; r++) {
      for (let c = start.col; c <= end.col; c++) {
        updates.push({
          id: cellKey(r, c),
          value: this._shiftFormula(value, r - origin.row, c - origin.col),
        });
      }
    }
    const cellCount = updates.length;
    this._modeAnnouncement = `Filled ${cellCount} cells`;

    const batch = this._buildCellBatch(updates, 'fill', this._snapshotSelection());
    const vetted = batch && this._vetUserBatch(batch);
    if (vetted) void this._runUserBatch(vetted);
    this.updateComplete.then(() => this._focusActiveCell());
  }

  // ─── Scroll Handling ────────────────────────────────

  private _handleScroll(e: Event): void {
    const target = e.target as HTMLElement;
    this._scrollTop = target.scrollTop;
    this._scrollLeft = target.scrollLeft;
  }

  // ─── Find Bar ───────────────────────────────────────

  private _handleFindInput(e: Event): void {
    this._findQuery = (e.target as HTMLInputElement).value;
    // Like browser find: jump to the first match at or after the active cell
    this._findStep(1, true);
  }

  private _handleFindKeydown(e: KeyboardEvent): void {
    const ctrl = e.ctrlKey || e.metaKey;
    if (e.key === 'Enter' || e.key === 'F3') {
      e.preventDefault();
      this._findStep(e.shiftKey ? -1 : 1);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      this.closeFind();
    } else if (ctrl && e.key.toLowerCase() === 'f') {
      e.preventDefault();
      this._findInput?.select();
    } else {
      return;
    }
    e.stopPropagation();
  }

  /**
   * Move the active cell to the next (1) or previous (-1) match after the
   * active cell, wrapping around. With `inclusive`, the active cell itself
   * counts. Focus stays in the find input.
   */
  private _findStep(direction: 1 | -1, inclusive = false): void {
    if (!this._findQuery) {
      this._findStatus = '';
      return;
    }
    const matches = this.findCells(this._findQuery);
    if (matches.length === 0) {
      this._findStatus = 'No matches';
      return;
    }
    const active = this._selection.activeCell;
    const current = active.row * this.cols + active.col;
    const position = (key: string) => {
      const { row, col } = parseKey(key);
      return row * this.cols + col;
    };
    let index = -1;
    if (direction === 1) {
      index = matches.findIndex((k) => (inclusive ? position(k) >= current : position(k) > current));
      if (index < 0) index = 0;
    } else {
      for (let i = matches.length - 1; i >= 0; i--) {
        if (position(matches[i]) < current) {
          index = i;
          break;
        }
      }
      if (index < 0) index = matches.length - 1;
    }
    const { row, col } = parseKey(matches[index]);
    this._selection.moveTo(row, col);
    this._ensureCellVisible(row, col);
    this._dispatchSelectionChange();
    this._findStatus = `${index + 1} of ${matches.length}`;
  }

  // ─── Editor Input ───────────────────────────────────

  private _handleEditorInput(e: Event): void {
    this._editValue = (e.target as HTMLInputElement).value;
    this._validationMessage = '';
    this._exitRefMode();
  }

  private _handleEditorBlur(): void {
    if (this._isEditing) {
      this._commitEdit();
    }
  }

  private _handleFormulaBarModeChange(
    e: CustomEvent<{ mode: FormulaBarMode }>
  ): void {
    this._formulaBarMode = e.detail.mode;
  }

  private _handleFormulaBarCommit(e: CustomEvent<{ value: string; cellRef: string }>): void {
    if (!this._applyFormulaBarCommit(e.detail)) {
      // Rejected: drop the draft so it can't be committed into another cell later
      this.shadowRoot?.querySelector<Y11nFormulaBar>('y11n-formula-bar')?.resetDraft();
    }
  }

  /** Apply a formula-bar commit; returns false if nothing was written. */
  private _applyFormulaBarCommit(detail: { value: string; cellRef: string }): boolean {
    if (this.readOnly) return false;

    // Use the cell key from the event (which cell the formula bar was showing)
    // rather than the current active cell, which may have changed due to clicking.
    let targetKey: string;
    if (this._editingCellKey) {
      targetKey = this._editingCellKey;
    } else {
      const coord = refToCoord(detail.cellRef);
      targetKey = coord ? cellKey(coord.row, coord.col) : this._getActiveCellKey();
    }

    if (this._isCellLocked(targetKey)) {
      this._modeAnnouncement = `${coordToRef(parseKey(targetKey))} is read-only`;
      return false;
    }
    const value = detail.value;
    const oldValue = this._internalData.get(targetKey)?.rawValue ?? '';
    const message = value === oldValue ? null : this._validate(targetKey, value);
    if (message) {
      this._dispatchValidationError({ errors: [{ cellId: targetKey, value, message }], operation: 'edit' });
      this._modeAnnouncement = `Invalid value: ${message}`;
      return false;
    }

    if (this._isEditing) {
      this._isEditing = false;
      this._editingCellKey = null;
      this._validationMessage = '';
    }

    // An unchanged value is a successful no-op, not a rejection
    return this._commitRawValue(targetKey, value, { validate: false }) || value === oldValue;
  }

  // ─── Event Dispatching ──────────────────────────────

  private _dispatchCellChange(detail: CellChangeDetail): void {
    this.dispatchEvent(new CustomEvent('cell-change', { detail, bubbles: true, composed: true }));
  }

  private _dispatchSelectionChange(): void {
    const detail: SelectionChangeDetail = { range: this._selection.range };
    this.dispatchEvent(
      new CustomEvent('selection-change', { detail, bubbles: true, composed: true })
    );
  }

  private _dispatchDataChange(detail: DataChangeDetail): void {
    this.dispatchEvent(
      new CustomEvent('data-change', { detail, bubbles: true, composed: true })
    );
  }

  private _dispatchFormatChange(detail: FormatChangeDetail): void {
    this.dispatchEvent(
      new CustomEvent('format-change', { detail, bubbles: true, composed: true })
    );
  }

  private _dispatchValidationError(detail: ValidationErrorDetail): void {
    this.dispatchEvent(
      new CustomEvent('validation-error', { detail, bubbles: true, composed: true })
    );
  }

  private _dispatchStructureChange(detail: StructureChangeDetail): void {
    this.dispatchEvent(
      new CustomEvent('structure-change', { detail, bubbles: true, composed: true })
    );
  }

  private _dispatchSortChange(detail: SortDetail): void {
    this.dispatchEvent(new CustomEvent('sort-change', { detail, bubbles: true, composed: true }));
  }

  // ─── Virtual Rendering ──────────────────────────────

  /**
   * Compute which rows and columns are visible based on scroll position.
   * We add a buffer of extra rows/cols for smooth scrolling.
   */
  private _getVisibleRange(): { rows: number[]; cols: number[] } {
    const cellHeight = this._cellSize.height;
    const cellWidth = this._cellSize.width;
    const gridEl = this._grid;
    const viewHeight =
      gridEl?.clientHeight || cellHeight * Y11nSpreadsheet.INITIAL_VIEW_ROWS;
    const viewWidth =
      gridEl?.clientWidth || cellWidth * Y11nSpreadsheet.INITIAL_VIEW_COLS;

    const buffer = 5;

    const startRow = Math.max(0, Math.floor(this._scrollTop / cellHeight) - buffer);
    const endRow = Math.min(this.rows, Math.ceil((this._scrollTop + viewHeight) / cellHeight) + buffer);
    const startCol = Math.max(0, this._colAt(this._scrollLeft) - buffer);
    const endCol = Math.min(this.cols, this._colAt(this._scrollLeft + viewWidth) + 1 + buffer);

    // The active cell (roving tabindex target) and the reference cursor must
    // always be in the DOM. Render them individually rather than stretching
    // the window, which would render every row between the viewport and a
    // far-away active cell after the user scrolls with the mouse.
    const { row: activeRow, col: activeCol } = this._selection.activeCell;
    const extraRows = [activeRow];
    const extraCols = [activeCol];
    // Frozen rows and columns are always on screen
    for (let r = 0; r < this._frozenRowCount; r++) extraRows.push(r);
    for (let c = 0; c < this._frozenColCount; c++) extraCols.push(c);
    if (this._refMode) {
      extraRows.push(this._refCursorRow);
      extraCols.push(this._refCursorCol);
    }

    return {
      rows: windowWithExtras(startRow, endRow, extraRows),
      cols: windowWithExtras(startCol, endCol, extraCols),
    };
  }

  private _ensureCellVisible(row: number, col: number): void {
    const gridEl = this._grid;
    if (!gridEl) return;

    const cellHeight = this._cellSize.height;
    const cellWidth = this._cellSize.width;
    const viewHeight =
      gridEl.clientHeight || cellHeight * Y11nSpreadsheet.INITIAL_VIEW_ROWS;
    const viewWidth =
      gridEl.clientWidth || cellWidth * Y11nSpreadsheet.INITIAL_VIEW_COLS;
    const dataViewportWidth = Math.max(
      this._colWidth(col),
      viewWidth - Y11nSpreadsheet.ROW_HEADER_WIDTH
    );
    // Subtract the sticky column header height so cells aren't scrolled
    // behind it (mirrors how dataViewportWidth subtracts the row header).
    const dataViewportHeight = Math.max(cellHeight, viewHeight - cellHeight);

    const cellTop = row * cellHeight;
    const cellBottom = cellTop + cellHeight;
    const cellLeft = this._colLeft(col);
    const cellRight = cellLeft + this._colWidth(col);

    // Frozen rows/columns never need scrolling; the others must clear the
    // frozen area, which covers the start of the viewport
    const frozenHeight = this._frozenHeight;
    const frozenWidth = this._frozenWidth;
    let nextScrollTop = this._scrollTop;
    if (row < this._frozenRowCount) {
      // stays put
    } else if (cellTop - frozenHeight < nextScrollTop) {
      nextScrollTop = Math.max(0, cellTop - frozenHeight);
    } else if (cellBottom > nextScrollTop + dataViewportHeight) {
      nextScrollTop = cellBottom - dataViewportHeight;
    }

    let nextScrollLeft = this._scrollLeft;
    if (col < this._frozenColCount) {
      // stays put
    } else if (cellLeft - frozenWidth < nextScrollLeft) {
      nextScrollLeft = Math.max(0, cellLeft - frozenWidth);
    } else if (cellRight > nextScrollLeft + dataViewportWidth) {
      nextScrollLeft = cellRight - dataViewportWidth;
    }

    if (nextScrollTop !== this._scrollTop || nextScrollLeft !== this._scrollLeft) {
      this._scrollTop = nextScrollTop;
      this._scrollLeft = nextScrollLeft;
      gridEl.scrollTop = nextScrollTop;
      gridEl.scrollLeft = nextScrollLeft;
      // The browser clamps to the scrollable size without firing a scroll
      // event when the position doesn't change: keep the state in sync
      this._scrollLeft = gridEl.scrollLeft;
    }
  }

  // ─── Cell Size ──────────────────────────────────────

  /**
   * Measure the rendered cell size from the hidden probe, which is sized with
   * the same `--_cell-width` / `--_cell-height` as the cells. This resolves
   * any CSS length (`7rem`, `2em`, `calc(...)`), not just px values.
   */
  private _measureCellSize(): void {
    const probe = this.renderRoot.querySelector<HTMLElement>('.ls-size-probe');
    if (!probe) return;
    const style = getComputedStyle(probe);
    this._setCellSize(parseFloat(style.width), parseFloat(style.height));
  }

  private _setCellSize(width: number, height: number): void {
    // A hidden (display: none) grid measures 0: keep the last good size
    if (!(width > 0) || !(height > 0)) return;
    if (width === this._cellSize.width && height === this._cellSize.height) return;
    this._cellSize = { width, height };
  }

  /** Re-measure whenever the probe resizes. Returns false where ResizeObserver is unavailable. */
  private _observeCellSize(): boolean {
    const probe = this.renderRoot.querySelector<HTMLElement>('.ls-size-probe');
    if (!probe || typeof ResizeObserver === 'undefined') return false;
    this._cellSizeObserver ??= new ResizeObserver((entries) => {
      const entry = entries[entries.length - 1];
      const box = entry.borderBoxSize?.[0];
      if (box) this._setCellSize(box.inlineSize, box.blockSize);
      else this._setCellSize(entry.contentRect.width, entry.contentRect.height);
      if (this._cellSize !== this._renderedCellSize) this.requestUpdate();
    });
    this._cellSizeObserver.observe(probe);
    return true;
  }

  // ─── Editor Positioning ─────────────────────────────

  private _getEditorStyle(): Record<string, string> {
    if (!this._isEditing) return { display: 'none' };

    const { row, col } = this._selection.activeCell;
    const cellHeight = this._cellSize.height;
    const cellWidth = this._colWidth(col);
    const headerWidth = Y11nSpreadsheet.ROW_HEADER_WIDTH;

    // +1 for the header row; frozen cells don't move with the scroll position
    const top = (row + 1) * cellHeight - (row < this._frozenRowCount ? 0 : this._scrollTop);
    const left = headerWidth + this._colLeft(col) - (col < this._frozenColCount ? 0 : this._scrollLeft);

    const styles: Record<string, string> = {
      display: 'block',
      position: 'absolute',
      top: `${top}px`,
      left: `${left}px`,
      width: `${cellWidth}px`,
      height: `${cellHeight}px`,
      'z-index': '10',
    };
    // A scrolling cell can slide under the headers and frozen panes: clip the
    // editor to what's left of the cell, so it never covers them
    const clipTop = row < this._frozenRowCount ? 0 : Math.max(0, cellHeight + this._frozenHeight - top);
    const clipLeft = col < this._frozenColCount ? 0 : Math.max(0, headerWidth + this._frozenWidth - left);
    if (clipTop > 0 || clipLeft > 0) styles['clip-path'] = `inset(${clipTop}px 0 0 ${clipLeft}px)`;

    // Inherit cell formatting into the editor (via styleMap for safety)
    const fmtStyles = this._getCellStyles(row, col);
    for (const [k, v] of Object.entries(fmtStyles)) {
      if (!['justify-content', 'background-color', 'align-items', 'white-space', 'overflow-wrap'].includes(k)) {
        styles[k] = v;
      }
    }

    return styles;
  }

  /** Position the validation message just below the editor. */
  private _getValidationMessageStyle(): Record<string, string> {
    const editor = this._getEditorStyle();
    const cellHeight = this._cellSize.height;
    return {
      top: `${parseFloat(editor.top) + cellHeight}px`,
      left: editor.left,
    };
  }

  // ─── Render ─────────────────────────────────────────

  static styles = css`
    :host {
      display: block;
      height: 100%;
      position: relative;
      font-family: var(--ls-font-family, system-ui, -apple-system, sans-serif);
      font-size: var(--ls-font-size, 13px);
      color: var(--ls-text-color, #333);
      --_cell-width: var(--ls-cell-width, 100px);
      --_cell-height: var(--ls-cell-height, 28px);
      --_header-bg: var(--ls-header-bg, #f3f3f3);
      --_border-color: var(--ls-border-color, #e0e0e0);
      --_selection-border: var(--ls-selection-border, 2px solid #1a73e8);
      --_selection-bg: var(--ls-selection-bg, rgba(26, 115, 232, 0.2));
      --_focus-ring: var(--ls-focus-ring, 2px solid #1a73e8);
      --_editor-bg: var(--ls-editor-bg, #fff);
      /* Background of frozen cells, which must be opaque over scrolled cells */
      --_cell-bg: var(--ls-cell-bg, #fff);
      --_frozen-line: var(--ls-frozen-line-color, #9aa0a6);
      --_editor-shadow: var(--ls-editor-shadow, 0 2px 6px rgba(0,0,0,0.2));
      --_row-header-width: 50px;
    }

    *,
    *::before,
    *::after {
      box-sizing: border-box;
    }

    .ls-wrapper {
      position: relative;
      display: flex;
      flex-direction: column;
      width: 100%;
      height: 100%;
      overflow: hidden;
    }

    .ls-grid-shell {
      position: relative;
      flex: 1 1 auto;
      min-height: 0;
      overflow: hidden;
    }

    /* Sized like a cell so the component can read the resolved px size */
    .ls-size-probe {
      position: absolute;
      top: 0;
      left: 0;
      width: var(--_cell-width);
      height: var(--_cell-height);
      visibility: hidden;
      pointer-events: none;
    }

    .ls-grid {
      display: grid;
      grid-template-columns: var(--_row-header-width) repeat(var(--cols), var(--_cell-width));
      align-content: start;
      overflow: auto;
      position: relative;
      width: 100%;
      height: round(down, 100%, var(--_cell-height));
      scroll-snap-type: y mandatory;
      outline: none;
    }

    .ls-header-row {
      display: contents;
    }

    .ls-row {
      display: contents;
    }

    /*
     * Stacking, bottom to top: cells (focused/active: 1), frozen cells (2),
     * cells frozen on both axes (3), row headers (4), frozen row headers (5),
     * column headers (6), frozen column headers (7), the corner (8).
     */
    .ls-corner-header,
    .ls-col-header {
      position: sticky;
      top: 0;
      z-index: 6;
      display: flex;
      align-items: center;
      justify-content: center;
      height: var(--_cell-height);
      background: var(--_header-bg);
      border-right: 1px solid var(--_border-color);
      border-bottom: 2px solid var(--_border-color);
      font-weight: 600;
      user-select: none;
    }

    .ls-corner-header {
      position: sticky;
      left: 0;
      z-index: 8;
      scroll-snap-align: start;
    }

    .ls-col-header,
    .ls-row-header {
      cursor: pointer;
    }

    /* Drag handle on a column header's right border */
    .ls-col-resizer {
      position: absolute;
      top: 0;
      right: -1px;
      width: 7px;
      height: 100%;
      cursor: col-resize;
      touch-action: none;
      z-index: 1;
    }

    .ls-col-resizer:hover {
      background: linear-gradient(to right, transparent 2px, var(--ls-resizer-color, #1a73e8) 2px, var(--ls-resizer-color, #1a73e8) 4px, transparent 4px);
    }

    .ls-col-header.in-selection,
    .ls-row-header.in-selection {
      background: var(--ls-header-selected-bg, #dbe7fb);
    }

    .ls-row-header {
      position: sticky;
      left: 0;
      z-index: 4;
      display: flex;
      align-items: center;
      justify-content: center;
      height: var(--_cell-height);
      background: var(--_header-bg);
      border-right: 2px solid var(--_border-color);
      border-bottom: 1px solid var(--_border-color);
      font-weight: 600;
      user-select: none;
      scroll-snap-align: start;
    }

    .ls-cell {
      display: flex;
      align-items: center;
      padding: 0 4px;
      height: var(--_cell-height);
      border-right: 1px solid var(--_border-color);
      border-bottom: 1px solid var(--_border-color);
      overflow: hidden;
      white-space: nowrap;
      text-overflow: ellipsis;
      cursor: cell;
      outline: none;
      user-select: none;
      position: relative;
      /* Contain the fill handle's z-index, so it can't rise above frozen cells */
      isolation: isolate;
    }

    .ls-cell.read-only-cell:not([aria-selected='true']) {
      background: var(--ls-readonly-cell-bg, transparent);
    }

    .ls-cell.fill-preview {
      background: var(--ls-fill-preview-bg, rgba(26, 115, 232, 0.08));
      outline: 1px dashed var(--ls-fill-preview-border, #1a73e8);
      outline-offset: -1px;
    }

    .ls-fill-handle {
      position: absolute;
      right: 0;
      bottom: 0;
      width: 7px;
      height: 7px;
      background: var(--ls-fill-handle-color, #1a73e8);
      border: 1px solid #fff;
      cursor: crosshair;
      touch-action: none;
      z-index: 2;
    }

    .ls-cell:focus {
      outline: var(--_focus-ring);
      outline-offset: -2px;
      z-index: 1;
    }

    .ls-cell[aria-selected="true"] {
      background: var(--_selection-bg);
    }

    .ls-cell.active-cell {
      outline: var(--_selection-border);
      outline-offset: -2px;
      z-index: 1;
    }

    .ls-cell.ref-highlight {
      outline: 2px dashed var(--ls-ref-highlight-border, #1a73e8);
      outline-offset: -2px;
      background: var(--ls-ref-highlight-bg, rgba(26, 115, 232, 0.15));
      z-index: 1;
    }

    .ls-cell .cell-text {
      overflow: hidden;
      text-overflow: ellipsis;
    }

    /* Frozen panes: sticky (top/left set inline), opaque, above scrolled cells */
    .ls-cell.frozen-row,
    .ls-cell.frozen-col {
      position: sticky;
      z-index: 2;
      background-color: var(--_cell-bg);
    }

    .ls-cell.frozen-row.frozen-col {
      z-index: 3;
    }

    /* Translucent states are layered over the opaque background */
    .ls-cell.frozen-row[aria-selected='true'],
    .ls-cell.frozen-col[aria-selected='true'] {
      background: linear-gradient(var(--_selection-bg), var(--_selection-bg)), var(--_cell-bg);
    }

    .ls-cell.frozen-row.read-only-cell:not([aria-selected='true']),
    .ls-cell.frozen-col.read-only-cell:not([aria-selected='true']) {
      background: linear-gradient(var(--ls-readonly-cell-bg, transparent), var(--ls-readonly-cell-bg, transparent)), var(--_cell-bg);
    }

    .ls-cell.frozen-row.fill-preview,
    .ls-cell.frozen-col.fill-preview {
      background: linear-gradient(var(--ls-fill-preview-bg, rgba(26, 115, 232, 0.08)), var(--ls-fill-preview-bg, rgba(26, 115, 232, 0.08))), var(--_cell-bg);
    }

    .ls-cell.frozen-row.ref-highlight,
    .ls-cell.frozen-col.ref-highlight {
      background: linear-gradient(var(--ls-ref-highlight-bg, rgba(26, 115, 232, 0.15)), var(--ls-ref-highlight-bg, rgba(26, 115, 232, 0.15))), var(--_cell-bg);
    }

    .ls-row-header.frozen-row {
      z-index: 5;
      /* Sticky snap areas would fight the scrolling rows' snap points */
      scroll-snap-align: none;
    }

    .ls-col-header.frozen-col {
      z-index: 7;
    }

    /* The freeze lines */
    .frozen-row-edge {
      border-bottom: 2px solid var(--_frozen-line);
    }

    .ls-cell.frozen-col-edge,
    .ls-col-header.frozen-col-edge {
      border-right: 2px solid var(--_frozen-line);
    }

    #editor {
      position: absolute;
      box-sizing: border-box;
      padding: 0 4px;
      border: var(--_selection-border);
      background: var(--_editor-bg);
      box-shadow: var(--_editor-shadow);
      font-family: inherit;
      font-size: inherit;
      color: inherit;
      outline: none;
      z-index: 10;
    }

    .ls-validation-message {
      position: absolute;
      z-index: 11;
      max-width: 280px;
      padding: 4px 8px;
      background: var(--ls-validation-bg, #fdecea);
      color: var(--ls-validation-color, #8a1c12);
      border: 1px solid var(--ls-validation-border, #d93025);
      font-size: 12px;
      white-space: normal;
      pointer-events: none;
    }

    .ls-find {
      position: absolute;
      top: 4px;
      right: 20px;
      z-index: 20;
      display: flex;
      align-items: center;
      gap: 4px;
      padding: 4px 6px;
      background: var(--ls-find-bg, #fff);
      border: 1px solid var(--_border-color);
      border-radius: 4px;
      box-shadow: var(--_editor-shadow);
    }

    .ls-find-input {
      width: 160px;
      padding: 2px 4px;
      font: inherit;
      border: 1px solid var(--_border-color);
      border-radius: 2px;
    }

    .ls-find-input:focus-visible,
    .ls-find-button:focus-visible {
      outline: var(--_focus-ring);
      outline-offset: 1px;
    }

    .ls-find-status {
      min-width: 5em;
      font-size: 12px;
      color: var(--ls-find-status-color, #555);
      text-align: center;
    }

    .ls-find-button {
      min-width: 24px;
      height: 24px;
      padding: 0 4px;
      font: inherit;
      color: inherit;
      background: transparent;
      border: 1px solid transparent;
      border-radius: 2px;
      cursor: pointer;
    }

    .ls-find-button:hover {
      background: var(--_header-bg);
      border-color: var(--_border-color);
    }

    /* Visually hidden live region for screen reader announcements */
    .sr-only {
      position: absolute;
      width: 1px;
      height: 1px;
      padding: 0;
      margin: -1px;
      overflow: hidden;
      clip: rect(0, 0, 0, 0);
      border: 0;
    }
  `;

  protected render() {
    this._renderedCellSize = this._cellSize;
    const { rows, cols } = this._getVisibleRange();
    const activeCell = this._selection.activeCell;
    const activeRawValue = this._getCellRaw(activeCell.row, activeCell.col);
    const activeDisplayValue = this._getCellDisplay(activeCell.row, activeCell.col);

    return html`
      <div class="ls-wrapper">
        <y11n-formula-bar
          cell-ref="${this._getActiveCellRef()}"
          raw-value="${activeRawValue}"
          display-value="${activeDisplayValue}"
          .mode="${this._formulaBarMode}"
          .readOnly="${this.readOnly || this._isCellLocked(this._getActiveCellKey())}"
          @formula-bar-mode-change="${this._handleFormulaBarModeChange}"
          @formula-bar-commit="${this._handleFormulaBarCommit}"
        ></y11n-formula-bar>

        <div class="ls-grid-shell">
          <div class="ls-size-probe" aria-hidden="true"></div>
          <input
            id="editor"
            part="editor"
            style=${styleMap(this._getEditorStyle())}
            .value="${this._editValue}"
            @input="${this._handleEditorInput}"
            @keydown="${this._handleEditorKeydown}"
            @blur="${this._handleEditorBlur}"
            aria-label="Cell editor"
            aria-invalid="${this._validationMessage ? 'true' : nothing}"
            aria-describedby="${this._validationMessage ? 'editor-error' : nothing}"
          />
          ${this._isEditing && this._validationMessage
            ? html`<div
                id="editor-error"
                class="ls-validation-message"
                part="validation-message"
                style=${styleMap(this._getValidationMessageStyle())}
              >${this._validationMessage}</div>`
            : nothing}
          ${this._findOpen ? this._renderFindBar() : nothing}

          <div
            class="ls-grid"
            role="grid"
            aria-label="${this.gridLabel}"
            aria-rowcount="${this.rows + 1}"
            aria-colcount="${this.cols + 1}"
            aria-readonly="${this.readOnly}"
            aria-multiselectable="true"
            style="--cols: ${this.cols}; grid-template-columns: ${this._gridTemplateColumns()}; ${this._scrollPadding()}"
            tabindex="-1"
            @keydown="${this._handleGridKeydown}"
            @scroll="${this._handleScroll}"
          >
            ${this._renderHeaderRow(cols)}
            ${this._renderRows(rows, cols)}
          </div>
        </div>

        <div class="sr-only" role="status" aria-live="polite" aria-atomic="true">
          ${this._getAnnouncement(activeCell)}
        </div>
        <div class="sr-only" role="log" aria-live="assertive" aria-atomic="true">
          ${this._modeAnnouncement}
        </div>
      </div>
    `;
  }

  private _renderFindBar() {
    return html`
      <div class="ls-find" role="search" aria-label="Find in spreadsheet">
        <input
          id="find-input"
          class="ls-find-input"
          type="text"
          aria-label="Find"
          aria-describedby="find-status"
          placeholder="Find"
          autocomplete="off"
          spellcheck="false"
          .value="${this._findQuery}"
          @input="${this._handleFindInput}"
          @keydown="${this._handleFindKeydown}"
        />
        <span id="find-status" class="ls-find-status" role="status">${this._findStatus}</span>
        <button
          type="button"
          class="ls-find-button"
          aria-label="Previous match"
          title="Previous match (Shift+Enter)"
          @click="${() => this._findStep(-1)}"
        ><span aria-hidden="true">&#x2191;</span></button>
        <button
          type="button"
          class="ls-find-button"
          aria-label="Next match"
          title="Next match (Enter)"
          @click="${() => this._findStep(1)}"
        ><span aria-hidden="true">&#x2193;</span></button>
        <button
          type="button"
          class="ls-find-button"
          aria-label="Close find"
          title="Close (Escape)"
          @click="${() => this.closeFind()}"
        ><span aria-hidden="true">&#x2715;</span></button>
      </div>
    `;
  }

  private _renderColumnSpacer(size: number) {
    return html`<div style="grid-column: span ${size};"></div>`;
  }

  private _renderHeaderRow(cols: number[]) {
    const range = this._selection.range;
    const fullColumns = range.start.row === 0 && range.end.row === this.rows - 1;
    return html`
      <div class="ls-header-row" role="row" aria-rowindex="1">
        <div
          class="ls-corner-header"
          role="columnheader"
          aria-colindex="1"
          aria-label="Select all"
          tabindex="-1"
          @click="${this._handleSelectAll}"
          @keydown="${this._handleCornerKeydown}"
        ><span class="sr-only">Select all</span></div>
        ${layoutWithGaps(
          cols,
          this.cols,
          (c) => html`
            <div
              class="ls-col-header ${c >= range.start.col && c <= range.end.col ? 'in-selection' : ''} ${this._frozenClasses(-1, c)}"
              style=${c < this._frozenColCount ? `left: ${this._stickyLeft(c)}` : nothing}
              role="columnheader"
              aria-colindex="${c + 2}"
              aria-label="Column ${colToLetter(c)}"
              aria-selected="${fullColumns && c >= range.start.col && c <= range.end.col ? 'true' : nothing}"
              data-header-col="${c}"
              @pointerdown="${this._handleHeaderPointerDown}"
            >
              ${colToLetter(c)}
              <div
                class="ls-col-resizer"
                aria-hidden="true"
                title="Drag to resize, double-click to fit (Alt+Shift+Left/Right, Alt+Shift+0)"
                data-resize-col="${c}"
                @pointerdown="${this._handleResizerPointerDown}"
                @dblclick="${this._handleResizerDblClick}"
              ></div>
            </div>
          `,
          (_start, size) => this._renderColumnSpacer(size)
        )}
      </div>
    `;
  }

  private _renderRows(rows: number[], cols: number[]) {
    const cellHeight = this._cellSize.height;
    type Segment = { key: string; row: number; gap: number };
    const segments = layoutWithGaps<Segment>(
      rows,
      this.rows,
      (row) => ({ key: `r${row}`, row, gap: 0 }),
      (start, size) => ({ key: `g${start}`, row: start, gap: size })
    );

    return html`
      ${repeat(
        segments,
        (seg) => seg.key,
        (seg) =>
          seg.gap > 0
            ? html`<div style="grid-column: 1 / -1; height: ${seg.gap * cellHeight}px;"></div>`
            : this._renderRow(seg.row, cols)
      )}
    `;
  }

  private _renderCell(row: number, c: number) {
    const isSelected = this._selection.isCellSelected(row, c);
    const isActive = this._selection.isCellActive(row, c);
    const isRefTarget = this._refMode && row === this._refCursorRow && c === this._refCursorCol;
    const display = this._getCellDisplay(row, c);
    const key = cellKey(row, c);
    const locked = this._isCellLocked(key);
    const range = this._selection.range;
    const showFillHandle =
      !this.readOnly && !this._isEditing && row === range.end.row && c === range.end.col;
    const preview = this._fillPreview;
    const inFillPreview =
      preview !== null &&
      !isSelected &&
      row >= preview.start.row &&
      row <= preview.end.row &&
      c >= preview.start.col &&
      c <= preview.end.col;

    const cellStyles = this._getCellStyles(row, c);
    if (row < this._frozenRowCount) cellStyles.top = this._stickyTop(row);
    if (c < this._frozenColCount) cellStyles.left = this._stickyLeft(c);

    return html`
      <div
        class="ls-cell ${isActive ? 'active-cell' : ''} ${isRefTarget ? 'ref-highlight' : ''} ${locked ? 'read-only-cell' : ''} ${inFillPreview ? 'fill-preview' : ''} ${this._frozenClasses(row, c)}"
        role="gridcell"
        aria-colindex="${c + 2}"
        aria-selected="${isSelected}"
        aria-readonly="${this.readOnly || locked}"
        aria-current="${isActive ? 'true' : nothing}"
        data-row="${row}"
        data-col="${c}"
        data-key="${key}"
        tabindex="${isActive ? 0 : -1}"
        style=${styleMap(cellStyles)}
        @pointerdown="${this._handleCellPointerDown}"
        @dblclick="${this._handleCellDblClick}"
      >
        <span class="cell-text">${display}</span>
        ${showFillHandle
          ? html`<div
              class="ls-fill-handle"
              aria-hidden="true"
              @pointerdown="${this._handleFillHandlePointerDown}"
            ></div>`
          : nothing}
      </div>
    `;
  }

  private _renderRow(row: number, cols: number[]) {
    const range = this._selection.range;
    const inSelection = row >= range.start.row && row <= range.end.row;
    const fullRow = inSelection && range.start.col === 0 && range.end.col === this.cols - 1;
    return html`
      <div class="ls-row" role="row" aria-rowindex="${row + 2}">
        <div
          class="ls-row-header ${inSelection ? 'in-selection' : ''} ${this._frozenClasses(row, -1)}"
          style=${row < this._frozenRowCount ? `top: ${this._stickyTop(row)}` : nothing}
          role="rowheader"
          aria-colindex="1"
          aria-label="Row ${row + 1}"
          aria-selected="${fullRow ? 'true' : nothing}"
          data-header-row="${row}"
          @pointerdown="${this._handleHeaderPointerDown}"
        >${row + 1}</div>
        ${repeat(
          layoutWithGaps<{ key: string; tpl: unknown }>(
            cols,
            this.cols,
            (c) => ({ key: `c${c}`, tpl: this._renderCell(row, c) }),
            (start, size) => ({ key: `g${start}`, tpl: this._renderColumnSpacer(size) })
          ),
          // Keyed so the focused cell's element is never reused for another column
          (seg) => seg.key,
          (seg) => seg.tpl
        )}
      </div>
    `;
  }

  private _getAnnouncement(activeCell: CellCoord): string {
    const ref = `${colToLetter(activeCell.col)}${activeCell.row + 1}`;
    const display = this._getCellDisplay(activeCell.row, activeCell.col);
    const base = display ? `${ref}: ${display}` : ref;

    const range = this._selection.range;
    const selRows = range.end.row - range.start.row + 1;
    const selCols = range.end.col - range.start.col + 1;
    if (selRows > 1 || selCols > 1) {
      const allRows = selRows === this.rows;
      const allCols = selCols === this.cols;
      if (allRows && !allCols) {
        const first = colToLetter(range.start.col);
        const cols = selCols > 1 ? `Columns ${first} to ${colToLetter(range.end.col)}` : `Column ${first}`;
        return `${base}, ${cols} selected`;
      }
      if (allCols && !allRows) {
        const rows = selRows > 1 ? `Rows ${range.start.row + 1} to ${range.end.row + 1}` : `Row ${range.start.row + 1}`;
        return `${base}, ${rows} selected`;
      }
      return `${base}, ${selRows} row${selRows > 1 ? 's' : ''}, ${selCols} column${selCols > 1 ? 's' : ''} selected`;
    }

    return base;
  }

  protected updated(): void {
    if (this._isEditing && this._editor) {
      this._editor.focus();
      if (this._refMode) {
        // Place cursor after the inserted reference
        this._editor.setSelectionRange(this._refInsertEnd, this._refInsertEnd);
      } else if (this._editorNeedsCursorReset) {
        // Place cursor at the end only when editing first starts
        this._editorNeedsCursorReset = false;
        this._editor.setSelectionRange(this._editValue.length, this._editValue.length);
      }
    }
  }
}

declare global {
  interface HTMLElementTagNameMap {
    'y11n-spreadsheet': Y11nSpreadsheet;
  }
}
