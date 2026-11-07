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
}

interface CommandBatch {
  deltas: CellDelta[];
  selectionBefore: SelectionSnapshot;
  selectionAfter: SelectionSnapshot;
  operation: DataChangeOperation;
  /** Set for row/column insertions and deletions, which replace whole cells */
  structure?: StructureBatch;
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

  // ─── Lifecycle ──────────────────────────────────────

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
    const generation = this._dataGeneration;
    this._applyBatch(batch, 'programmatic');
    this._pushHistory(batch, generation);
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
      const generation = this._dataGeneration;
      this._applyBatch(batch, source);
      this._pushHistory(batch, generation);
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
      const generation = this._dataGeneration;
      this._applyBatch(batch, source);
      this._pushHistory(batch, generation);
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
  private _applyNumberFormatToCell(_key: string, cell: CellData): void {
    if (cell.type === 'number' && cell.format?.numberFormat) {
      const numericValue = cell.rawValue.startsWith('=')
        ? Number(this._formulaEngine.evaluate(cell.rawValue).displayValue)
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
    this._mutationCount++;
    if (batch.structure) {
      this._applyStructureBatch(batch, 'after', source);
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
    this._mutationCount++;
    if (batch.structure) {
      this._applyStructureBatch(batch, 'before', 'undo');
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
      // For formula cells, re-evaluate to get the raw numeric value
      // (displayValue may already be formatted, e.g., "$1,234.50")
      if (cell.rawValue.startsWith('=')) {
        const evaluated = this._formulaEngine.evaluate(cell.rawValue);
        cell.displayValue = formatNumber(Number(evaluated.displayValue), cell.format.numberFormat);
      } else {
        cell.displayValue = formatNumber(Number(cell.rawValue), cell.format.numberFormat);
      }
    } else if (cell.type === 'number' && !cell.format?.numberFormat) {
      // Remove number formatting — re-evaluate to get plain display
      const evaluated = this._formulaEngine.evaluate(cell.rawValue);
      cell.displayValue = evaluated.displayValue;
    }
  }

  private _executeUserBatch(batch: CommandBatch): void {
    const generation = this._dataGeneration;
    this._applyBatch(batch, 'user');
    this._pushHistory(batch, generation);
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
      },
    };
    const generation = this._dataGeneration;
    this._applyBatch(batch, source);
    this._pushHistory(batch, generation);

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
    this.rows = side === 'after' ? structure.rowsAfter : structure.rowsBefore;
    this.cols = side === 'after' ? structure.colsAfter : structure.colsBefore;
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
    const { change } = structure;
    const kind = side === 'after' ? change.kind : change.kind === 'insert' ? 'delete' : 'insert';
    this._dispatchStructureChange({ ...change, kind, source });
    this.requestUpdate();
  }

  private _undo(): void {
    if (this.readOnly || this._pasteInProgress) return;

    const batch = this._undoStack.pop();
    if (!batch) return;

    const generation = this._dataGeneration;
    this._revertBatch(batch);
    if (generation === this._dataGeneration) this._redoStack.push(batch);
    this._dispatchFormatChangeForBatch(batch, 'undo');
  }

  private _redo(): void {
    if (this.readOnly || this._pasteInProgress) return;

    const batch = this._redoStack.pop();
    if (!batch) return;

    const generation = this._dataGeneration;
    this._applyBatch(batch, 'redo');
    if (generation === this._dataGeneration) this._undoStack.push(batch);
    if (this._undoStack.length > this._maxHistory) {
      this._undoStack.shift();
    }
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
          this._selection.moveTo(0, 0, shift);
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
        const cellHeight = this._getCSSVarPx('--ls-cell-height', 28);
        const gridEl = this._grid;
        const viewHeight = gridEl?.clientHeight || cellHeight * Y11nSpreadsheet.INITIAL_VIEW_ROWS;
        const pageRows = Math.max(1, Math.floor(viewHeight / cellHeight) - 1);
        this._selection.move(pageRows, 0, shift);
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;
      }

      case 'PageUp': {
        e.preventDefault();
        const cellHeight = this._getCSSVarPx('--ls-cell-height', 28);
        const gridEl = this._grid;
        const viewHeight = gridEl?.clientHeight || cellHeight * Y11nSpreadsheet.INITIAL_VIEW_ROWS;
        const pageRows = Math.max(1, Math.floor(viewHeight / cellHeight) - 1);
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

  // ─── Virtual Rendering ──────────────────────────────

  /**
   * Compute which rows and columns are visible based on scroll position.
   * We add a buffer of extra rows/cols for smooth scrolling.
   */
  private _getVisibleRange(): { rows: number[]; cols: number[] } {
    const cellHeight = this._getCSSVarPx('--ls-cell-height', 28);
    const cellWidth = this._getCSSVarPx('--ls-cell-width', 100);
    const gridEl = this._grid;
    const viewHeight =
      gridEl?.clientHeight || cellHeight * Y11nSpreadsheet.INITIAL_VIEW_ROWS;
    const viewWidth =
      gridEl?.clientWidth || cellWidth * Y11nSpreadsheet.INITIAL_VIEW_COLS;

    const buffer = 5;

    const startRow = Math.max(0, Math.floor(this._scrollTop / cellHeight) - buffer);
    const endRow = Math.min(this.rows, Math.ceil((this._scrollTop + viewHeight) / cellHeight) + buffer);
    const startCol = Math.max(0, Math.floor(this._scrollLeft / cellWidth) - buffer);
    const endCol = Math.min(this.cols, Math.ceil((this._scrollLeft + viewWidth) / cellWidth) + buffer);

    // The active cell (roving tabindex target) and the reference cursor must
    // always be in the DOM. Render them individually rather than stretching
    // the window, which would render every row between the viewport and a
    // far-away active cell after the user scrolls with the mouse.
    const { row: activeRow, col: activeCol } = this._selection.activeCell;
    const extraRows = [activeRow];
    const extraCols = [activeCol];
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

    const cellHeight = this._getCSSVarPx('--ls-cell-height', 28);
    const cellWidth = this._getCSSVarPx('--ls-cell-width', 100);
    const viewHeight =
      gridEl.clientHeight || cellHeight * Y11nSpreadsheet.INITIAL_VIEW_ROWS;
    const viewWidth =
      gridEl.clientWidth || cellWidth * Y11nSpreadsheet.INITIAL_VIEW_COLS;
    const dataViewportWidth = Math.max(
      cellWidth,
      viewWidth - Y11nSpreadsheet.ROW_HEADER_WIDTH
    );
    // Subtract the sticky column header height so cells aren't scrolled
    // behind it (mirrors how dataViewportWidth subtracts the row header).
    const dataViewportHeight = Math.max(cellHeight, viewHeight - cellHeight);

    const cellTop = row * cellHeight;
    const cellBottom = cellTop + cellHeight;
    const cellLeft = col * cellWidth;
    const cellRight = cellLeft + cellWidth;

    let nextScrollTop = this._scrollTop;
    if (cellTop < nextScrollTop) {
      nextScrollTop = cellTop;
    } else if (cellBottom > nextScrollTop + dataViewportHeight) {
      nextScrollTop = cellBottom - dataViewportHeight;
    }

    let nextScrollLeft = this._scrollLeft;
    if (cellLeft < nextScrollLeft) {
      nextScrollLeft = cellLeft;
    } else if (cellRight > nextScrollLeft + dataViewportWidth) {
      nextScrollLeft = cellRight - dataViewportWidth;
    }

    if (nextScrollTop !== this._scrollTop || nextScrollLeft !== this._scrollLeft) {
      this._scrollTop = nextScrollTop;
      this._scrollLeft = nextScrollLeft;
      gridEl.scrollTop = nextScrollTop;
      gridEl.scrollLeft = nextScrollLeft;
    }
  }

  private _getCSSVarPx(name: string, fallback: number): number {
    const val = getComputedStyle(this).getPropertyValue(name);
    if (val) {
      const parsed = parseInt(val, 10);
      if (!isNaN(parsed)) return parsed;
    }
    return fallback;
  }

  // ─── Editor Positioning ─────────────────────────────

  private _getEditorStyle(): Record<string, string> {
    if (!this._isEditing) return { display: 'none' };

    const { row, col } = this._selection.activeCell;
    const cellHeight = this._getCSSVarPx('--ls-cell-height', 28);
    const cellWidth = this._getCSSVarPx('--ls-cell-width', 100);
    const headerWidth = Y11nSpreadsheet.ROW_HEADER_WIDTH;

    const top = (row + 1) * cellHeight - this._scrollTop; // +1 for header, offset by scroll
    const left = headerWidth + col * cellWidth - this._scrollLeft;

    const styles: Record<string, string> = {
      display: 'block',
      position: 'absolute',
      top: `${top}px`,
      left: `${left}px`,
      width: `${cellWidth}px`,
      height: `${cellHeight}px`,
      'z-index': '10',
    };

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
    const cellHeight = this._getCSSVarPx('--ls-cell-height', 28);
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

    .ls-corner-header,
    .ls-col-header {
      position: sticky;
      top: 0;
      z-index: 3;
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
      z-index: 4;
      scroll-snap-align: start;
    }

    .ls-col-header,
    .ls-row-header {
      cursor: pointer;
    }

    .ls-col-header.in-selection,
    .ls-row-header.in-selection {
      background: var(--ls-header-selected-bg, #dbe7fb);
    }

    .ls-row-header {
      position: sticky;
      left: 0;
      z-index: 2;
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
            style="--cols: ${this.cols};"
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
              class="ls-col-header ${c >= range.start.col && c <= range.end.col ? 'in-selection' : ''}"
              role="columnheader"
              aria-colindex="${c + 2}"
              aria-label="Column ${colToLetter(c)}"
              aria-selected="${fullColumns && c >= range.start.col && c <= range.end.col ? 'true' : nothing}"
              data-header-col="${c}"
              @pointerdown="${this._handleHeaderPointerDown}"
            >
              ${colToLetter(c)}
            </div>
          `,
          (_start, size) => this._renderColumnSpacer(size)
        )}
      </div>
    `;
  }

  private _renderRows(rows: number[], cols: number[]) {
    const cellHeight = this._getCSSVarPx('--ls-cell-height', 28);
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

    return html`
      <div
        class="ls-cell ${isActive ? 'active-cell' : ''} ${isRefTarget ? 'ref-highlight' : ''} ${locked ? 'read-only-cell' : ''} ${inFillPreview ? 'fill-preview' : ''}"
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
          class="ls-row-header ${inSelection ? 'in-selection' : ''}"
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
