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
import { FormulaEngine } from './engine/formula-engine.js';
import { ClipboardManager } from './controllers/clipboard-manager.js';
import './components/y11n-formula-bar.js';
import type { FormulaBarMode } from './components/y11n-formula-bar.js';

type DataChangeOperation = NonNullable<DataChangeDetail['operation']>;
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

interface CommandBatch {
  deltas: CellDelta[];
  selectionBefore: SelectionSnapshot;
  selectionAfter: SelectionSnapshot;
  operation: DataChangeOperation;
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
 *
 * @csspart editor - The inline cell editor input element
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

  /** Track scroll position for virtual rendering */
  @state() private _scrollTop = 0;
  @state() private _scrollLeft = 0;

  // ─── DOM Refs ───────────────────────────────────────

  @query('#editor') private _editor!: HTMLInputElement;
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
    this._applyBatch(batch, 'programmatic');
    this._pushHistory(batch);
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

    for (let r = range.start.row; r <= range.end.row; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        const id = cellKey(r, c);
        cellIds.push(id);
        const cell = this._internalData.get(id);
        if (!cell?.format?.[prop]) {
          allHaveProp = false;
        }
      }
    }

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
      this._applyBatch(batch, source);
      this._pushHistory(batch);
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
      this._applyBatch(batch, source);
      this._pushHistory(batch);
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

    this._internalData.set(key, {
      rawValue,
      displayValue,
      type: evaluated.type,
      format: existing?.format,
    });
  }

  private _applyRawValueByString(key: string, rawValue: string): void {
    if (rawValue === '') {
      const existing = this._internalData.get(key);
      if (existing?.format) {
        // Preserve format on empty cells
        this._internalData.set(key, {
          rawValue: '',
          displayValue: '',
          type: 'text',
          format: existing.format,
        });
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
    try {
      await this._applyBatchProgressive(batch);
      this._pushHistory(batch);
    } finally {
      this._pasteInProgress = false;
    }
  }

  private _pushHistory(batch: CommandBatch): void {
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
      if (cell.rawValue === '') {
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
    this._applyBatch(batch, 'user');
    this._pushHistory(batch);
  }

  private _commitRawValue(key: string, newValue: string): boolean {
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

    if (!batch) return false;
    this._executeUserBatch(batch);
    this._dispatchCellChange({ cellId: key, value: newValue, oldValue });
    return true;
  }

  private _undo(): void {
    if (this.readOnly || this._pasteInProgress) return;

    const batch = this._undoStack.pop();
    if (!batch) return;

    this._revertBatch(batch);
    this._redoStack.push(batch);
    this._dispatchFormatChangeForBatch(batch, 'undo');
  }

  private _redo(): void {
    if (this.readOnly || this._pasteInProgress) return;

    const batch = this._redoStack.pop();
    if (!batch) return;

    this._applyBatch(batch, 'redo');
    this._undoStack.push(batch);
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
    this._editingCellKey = key;

    const cell = this._internalData.get(key);
    this._editValue = initialValue ?? cell?.rawValue ?? '';
    this._isEditing = true;
    this._editorNeedsCursorReset = true;
    this._modeAnnouncement = `Editing ${coordToRef(this._selection.activeCell)}`;
  }

  private _commitEdit(): void {
    if (!this._isEditing || !this._editingCellKey) return;

    const key = this._editingCellKey;
    const newValue = this._editValue;

    this._isEditing = false;
    this._editingCellKey = null;
    this._modeAnnouncement = 'Edit committed';

    this._commitRawValue(key, newValue);

    // Return focus to the grid
    this.updateComplete.then(() => this._focusActiveCell());
  }

  private _cancelEdit(): void {
    this._isEditing = false;
    this._editingCellKey = null;
    this._modeAnnouncement = 'Edit cancelled';
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

      case 'd':
        if (ctrl && !shift) {
          e.preventDefault();
          this._fillSelection('down');
        } else {
          handled = false;
        }
        break;

      case 'r':
        if (ctrl && !shift) {
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
        this._commitEdit();
        // Enter moves down after committing, Shift+Enter moves up
        this._selection.move(e.shiftKey ? -1 : 1, 0);
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'Tab':
        e.preventDefault();
        this._exitRefMode();
        this._commitEdit();
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

      case 'F4':
        if (this._editValue.startsWith('=')) {
          e.preventDefault();
          this._cycleReferenceMode();
        }
        break;
    }
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

    const keysToDelete = await this._clipboardManager.cut(
      this._internalData,
      this._selection.range
    );

    const updates: Array<{ id: string; value: string }> = [];
    const selection = this._snapshotSelection();
    for (const key of keysToDelete) {
      if (this._internalData.has(key)) {
        updates.push({ id: key, value: '' });
      }
    }

    const batch = this._buildCommandBatch(updates, 'cut', selection, selection);
    if (batch) {
      this._executeUserBatch(batch);
    }
  }

  private async _handlePaste(): Promise<void> {
    if (this.readOnly || this._pasteInProgress) return;

    // Paste anchors at the top-left of the selection (not the head, which may
    // be the bottom-right corner after a Shift+arrow or drag selection).
    const { row, col } = this._selection.range.start;
    const updates = await this._clipboardManager.paste(row, col, this.rows, this.cols);

    if (!updates || updates.length === 0) return;

    const selection = this._snapshotSelection();
    // Pasted formats replace the target's; cells pasted without one keep theirs.
    const batch = this._buildCellBatch(
      updates.map((u) => ({ id: u.id, value: u.value, format: u.format })),
      'paste',
      selection
    );
    if (!batch) return;
    await this._runUserBatch(batch);
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
    if (batch) {
      this._executeUserBatch(batch);
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
    if (batch) void this._runUserBatch(batch);
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
      this._commitEdit();
      return;
    }

    this._isEditing = false;
    this._editingCellKey = null;

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
    if (batch) void this._runUserBatch(batch);
    this.updateComplete.then(() => this._focusActiveCell());
  }

  // ─── Scroll Handling ────────────────────────────────

  private _handleScroll(e: Event): void {
    const target = e.target as HTMLElement;
    this._scrollTop = target.scrollTop;
    this._scrollLeft = target.scrollLeft;
  }

  // ─── Editor Input ───────────────────────────────────

  private _handleEditorInput(e: Event): void {
    this._editValue = (e.target as HTMLInputElement).value;
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
    if (this.readOnly) return;

    // Use the cell key from the event (which cell the formula bar was showing)
    // rather than the current active cell, which may have changed due to clicking.
    let targetKey: string;
    if (this._editingCellKey) {
      targetKey = this._editingCellKey;
    } else {
      const coord = refToCoord(e.detail.cellRef);
      targetKey = coord ? cellKey(coord.row, coord.col) : this._getActiveCellKey();
    }

    if (this._isEditing) {
      this._isEditing = false;
      this._editingCellKey = null;
    }

    this._commitRawValue(targetKey, e.detail.value);
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
          .readOnly="${this.readOnly}"
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
          />

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

    const cellStyles = this._getCellStyles(row, c);

    return html`
      <div
        class="ls-cell ${isActive ? 'active-cell' : ''} ${isRefTarget ? 'ref-highlight' : ''}"
        role="gridcell"
        aria-colindex="${c + 2}"
        aria-selected="${isSelected}"
        aria-readonly="${this.readOnly}"
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
