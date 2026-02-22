import { LitElement, html, css, nothing } from 'lit';
import { customElement, property, state, query } from 'lit/decorators.js';
import { repeat } from 'lit/directives/repeat.js';
import {
  type CellCoord,
  type CellData,
  type GridData,
  type FormulaFunction,
  type CellChangeDetail,
  type SelectionChangeDetail,
  type DataChangeDetail,
  cellKey,
  colToLetter,
  coordToRef,
} from './types.js';
import { SelectionManager } from './controllers/selection-manager.js';
import { FormulaEngine } from './engine/formula-engine.js';
import { ClipboardManager } from './controllers/clipboard-manager.js';
import './components/y11n-formula-bar.js';
import type { FormulaBarMode } from './components/y11n-formula-bar.js';

type DataChangeOperation = NonNullable<DataChangeDetail['operation']>;
type ChangeSource = NonNullable<DataChangeDetail['source']>;

interface CellDelta {
  id: string;
  before: string;
  after: string;
}

interface SelectionSnapshot {
  anchor: CellCoord;
  head: CellCoord;
}

interface CommandBatch {
  deltas: CellDelta[];
  selectionBefore: SelectionSnapshot;
  selectionAfter: SelectionSnapshot;
  operation: DataChangeOperation;
}

/**
 * <y11n-spreadsheet> - An accessible spreadsheet web component.
 *
 * Implements the WAI-ARIA Grid pattern with roving tabindex,
 * keyboard navigation, formula support, and clipboard operations.
 */
@customElement('y11n-spreadsheet')
export class Y11nSpreadsheet extends LitElement {
  // ─── Public Properties ──────────────────────────────

  @property({ type: Number }) rows = 100;
  @property({ type: Number }) cols = 26;
  @property({ attribute: false }) data: GridData = new Map();
  @property({ type: Boolean, attribute: 'read-only', reflect: true }) readOnly = false;
  @property({ attribute: false }) functions: Record<string, FormulaFunction> = {};

  // ─── Internal State ─────────────────────────────────

  @state() private _isEditing = false;
  @state() private _editValue = '';
  @state() private _formulaBarMode: FormulaBarMode = 'raw';

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

  /** Track which cell started editing for commit logic */
  private _editingCellKey: string | null = null;

  private _undoStack: CommandBatch[] = [];
  private _redoStack: CommandBatch[] = [];
  private readonly _maxHistory = 100;

  // ─── Lifecycle ──────────────────────────────────────

  connectedCallback(): void {
    super.connectedCallback();
    this._syncData();
  }

  willUpdate(changedProps: Map<string, unknown>): void {
    if (changedProps.has('data')) {
      this._syncData();
    }
    if (changedProps.has('rows') || changedProps.has('cols')) {
      this.rows = Math.max(1, Math.round(this.rows));
      this.cols = Math.max(1, Math.round(this.cols));
      this._selection.setBounds(this.rows, this.cols);
    }
    if (changedProps.has('functions')) {
      for (const [name, fn] of Object.entries(this.functions)) {
        this._formulaEngine.registerFunction(name, fn as FormulaFunction);
      }
    }
  }

  // ─── Public API Methods ─────────────────────────────

  /** Returns the current state of the grid data */
  getData(): GridData {
    return new Map(this._internalData);
  }

  /** Hard reset of the grid data */
  setData(data: GridData): void {
    this.data = data;
    this._syncData();
    this.requestUpdate();
  }

  /** Dynamically register a formula function */
  registerFunction(name: string, fn: FormulaFunction): void {
    this._formulaEngine.registerFunction(name, fn);
    this._recalcAll();
  }

  // ─── Data Sync ──────────────────────────────────────

  private _syncData(): void {
    this._internalData = new Map(this.data);
    this._undoStack = [];
    this._redoStack = [];
    this._formulaEngine.setData(this._internalData);
    this._recalcAll();
  }

  private _recalcAll(): void {
    this._formulaEngine.recalculate();
  }

  private _recalcAffected(changedKeys: string[]): void {
    this._formulaEngine.recalculateAffected(changedKeys);
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

    this._internalData.set(key, {
      rawValue,
      displayValue: evaluated.displayValue,
      type: evaluated.type,
      style: existing?.style,
    });
  }

  private _applyRawValueByString(key: string, rawValue: string): void {
    if (rawValue === '') {
      this._internalData.delete(key);
      return;
    }
    this._setCellRaw(key, rawValue);
  }

  private _snapshotSelection(): SelectionSnapshot {
    return {
      anchor: { ...this._selection.anchor },
      head: { ...this._selection.head },
    };
  }

  private _restoreSelection(snapshot: SelectionSnapshot): void {
    this._selection.moveTo(snapshot.anchor.row, snapshot.anchor.col);
    this._selection.moveTo(snapshot.head.row, snapshot.head.col, true);
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
    }
    this._restoreSelection(batch.selectionAfter);
    this._finalizeBatch(batch, source, 'after');
  }

  private _revertBatch(batch: CommandBatch): void {
    for (const delta of batch.deltas) {
      this._applyRawValueByString(delta.id, delta.before);
    }
    this._restoreSelection(batch.selectionBefore);
    this._finalizeBatch(batch, 'undo', 'before');
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
    if (this.readOnly) return;

    const batch = this._undoStack.pop();
    if (!batch) return;

    this._revertBatch(batch);
    this._redoStack.push(batch);
  }

  private _redo(): void {
    if (this.readOnly) return;

    const batch = this._redoStack.pop();
    if (!batch) return;

    this._applyBatch(batch, 'redo');
    this._undoStack.push(batch);
    if (this._undoStack.length > this._maxHistory) {
      this._undoStack.shift();
    }
  }

  // ─── Editing ────────────────────────────────────────

  private _startEditing(initialValue?: string): void {
    if (this.readOnly || this._isEditing) return;

    const { row, col } = this._selection.activeCell;
    const key = cellKey(row, col);
    this._editingCellKey = key;

    const cell = this._internalData.get(key);
    this._editValue = initialValue ?? cell?.rawValue ?? '';
    this._isEditing = true;
  }

  private _commitEdit(): void {
    if (!this._isEditing || !this._editingCellKey) return;

    const key = this._editingCellKey;
    const newValue = this._editValue;

    this._isEditing = false;
    this._editingCellKey = null;

    this._commitRawValue(key, newValue);

    // Return focus to the grid
    this.updateComplete.then(() => this._focusActiveCell());
  }

  private _cancelEdit(): void {
    this._isEditing = false;
    this._editingCellKey = null;
    this.updateComplete.then(() => this._focusActiveCell());
  }

  // ─── Focus Management ───────────────────────────────

  private _focusActiveCell(): void {
    const { row, col } = this._selection.activeCell;
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

    // Stop propagation for all handled keys so parent contexts
    // (e.g. Storybook shortcuts) don't intercept them.
    e.stopPropagation();

    switch (e.key) {
      case 'ArrowUp':
        e.preventDefault();
        this._selection.move(-1, 0, shift);
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'ArrowDown':
        e.preventDefault();
        this._selection.move(1, 0, shift);
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'ArrowLeft':
        e.preventDefault();
        this._selection.move(0, -1, shift);
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'ArrowRight':
        e.preventDefault();
        this._selection.move(0, 1, shift);
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'Tab':
        e.preventDefault();
        if (shift) {
          this._selection.move(0, -1);
        } else {
          this._selection.move(0, 1);
        }
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'Enter':
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

      case 'a':
        if (ctrl) {
          e.preventDefault();
          this._selection.selectAll();
          this._dispatchSelectionChange();
        }
        break;

      case 'c':
        if (ctrl) {
          e.preventDefault();
          this._handleCopy();
        }
        break;

      case 'x':
        if (ctrl) {
          e.preventDefault();
          this._handleCut();
        }
        break;

      case 'v':
        if (ctrl) {
          e.preventDefault();
          this._handlePaste();
        }
        break;

      case 'z':
      case 'Z':
        if (ctrl) {
          e.preventDefault();
          if (shift) {
            this._redo();
          } else {
            this._undo();
          }
        }
        break;

      default:
        // Start editing on printable character input
        if (!ctrl && !e.altKey && e.key.length === 1 && !this.readOnly) {
          e.preventDefault();
          this._startEditing(e.key);
        }
        break;
    }
  }

  private _handleEditorKeydown(e: KeyboardEvent): void {
    e.stopPropagation();

    switch (e.key) {
      case 'Enter':
        e.preventDefault();
        this._commitEdit();
        // Move down after Enter commit
        this._selection.move(1, 0);
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'Tab':
        e.preventDefault();
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
        this._cancelEdit();
        break;
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
    if (this.readOnly) return;

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
    if (this.readOnly) return;

    const { row, col } = this._selection.activeCell;
    const updates = await this._clipboardManager.paste(row, col, this.rows, this.cols);

    if (updates && updates.length > 0) {
      const selection = this._snapshotSelection();
      const batch = this._buildCommandBatch(updates, 'paste', selection, selection);
      if (batch) {
        this._executeUserBatch(batch);
      }
    }
  }

  // ─── Cell Clearing ──────────────────────────────────

  private _clearSelectedCells(): void {
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

  // ─── Scroll Handling ────────────────────────────────

  private _handleScroll(e: Event): void {
    const target = e.target as HTMLElement;
    this._scrollTop = target.scrollTop;
    this._scrollLeft = target.scrollLeft;
  }

  // ─── Editor Input ───────────────────────────────────

  private _handleEditorInput(e: Event): void {
    this._editValue = (e.target as HTMLInputElement).value;
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

  private _handleFormulaBarCommit(e: CustomEvent<{ value: string }>): void {
    if (this.readOnly) return;

    if (this._isEditing) {
      this._isEditing = false;
      this._editingCellKey = null;
    }

    this._commitRawValue(this._getActiveCellKey(), e.detail.value);
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

  // ─── Virtual Rendering ──────────────────────────────

  /**
   * Compute which rows and columns are visible based on scroll position.
   * We add a buffer of extra rows/cols for smooth scrolling.
   */
  private _getVisibleRange(): { startRow: number; endRow: number; startCol: number; endCol: number } {
    const gridEl = this._grid;
    if (!gridEl) {
      return { startRow: 0, endRow: Math.min(this.rows, 40), startCol: 0, endCol: this.cols };
    }

    const cellHeight = this._getCSSVarPx('--ls-cell-height', 28);
    const cellWidth = this._getCSSVarPx('--ls-cell-width', 100);
    const viewHeight = gridEl.clientHeight;
    const viewWidth = gridEl.clientWidth;

    const buffer = 5;

    const startRow = Math.max(0, Math.floor(this._scrollTop / cellHeight) - buffer);
    const endRow = Math.min(this.rows, Math.ceil((this._scrollTop + viewHeight) / cellHeight) + buffer);
    const startCol = Math.max(0, Math.floor(this._scrollLeft / cellWidth) - buffer);
    const endCol = Math.min(this.cols, Math.ceil((this._scrollLeft + viewWidth) / cellWidth) + buffer);

    return { startRow, endRow, startCol, endCol };
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

  private _getEditorStyle(): string {
    if (!this._isEditing) return 'display:none;';

    const { row, col } = this._selection.activeCell;
    const cellHeight = this._getCSSVarPx('--ls-cell-height', 28);
    const cellWidth = this._getCSSVarPx('--ls-cell-width', 100);
    const headerWidth = 50; // row header width

    const top = (row + 1) * cellHeight; // +1 for the column header row
    const left = headerWidth + col * cellWidth;

    return `
      display: block;
      position: absolute;
      top: ${top}px;
      left: ${left}px;
      width: ${cellWidth}px;
      height: ${cellHeight}px;
      z-index: 10;
    `;
  }

  // ─── Render ─────────────────────────────────────────

  static styles = css`
    :host {
      display: block;
      position: relative;
      font-family: var(--ls-font-family, system-ui, -apple-system, sans-serif);
      font-size: var(--ls-font-size, 13px);
      color: var(--ls-text-color, #333);
      --_cell-width: var(--ls-cell-width, 100px);
      --_cell-height: var(--ls-cell-height, 28px);
      --_header-bg: var(--ls-header-bg, #f3f3f3);
      --_border-color: var(--ls-border-color, #e0e0e0);
      --_selection-border: var(--ls-selection-border, 2px solid #1a73e8);
      --_selection-bg: var(--ls-selection-bg, rgba(26, 115, 232, 0.1));
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
      overflow: auto;
      position: relative;
      width: 100%;
      height: 100%;
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
    const { startRow, endRow, startCol, endCol } = this._getVisibleRange();
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
            style="${this._getEditorStyle()}"
            .value="${this._editValue}"
            @input="${this._handleEditorInput}"
            @keydown="${this._handleEditorKeydown}"
            @blur="${this._handleEditorBlur}"
            aria-label="Cell editor"
          />

          <div
            class="ls-grid"
            role="grid"
            aria-rowcount="${this.rows + 1}"
            aria-colcount="${this.cols + 1}"
            aria-readonly="${this.readOnly}"
            style="--cols: ${this.cols};"
            tabindex="-1"
            @keydown="${this._handleGridKeydown}"
            @scroll="${this._handleScroll}"
          >
            ${this._renderHeaderRow(startCol, endCol)}
            ${this._renderRows(startRow, endRow, startCol, endCol)}
          </div>
        </div>

        <div class="sr-only" role="status" aria-live="polite" aria-atomic="true">
          ${this._getAnnouncement(activeCell)}
        </div>
      </div>
    `;
  }

  private _renderHeaderRow(startCol: number, endCol: number) {
    const headers = [];
    for (let c = startCol; c < endCol; c++) {
      headers.push({ letter: colToLetter(c), index: c });
    }

    const cellWidth = this._getCSSVarPx('--ls-cell-width', 100);

    return html`
      <div class="ls-header-row" role="row" aria-rowindex="1">
        <div class="ls-corner-header" role="columnheader"></div>
        ${startCol > 0
          ? html`<div style="width: ${startCol * cellWidth}px;"></div>`
          : nothing}
        ${headers.map(
          (h) => html`
            <div
              class="ls-col-header"
              role="columnheader"
              aria-colindex="${h.index + 2}"
            >
              ${h.letter}
            </div>
          `
        )}
        ${endCol < this.cols
          ? html`<div style="width: ${(this.cols - endCol) * cellWidth}px;"></div>`
          : nothing}
      </div>
    `;
  }

  private _renderRows(startRow: number, endRow: number, startCol: number, endCol: number) {
    const cellHeight = this._getCSSVarPx('--ls-cell-height', 28);
    const rowIndices = Array.from({ length: endRow - startRow }, (_, i) => startRow + i);

    return html`
      ${startRow > 0
        ? html`<div style="grid-column: 1 / -1; height: ${startRow * cellHeight}px;"></div>`
        : nothing}
      ${repeat(rowIndices, (r) => r, (r) => this._renderRow(r, startCol, endCol))}
      ${endRow < this.rows
        ? html`<div style="grid-column: 1 / -1; height: ${(this.rows - endRow) * cellHeight}px;"></div>`
        : nothing}
    `;
  }

  private _renderRow(row: number, startCol: number, endCol: number) {
    const cellWidth = this._getCSSVarPx('--ls-cell-width', 100);
    const cells = [];
    for (let c = startCol; c < endCol; c++) {
      const isSelected = this._selection.isCellSelected(row, c);
      const isActive = this._selection.isCellActive(row, c);
      const display = this._getCellDisplay(row, c);
      const key = cellKey(row, c);

      cells.push(html`
        <div
          class="ls-cell ${isActive ? 'active-cell' : ''}"
          role="gridcell"
          aria-colindex="${c + 2}"
          aria-selected="${isSelected}"
          aria-readonly="${this.readOnly}"
          data-row="${row}"
          data-col="${c}"
          data-key="${key}"
          tabindex="${isActive ? 0 : -1}"
          @pointerdown="${this._handleCellPointerDown}"
          @dblclick="${this._handleCellDblClick}"
        >
          <span class="cell-text">${display}</span>
        </div>
      `);
    }

    return html`
      <div class="ls-row" role="row" aria-rowindex="${row + 2}">
        <div class="ls-row-header" role="rowheader">${row + 1}</div>
        ${startCol > 0
          ? html`<div style="width: ${startCol * cellWidth}px;"></div>`
          : nothing}
        ${cells}
        ${endCol < this.cols
          ? html`<div style="width: ${(this.cols - endCol) * cellWidth}px;"></div>`
          : nothing}
      </div>
    `;
  }

  private _getAnnouncement(activeCell: CellCoord): string {
    const ref = `${colToLetter(activeCell.col)}${activeCell.row + 1}`;
    const display = this._getCellDisplay(activeCell.row, activeCell.col);
    return display ? `${ref}: ${display}` : ref;
  }

  protected updated(): void {
    if (this._isEditing && this._editor) {
      const editor = this._editor;
      const len = this._editValue.length;
      editor.focus();
      editor.setSelectionRange(len, len);
    }
  }
}

declare global {
  interface HTMLElementTagNameMap {
    'y11n-spreadsheet': Y11nSpreadsheet;
  }
}
