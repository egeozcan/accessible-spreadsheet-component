import { type ReactiveController, type ReactiveControllerHost } from 'lit';
import { type CellCoord, type SelectionRange } from '../types.js';

/**
 * SelectionManager - Reactive Controller for managing grid selection state.
 *
 * Implements anchor/head selection model:
 * - anchor: where the selection started
 * - head: where the selection currently extends to
 * - range: the normalized bounding box
 */
export class SelectionManager implements ReactiveController {
  host: ReactiveControllerHost;

  /** The cell where the selection was initiated */
  anchor: CellCoord = { row: 0, col: 0 };

  /** The current end of the selection */
  head: CellCoord = { row: 0, col: 0 };

  /**
   * Active cell when it differs from `head` (e.g. after selecting whole
   * rows/columns, where the active cell stays put while the range spans the
   * full grid). `null` means the active cell is `head`.
   */
  private _active: CellCoord | null = null;

  /** Whether a drag selection is in progress */
  private _isDragging = false;

  /** Grid bounds */
  private _maxRows: number;
  private _maxCols: number;

  constructor(host: ReactiveControllerHost, maxRows: number, maxCols: number) {
    this.host = host;
    this._maxRows = maxRows;
    this._maxCols = maxCols;
    host.addController(this);
  }

  hostConnected(): void {
    // No-op; state is initialized in the constructor
  }

  hostDisconnected(): void {
    this._isDragging = false;
  }

  /** Update grid bounds (e.g., when rows/cols properties change) */
  setBounds(maxRows: number, maxCols: number): void {
    this._maxRows = maxRows;
    this._maxCols = maxCols;
    this.anchor = this.clamp(this.anchor);
    this.head = this.clamp(this.head);
    if (this._active) this._active = this.clamp(this._active);
  }

  /** Get the normalized selection range (top-left to bottom-right) */
  get range(): SelectionRange {
    return {
      start: {
        row: Math.min(this.anchor.row, this.head.row),
        col: Math.min(this.anchor.col, this.head.col),
      },
      end: {
        row: Math.max(this.anchor.row, this.head.row),
        col: Math.max(this.anchor.col, this.head.col),
      },
    };
  }

  /** The active-cell override, or `null` when the active cell follows the head. */
  get activeOverride(): CellCoord | null {
    return this._active ? { ...this._active } : null;
  }

  /** Get the currently active cell (the head, unless rows/columns were selected) */
  get activeCell(): CellCoord {
    return { ...(this._active ?? this.head) };
  }

  /** Check if a cell is within the current selection */
  isCellSelected(row: number, col: number): boolean {
    const r = this.range;
    return row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col;
  }

  /** Check if a cell is the active cell (head) */
  isCellActive(row: number, col: number): boolean {
    const active = this._active ?? this.head;
    return row === active.row && col === active.col;
  }

  // ─── Mouse Interaction ──────────────────────────────

  /**
   * Begin a new selection (mousedown/pointerdown on a cell).
   *
   * @param row - Row index of the clicked cell
   * @param col - Column index of the clicked cell
   * @param extend - If true (Shift+click), keeps the anchor and moves the head
   */
  startSelection(row: number, col: number, extend = false): void {
    this._active = null;
    if (extend) {
      // Shift+click: keep anchor, move head
      this.head = this.clamp({ row, col });
    } else {
      const clamped = this.clamp({ row, col });
      this.anchor = { ...clamped };
      this.head = { ...clamped };
    }
    this._isDragging = true;
    this.host.requestUpdate();
  }

  /**
   * Extend the selection during a drag (mousemove while dragging).
   *
   * @param row - Row index the mouse is over
   * @param col - Column index the mouse is over
   */
  extendSelection(row: number, col: number): void {
    if (!this._isDragging) return;
    this.head = this.clamp({ row, col });
    this.host.requestUpdate();
  }

  /** Finalize the drag selection (mouseup). Clears the dragging flag. */
  endSelection(): void {
    this._isDragging = false;
  }

  get isDragging(): boolean {
    return this._isDragging;
  }

  // ─── Keyboard Navigation ────────────────────────────

  /**
   * Move the active cell by a delta offset.
   *
   * @param dRow - Row delta (-1 = up, +1 = down)
   * @param dCol - Column delta (-1 = left, +1 = right)
   * @param extend - If true, keeps the anchor and extends the selection (Shift+arrow)
   */
  move(dRow: number, dCol: number, extend = false): void {
    // Extending grows the range from its head; plain moves start from the
    // active cell, which can differ from the head after a row/column select.
    const base = extend ? this.head : this.activeCell;
    const newHead = this.clamp({
      row: base.row + dRow,
      col: base.col + dCol,
    });

    this.head = newHead;
    if (!extend) {
      this.anchor = { ...newHead };
    }
    this._syncActive(extend);
    this.host.requestUpdate();
  }

  /**
   * Move to a specific cell by absolute position.
   *
   * @param row - Target row index
   * @param col - Target column index
   * @param extend - If true, keeps the anchor and extends the selection
   */
  moveTo(row: number, col: number, extend = false): void {
    const clamped = this.clamp({ row, col });
    this.head = clamped;
    if (!extend) {
      this.anchor = { ...clamped };
    }
    this._syncActive(extend);
    this.host.requestUpdate();
  }

  /**
   * Restore a full selection state (e.g. from an undo snapshot).
   *
   * @param anchor - Selection anchor
   * @param head - Selection head
   * @param active - Active-cell override; omit or pass `null` for the head
   */
  setSelection(anchor: CellCoord, head: CellCoord, active?: CellCoord | null): void {
    this.anchor = this.clamp(anchor);
    this.head = this.clamp(head);
    this._active = active ? this.clamp(active) : null;
    this._syncActive(true);
    this.host.requestUpdate();
  }

  /** Select all cells */
  selectAll(): void {
    this._active = null;
    this.anchor = { row: 0, col: 0 };
    this.head = { row: this._maxRows - 1, col: this._maxCols - 1 };
    this.host.requestUpdate();
  }

  /**
   * Select every row of the columns between `fromCol` and `toCol` (inclusive).
   * The head sits on `toCol`, so Shift+Left/Right keeps growing whole columns.
   *
   * @param fromCol - Column the selection is anchored at
   * @param toCol - Column the selection extends to
   * @param active - Cell to keep active; defaults to the current active row in `toCol`
   */
  selectColumns(fromCol: number, toCol: number, active?: CellCoord): void {
    const nextActive = active ?? { row: this.activeCell.row, col: toCol };
    this.anchor = this.clamp({ row: 0, col: fromCol });
    this.head = this.clamp({ row: this._maxRows - 1, col: toCol });
    this._active = this.clamp(nextActive);
    this._syncActive(true);
    this.host.requestUpdate();
  }

  /**
   * Select every column of the rows between `fromRow` and `toRow` (inclusive).
   * The head sits on `toRow`, so Shift+Up/Down keeps growing whole rows.
   *
   * @param fromRow - Row the selection is anchored at
   * @param toRow - Row the selection extends to
   * @param active - Cell to keep active; defaults to the current active column in `toRow`
   */
  selectRows(fromRow: number, toRow: number, active?: CellCoord): void {
    const nextActive = active ?? { row: toRow, col: this.activeCell.col };
    this.anchor = this.clamp({ row: fromRow, col: 0 });
    this.head = this.clamp({ row: toRow, col: this._maxCols - 1 });
    this._active = this.clamp(nextActive);
    this._syncActive(true);
    this.host.requestUpdate();
  }

  /** Clear selection (collapse to active cell) */
  clearSelection(): void {
    const active = this.activeCell;
    this._active = null;
    this.anchor = { ...active };
    this.head = { ...active };
    this.host.requestUpdate();
  }

  // ─── Helpers ────────────────────────────────────────

  /**
   * Keep the active-cell override only while extending and only while it is
   * still inside the range; otherwise the head is the active cell again.
   */
  private _syncActive(extend: boolean): void {
    if (!this._active) return;
    if (!extend || !this.isCellSelected(this._active.row, this._active.col)) {
      this._active = null;
    }
  }

  /**
   * Clamp a coordinate to the valid grid bounds.
   *
   * @param coord - The coordinate to clamp
   * @returns A new coordinate within `[0, maxRows-1]` x `[0, maxCols-1]`
   */
  clamp(coord: CellCoord): CellCoord {
    return {
      row: Math.max(0, Math.min(coord.row, this._maxRows - 1)),
      col: Math.max(0, Math.min(coord.col, this._maxCols - 1)),
    };
  }
}

/**
 * Find the target of a Ctrl/Cmd+Arrow jump, following spreadsheet semantics:
 *
 * - From a filled cell whose neighbour is also filled: move to the last
 *   filled cell of that contiguous run.
 * - Otherwise: move to the next filled cell in that direction, or to the
 *   grid edge if there is none.
 *
 * @param from - Starting cell
 * @param dRow - Row direction (-1, 0, or 1)
 * @param dCol - Column direction (-1, 0, or 1)
 * @param rows - Grid row count
 * @param cols - Grid column count
 * @param isFilled - Whether the cell at (row, col) has content
 * @returns The cell to jump to (equal to `from` when already at the edge)
 */
export function findDataEdge(
  from: CellCoord,
  dRow: number,
  dCol: number,
  rows: number,
  cols: number,
  isFilled: (row: number, col: number) => boolean
): CellCoord {
  const inBounds = (r: number, c: number) => r >= 0 && r < rows && c >= 0 && c < cols;
  let { row, col } = from;
  if (!inBounds(row + dRow, col + dCol)) return { row, col };

  if (isFilled(row, col) && isFilled(row + dRow, col + dCol)) {
    while (inBounds(row + dRow, col + dCol) && isFilled(row + dRow, col + dCol)) {
      row += dRow;
      col += dCol;
    }
    return { row, col };
  }

  row += dRow;
  col += dCol;
  while (!isFilled(row, col) && inBounds(row + dRow, col + dCol)) {
    row += dRow;
    col += dCol;
  }
  return { row, col };
}
