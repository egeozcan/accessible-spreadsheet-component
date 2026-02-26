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

  /** Get the currently active cell (head of selection) */
  get activeCell(): CellCoord {
    return { ...this.head };
  }

  /** Check if a cell is within the current selection */
  isCellSelected(row: number, col: number): boolean {
    const r = this.range;
    return row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col;
  }

  /** Check if a cell is the active cell (head) */
  isCellActive(row: number, col: number): boolean {
    return row === this.head.row && col === this.head.col;
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
    const newHead = this.clamp({
      row: this.head.row + dRow,
      col: this.head.col + dCol,
    });

    this.head = newHead;
    if (!extend) {
      this.anchor = { ...newHead };
    }
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
    this.host.requestUpdate();
  }

  /** Select all cells */
  selectAll(): void {
    this.anchor = { row: 0, col: 0 };
    this.head = { row: this._maxRows - 1, col: this._maxCols - 1 };
    this.host.requestUpdate();
  }

  /** Clear selection (collapse to active cell) */
  clearSelection(): void {
    this.anchor = { ...this.head };
    this.host.requestUpdate();
  }

  // ─── Helpers ────────────────────────────────────────

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
