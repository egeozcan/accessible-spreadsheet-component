import { describe, it, expect, beforeEach, vi } from 'vitest';
import { SelectionManager } from '../selection-manager.js';

function createMockHost() {
  return {
    addController: vi.fn(),
    removeController: vi.fn(),
    requestUpdate: vi.fn(),
    updateComplete: Promise.resolve(true),
  };
}

describe('SelectionManager', () => {
  let manager: SelectionManager;
  let host: ReturnType<typeof createMockHost>;

  beforeEach(() => {
    host = createMockHost();
    manager = new SelectionManager(host, 100, 26);
  });

  describe('initialization', () => {
    it('starts at cell 0,0', () => {
      expect(manager.activeCell).toEqual({ row: 0, col: 0 });
    });

    it('anchor and head are at 0,0', () => {
      expect(manager.anchor).toEqual({ row: 0, col: 0 });
      expect(manager.head).toEqual({ row: 0, col: 0 });
    });

    it('registers controller with host', () => {
      expect(host.addController).toHaveBeenCalledWith(manager);
    });
  });

  describe('range', () => {
    it('returns a normalized range', () => {
      manager.anchor = { row: 5, col: 3 };
      manager.head = { row: 2, col: 1 };
      expect(manager.range).toEqual({
        start: { row: 2, col: 1 },
        end: { row: 5, col: 3 },
      });
    });

    it('returns single cell range when anchor equals head', () => {
      expect(manager.range).toEqual({
        start: { row: 0, col: 0 },
        end: { row: 0, col: 0 },
      });
    });
  });

  describe('isCellSelected', () => {
    it('returns true for cells within selection range', () => {
      manager.anchor = { row: 1, col: 1 };
      manager.head = { row: 3, col: 3 };

      expect(manager.isCellSelected(2, 2)).toBe(true);
      expect(manager.isCellSelected(1, 1)).toBe(true);
      expect(manager.isCellSelected(3, 3)).toBe(true);
    });

    it('returns false for cells outside selection range', () => {
      manager.anchor = { row: 1, col: 1 };
      manager.head = { row: 3, col: 3 };

      expect(manager.isCellSelected(0, 0)).toBe(false);
      expect(manager.isCellSelected(4, 4)).toBe(false);
    });
  });

  describe('isCellActive', () => {
    it('returns true for the head cell', () => {
      manager.head = { row: 2, col: 3 };
      expect(manager.isCellActive(2, 3)).toBe(true);
    });

    it('returns false for non-head cells', () => {
      manager.head = { row: 2, col: 3 };
      expect(manager.isCellActive(0, 0)).toBe(false);
    });
  });

  describe('move', () => {
    it('moves head and anchor together by default', () => {
      manager.move(1, 0);
      expect(manager.activeCell).toEqual({ row: 1, col: 0 });
      expect(manager.anchor).toEqual({ row: 1, col: 0 });
    });

    it('moves only head when extending selection', () => {
      manager.move(2, 0, true);
      expect(manager.head).toEqual({ row: 2, col: 0 });
      expect(manager.anchor).toEqual({ row: 0, col: 0 });
    });

    it('clamps at top-left edge (row 0, col 0)', () => {
      manager.move(-1, -1);
      expect(manager.activeCell).toEqual({ row: 0, col: 0 });
    });

    it('clamps at bottom-right edge (maxRows-1, maxCols-1)', () => {
      manager.moveTo(99, 25);
      manager.move(1, 1);
      expect(manager.activeCell).toEqual({ row: 99, col: 25 });
    });

    it('clamps at top edge only', () => {
      manager.moveTo(0, 5);
      manager.move(-1, 0);
      expect(manager.activeCell).toEqual({ row: 0, col: 5 });
    });

    it('clamps at left edge only', () => {
      manager.moveTo(5, 0);
      manager.move(0, -1);
      expect(manager.activeCell).toEqual({ row: 5, col: 0 });
    });

    it('clamps at bottom edge only', () => {
      manager.moveTo(99, 5);
      manager.move(1, 0);
      expect(manager.activeCell).toEqual({ row: 99, col: 5 });
    });

    it('clamps at right edge only', () => {
      manager.moveTo(5, 25);
      manager.move(0, 1);
      expect(manager.activeCell).toEqual({ row: 5, col: 25 });
    });

    it('requests host update', () => {
      host.requestUpdate.mockClear();
      manager.move(1, 0);
      expect(host.requestUpdate).toHaveBeenCalled();
    });
  });

  describe('moveTo', () => {
    it('moves to a specific cell', () => {
      manager.moveTo(5, 3);
      expect(manager.activeCell).toEqual({ row: 5, col: 3 });
      expect(manager.anchor).toEqual({ row: 5, col: 3 });
    });

    it('extends selection when extend is true', () => {
      manager.moveTo(5, 3, true);
      expect(manager.head).toEqual({ row: 5, col: 3 });
      expect(manager.anchor).toEqual({ row: 0, col: 0 });
    });

    it('clamps to bounds when out of range', () => {
      manager.moveTo(200, 50);
      expect(manager.activeCell).toEqual({ row: 99, col: 25 });
    });

    it('clamps negative values to 0', () => {
      manager.moveTo(-5, -3);
      expect(manager.activeCell).toEqual({ row: 0, col: 0 });
    });
  });

  describe('selectAll', () => {
    it('selects all cells in the grid', () => {
      manager.selectAll();
      expect(manager.anchor).toEqual({ row: 0, col: 0 });
      expect(manager.head).toEqual({ row: 99, col: 25 });
    });
  });

  describe('clearSelection', () => {
    it('collapses selection to the head cell', () => {
      manager.anchor = { row: 0, col: 0 };
      manager.head = { row: 5, col: 5 };

      manager.clearSelection();
      expect(manager.anchor).toEqual({ row: 5, col: 5 });
      expect(manager.head).toEqual({ row: 5, col: 5 });
    });
  });

  describe('startSelection', () => {
    it('sets anchor and head to the cell', () => {
      manager.startSelection(3, 4);
      expect(manager.anchor).toEqual({ row: 3, col: 4 });
      expect(manager.head).toEqual({ row: 3, col: 4 });
    });

    it('extends selection when extend is true', () => {
      manager.moveTo(2, 2);
      manager.startSelection(5, 5, true);
      expect(manager.anchor).toEqual({ row: 2, col: 2 });
      expect(manager.head).toEqual({ row: 5, col: 5 });
    });

    it('sets dragging state', () => {
      expect(manager.isDragging).toBe(false);
      manager.startSelection(0, 0);
      expect(manager.isDragging).toBe(true);
    });
  });

  describe('extendSelection', () => {
    it('extends head during drag', () => {
      manager.startSelection(0, 0);
      manager.extendSelection(3, 3);
      expect(manager.head).toEqual({ row: 3, col: 3 });
    });

    it('does nothing when not dragging', () => {
      manager.extendSelection(3, 3);
      expect(manager.head).toEqual({ row: 0, col: 0 });
    });
  });

  describe('endSelection', () => {
    it('stops dragging', () => {
      manager.startSelection(0, 0);
      expect(manager.isDragging).toBe(true);
      manager.endSelection();
      expect(manager.isDragging).toBe(false);
    });
  });

  describe('setBounds', () => {
    it('updates grid bounds', () => {
      manager.setBounds(10, 5);
      manager.moveTo(20, 10);
      expect(manager.activeCell).toEqual({ row: 9, col: 4 });
    });
  });

  describe('hostDisconnected', () => {
    it('stops dragging', () => {
      manager.startSelection(0, 0);
      expect(manager.isDragging).toBe(true);
      manager.hostDisconnected();
      expect(manager.isDragging).toBe(false);
    });
  });
});
