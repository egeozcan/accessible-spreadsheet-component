// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';
import type { FormulaFunction } from '../types.js';
import type { Y11nSpreadsheet } from '../y11n-spreadsheet.js';
import '../y11n-spreadsheet.js';

async function createSpreadsheet(
  attrs: Partial<{
    rows: number;
    cols: number;
  }> = {}
): Promise<Y11nSpreadsheet> {
  const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
  if (attrs.rows !== undefined) el.rows = attrs.rows;
  if (attrs.cols !== undefined) el.cols = attrs.cols;
  document.body.appendChild(el);
  await el.updateComplete;
  return el;
}

function getRenderedColumnCount(el: Y11nSpreadsheet): number {
  const cols = new Set(
    Array.from(el.shadowRoot!.querySelectorAll<HTMLElement>('[data-col]')).map(
      (cell) => cell.dataset.col
    )
  );
  return cols.size;
}

function getCellText(el: Y11nSpreadsheet, row: number, col: number): string {
  const cell = el.shadowRoot!.querySelector<HTMLElement>(
    `[data-row="${row}"][data-col="${col}"] .cell-text`
  );
  return cell?.textContent ?? '';
}

describe('Y11nSpreadsheet', () => {
  beforeEach(() => {
    document.body.innerHTML = '';
  });

  it('does not render every column before the grid is measured', async () => {
    const el = await createSpreadsheet({ rows: 10, cols: 200 });

    expect(getRenderedColumnCount(el)).toBeLessThan(200);
    expect(getRenderedColumnCount(el)).toBeGreaterThan(0);
  });

  it('keeps the formula reference target rendered while arrowing in edit mode', async () => {
    const el = await createSpreadsheet({ rows: 200, cols: 20 });
    const grid = el.shadowRoot!.querySelector('.ls-grid') as HTMLDivElement;
    Object.defineProperty(grid, 'clientHeight', { configurable: true, value: 140 });
    Object.defineProperty(grid, 'clientWidth', { configurable: true, value: 320 });

    const sheet = el as unknown as {
      _isEditing: boolean;
      _editValue: string;
      _handleRefArrow: (key: string) => void;
      _refCursorRow: number;
      requestUpdate: () => void;
    };

    sheet._isEditing = true;
    sheet._editValue = '=';

    for (let i = 0; i < 25; i++) {
      sheet._handleRefArrow('ArrowDown');
    }

    sheet.requestUpdate();
    await el.updateComplete;

    expect(grid.scrollTop).toBeGreaterThan(0);

    const refCell = el.shadowRoot!.querySelector<HTMLElement>(
      `[data-row="${sheet._refCursorRow}"][data-col="0"]`
    );
    expect(refCell).not.toBeNull();
    expect(refCell?.classList.contains('ref-highlight')).toBe(true);
  });

  it('unregisters removed custom functions when the functions property is replaced', async () => {
    const el = await createSpreadsheet({ rows: 5, cols: 5 });
    const double: FormulaFunction = (_ctx, value) => Number(value) * 2;

    el.functions = { DOUBLE: double };
    el.setData(
      new Map([
        ['0:0', { rawValue: '=DOUBLE(2)', displayValue: '', type: 'text' as const }],
      ])
    );
    await el.updateComplete;

    expect(getCellText(el, 0, 0)).toBe('4');

    el.functions = {};
    await el.updateComplete;

    expect(getCellText(el, 0, 0)).toBe('#NAME?');
  });

  it('preserves pre-connected format-only cells when the element is attached', async () => {
    const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
    el.rows = 10;
    el.cols = 10;
    el.setCellFormat('0:0', { backgroundColor: '#ff0000' });

    document.body.appendChild(el);
    await el.updateComplete;

    const styledCell = el.shadowRoot!.querySelector<HTMLElement>('[data-row="0"][data-col="0"]');
    expect(styledCell?.style.backgroundColor).toBe('rgb(255, 0, 0)');
  });
});
