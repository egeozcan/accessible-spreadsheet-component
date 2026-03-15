import { test, expect } from './fixtures';

test.describe('Virtual Scrolling', () => {
  test('only renders a subset of rows (not all 50)', async ({ spreadsheet }) => {
    // The grid should not render all 50 rows at once
    const totalRenderedRows = await spreadsheet.shadow('[role="row"]').count();

    // Header row + visible rows + buffer (not all 50 data rows)
    // Exact count depends on viewport, but should be less than 51
    expect(totalRenderedRows).toBeLessThan(51);
    expect(totalRenderedRows).toBeGreaterThan(1); // at least header + some rows
  });

  test('scrolling reveals new rows', async ({ spreadsheet }) => {
    const initialVisibleRows = await spreadsheet.shadow('[data-row]').evaluateAll(
      (elements) => elements.map((el) => parseInt((el as HTMLElement).dataset.row!))
    );
    const initialMaxRow = Math.max(...initialVisibleRows);

    // Scroll well down the grid and wait for virtualization to catch up.
    await spreadsheet.grid.evaluate((grid, top) => {
      grid.scrollTop = top as number;
    }, 1000);
    await spreadsheet.page.waitForTimeout(200);

    const visibleRowsAfterScroll = await spreadsheet.shadow('[data-row]').evaluateAll(
      (elements) => elements.map((el) => parseInt((el as HTMLElement).dataset.row!))
    );
    const maxRowAfterScroll = Math.max(...visibleRowsAfterScroll);

    expect(maxRowAfterScroll).toBeGreaterThan(initialMaxRow);
    expect(maxRowAfterScroll).toBeGreaterThan(30);
  });

  test('scrolled-to cells are accessible and clickable', async ({ spreadsheet }) => {
    // Scroll to make a lower row visible
    await spreadsheet.grid.evaluate((grid) => {
      grid.scrollTop = 400;
    });
    await spreadsheet.page.waitForTimeout(200);

    // Find a row that's now rendered
    const visibleRows = await spreadsheet.shadow('[data-row]').evaluateAll(
      (elements) => elements.map((el) => parseInt((el as HTMLElement).dataset.row!))
    );

    const highRow = Math.max(...visibleRows);

    // Click on the high row cell
    await spreadsheet.clickCell(highRow, 0);
    await expect(spreadsheet.cell(highRow, 0)).toHaveClass(/active-cell/);
  });

  test('navigating with arrow keys scrolls the view', async ({ spreadsheet }) => {
    // Click first cell
    await spreadsheet.clickCell(0, 0);

    // Press arrow down many times to go past visible area
    for (let i = 0; i < 30; i++) {
      const activeRow = await spreadsheet.page.evaluate(() => {
        const sheet = document.querySelector('y11n-spreadsheet');
        const active = sheet?.shadowRoot?.querySelector('.active-cell') as HTMLElement;
        return active ? parseInt(active.dataset.row!) : -1;
      });

      const nextCell = spreadsheet.cell(activeRow, 0);
      await nextCell.press('ArrowDown');
    }

    // The active cell should be at row 30
    await expect(spreadsheet.cell(30, 0)).toHaveClass(/active-cell/);
  });

  test('ensureCellVisible scrolls cell fully into viewport below sticky header', async ({ spreadsheet }) => {
    // Programmatically move the selection to a far-off row via the component's
    // internal API, which calls _ensureCellVisible but does NOT trigger the
    // browser's native focus-scroll. This isolates the scroll calculation.
    await spreadsheet.clickCell(0, 0);

    const visibility = await spreadsheet.page.evaluate(async () => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      // Use the public keyboard-like API: set selection then request update
      // We access the selection manager to jump directly
      const sel = sheet._selection;
      sel.moveTo(30, 0);
      // Trigger the internal ensureCellVisible (same as keyboard nav does)
      sheet._ensureCellVisible(30, 0);
      sheet.requestUpdate();
      await sheet.updateComplete;

      const grid = sheet.shadowRoot!.querySelector('[role="grid"]') as HTMLElement;
      const activeCell = sheet.shadowRoot!.querySelector('[data-row="30"][data-col="0"]') as HTMLElement;
      if (!activeCell || !grid) return { visible: false, reason: 'missing element' };

      const gridRect = grid.getBoundingClientRect();
      const cellRect = activeCell.getBoundingClientRect();

      // Find the sticky header height
      const header = sheet.shadowRoot!.querySelector('.ls-col-header') as HTMLElement;
      const headerHeight = header ? header.getBoundingClientRect().height : 0;

      const dataTop = gridRect.top + headerHeight;
      const dataBottom = gridRect.bottom;

      return {
        visible: cellRect.top >= dataTop - 1 && cellRect.bottom <= dataBottom + 1,
        cellTop: Math.round(cellRect.top),
        cellBottom: Math.round(cellRect.bottom),
        dataTop: Math.round(dataTop),
        dataBottom: Math.round(dataBottom),
      };
    });

    expect(visibility.visible).toBe(true);
  });

  test('large grid renders efficiently with spacers', async ({ spreadsheet }) => {
    // Set the grid to have many rows
    await spreadsheet.setProperty('rows', 500);
    await spreadsheet.page.waitForTimeout(200);

    // Count distinct rendered rows using data-row attributes
    const renderedRowNumbers = await spreadsheet.shadow('[data-row]').evaluateAll(
      (elements) => {
        const rows = new Set(elements.map((el) => (el as HTMLElement).dataset.row));
        return rows.size;
      }
    );
    // Should render far fewer rows than 500 due to virtualization
    expect(renderedRowNumbers).toBeLessThan(100);
    expect(renderedRowNumbers).toBeGreaterThan(0);

    // aria-rowcount should reflect the full size
    await expect(spreadsheet.grid).toHaveAttribute('aria-rowcount', '501');
  });
});
