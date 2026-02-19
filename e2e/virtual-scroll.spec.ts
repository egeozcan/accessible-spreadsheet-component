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
    // Get initial visible rows
    const initialRows = await spreadsheet.shadow('[data-row]').evaluateAll(
      (elements) => elements.map((el) => parseInt((el as HTMLElement).dataset.row!))
    );

    // Scroll down
    await spreadsheet.grid.evaluate((grid) => {
      grid.scrollTop = 500;
    });

    // Wait for scroll event to process
    await spreadsheet.page.waitForTimeout(200);

    // Get new visible rows
    const scrolledRows = await spreadsheet.shadow('[data-row]').evaluateAll(
      (elements) => elements.map((el) => parseInt((el as HTMLElement).dataset.row!))
    );

    // The scrolled view should contain rows that weren't in the initial view
    const maxInitial = Math.max(...initialRows);
    const maxScrolled = Math.max(...scrolledRows);
    expect(maxScrolled).toBeGreaterThan(maxInitial);
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
