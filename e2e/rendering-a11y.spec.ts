import { test, expect } from './fixtures';

test.describe('Rendering & Accessibility', () => {
  test('renders the grid with correct ARIA roles', async ({ spreadsheet }) => {
    const grid = spreadsheet.grid;
    await expect(grid).toHaveRole('grid');
    await expect(grid).toHaveAttribute('aria-rowcount');
    await expect(grid).toHaveAttribute('aria-colcount');
  });

  test('grid has correct aria-rowcount and aria-colcount', async ({ spreadsheet }) => {
    // Default demo has 50 rows, 26 cols
    // aria-rowcount = rows + 1 (header row), aria-colcount = cols + 1 (row header)
    await expect(spreadsheet.grid).toHaveAttribute('aria-rowcount', '51');
    await expect(spreadsheet.grid).toHaveAttribute('aria-colcount', '27');
  });

  test('renders column headers A through Z', async ({ spreadsheet }) => {
    const headers = spreadsheet.columnHeaders;
    // First columnheader is the corner header (empty), then A-Z
    const count = await headers.count();
    expect(count).toBe(27); // corner + 26 columns

    // Check first few column labels
    await expect(headers.nth(1)).toHaveText('A');
    await expect(headers.nth(2)).toHaveText('B');
    await expect(headers.nth(3)).toHaveText('C');
    await expect(headers.nth(26)).toHaveText('Z');
  });

  test('renders row headers with 1-based numbers', async ({ spreadsheet }) => {
    const rowHeaders = spreadsheet.rowHeaders;
    await expect(rowHeaders.first()).toHaveText('1');
    await expect(rowHeaders.nth(1)).toHaveText('2');
    await expect(rowHeaders.nth(2)).toHaveText('3');
  });

  test('cells have correct ARIA attributes', async ({ spreadsheet }) => {
    const cell = spreadsheet.cell(0, 0);
    await expect(cell).toHaveRole('gridcell');
    await expect(cell).toHaveAttribute('aria-colindex');
    await expect(cell).toHaveAttribute('aria-selected');
  });

  test('active cell has tabindex=0, others have tabindex=-1', async ({ spreadsheet }) => {
    // Default active cell is 0,0
    const activeCell = spreadsheet.cell(0, 0);
    await expect(activeCell).toHaveAttribute('tabindex', '0');

    const otherCell = spreadsheet.cell(1, 1);
    await expect(otherCell).toHaveAttribute('tabindex', '-1');
  });

  test('renders pre-populated demo data correctly', async ({ spreadsheet }) => {
    // The demo page sets data: Item, Qty, Price, Total headers
    await expect(spreadsheet.cell(0, 0).locator('.cell-text')).toHaveText('Item');
    await expect(spreadsheet.cell(0, 1).locator('.cell-text')).toHaveText('Qty');
    await expect(spreadsheet.cell(0, 2).locator('.cell-text')).toHaveText('Price');
    await expect(spreadsheet.cell(0, 3).locator('.cell-text')).toHaveText('Total');
  });

  test('renders data row values', async ({ spreadsheet }) => {
    await expect(spreadsheet.cell(1, 0).locator('.cell-text')).toHaveText('Widget A');
    await expect(spreadsheet.cell(1, 1).locator('.cell-text')).toHaveText('10');
    await expect(spreadsheet.cell(1, 2).locator('.cell-text')).toHaveText('5.99');
  });

  test('row elements have correct aria-rowindex', async ({ spreadsheet }) => {
    // Data rows have aria-rowindex starting at 2 (header is 1)
    const rows = spreadsheet.shadow('[role="row"]');
    // Header row
    await expect(rows.first()).toHaveAttribute('aria-rowindex', '1');
    // First data row
    await expect(rows.nth(1)).toHaveAttribute('aria-rowindex', '2');
  });

  test('has a live region for screen reader announcements', async ({ spreadsheet }) => {
    const liveRegion = spreadsheet.liveRegion;
    await expect(liveRegion).toHaveAttribute('aria-live', 'polite');
    await expect(liveRegion).toHaveAttribute('aria-atomic', 'true');
  });

  test('active cell gets the active-cell class', async ({ spreadsheet }) => {
    const cell = spreadsheet.cell(0, 0);
    await expect(cell).toHaveClass(/active-cell/);
  });

  test('editor input has aria-label', async ({ spreadsheet }) => {
    await expect(spreadsheet.editor).toHaveAttribute('aria-label', 'Cell editor');
  });
});
