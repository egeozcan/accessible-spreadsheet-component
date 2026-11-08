import { test, expect } from './fixtures';

// The demo page seeds A1:D4 as a table (Item / Qty / Price / Total, with
// =B2*C2 style totals) and a grand total =SUM(D2:D4) in D6, below a blank row.
test.describe('Sorting', () => {
  test('Alt+Shift+Down sorts the table around the active cell, keeping its header', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(1, 1);
    await spreadsheet.cell(1, 1).press('Alt+Shift+ArrowDown');

    await spreadsheet.waitForCellText(1, 0, 'Widget B');
    expect(await spreadsheet.getCellText(0, 0)).toBe('Item');
    expect(await spreadsheet.getCellText(2, 0)).toBe('Widget A');
    expect(await spreadsheet.getCellText(3, 0)).toBe('Widget C');
    // Totals moved with their rows; the grand total below the blank row stayed put
    expect(await spreadsheet.getCellText(1, 3)).toBe('87.5');
    expect(await spreadsheet.getCellText(5, 0)).toBe('Grand Total');
    expect(await spreadsheet.getCellText(5, 3)).toBe('231.4');

    await expect(spreadsheet.shadow('[role="log"]')).toHaveText(
      'Sorted A2:D4 by column B, descending. Row 1 kept as the header'
    );
    // Focus stays on the active cell, so keyboard users can keep going
    await expect(spreadsheet.cell(1, 1)).toBeFocused();
  });

  test('Alt+Shift+Up sorts ascending, and Ctrl+Z undoes the sort', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(2, 2);
    await spreadsheet.cell(2, 2).press('Alt+Shift+ArrowUp');
    // By price: Widget B (3.50), Widget A (5.99), Widget C (12.00)
    await spreadsheet.waitForCellText(1, 0, 'Widget B');
    expect(await spreadsheet.getCellText(3, 0)).toBe('Widget C');

    await spreadsheet.cell(2, 2).press('Control+z');
    await spreadsheet.waitForCellText(1, 0, 'Widget A');
    expect(await spreadsheet.getCellText(1, 3)).toBe('59.9');
  });

  test('a before-sort listener can cancel the sort', async ({ spreadsheet }) => {
    await spreadsheet.page.evaluate(() => {
      document.querySelector('y11n-spreadsheet')!.addEventListener('before-sort', (e) => e.preventDefault());
    });
    await spreadsheet.clickCell(1, 1);
    await spreadsheet.cell(1, 1).press('Alt+Shift+ArrowDown');

    await expect(spreadsheet.shadow('[role="log"]')).toHaveText('Sort cancelled');
    expect(await spreadsheet.getCellText(1, 0)).toBe('Widget A');
  });

  test('nothing is sorted in read-only mode', async ({ spreadsheet }) => {
    await spreadsheet.setReadOnly(true);
    await spreadsheet.clickCell(1, 1);
    await spreadsheet.cell(1, 1).press('Alt+Shift+ArrowDown');
    await spreadsheet.page.waitForTimeout(100);
    expect(await spreadsheet.getCellText(1, 0)).toBe('Widget A');
  });

  test('sort() API sorts below a header row and fires sort-change', async ({ spreadsheet }) => {
    const detail = await spreadsheet.page.evaluate(() => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      let seen: unknown = null;
      sheet.addEventListener('sort-change', (e: CustomEvent) => (seen = e.detail));
      sheet.sort(
        { start: { row: 0, col: 0 }, end: { row: 3, col: 3 } },
        [{ col: 3, direction: 'descending' }],
        { hasHeader: true }
      );
      return seen;
    });
    expect(detail).toEqual({
      range: { start: { row: 0, col: 0 }, end: { row: 3, col: 3 } },
      keys: [{ col: 3, direction: 'descending' }],
      hasHeader: true,
      source: 'programmatic',
    });
    await spreadsheet.waitForCellText(1, 0, 'Widget B');
    expect(await spreadsheet.getCellText(2, 0)).toBe('Widget C');
    expect(await spreadsheet.getCellText(3, 0)).toBe('Widget A');
  });
});
