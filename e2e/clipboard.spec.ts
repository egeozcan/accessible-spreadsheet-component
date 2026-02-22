import { test, expect } from './fixtures';

test.describe('Clipboard Operations', () => {
  test.use({
    permissions: ['clipboard-read', 'clipboard-write'],
  });

  test('Ctrl+C copies cell value to clipboard', async ({ spreadsheet }) => {
    // Cell A1 has "Item"
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Control+c');

    const clipboardText = await spreadsheet.page.evaluate(() =>
      navigator.clipboard.readText()
    );
    expect(clipboardText).toBe('Item');
  });

  test('Ctrl+V pastes clipboard content into cell', async ({ spreadsheet }) => {
    await spreadsheet.page.evaluate(() =>
      navigator.clipboard.writeText('Pasted Value')
    );

    await spreadsheet.clickCell(7, 0);
    await spreadsheet.cell(7, 0).press('Control+v');

    await spreadsheet.waitForCellText(7, 0, 'Pasted Value');
  });

  test('Ctrl+X cuts cell value (copies and clears)', async ({ spreadsheet }) => {
    // Use an existing demo data cell
    await spreadsheet.waitForCellText(1, 0, 'Widget A');

    await spreadsheet.clickCell(1, 0);
    await spreadsheet.cell(1, 0).press('Control+x');

    // Cell should be cleared
    await spreadsheet.waitForCellText(1, 0, '');

    // Clipboard should have the value
    const clipboardText = await spreadsheet.page.evaluate(() =>
      navigator.clipboard.readText()
    );
    expect(clipboardText).toBe('Widget A');
  });

  test('copy and paste a range of cells', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'A', displayValue: 'A', type: 'text' },
      '0:1': { rawValue: 'B', displayValue: 'B', type: 'text' },
      '1:0': { rawValue: 'C', displayValue: 'C', type: 'text' },
      '1:1': { rawValue: 'D', displayValue: 'D', type: 'text' },
    });

    // Select the 2x2 range
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Shift+ArrowRight');
    await spreadsheet.cell(0, 1).press('Shift+ArrowDown');

    // Copy
    await spreadsheet.cell(1, 1).press('Control+c');

    // Verify clipboard has TSV format
    const clipboardText = await spreadsheet.page.evaluate(() =>
      navigator.clipboard.readText()
    );
    expect(clipboardText).toContain('A\tB');
    expect(clipboardText).toContain('C\tD');

    // Paste to a new location
    await spreadsheet.clickCell(0, 4);
    await spreadsheet.cell(0, 4).press('Control+v');

    // Verify pasted data
    await spreadsheet.waitForCellText(0, 4, 'A');
    await spreadsheet.waitForCellText(0, 5, 'B');
    await spreadsheet.waitForCellText(1, 4, 'C');
    await spreadsheet.waitForCellText(1, 5, 'D');
  });

  test('pasting TSV data from external source works', async ({ spreadsheet }) => {
    await spreadsheet.page.evaluate(() =>
      navigator.clipboard.writeText('X1\tX2\tX3\nY1\tY2\tY3')
    );

    await spreadsheet.clickCell(8, 0);
    await spreadsheet.cell(8, 0).press('Control+v');

    await spreadsheet.waitForCellText(8, 0, 'X1');
    await spreadsheet.waitForCellText(8, 1, 'X2');
    await spreadsheet.waitForCellText(8, 2, 'X3');
    await spreadsheet.waitForCellText(9, 0, 'Y1');
    await spreadsheet.waitForCellText(9, 1, 'Y2');
    await spreadsheet.waitForCellText(9, 2, 'Y3');
  });

  test('cut a range of cells clears the source', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '3:4': { rawValue: 'X', displayValue: 'X', type: 'text' },
      '3:5': { rawValue: 'Y', displayValue: 'Y', type: 'text' },
    });

    // Select range
    await spreadsheet.clickCell(3, 4);
    await spreadsheet.cell(3, 4).press('Shift+ArrowRight');

    // Cut
    await spreadsheet.cell(3, 5).press('Control+x');

    // Source cells should be cleared
    await spreadsheet.waitForCellText(3, 4, '');
    await spreadsheet.waitForCellText(3, 5, '');
  });
});
