import { test, expect } from './fixtures';
import AxeBuilder from '@axe-core/playwright';

test.use({ permissions: ['clipboard-read', 'clipboard-write'] });

// The demo page seeds A1:D4 as a table (Item / Qty / Price / Total, with
// =B2*C2 style totals) and a grand total =SUM(D2:D4) in D6, below a blank row.
test.describe('Excel quick wins', () => {
  test('typed percentages are stored as numbers and edited as typed', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 5);
    await spreadsheet.page.keyboard.type('15%');
    await spreadsheet.page.keyboard.press('Enter');
    await spreadsheet.waitForCellText(0, 5, '15%');
    const data = await spreadsheet.getData();
    expect(data['0:5'].rawValue).toBe('0.15');

    await spreadsheet.clickCell(0, 5);
    await spreadsheet.page.keyboard.press('F2');
    await expect(spreadsheet.editor).toHaveValue('15%');
  });

  test('Alt+= starts a SUM of the numbers above', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(4, 1);
    await spreadsheet.page.keyboard.press('Alt+Equal');
    await expect(spreadsheet.editor).toHaveValue('=SUM(B2:B4)');
  });

  test('the status bar sums the selection and Alt+Shift+S reads it out', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(1, 1);
    await spreadsheet.cell(3, 1).click({ modifiers: ['Shift'] });
    const status = spreadsheet.shadow('.ls-status');
    await expect(status).toContainText('Count: 3');
    await expect(status).toContainText('Sum:');
    await spreadsheet.page.keyboard.press('Alt+Shift+KeyS');
    await expect(spreadsheet.shadow('[role="log"]')).toContainText('Count 3');
  });

  test('Ctrl+Shift+V pastes the value of a formula', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(1, 3);
    await spreadsheet.copyFromCell(1, 3);
    await spreadsheet.clickCell(1, 5);
    await spreadsheet.page.keyboard.press('Control+Shift+KeyV');
    const expected = await spreadsheet.getCellText(1, 3);
    await spreadsheet.waitForCellText(1, 5, expected);
    const data = await spreadsheet.getData();
    expect(data['1:5'].rawValue.startsWith('=')).toBe(false);
  });

  test('double-clicking the fill handle continues day names down the table', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(1, 4);
    await spreadsheet.page.keyboard.type('Mon');
    await spreadsheet.page.keyboard.press('Enter');
    await spreadsheet.clickCell(1, 4);
    await spreadsheet.shadow('.ls-fill-handle').dblclick();
    await spreadsheet.waitForCellText(3, 4, 'Wed');
    expect(await spreadsheet.getCellText(2, 4)).toBe('Tue');
    expect(await spreadsheet.getCellText(4, 4)).toBe('');
  });

  test('Ctrl+Alt+V opens an accessible Paste Special dialog', async ({ spreadsheet, page }) => {
    await spreadsheet.clickCell(1, 0);
    await spreadsheet.copyFromCell(1, 0);
    await spreadsheet.clickCell(8, 0);
    await page.keyboard.press('Control+Alt+KeyV');
    const dialog = spreadsheet.shadow('.ls-paste-special');
    await expect(dialog).toBeVisible();
    await expect(dialog.locator('input[value="all"]')).toBeFocused();

    const results = await new AxeBuilder({ page }).include('y11n-spreadsheet').analyze();
    expect(results.violations).toEqual([]);

    await page.keyboard.press('Escape');
    await expect(dialog).toHaveCount(0);
    await expect(spreadsheet.cell(8, 0)).toBeFocused();
  });
});
