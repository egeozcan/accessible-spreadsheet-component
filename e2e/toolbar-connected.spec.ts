import { test, expect } from './fixtures';
import type { Page } from '@playwright/test';

/**
 * `<y11n-toolbar>` pointed at a grid (`for` / `target`): it follows
 * the grid's state and its controls act on the selection.
 */
test.describe('Toolbar connected to a grid', () => {
  const DATA = {
    '0:0': { rawValue: 'Name', displayValue: 'Name', type: 'text' },
    '0:1': { rawValue: 'Score', displayValue: 'Score', type: 'text' },
    '1:0': { rawValue: 'Cy', displayValue: 'Cy', type: 'text' },
    '1:1': { rawValue: '30', displayValue: '30', type: 'number' },
    '2:0': { rawValue: 'Ann', displayValue: 'Ann', type: 'text' },
    '2:1': { rawValue: '10.5', displayValue: '10.5', type: 'number' },
    '3:0': { rawValue: 'Bob', displayValue: 'Bob', type: 'text' },
    '3:1': { rawValue: '20', displayValue: '20', type: 'number' },
  };

  async function addToolbar(page: Page, items?: string) {
    await page.evaluate((items) => {
      const toolbar = document.createElement('y11n-toolbar');
      toolbar.id = 'tb';
      toolbar.setAttribute('for', 'sheet');
      if (items) toolbar.setAttribute('items', items);
      const container = document.querySelector('.demo-container')!;
      container.parentElement!.insertBefore(toolbar, container);
    }, items);
    await page.waitForTimeout(100);
  }

  const control = (page: Page, label: string) => page.locator('#tb').locator(`[aria-label="${label}"]`);
  const menuItem = (page: Page, text: string) => page.locator('#tb').locator('[role="menuitem"]', { hasText: text });
  const sheetEval = <T>(page: Page, fn: (sheet: any) => T) =>
    page.evaluate((source) => new Function('sheet', `return (${source})(sheet)`)(document.querySelector('y11n-spreadsheet')), fn.toString()) as Promise<T>;

  test('shows the format of the selection and formats it', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.setCellFormat('0:0', { bold: true, textAlign: 'center', numberFormat: { type: 'percent', decimals: 0 } });

    await spreadsheet.clickCell(0, 0);
    await expect(control(page, 'Bold')).toHaveAttribute('aria-pressed', 'true');
    await expect(control(page, 'Align center')).toHaveAttribute('aria-pressed', 'true');
    await expect(control(page, 'Number format')).toHaveValue('percent');

    await spreadsheet.clickCell(1, 1);
    await expect(control(page, 'Bold')).toHaveAttribute('aria-pressed', 'false');
    await expect(control(page, 'Number format')).toHaveValue('general');

    await control(page, 'Italic').click();
    await control(page, 'Wrap text').click();
    await control(page, 'Align top').click();
    await control(page, 'Number format').selectOption('currency');
    expect(await spreadsheet.getCellFormat('1:1')).toEqual({
      italic: true,
      wrap: true,
      verticalAlign: 'top',
      numberFormat: { type: 'currency', decimals: 2 },
    });
    await expect(control(page, 'Italic')).toHaveAttribute('aria-pressed', 'true');
    await expect(control(page, 'Wrap text')).toHaveAttribute('aria-pressed', 'true');
    await expect(control(page, 'Number format')).toHaveValue('currency');

    await control(page, 'Clear formatting').click();
    expect(await spreadsheet.getCellFormat('1:1')).toBeUndefined();
    await expect(control(page, 'Italic')).toHaveAttribute('aria-pressed', 'false');
  });

  test('changes decimals from what the cell shows', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.clickCell(2, 1); // 10.5
    await control(page, 'Increase decimals').click();
    await spreadsheet.waitForCellText(2, 1, '10.50');
    await control(page, 'Decrease decimals').click();
    await control(page, 'Decrease decimals').click();
    expect((await spreadsheet.getCellFormat('2:1'))?.numberFormat.decimals).toBe(0);
  });

  test('undo and redo follow the history', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await expect(control(page, 'Undo')).toBeDisabled();
    await expect(control(page, 'Redo')).toBeDisabled();

    await spreadsheet.clickCell(1, 0);
    await page.keyboard.type('Zed');
    await page.keyboard.press('Enter');
    await expect(control(page, 'Undo')).toBeEnabled();

    await control(page, 'Undo').click();
    await spreadsheet.waitForCellText(1, 0, 'Cy');
    await expect(control(page, 'Undo')).toBeDisabled();
    await expect(control(page, 'Redo')).toBeEnabled();

    await control(page, 'Redo').click();
    await spreadsheet.waitForCellText(1, 0, 'Zed');
    await expect(control(page, 'Redo')).toBeDisabled();
  });

  test('sorts, filters and reapplies the filter', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.clickCell(1, 0);

    await control(page, 'Sort ascending').click();
    await spreadsheet.waitForCellText(1, 0, 'Ann');
    await control(page, 'Sort descending').click();
    await spreadsheet.waitForCellText(1, 0, 'Cy');

    await expect(control(page, 'Filter')).toHaveAttribute('aria-pressed', 'false');
    await expect(control(page, 'Reapply filter')).toBeDisabled();
    await control(page, 'Filter').click();
    await expect(control(page, 'Filter')).toHaveAttribute('aria-pressed', 'true');
    await expect(control(page, 'Reapply filter')).toBeEnabled();
    expect(await sheetEval(page, (sheet) => sheet.getFilter() !== null)).toBe(true);

    await control(page, 'Filter').click();
    await expect(control(page, 'Filter')).toHaveAttribute('aria-pressed', 'false');
  });

  test('inserts and deletes rows and columns at the selection', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.clickCell(1, 0);

    await control(page, 'Insert rows or columns').click();
    await menuItem(page, 'Insert row above').click();
    await spreadsheet.waitForCellText(2, 0, 'Cy');
    expect(await spreadsheet.getCellText(1, 0)).toBe('');

    await control(page, 'Delete rows or columns').click();
    await menuItem(page, 'Delete row').click();
    await spreadsheet.waitForCellText(1, 0, 'Cy');

    await control(page, 'Insert rows or columns').click();
    await menuItem(page, 'Insert column right').click();
    await spreadsheet.waitForCellText(0, 2, 'Score');

    await spreadsheet.clickCell(0, 1);
    await control(page, 'Delete rows or columns').click();
    await menuItem(page, 'Delete column').click();
    await spreadsheet.waitForCellText(0, 1, 'Score');
  });

  test('freezes and unfreezes panes', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    const frozen = () => sheetEval(page, (sheet) => [sheet.frozenRows, sheet.frozenCols]);

    await control(page, 'Freeze panes').click();
    await expect(menuItem(page, 'Unfreeze panes')).toBeDisabled();
    await menuItem(page, 'Freeze first row').click();
    expect(await frozen()).toEqual([1, 0]);
    await expect(spreadsheet.shadow('[role="log"]')).toHaveText('Froze 1 row');

    await spreadsheet.clickCell(2, 1);
    await control(page, 'Freeze panes').click();
    await menuItem(page, 'Freeze up to the active cell').click();
    expect(await frozen()).toEqual([2, 1]);

    await control(page, 'Freeze panes').click();
    await menuItem(page, 'Unfreeze panes').click();
    expect(await frozen()).toEqual([0, 0]);
  });

  test('opens find, the name manager and Paste Special', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.clickCell(1, 0);

    await control(page, 'Find').click();
    await expect(spreadsheet.shadow('#find-input')).toBeFocused();
    await page.keyboard.press('Escape');

    await control(page, 'Name manager').click();
    await expect(spreadsheet.shadow('y11n-name-manager')).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(spreadsheet.shadow('y11n-name-manager')).toHaveCount(0);

    await control(page, 'Paste Special').click();
    await expect(spreadsheet.shadow('.ls-paste-special')).toBeVisible();
  });

  test('AutoSum and lists', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData({ ...DATA, '5:0': { rawValue: 'a, b', displayValue: 'a, b', type: 'text' } });
    await spreadsheet.clickCell(4, 1);
    await control(page, 'AutoSum').click();
    await expect(spreadsheet.editor).toHaveValue('=SUM(B2:B4)');
    await page.keyboard.press('Enter');
    await spreadsheet.waitForCellText(4, 1, '60.5');

    await spreadsheet.clickCell(5, 0);
    await control(page, 'List in cell').click();
    expect(await sheetEval(page, (sheet) => sheet.getCellValues('5:0')?.values)).toEqual(['a', 'b']);
  });

  test('copy and cut use the clipboard', async ({ spreadsheet, page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.clickCell(1, 0);
    await control(page, 'Copy').click();
    await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('Cy');

    await spreadsheet.clickCell(2, 0);
    await control(page, 'Cut').click();
    await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('Ann');
    await spreadsheet.waitForCellText(2, 0, '');

    await spreadsheet.clickCell(6, 0);
    await control(page, 'Paste').click();
    await spreadsheet.waitForCellText(6, 0, 'Ann');
  });

  test('read-only: only the controls that change cells are disabled', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.setReadOnly(true);

    for (const label of ['Bold', 'Cut', 'Paste', 'Sort ascending', 'AutoSum', 'Insert rows or columns', 'Number format', 'Conditional formatting']) {
      await expect(control(page, label)).toBeDisabled();
    }
    for (const label of ['Copy', 'Filter', 'Freeze panes', 'Find', 'Fit column width', 'Name manager']) {
      await expect(control(page, label)).toBeEnabled();
    }

    await spreadsheet.setReadOnly(false);
    await expect(control(page, 'Bold')).toBeEnabled();
  });

  test('items limits and orders the controls', async ({ spreadsheet, page }) => {
    await addToolbar(page, 'history | filter bold');
    await spreadsheet.setData(DATA);
    const labels = await page
      .locator('#tb')
      .locator('[role="toolbar"] [data-item]')
      .evaluateAll((els) => els.map((el) => el.getAttribute('aria-label')));
    expect(labels).toEqual(['Undo', 'Redo', 'Filter', 'Bold']);
    await expect(page.locator('#tb').locator('[role="separator"]')).toHaveCount(1);
  });

  test('preventDefault on an event stops the control', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await page.evaluate(() => {
      document.querySelector('#tb')!.addEventListener('format-action', (e) => e.preventDefault());
    });
    await spreadsheet.clickCell(1, 0);
    await control(page, 'Bold').click();
    await page.waitForTimeout(100);
    expect(await spreadsheet.getCellFormat('1:0')).toBeUndefined();
  });

  test('one Tab stop; arrow keys move along the controls', async ({ spreadsheet, page }) => {
    await addToolbar(page, 'bold italic fontSize textColor filter');
    await spreadsheet.setData(DATA);
    await control(page, 'Bold').focus();
    await page.keyboard.press('ArrowRight');
    await expect(control(page, 'Italic')).toBeFocused();
    await page.keyboard.press('ArrowRight');
    await expect(control(page, 'Font size')).toBeFocused();
    await page.keyboard.press('ArrowRight');
    await expect(control(page, 'Text color')).toBeFocused();
    await page.keyboard.press('End');
    await expect(control(page, 'Filter')).toBeFocused();

    // Tab leaves the toolbar, Shift+Tab comes back to the control it was left on
    await page.keyboard.press('Tab');
    await expect(control(page, 'Filter')).not.toBeFocused();
    expect(await page.evaluate(() => document.activeElement?.id)).not.toBe('tb');
    await page.keyboard.press('Shift+Tab');
    await expect(control(page, 'Filter')).toBeFocused();
  });
});
