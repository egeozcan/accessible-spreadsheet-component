import { test, expect } from './fixtures';
import type { Page } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';

const log = (page: Page) => page.locator('y11n-spreadsheet').locator('[role="log"]');
const tabs = (page: Page) => page.locator('y11n-sheet-tabs');
const nameBox = (page: Page) => page.locator('y11n-formula-bar').locator('.name-box');
const manager = (page: Page) => page.locator('y11n-name-manager');
const sheet = <T>(page: Page, fn: (el: any) => T) =>
  page.evaluate((source) => new Function('el', `return (${source})(el)`)(document.querySelector('y11n-spreadsheet')), fn.toString()) as Promise<T>;

async function seed(page: Page, rows: string[][]) {
  await page.evaluate((r) => {
    const el = document.querySelector('y11n-spreadsheet') as any;
    const data = new Map();
    r.forEach((row, i) =>
      row.forEach((value, j) => {
        if (value !== '') data.set(`${i}:${j}`, { rawValue: value, displayValue: value, type: 'text' });
      })
    );
    el.setData(data);
  }, rows);
}

test.describe('Name Box', () => {
  test('goes to a typed reference and names the selection', async ({ spreadsheet, page }) => {
    await seed(page, [['10', '', '=SUM(sales)'], ['20']]);
    await spreadsheet.clickCell(0, 0);
    await expect(nameBox(page)).toHaveValue('A1');

    // Ctrl+G moves to the Name Box; a reference goes there
    await page.keyboard.press('Control+g');
    await expect(nameBox(page)).toBeFocused();
    await page.keyboard.type('c5');
    await page.keyboard.press('Enter');
    await expect(spreadsheet.cell(4, 2)).toBeFocused();
    await expect(nameBox(page)).toHaveValue('C5');

    // A new name names the selection
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.press('Shift+ArrowDown');
    await page.keyboard.press('Control+g');
    await page.keyboard.type('sales');
    await page.keyboard.press('Enter');
    await expect(log(page)).toHaveText('Named Sheet1!$A$1:$A$2 sales');
    await expect(spreadsheet.cell(0, 2).locator('.cell-text')).toHaveText('30');
    await expect(nameBox(page)).toHaveValue('sales');
    // Focus is back on the active cell (A2: the selection was extended down to it)
    await expect(spreadsheet.cell(1, 0)).toBeFocused();

    // And typing the name later selects its cells
    await spreadsheet.clickCell(5, 5);
    await expect(nameBox(page)).toHaveValue('F6');
    await nameBox(page).click();
    await page.keyboard.type('SALES');
    await page.keyboard.press('Enter');
    await expect(spreadsheet.liveRegion).toHaveText('A1: 10, 2 rows, 1 column selected');

    // Escape gives the focus back without changing anything
    await page.keyboard.press('Control+g');
    await page.keyboard.type('zzz');
    await page.keyboard.press('Escape');
    await expect(spreadsheet.cell(0, 0)).toBeFocused();
    await expect(nameBox(page)).toHaveValue('sales');
    expect(await sheet(page, (el) => el.getNames().length)).toBe(1);
  });
});

test.describe('Name manager', () => {
  test('keyboard: add, edit and delete a name', async ({ spreadsheet, page }) => {
    await seed(page, [['10', '', '=SUM(sales)*rate'], ['20']]);
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.press('Shift+ArrowDown');
    await page.keyboard.press('Control+F3');

    await expect(manager(page).locator('[role="dialog"]')).toBeVisible();
    await expect(manager(page).locator('#name-new')).toBeFocused();
    await page.keyboard.press('Enter');
    await expect(manager(page).locator('#name-name')).toBeFocused();
    // "Refers to" starts as the selection
    await expect(manager(page).locator('#name-refers')).toHaveValue('=Sheet1!$A$1:$A$2');
    await page.keyboard.type('sales');
    await page.keyboard.press('Enter');
    await expect(manager(page).locator('#name-status')).toHaveText('Name added: sales');
    await expect(manager(page).locator('#name-new')).toBeFocused();

    // A second name: a constant; an invalid name is refused first
    await page.keyboard.press('Enter');
    await page.keyboard.type('A1');
    await page.keyboard.press('Enter');
    await expect(manager(page).locator('#name-name')).toHaveAttribute('aria-invalid', 'true');
    await manager(page).locator('#name-name').fill('rate');
    await manager(page).locator('#name-refers').fill('0.5');
    await page.keyboard.press('Enter');
    await expect(manager(page).locator('#name-status')).toHaveText('Name added: rate');
    await expect(spreadsheet.cell(0, 2).locator('.cell-text')).toHaveText('15');

    // Rename: the formula follows
    await manager(page).getByRole('button', { name: 'Edit sales' }).click();
    await manager(page).locator('#name-name').fill('revenue');
    await page.keyboard.press('Enter');
    await expect(manager(page).locator('#name-status')).toHaveText('Name changed: revenue');
    expect(await sheet(page, (el) => el.getData().get('0:2').rawValue)).toBe('=SUM(revenue)*rate');

    // Delete asks first
    await manager(page).getByRole('button', { name: 'Delete rate' }).click();
    await expect(manager(page).locator('#name-delete-cancel')).toBeFocused();
    await manager(page).locator('#name-delete-confirm').click();
    await expect(manager(page).locator('#name-status')).toHaveText('Name deleted: rate');
    await expect(spreadsheet.cell(0, 2).locator('.cell-text')).toHaveText('#NAME?');

    // Go to closes the manager and selects the cells
    await manager(page).getByRole('button', { name: 'Go to revenue' }).click();
    await expect(manager(page)).toHaveCount(0);
    await expect(spreadsheet.liveRegion).toHaveText('A1: 10, 2 rows, 1 column selected');
  });

  test('Escape closes it and has no axe violations', async ({ spreadsheet, page }) => {
    await sheet(page, (el) => {
      el.defineName('first', '=A1');
      el.defineName('tax', '=0.19');
    });
    await spreadsheet.clickCell(2, 2);
    await page.keyboard.press('Control+F3');
    await expect(manager(page).locator('table')).toBeVisible();
    const results = await new AxeBuilder({ page }).include('y11n-spreadsheet').analyze();
    expect(results.violations).toEqual([]);
    await page.keyboard.press('Escape');
    await expect(manager(page)).toHaveCount(0);
    await expect(spreadsheet.cell(2, 2)).toBeFocused();
  });
});

test.describe('Picking references while typing a formula', () => {
  test('a click picks a cell and a drag picks a range', async ({ spreadsheet, page }) => {
    await seed(page, [['1', '2'], ['3', '4']]);
    await spreadsheet.clickCell(3, 3);
    await page.keyboard.type('=SUM(');
    await spreadsheet.cell(0, 0).hover();
    await page.mouse.down();
    await spreadsheet.cell(1, 1).hover();
    await page.mouse.up();
    await expect(spreadsheet.editor).toHaveValue('=SUM(A1:B2');
    await expect(spreadsheet.editor).toBeFocused();
    await page.keyboard.type(')+');
    await spreadsheet.clickCell(0, 1);
    await expect(spreadsheet.editor).toHaveValue('=SUM(A1:B2)+B1');
    await page.keyboard.press('Enter');
    await expect(spreadsheet.cell(3, 3).locator('.cell-text')).toHaveText('12');
  });

  test('a reference is picked on another sheet with the mouse', async ({ spreadsheet, page }) => {
    await seed(page, [['1']]);
    await sheet(page, (el) => {
      const data = new Map([['1:1', { rawValue: '7', displayValue: '7', type: 'number' }]]);
      el.addSheet('My data', { data });
    });
    await spreadsheet.clickCell(0, 2);
    await page.keyboard.type('=A1+');
    await tabs(page).locator('[role="tab"]', { hasText: 'My data' }).click();
    await expect(log(page)).toHaveText('My data. Pick a cell for the formula in Sheet1!C1; Enter finishes');
    await expect(spreadsheet.editor).toBeVisible();
    await expect(spreadsheet.editor).toBeFocused();
    await spreadsheet.clickCell(1, 1);
    await expect(spreadsheet.editor).toHaveValue("=A1+'My data'!B2");
    await page.keyboard.press('Enter');

    await expect(tabs(page).locator('[role="tab"][aria-selected="true"]')).toHaveText('Sheet1');
    await expect(spreadsheet.cell(0, 2).locator('.cell-text')).toHaveText('8');
    expect(await sheet(page, (el) => el.getData().get('0:2').rawValue)).toBe("=A1+'My data'!B2");
  });

  test('keyboard: Alt+PageDown in a formula, arrows pick, Escape comes back', async ({ spreadsheet, page }) => {
    await seed(page, [['1']]);
    await sheet(page, (el) => el.addSheet('Data'));
    await spreadsheet.clickCell(0, 2);
    await page.keyboard.type('=');
    await page.keyboard.press('Alt+PageDown');
    await expect(tabs(page).locator('[role="tab"][aria-selected="true"]')).toHaveText('Data');
    await page.keyboard.press('ArrowDown');
    await expect(spreadsheet.editor).toHaveValue('=Data!A2');
    await page.keyboard.press('Escape');
    await expect(tabs(page).locator('[role="tab"][aria-selected="true"]')).toHaveText('Sheet1');
    await expect(spreadsheet.editor).toBeHidden();
    expect(await sheet(page, (el) => el.getData().has('0:2'))).toBe(false);
  });
});

test.describe('Values of a list in a filter header', () => {
  test('Alt+Shift+V opens the values where Alt+Down opens the filter menu', async ({ spreadsheet, page }) => {
    await seed(page, [['={{"Ann";"Bob"}}'], ['x']]);
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.press('Control+Shift+L');
    await expect(spreadsheet.liveRegion).toContainText('Alt+Shift+V to show the values');
    await page.keyboard.press('Alt+Shift+V');
    await expect(spreadsheet.shadow('.ls-array-card li')).toHaveText(['Ann', 'Bob']);
    await page.keyboard.press('Escape');
    await page.keyboard.press('Alt+ArrowDown');
    await expect(spreadsheet.shadow('.ls-filter-menu')).toBeVisible();
  });
});

test.describe('Reordering sheet tabs', () => {
  test('a tab is dragged to a new position with the mouse', async ({ spreadsheet, page }) => {
    await sheet(page, (el) => {
      el.addSheet('Two');
      el.addSheet('Three');
    });
    await expect(tabs(page).locator('[role="tab"]')).toHaveText(['Sheet1', 'Two', 'Three']);
    const first = await tabs(page).locator('[role="tab"]').nth(0).boundingBox();
    const last = await tabs(page).locator('[role="tab"]').nth(2).boundingBox();
    await page.mouse.move(first!.x + first!.width / 2, first!.y + first!.height / 2);
    await page.mouse.down();
    await page.mouse.move(last!.x + last!.width - 4, last!.y + last!.height / 2, { steps: 8 });
    await expect(tabs(page).locator('.drop-indicator')).toBeVisible();
    await page.mouse.up();
    await expect(tabs(page).locator('[role="tab"]')).toHaveText(['Two', 'Three', 'Sheet1']);
    expect(await sheet(page, (el) => el.sheets)).toEqual(['Two', 'Three', 'Sheet1']);
    // A plain click still just selects
    await tabs(page).locator('[role="tab"]', { hasText: 'Two' }).click();
    expect(await sheet(page, (el) => [el.activeSheet, el.sheets.join()])).toEqual(['Two', 'Two,Three,Sheet1']);
    await expect(spreadsheet.grid).toBeVisible();
  });
});
