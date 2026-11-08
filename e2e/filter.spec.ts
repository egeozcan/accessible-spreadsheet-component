import { test, expect } from './fixtures';
import type { Page } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';

test.use({ permissions: ['clipboard-read', 'clipboard-write'] });

const log = (page: Page) => page.locator('y11n-spreadsheet').locator('[role="log"]');
const menu = (page: Page) => page.locator('y11n-spreadsheet').locator('.ls-filter-menu');

/** A table in A1:C201: a header row, then 200 rows whose column B cycles Red, Green, Blue. */
async function seedTable(page: Page) {
  await page.evaluate(() => {
    const sheet = document.querySelector('y11n-spreadsheet') as any;
    sheet.rows = 250;
    const colors = ['Red', 'Green', 'Blue'];
    const data = new Map();
    const set = (key: string, value: string) =>
      data.set(key, { rawValue: value, displayValue: value, type: /^\d+$/.test(value) ? 'number' : 'text' });
    set('0:0', 'Id');
    set('0:1', 'Color');
    set('0:2', 'Size');
    for (let r = 1; r <= 200; r++) {
      set(`${r}:0`, String(r));
      set(`${r}:1`, colors[(r - 1) % 3]);
      set(`${r}:2`, String((r * 7) % 10));
    }
    sheet.setData(data);
  });
}

/** Bounds of an element relative to the grid's visible area */
async function boxInGrid(page: Page, selector: string) {
  return page.evaluate((sel) => {
    const root = document.querySelector('y11n-spreadsheet')!.shadowRoot!;
    const grid = root.querySelector<HTMLElement>('[role="grid"]')!.getBoundingClientRect();
    const r = root.querySelector<HTMLElement>(sel)!.getBoundingClientRect();
    return { top: r.top - grid.top, left: r.left - grid.left, bottom: r.bottom - grid.top, right: r.right - grid.left };
  }, selector);
}

test.describe('Filter', () => {
  test('keyboard: add a filter, pick values in the menu, and move over the remaining rows', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await spreadsheet.clickCell(3, 1);
    await page.keyboard.press('Control+Shift+L');
    await expect(log(page)).toHaveText(
      'Filter added to A1:C201. Press Alt+Down on a header cell for filter options'
    );

    await page.keyboard.press('Control+Home');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Alt+ArrowDown');
    await expect(menu(page)).toBeVisible();
    await expect(spreadsheet.shadow('#filter-search')).toBeFocused();
    await expect(menu(page).locator('.ls-filter-option')).toHaveText([
      '(Select all)',
      'Blue',
      'Green',
      'Red',
    ]);

    // Down into the list, past Select all and Blue, and uncheck Green and Red
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('Space');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('Space');
    await page.keyboard.press('Enter');

    await expect(menu(page)).toBeHidden();
    await expect(log(page)).toHaveText('Showing 66 of 200 rows');
    await expect(spreadsheet.cell(0, 1)).toBeFocused();
    await expect(spreadsheet.cell(0, 1).locator('.ls-filter-button')).toHaveClass(/filtered/);

    // Row 3 (Blue) sits right under the header row; rows 1 and 2 are gone
    await expect(spreadsheet.cell(1, 0)).toHaveCount(0);
    const header = await boxInGrid(page, '[data-row="0"][data-col="0"]');
    const first = await boxInGrid(page, '[data-row="3"][data-col="0"]');
    expect(first.top).toBeCloseTo(header.bottom, 0);

    await page.keyboard.press('ArrowDown');
    await expect(spreadsheet.cell(3, 1)).toBeFocused();
    await page.keyboard.press('ArrowDown');
    await expect(spreadsheet.cell(6, 1)).toBeFocused();
  });

  test('scrolling and Ctrl+End reach the last visible row', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await page.evaluate(() => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      sheet.rows = 201;
      sheet.setFilter({ start: { row: 0, col: 0 }, end: { row: 200, col: 2 } }, [{ col: 1, values: ['Red'] }]);
    });
    await spreadsheet.clickCell(1, 0);
    await page.keyboard.press('Control+End');
    // Row 199 is the last Red row
    await expect(spreadsheet.cell(199, 25)).toBeFocused();
    const box = await boxInGrid(page, '[data-row="199"][data-col="25"]');
    const gridHeight = await spreadsheet.grid.evaluate((g) => g.clientHeight);
    expect(box.top).toBeGreaterThan(0);
    expect(box.bottom).toBeLessThanOrEqual(gridHeight + 1);

    // The scroll range covers only the 67 visible data rows plus the header
    const { scrollHeight, cellHeight } = await spreadsheet.grid.evaluate((g) => ({
      scrollHeight: g.scrollHeight,
      cellHeight: g.querySelector<HTMLElement>('[data-row]')!.getBoundingClientRect().height,
    }));
    expect(scrollHeight).toBeCloseTo((1 + 1 + 67) * cellHeight, -1);

    // Every rendered row sits one row height below the previous one
    const tops = await page.evaluate(() => {
      const root = document.querySelector('y11n-spreadsheet')!.shadowRoot!;
      return [...root.querySelectorAll<HTMLElement>('[data-col="0"][data-row]')].map((c) => c.getBoundingClientRect().top);
    });
    for (let i = 1; i < tops.length; i++) expect(tops[i] - tops[i - 1]).toBeCloseTo(cellHeight, 0);
  });

  test('mouse: the filter button opens the menu under its header, an outside click closes it', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await page.evaluate(() => {
      (document.querySelector('y11n-spreadsheet') as any).setFilter({ start: { row: 0, col: 0 }, end: { row: 200, col: 2 } });
    });
    await spreadsheet.cell(0, 2).locator('.ls-filter-button').click();
    await expect(menu(page)).toBeVisible();
    const headerBox = await boxInGrid(page, '[data-row="0"][data-col="2"]');
    const menuBox = await boxInGrid(page, '.ls-filter-menu');
    expect(menuBox.top).toBeCloseTo(headerBox.bottom, 0);
    expect(menuBox.left).toBeCloseTo(headerBox.left, 0);

    await menu(page).getByText('(Select all)').click();
    await menu(page).getByText('7', { exact: true }).click();
    await menu(page).getByRole('button', { name: 'OK' }).click();
    await expect.poll(() => page.evaluate(() => (document.querySelector('y11n-spreadsheet') as any).getFilter().columns))
      .toEqual([{ col: 2, values: ['7'] }]);

    await spreadsheet.cell(0, 2).locator('.ls-filter-button').click();
    await expect(menu(page)).toBeVisible();
    await spreadsheet.cell(0, 0).click();
    await expect(menu(page)).toBeHidden();
    await expect(spreadsheet.cell(0, 0)).toBeFocused();
  });

  test('copy and paste skip hidden rows', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await page.evaluate(() => {
      (document.querySelector('y11n-spreadsheet') as any).setFilter(
        { start: { row: 0, col: 0 }, end: { row: 200, col: 2 } },
        [{ col: 1, values: ['Red'] }]
      );
    });
    // Copy A2:A8 (visible: rows 1, 4, 7)
    await spreadsheet.clickCell(1, 0);
    await page.keyboard.press('Shift+ArrowDown');
    await page.keyboard.press('Shift+ArrowDown');
    await spreadsheet.copyFromCell(7, 0);
    expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('1\n4\n7');

    // Paste into column D from row 2: lands on rows 1, 4 and 7
    await spreadsheet.clickCell(1, 3);
    await page.keyboard.press('Control+v');
    await spreadsheet.waitForCellText(7, 3, '7');
    const column = await page.evaluate(() => {
      const data = (document.querySelector('y11n-spreadsheet') as any).getData();
      return [1, 2, 3, 4, 5, 6, 7].map((r) => data.get(`${r}:3`)?.rawValue ?? '');
    });
    expect(column).toEqual(['1', '', '', '4', '', '', '7']);
  });

  test('a frozen header row stays on top of a filtered table', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await page.evaluate(() => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      sheet.freezePanes(1, 0);
      sheet.setFilter({ start: { row: 0, col: 0 }, end: { row: 200, col: 2 } }, [{ col: 1, values: ['Green'] }]);
    });
    const before = await boxInGrid(page, '[data-row="0"][data-col="1"]');
    await spreadsheet.grid.evaluate((g) => (g.scrollTop = 28 * 30));
    await page.waitForTimeout(200);
    expect(await boxInGrid(page, '[data-row="0"][data-col="1"]')).toEqual(before);
    await spreadsheet.cell(0, 1).locator('.ls-filter-button').click();
    await expect(menu(page)).toBeVisible();
    expect((await boxInGrid(page, '.ls-filter-menu')).top).toBeCloseTo(before.bottom, 0);
  });

  test('the open filter menu has no axe violations', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await spreadsheet.clickCell(0, 1);
    await page.keyboard.press('Control+Shift+L');
    await page.keyboard.press('Alt+ArrowDown');
    await expect(menu(page)).toBeVisible();
    const results = await new AxeBuilder({ page }).include('y11n-spreadsheet').disableRules(['color-contrast']).analyze();
    expect(results.violations.map((v) => `${v.id}: ${v.help}`)).toEqual([]);
  });

  test('undo restores the rows a filter hid', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await spreadsheet.clickCell(0, 1);
    await page.keyboard.press('Control+Shift+L');
    await page.keyboard.press('Alt+ArrowDown');
    await page.keyboard.type('Red');
    await page.keyboard.press('Enter');
    await expect(spreadsheet.cell(2, 0)).toHaveCount(0);
    await page.keyboard.press('Control+z');
    await expect(spreadsheet.cell(2, 0)).toHaveCount(1);
    await page.keyboard.press('Control+z');
    await expect(spreadsheet.cell(0, 1).locator('.ls-filter-button')).toHaveCount(0);
  });
});
