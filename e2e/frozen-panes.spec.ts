import { test, expect } from './fixtures';
import type { Page } from '@playwright/test';

const log = (page: Page) => page.locator('y11n-spreadsheet').locator('[role="log"]');

/** Fill rows 7–50 so there is something to scroll through */
async function fillGrid(page: Page) {
  await page.evaluate(() => {
    const sheet = document.querySelector('y11n-spreadsheet') as any;
    const data = sheet.getData();
    for (let r = 6; r < 50; r++) {
      for (let c = 0; c < 26; c++) data.set(`${r}:${c}`, { rawValue: `r${r + 1}c${c}`, displayValue: '', type: 'text' });
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

test.describe('Frozen panes', () => {
  test('Alt+Shift+F freezes the header row and first column, which stay put while scrolling', async ({ spreadsheet, page }) => {
    await fillGrid(page);
    await spreadsheet.clickCell(1, 1);
    await spreadsheet.cell(1, 1).press('Alt+Shift+F');
    await expect(log(page)).toHaveText('Froze row 1 and column A');
    await expect(spreadsheet.host).toHaveAttribute('frozen-rows', '1');

    const a1Before = await boxInGrid(page, '[data-row="0"][data-col="0"]');
    await spreadsheet.grid.evaluate((grid) => {
      grid.scrollTop = 28 * 20;
      grid.scrollLeft = 600;
    });
    await page.waitForTimeout(200);

    expect(await boxInGrid(page, '[data-row="0"][data-col="0"]')).toEqual(a1Before);
    // Row 1 follows horizontal scrolling only; column A vertical scrolling only
    const c1 = await boxInGrid(page, '[data-row="0"][data-col="10"]');
    expect(c1.top).toBeCloseTo(a1Before.top, 0);
    const a30 = await boxInGrid(page, '[data-row="30"][data-col="0"]');
    expect(a30.left).toBeCloseTo(a1Before.left, 0);

    // What's on top at A1's position is A1 itself, not a scrolled cell
    const hit = await page.evaluate(() => {
      const root = document.querySelector('y11n-spreadsheet')!.shadowRoot!;
      const r = root.querySelector<HTMLElement>('[data-row="0"][data-col="0"]')!.getBoundingClientRect();
      return (root.elementFromPoint(r.left + 10, r.top + 10) as HTMLElement).closest<HTMLElement>('[data-key]')?.dataset.key;
    });
    expect(hit).toBe('0:0');
    await spreadsheet.cell(0, 0).click();
    await expect(spreadsheet.cell(0, 0)).toBeFocused();
  });

  test('moving up keeps the active cell clear of the frozen rows', async ({ spreadsheet, page }) => {
    await fillGrid(page);
    await page.evaluate(() => (document.querySelector('y11n-spreadsheet') as any).freezePanes(3, 0));
    await spreadsheet.clickCell(3, 1);
    await page.keyboard.press('Control+End');
    await page.keyboard.press('Home');
    await page.keyboard.press('ArrowRight');
    await expect(spreadsheet.cell(49, 1)).toBeFocused();
    for (let i = 0; i < 30; i++) await page.keyboard.press('ArrowUp');
    await expect(spreadsheet.cell(19, 1)).toBeFocused();

    const frozenBottom = (await boxInGrid(page, '[data-row="2"][data-col="1"]')).bottom;
    const active = await boxInGrid(page, '[data-row="19"][data-col="1"]');
    expect(active.top).toBeGreaterThanOrEqual(frozenBottom - 1);
  });

  test('moving left keeps the active cell clear of the frozen columns', async ({ spreadsheet, page }) => {
    await fillGrid(page);
    await page.evaluate(() => (document.querySelector('y11n-spreadsheet') as any).freezePanes(0, 2));
    await spreadsheet.clickCell(10, 2);
    await spreadsheet.cell(10, 2).press('End');
    for (let i = 0; i < 15; i++) await page.keyboard.press('ArrowLeft');
    await expect(spreadsheet.cell(10, 10)).toBeFocused();
    const frozenRight = (await boxInGrid(page, '[data-row="10"][data-col="1"]')).right;
    const active = await boxInGrid(page, '[data-row="10"][data-col="10"]');
    expect(active.left).toBeGreaterThanOrEqual(frozenRight - 1);
  });

  test('Ctrl+Home goes to the first unfrozen cell, and editing a frozen cell after scrolling lines up', async ({ spreadsheet, page }) => {
    await fillGrid(page);
    await page.evaluate(() => (document.querySelector('y11n-spreadsheet') as any).freezePanes(1, 1));
    await spreadsheet.clickCell(20, 5);
    await spreadsheet.cell(20, 5).press('Control+Home');
    await expect(spreadsheet.cell(1, 1)).toBeFocused();

    await spreadsheet.grid.evaluate((grid) => {
      grid.scrollTop = 28 * 15;
      grid.scrollLeft = 500;
    });
    await page.waitForTimeout(150);
    await spreadsheet.cell(0, 0).dblclick();
    const cellBox = (await spreadsheet.cell(0, 0).boundingBox())!;
    const editorBox = (await spreadsheet.editor.boundingBox())!;
    expect(Math.abs(editorBox.x - cellBox.x)).toBeLessThan(1);
    expect(Math.abs(editorBox.y - cellBox.y)).toBeLessThan(1);
  });

  test('freezing while scrolled keeps the active cell in view', async ({ spreadsheet, page }) => {
    await fillGrid(page);
    await spreadsheet.grid.evaluate((grid) => {
      grid.scrollTop = 28 * 10;
    });
    await page.waitForTimeout(150);
    await spreadsheet.clickCell(13, 2);
    await spreadsheet.cell(13, 2).press('Alt+Shift+F');
    await expect(log(page)).toHaveText('Froze rows 1 to 13 and columns A to B');
    await expect
      .poll(() =>
        page.evaluate(() => {
          const root = document.querySelector('y11n-spreadsheet')!.shadowRoot!;
          const r = root.querySelector<HTMLElement>('[data-row="13"][data-col="2"]')!.getBoundingClientRect();
          return (root.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2) as HTMLElement)
            .closest<HTMLElement>('[data-key]')?.dataset.key;
        })
      )
      .toBe('13:2');
    await expect(spreadsheet.cell(13, 2)).toBeFocused();
  });

  test('the fill handle stays under frozen rows', async ({ spreadsheet, page }) => {
    await fillGrid(page);
    await page.evaluate(() => (document.querySelector('y11n-spreadsheet') as any).freezePanes(2, 0));
    await spreadsheet.clickCell(3, 1);
    await spreadsheet.cell(2, 0).click({ modifiers: ['Shift'] });
    await spreadsheet.grid.evaluate((grid) => {
      grid.scrollTop = 28 * 3;
    });
    await page.waitForTimeout(150);
    const top = await page.evaluate(() => {
      const root = document.querySelector('y11n-spreadsheet')!.shadowRoot!;
      const r = root.querySelector<HTMLElement>('.ls-fill-handle')!.getBoundingClientRect();
      return (root.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2) as HTMLElement).className;
    });
    expect(top).not.toContain('ls-fill-handle');
  });

  test('inserting a row inside the frozen area grows it', async ({ spreadsheet, page }) => {
    await page.evaluate(() => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      sheet.freezePanes(1, 0);
      sheet.insertRows(0);
    });
    await expect(spreadsheet.host).toHaveAttribute('frozen-rows', '2');
    // The header row moved down and is still frozen
    await expect(spreadsheet.cell(1, 0)).toHaveClass(/frozen-row-edge/);
    expect(await spreadsheet.getCellText(1, 0)).toBe('Item');
  });
});
