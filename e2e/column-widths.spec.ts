import { test, expect } from './fixtures';
import type { Page } from '@playwright/test';

const log = (page: Page) => page.locator('y11n-spreadsheet').locator('[role="log"]');

test.describe('Column widths', () => {
  test('dragging a header border resizes the column, and Ctrl+Z undoes it', async ({ spreadsheet, page }) => {
    const handle = spreadsheet.shadow('[data-resize-col="0"]');
    const box = (await handle.boundingBox())!;
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2 + 60, box.y + box.height / 2, { steps: 5 });
    // Live preview while dragging
    await expect.poll(async () => (await spreadsheet.cell(0, 0).boundingBox())?.width).toBeCloseTo(160, 0);
    await page.mouse.up();

    await expect(log(page)).toHaveText('Column A width 160 pixels');
    const a = (await spreadsheet.cell(0, 0).boundingBox())!;
    const b = (await spreadsheet.cell(0, 1).boundingBox())!;
    expect(b.x).toBeCloseTo(a.x + 160, 0);
    // The column header was not selected by the drag
    await expect(spreadsheet.shadow('[data-header-col="0"]')).not.toHaveAttribute('aria-selected', 'true');

    await spreadsheet.clickCell(2, 2);
    await spreadsheet.cell(2, 2).press('Control+z');
    await expect.poll(async () => (await spreadsheet.cell(0, 0).boundingBox())?.width).toBeCloseTo(100, 0);
  });

  test('Escape cancels a drag', async ({ spreadsheet, page }) => {
    const box = (await spreadsheet.shadow('[data-resize-col="1"]').boundingBox())!;
    await page.mouse.move(box.x + 3, box.y + 5);
    await page.mouse.down();
    await page.mouse.move(box.x + 83, box.y + 5, { steps: 3 });
    await page.keyboard.press('Escape');
    await page.mouse.up();
    await expect.poll(async () => (await spreadsheet.cell(0, 1).boundingBox())?.width).toBeCloseTo(100, 0);
  });

  test('Alt+Shift+Right widens the active column and announces it; works in read-only mode', async ({ spreadsheet, page }) => {
    await spreadsheet.setReadOnly(true);
    await spreadsheet.clickCell(1, 1);
    await spreadsheet.cell(1, 1).press('Alt+Shift+ArrowRight');
    await spreadsheet.cell(1, 1).press('Alt+Shift+ArrowRight');
    await expect(log(page)).toHaveText('Column B width 116 pixels');
    expect((await spreadsheet.cell(1, 1).boundingBox())!.width).toBeCloseTo(116, 0);
    await expect(spreadsheet.cell(1, 1)).toBeFocused();
  });

  test('double-clicking a header border fits the column to its content', async ({ spreadsheet }) => {
    await spreadsheet.page.evaluate(() => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      const data = sheet.getData();
      data.set('30:0', { rawValue: 'A considerably longer item name', displayValue: '', type: 'text' });
      sheet.setData(data);
    });
    await spreadsheet.shadow('[data-resize-col="0"]').dblclick();

    // The long value (rendered or not) fits: scroll it into view and check it isn't clipped
    await spreadsheet.page.evaluate(() => {
      const grid = document.querySelector('y11n-spreadsheet')!.shadowRoot!.querySelector('[role="grid"]')!;
      grid.scrollTop = 30 * 28;
    });
    const text = spreadsheet.cell(30, 0).locator('.cell-text');
    await expect(text).toBeVisible();
    const fits = await spreadsheet.cell(30, 0).evaluate((cell) => cell.scrollWidth <= cell.clientWidth);
    expect(fits).toBe(true);
    expect((await spreadsheet.cell(30, 0).boundingBox())!.width).toBeGreaterThan(150);
  });

  test('Alt+Shift+0 fits the selected columns', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 1);
    await spreadsheet.cell(0, 1).press('Alt+Shift+0');
    // "Qty" and small numbers: narrower than the default
    expect((await spreadsheet.cell(0, 1).boundingBox())!.width).toBeLessThan(60);
  });

  test('fitting the last column at the far right keeps it fully in view', async ({ spreadsheet }) => {
    await spreadsheet.page.evaluate(() => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      const data = sheet.getData();
      data.set('0:25', { rawValue: 'x'.repeat(60), displayValue: '', type: 'text' });
      sheet.setData(data);
    });
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('End');
    await spreadsheet.cell(0, 25).press('Alt+Shift+0');

    await expect
      .poll(() =>
        spreadsheet.page.evaluate(() => {
          const root = document.querySelector('y11n-spreadsheet')!.shadowRoot!;
          const grid = root.querySelector<HTMLElement>('[role="grid"]')!;
          const c = root.querySelector<HTMLElement>('[data-row="0"][data-col="25"]')!.getBoundingClientRect();
          return Math.round(grid.getBoundingClientRect().left + grid.clientWidth - c.right);
        })
      )
      .toBeGreaterThanOrEqual(-1);
    expect((await spreadsheet.cell(0, 25).boundingBox())!.width).toBeGreaterThan(300);

    await spreadsheet.cell(0, 25).press('F2');
    const cellBox = (await spreadsheet.cell(0, 25).boundingBox())!;
    const editorBox = (await spreadsheet.editor.boundingBox())!;
    expect(Math.abs(editorBox.x - cellBox.x)).toBeLessThan(1);
  });

  test('resizing a column far from the active cell does not scroll back to it', async ({ spreadsheet, page }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.grid.evaluate((grid) => {
      grid.scrollLeft = 1200;
    });
    await page.waitForTimeout(100);
    const header = spreadsheet.shadow('[data-resize-col="15"]');
    await header.scrollIntoViewIfNeeded();
    const before = await spreadsheet.grid.evaluate((grid) => grid.scrollLeft);
    const box = (await header.boundingBox())!;
    await page.mouse.move(box.x + 3, box.y + 5);
    await page.mouse.down();
    await page.mouse.move(box.x + 43, box.y + 5, { steps: 3 });
    await page.mouse.up();
    await expect.poll(() => spreadsheet.page.evaluate(() => (document.querySelector('y11n-spreadsheet') as any).getColumnWidth(15))).toBe(140);
    await page.waitForTimeout(100);
    expect(await spreadsheet.grid.evaluate((grid) => grid.scrollLeft)).toBe(before);
  });

  test('scrolling, keyboard navigation and the editor follow mixed widths', async ({ spreadsheet }) => {
    await spreadsheet.page.evaluate(() => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      const widths: Record<number, number> = {};
      for (let c = 0; c < 26; c++) widths[c] = c % 3 === 0 ? 260 : c % 3 === 1 ? 45 : 130;
      sheet.setColumnWidths(widths);
    });
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('End');
    await expect(spreadsheet.cell(0, 25)).toBeFocused();

    const gaps = await spreadsheet.page.evaluate(() => {
      const root = document.querySelector('y11n-spreadsheet')!.shadowRoot!;
      const grid = root.querySelector<HTMLElement>('[role="grid"]')!;
      const cell = root.querySelector<HTMLElement>('[data-row="0"][data-col="25"]')!;
      const g = grid.getBoundingClientRect();
      const c = cell.getBoundingClientRect();
      return { left: c.left - (g.left + 50), right: g.left + grid.clientWidth - c.right };
    });
    expect(gaps.left).toBeGreaterThanOrEqual(-1);
    expect(gaps.right).toBeGreaterThanOrEqual(-1);

    await spreadsheet.cell(0, 25).press('F2');
    const cellBox = (await spreadsheet.cell(0, 25).boundingBox())!;
    const editorBox = (await spreadsheet.editor.boundingBox())!;
    expect(Math.abs(editorBox.x - cellBox.x)).toBeLessThan(1);
    expect(editorBox.width).toBeCloseTo(cellBox.width, 0);

    // A cell in the middle of the scrolled view is rendered where the layout puts it
    const centre = await spreadsheet.page.evaluate(() => {
      const root = document.querySelector('y11n-spreadsheet')!.shadowRoot!;
      const grid = root.querySelector<HTMLElement>('[role="grid"]')!;
      grid.scrollLeft = 1500;
      return new Promise<string>((resolve) =>
        requestAnimationFrame(() =>
          requestAnimationFrame(() => {
            const r = grid.getBoundingClientRect();
            const hit = root.elementFromPoint(r.left + r.width / 2, r.top + 60) as HTMLElement | null;
            // Which column the widths put there: content x minus the 50px row header
            const sheet = document.querySelector('y11n-spreadsheet') as any;
            const x = grid.scrollLeft + r.width / 2 - 50;
            let col = 0;
            for (let left = 0; left + sheet.getColumnWidth(col) <= x; col++) left += sheet.getColumnWidth(col);
            resolve(`${hit?.closest<HTMLElement>('[data-col]')?.dataset.col} vs ${col}`);
          })
        )
      );
    });
    const [rendered, expected] = centre!.split(' vs ');
    expect(rendered).toBe(expected);
  });
});
