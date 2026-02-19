import { test, expect } from './fixtures';

test.describe('Custom Events', () => {
  test('cell-change event fires when a cell value is committed', async ({ spreadsheet }) => {
    const eventPromise = spreadsheet.page.evaluate(() => {
      return new Promise<any>((resolve) => {
        const sheet = document.querySelector('y11n-spreadsheet')!;
        sheet.addEventListener(
          'cell-change',
          (e: any) => resolve(e.detail),
          { once: true }
        );
      });
    });

    await spreadsheet.clickCell(6, 0);
    await spreadsheet.cell(6, 0).press('Enter');
    await spreadsheet.typeInEditor('Event Test');
    await spreadsheet.commitWithEnter();

    const detail = await eventPromise;
    expect(detail.cellId).toBe('6:0');
    expect(detail.value).toBe('Event Test');
    expect(detail.oldValue).toBe('');
  });

  test('cell-change event includes old value', async ({ spreadsheet }) => {
    // Use a cell that has data in the demo (row 0, col 0 = "Item")
    const eventPromise = spreadsheet.page.evaluate(() => {
      return new Promise<any>((resolve) => {
        const sheet = document.querySelector('y11n-spreadsheet')!;
        sheet.addEventListener(
          'cell-change',
          (e: any) => resolve(e.detail),
          { once: true }
        );
      });
    });

    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Enter');
    await spreadsheet.typeInEditor('New Value');
    await spreadsheet.commitWithEnter();

    const detail = await eventPromise;
    expect(detail.oldValue).toBe('Item');
    expect(detail.value).toBe('New Value');
  });

  test('cell-change does not fire when value is unchanged', async ({ spreadsheet }) => {
    // Set up a cell with known data
    await spreadsheet.page.evaluate(() => {
      const sheet = document.querySelector('y11n-spreadsheet')!;
      (window as any).__cellChangeCount = 0;
      sheet.addEventListener('cell-change', () => {
        (window as any).__cellChangeCount++;
      });
    });

    // Edit cell 0:0 which has "Item" and commit without changing
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Enter');
    // Don't change the value - just commit
    await spreadsheet.commitWithEnter();

    await spreadsheet.page.waitForTimeout(200);
    const count = await spreadsheet.page.evaluate(() => (window as any).__cellChangeCount);
    expect(count).toBe(0);
  });

  test('selection-change event fires on cell click', async ({ spreadsheet }) => {
    const eventPromise = spreadsheet.page.evaluate(() => {
      return new Promise<any>((resolve) => {
        const sheet = document.querySelector('y11n-spreadsheet')!;
        sheet.addEventListener(
          'selection-change',
          (e: any) => resolve(e.detail),
          { once: true }
        );
      });
    });

    await spreadsheet.clickCell(3, 2);
    const detail = await eventPromise;

    expect(detail.range).toBeDefined();
    expect(detail.range.start).toBeDefined();
    expect(detail.range.end).toBeDefined();
  });

  test('selection-change event fires on keyboard navigation', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);

    const eventPromise = spreadsheet.page.evaluate(() => {
      return new Promise<any>((resolve) => {
        const sheet = document.querySelector('y11n-spreadsheet')!;
        sheet.addEventListener(
          'selection-change',
          (e: any) => resolve(e.detail),
          { once: true }
        );
      });
    });

    await spreadsheet.cell(0, 0).press('ArrowDown');
    const detail = await eventPromise;

    expect(detail.range.start.row).toBe(1);
    expect(detail.range.start.col).toBe(0);
  });

  test('data-change event fires on delete', async ({ spreadsheet }) => {
    // Use an existing cell from demo data (row 1, col 0 = "Widget A")
    const eventPromise = spreadsheet.page.evaluate(() => {
      return new Promise<any>((resolve) => {
        const sheet = document.querySelector('y11n-spreadsheet')!;
        sheet.addEventListener(
          'data-change',
          (e: any) => resolve(e.detail),
          { once: true }
        );
      });
    });

    await spreadsheet.clickCell(1, 0);
    await spreadsheet.cell(1, 0).press('Delete');

    const detail = await eventPromise;
    expect(detail.updates).toBeDefined();
    expect(detail.updates.length).toBeGreaterThan(0);
    expect(detail.updates[0].id).toBe('1:0');
    expect(detail.updates[0].value).toBe('');
  });

  test('data-change event fires on paste', async ({ spreadsheet, page }) => {
    await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);

    await page.evaluate(() =>
      navigator.clipboard.writeText('Pasted')
    );

    const eventPromise = page.evaluate(() => {
      return new Promise<any>((resolve) => {
        const sheet = document.querySelector('y11n-spreadsheet')!;
        sheet.addEventListener(
          'data-change',
          (e: any) => resolve(e.detail),
          { once: true }
        );
      });
    });

    await spreadsheet.clickCell(7, 0);
    await spreadsheet.cell(7, 0).press('Control+v');

    const detail = await eventPromise;
    expect(detail.updates).toBeDefined();
    expect(detail.updates.length).toBeGreaterThan(0);
  });

  test('events bubble through shadow DOM (composed: true)', async ({ spreadsheet }) => {
    const eventPromise = spreadsheet.page.evaluate(() => {
      return new Promise<boolean>((resolve) => {
        document.addEventListener(
          'cell-change',
          () => resolve(true),
          { once: true }
        );
        setTimeout(() => resolve(false), 3000);
      });
    });

    await spreadsheet.clickCell(6, 5);
    await spreadsheet.cell(6, 5).press('Enter');
    await spreadsheet.typeInEditor('Bubble Test');
    await spreadsheet.commitWithEnter();

    const didBubble = await eventPromise;
    expect(didBubble).toBe(true);
  });
});
