import { test, expect } from './fixtures';

// The demo page seeds A1:D4 (header + 3 rows of items) plus A6 and D6,
// in a 50-row × 26-column grid.

test.describe('Ctrl/Cmd+Arrow data-edge navigation', () => {
  test('Ctrl+ArrowDown walks data runs, gaps, and the grid edge', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Control+ArrowDown');
    await expect(spreadsheet.cell(3, 0)).toHaveClass(/active-cell/);

    await spreadsheet.cell(3, 0).press('Control+ArrowDown');
    await expect(spreadsheet.cell(5, 0)).toHaveClass(/active-cell/);

    await spreadsheet.cell(5, 0).press('Control+ArrowDown');
    await expect(spreadsheet.cell(49, 0)).toHaveClass(/active-cell/);
    await expect(spreadsheet.cell(49, 0)).toBeFocused();

    await spreadsheet.cell(49, 0).press('Control+ArrowUp');
    await expect(spreadsheet.cell(5, 0)).toHaveClass(/active-cell/);
  });

  test('Ctrl+ArrowRight jumps to the end of the row data', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Control+ArrowRight');
    await expect(spreadsheet.cell(0, 3)).toHaveClass(/active-cell/);
  });

  test('Ctrl+Shift+ArrowDown extends the selection to the data edge', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 1);
    await spreadsheet.cell(0, 1).press('Control+Shift+ArrowDown');
    for (const row of [0, 1, 2, 3]) {
      await expect(spreadsheet.cell(row, 1)).toHaveAttribute('aria-selected', 'true');
    }
    await expect(spreadsheet.cell(4, 1)).toHaveAttribute('aria-selected', 'false');
  });
});

test.describe('Row and column selection', () => {
  test('clicking a column header selects the whole column', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(2, 0);
    await spreadsheet.shadow('[data-header-col="1"]').click();

    await expect(spreadsheet.shadow('[data-header-col="1"]')).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(0, 1)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(10, 1)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(0, 0)).toHaveAttribute('aria-selected', 'false');
    // The active cell moves into the column but keeps its row
    await expect(spreadsheet.cell(2, 1)).toHaveClass(/active-cell/);
    await expect(spreadsheet.cell(2, 1)).toBeFocused();
  });

  test('Shift+click on a column header extends to a span of columns', async ({ spreadsheet }) => {
    await spreadsheet.shadow('[data-header-col="1"]').click();
    await spreadsheet.shadow('[data-header-col="3"]').click({ modifiers: ['Shift'] });
    for (const col of [1, 2, 3]) {
      await expect(spreadsheet.cell(5, col)).toHaveAttribute('aria-selected', 'true');
    }
    await expect(spreadsheet.cell(5, 4)).toHaveAttribute('aria-selected', 'false');
  });

  test('dragging across column headers selects the span', async ({ spreadsheet, page }) => {
    const from = await spreadsheet.shadow('[data-header-col="0"]').boundingBox();
    const to = await spreadsheet.shadow('[data-header-col="2"]').boundingBox();
    await page.mouse.move(from!.x + 5, from!.y + 5);
    await page.mouse.down();
    await page.mouse.move(to!.x + 5, to!.y + 5, { steps: 5 });
    await page.mouse.up();
    await expect(spreadsheet.shadow('[data-header-col="2"]')).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(7, 2)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(7, 3)).toHaveAttribute('aria-selected', 'false');
  });

  test('clicking a row header selects the whole row', async ({ spreadsheet }) => {
    await spreadsheet.shadow('[data-header-row="2"]').click();
    await expect(spreadsheet.shadow('[data-header-row="2"]')).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(2, 0)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(2, 7)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(1, 0)).toHaveAttribute('aria-selected', 'false');
  });

  test('Ctrl+Space selects the column, Shift+Space the row', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(4, 2);
    await spreadsheet.cell(4, 2).press('Control+Space');
    await expect(spreadsheet.cell(0, 2)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(4, 2)).toHaveClass(/active-cell/);
    await expect(spreadsheet.liveRegion).toContainText('Column C selected');

    await spreadsheet.clickCell(4, 2);
    await spreadsheet.cell(4, 2).press('Shift+Space');
    await expect(spreadsheet.cell(4, 0)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(3, 2)).toHaveAttribute('aria-selected', 'false');
    await expect(spreadsheet.liveRegion).toContainText('Row 5 selected');
    expect(await spreadsheet.isEditorVisible()).toBe(false);
  });
});

test.describe('Fill and multi-cell entry', () => {
  test('Ctrl+D fills a formula down with shifted references', async ({ spreadsheet }) => {
    await spreadsheet.dblClickCell(1, 4);
    await spreadsheet.typeInEditor('=D2*2');
    await spreadsheet.commitWithEnter();

    await spreadsheet.clickCell(1, 4);
    await spreadsheet.cell(3, 4).click({ modifiers: ['Shift'] });
    await spreadsheet.cell(3, 4).press('Control+d');

    await expect(spreadsheet.cell(2, 4).locator('.cell-text')).toHaveText('175');
    await expect(spreadsheet.cell(3, 4).locator('.cell-text')).toHaveText('168');

    await spreadsheet.cell(3, 4).press('Control+z');
    await expect(spreadsheet.cell(2, 4).locator('.cell-text')).toHaveText('');
  });

  test('Ctrl+R fills right', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 2).click({ modifiers: ['Shift'] });
    await spreadsheet.cell(0, 2).press('Control+r');
    await expect(spreadsheet.cell(0, 1).locator('.cell-text')).toHaveText('Item');
    await expect(spreadsheet.cell(0, 2).locator('.cell-text')).toHaveText('Item');
  });

  test('Ctrl+Enter writes the entry into every selected cell', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(10, 5);
    await spreadsheet.cell(12, 5).click({ modifiers: ['Shift'] });
    await spreadsheet.cell(12, 5).press('x');
    await expect(spreadsheet.editor).toBeVisible();
    await spreadsheet.editor.press('Control+Enter');

    for (const row of [10, 11, 12]) {
      await expect(spreadsheet.cell(row, 5).locator('.cell-text')).toHaveText('x');
    }
    await expect(spreadsheet.cell(12, 5)).toHaveClass(/active-cell/);
    await expect(spreadsheet.cell(10, 5)).toHaveAttribute('aria-selected', 'true');
  });

  test('Shift+Enter commits and moves up', async ({ spreadsheet }) => {
    await spreadsheet.dblClickCell(8, 1);
    await spreadsheet.typeInEditor('up');
    await spreadsheet.editor.press('Shift+Enter');
    await expect(spreadsheet.cell(8, 1).locator('.cell-text')).toHaveText('up');
    await expect(spreadsheet.cell(7, 1)).toHaveClass(/active-cell/);
  });
});

test.describe('CSV API', () => {
  test('toCSV and fromCSV round-trip', async ({ spreadsheet, page }) => {
    const csv = await page.evaluate(() => (document.querySelector('y11n-spreadsheet') as any).toCSV());
    expect(csv.split('\n')[0]).toBe('Item,Qty,Price,Total');
    expect(csv.split('\n')[1]).toBe('Widget A,10,5.99,59.9');

    await page.evaluate(() =>
      (document.querySelector('y11n-spreadsheet') as any).fromCSV('a,"b,c"\n1,=A12+1', {
        start: { row: 10, col: 0 },
      })
    );
    await expect(spreadsheet.cell(10, 1).locator('.cell-text')).toHaveText('b,c');
    await expect(spreadsheet.cell(11, 1).locator('.cell-text')).toHaveText('2');
  });
});
