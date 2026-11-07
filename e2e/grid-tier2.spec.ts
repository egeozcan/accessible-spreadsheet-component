import AxeBuilder from '@axe-core/playwright';
import { test, expect } from './fixtures';

// The demo page seeds A1:D4 (header + 3 items with Qty × Price totals in D),
// A6 "Grand Total" and D6 =SUM(D2:D4) = 231.4, in a 50-row × 26-column grid.


test.describe('Fill handle', () => {
  test('dragging the handle continues a numeric series', async ({ spreadsheet, page }) => {
    for (const [row, value] of [[0, '1'], [1, '2']] as const) {
      await spreadsheet.clickCell(row, 5);
      await spreadsheet.cell(row, 5).press('F2');
      await spreadsheet.typeInEditor(value);
      await spreadsheet.commitWithEnter();
    }
    await spreadsheet.clickCell(0, 5);
    await spreadsheet.cell(0, 5).press('Shift+ArrowDown');

    const handle = spreadsheet.cell(1, 5).locator('.ls-fill-handle');
    await expect(handle).toBeVisible();
    const from = (await handle.boundingBox())!;
    const to = (await spreadsheet.cell(4, 5).boundingBox())!;
    await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2);
    await page.mouse.down();
    await page.mouse.move(to.x + to.width / 2, to.y + to.height / 2, { steps: 5 });
    await expect(spreadsheet.cell(4, 5)).toHaveClass(/fill-preview/);
    await page.mouse.up();

    await expect(spreadsheet.cell(2, 5).locator('.cell-text')).toHaveText('3');
    await expect(spreadsheet.cell(3, 5).locator('.cell-text')).toHaveText('4');
    await expect(spreadsheet.cell(4, 5).locator('.cell-text')).toHaveText('5');
    await expect(spreadsheet.cell(4, 5)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(4, 5)).not.toHaveClass(/fill-preview/);
  });

  test('dragging a formula down shifts its references', async ({ spreadsheet, page }) => {
    await spreadsheet.clickCell(1, 3);
    const handle = spreadsheet.cell(1, 3).locator('.ls-fill-handle');
    const from = (await handle.boundingBox())!;
    const to = (await spreadsheet.cell(3, 3).boundingBox())!;
    await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2);
    await page.mouse.down();
    await page.mouse.move(to.x + to.width / 2, to.y + to.height / 2, { steps: 5 });
    await page.mouse.up();
    const data = await spreadsheet.getData();
    expect(data['2:3'].rawValue).toBe('=B3*C3');
    expect(data['3:3'].rawValue).toBe('=B4*C4');
  });
});

test.describe('Find', () => {
  test('Ctrl+F finds and steps through matches', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Control+f');
    const input = spreadsheet.shadow('#find-input');
    await expect(input).toBeFocused();

    await input.fill('widget');
    await expect(spreadsheet.cell(1, 0)).toHaveClass(/active-cell/);
    await expect(spreadsheet.shadow('#find-status')).toHaveText('1 of 3');

    await input.press('Enter');
    await expect(spreadsheet.cell(2, 0)).toHaveClass(/active-cell/);
    await input.press('Shift+Enter');
    await expect(spreadsheet.cell(1, 0)).toHaveClass(/active-cell/);

    await spreadsheet.shadow('button[aria-label="Next match"]').click();
    await expect(spreadsheet.cell(2, 0)).toHaveClass(/active-cell/);

    await input.press('Escape');
    await expect(spreadsheet.shadow('#find-input')).toHaveCount(0);
    await expect(spreadsheet.cell(2, 0)).toBeFocused();
  });
});

test.describe('Insert and delete rows/columns', () => {
  test('keyboard insert and delete of whole rows keeps formulas pointing at their cells', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(1, 0);
    await spreadsheet.cell(1, 0).press('Shift+Space');
    await spreadsheet.cell(1, 0).press('Control+Shift+Equal');

    await expect(spreadsheet.cell(1, 0).locator('.cell-text')).toHaveText('');
    await expect(spreadsheet.cell(2, 0).locator('.cell-text')).toHaveText('Widget A');
    await expect(spreadsheet.cell(6, 3).locator('.cell-text')).toHaveText('231.4');
    let data = await spreadsheet.getData();
    expect(data['6:3'].rawValue).toBe('=SUM(D3:D5)');
    expect(data['2:3'].rawValue).toBe('=B3*C3');
    await expect(spreadsheet.grid).toHaveAttribute('aria-rowcount', '52');

    // Delete the (now empty) selected row again
    await spreadsheet.cell(1, 0).press('Control+Minus');
    await expect(spreadsheet.cell(1, 0).locator('.cell-text')).toHaveText('Widget A');
    data = await spreadsheet.getData();
    expect(data['5:3'].rawValue).toBe('=SUM(D2:D4)');
    await expect(spreadsheet.grid).toHaveAttribute('aria-rowcount', '51');
  });

  test('deleting a referenced column produces #REF! and undo restores it', async ({ spreadsheet, page }) => {
    await page.evaluate(() => (document.querySelector('y11n-spreadsheet') as any).deleteColumns(1));
    await expect(spreadsheet.cell(1, 2).locator('.cell-text')).toHaveText('#REF!');
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Control+z');
    await expect(spreadsheet.cell(1, 3).locator('.cell-text')).toHaveText('59.9');
    await expect(spreadsheet.cell(0, 1).locator('.cell-text')).toHaveText('Qty');
  });
});

test.describe('Read-only cells and validation', () => {
  test('a read-only cell refuses edits and is announced as read-only', async ({ spreadsheet, page }) => {
    await page.evaluate(() => {
      const el = document.querySelector('y11n-spreadsheet') as any;
      const data = el.getData();
      data.get('0:0').readOnly = true;
      el.setData(data);
    });
    await expect(spreadsheet.cell(0, 0)).toHaveAttribute('aria-readonly', 'true');
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('x');
    await expect(spreadsheet.editor).toBeHidden();
    await expect(spreadsheet.shadow('[role="log"]')).toHaveText('A1 is read-only');
    await spreadsheet.cell(0, 0).press('Delete');
    await expect(spreadsheet.cell(0, 0).locator('.cell-text')).toHaveText('Item');
  });

  test('validateCell keeps the editor open with a message', async ({ spreadsheet, page }) => {
    await page.evaluate(() => {
      (document.querySelector('y11n-spreadsheet') as any).validateCell = (_id: string, value: string) =>
        /^\d+$/.test(value) ? null : 'Enter a whole number';
    });
    await spreadsheet.clickCell(1, 1);
    await spreadsheet.cell(1, 1).press('F2');
    await spreadsheet.typeInEditor('ten');
    await spreadsheet.commitWithEnter();

    await expect(spreadsheet.editor).toBeVisible();
    await expect(spreadsheet.editor).toHaveAttribute('aria-invalid', 'true');
    await expect(spreadsheet.shadow('#editor-error')).toHaveText('Enter a whole number');

    await spreadsheet.typeInEditor('12');
    await spreadsheet.commitWithEnter();
    await expect(spreadsheet.editor).toBeHidden();
    await expect(spreadsheet.cell(1, 1).locator('.cell-text')).toHaveText('12');
  });

  test('before-change can rewrite a value', async ({ spreadsheet, page }) => {
    await page.evaluate(() => {
      document.querySelector('y11n-spreadsheet')!.addEventListener('before-change', (e: any) => {
        for (const change of e.detail.changes) change.value = change.value.trim();
      });
    });
    await spreadsheet.clickCell(8, 0);
    await spreadsheet.cell(8, 0).press('F2');
    await spreadsheet.typeInEditor('   padded   ');
    await spreadsheet.commitWithEnter();
    const data = await spreadsheet.getData();
    expect(data['8:0'].rawValue).toBe('padded');
  });
});

test.describe('Accessibility of the new UI', () => {
  test('the validation message and find bar have no axe violations', async ({ spreadsheet, page }) => {
    await page.evaluate(() => {
      (document.querySelector('y11n-spreadsheet') as any).validateCell = () => 'Always invalid';
    });
    await spreadsheet.clickCell(1, 1);
    await spreadsheet.cell(1, 1).press('F2');
    await spreadsheet.typeInEditor('x');
    await spreadsheet.commitWithEnter();
    await expect(spreadsheet.shadow('#editor-error')).toBeVisible();

    let results = await new AxeBuilder({ page }).include('y11n-spreadsheet').disableRules(['color-contrast']).analyze();
    expect(results.violations.map((v) => v.id)).toEqual([]);

    await spreadsheet.editor.press('Escape');
    await spreadsheet.cell(1, 1).press('Control+f');
    await spreadsheet.shadow('#find-input').fill('widget');
    await expect(spreadsheet.shadow('#find-status')).toHaveText('2 of 3');
    results = await new AxeBuilder({ page }).include('y11n-spreadsheet').analyze();
    expect(results.violations.map((v) => v.id)).toEqual([]);
  });
});
