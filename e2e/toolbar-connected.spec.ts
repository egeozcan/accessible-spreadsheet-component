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
    await expect(spreadsheet.shadow('[role="log"]')).toHaveText('Froze row 1');

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

  /** The label of the toolbar control that has the focus, or null when the focus is elsewhere. */
  const focusedControl = (page: Page) =>
    page.evaluate(() => document.querySelector('#tb')?.shadowRoot?.activeElement?.getAttribute('aria-label') ?? null);

  test('the focus stays on the control after it acts', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.clickCell(1, 1);

    for (const label of ['Bold', 'Italic', 'Increase decimals', 'Sort descending', 'Clear formatting', 'Undo', 'Undo', 'Redo']) {
      await control(page, label).focus();
      await page.keyboard.press('Enter');
      await page.waitForTimeout(80);
      expect(await focusedControl(page), label).toBe(label);
    }

    // Menus return to their button
    await control(page, 'Insert rows or columns').focus();
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(80);
    expect(await focusedControl(page)).toBe('Insert rows or columns');
    expect(await sheetEval(page, (sheet) => sheet.rows)).toBe(51);

    // Two toggles in a row from the keyboard
    await spreadsheet.clickCell(0, 0);
    await control(page, 'Bold').focus();
    await page.keyboard.press('Enter');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Enter');
    expect(await spreadsheet.getCellFormat('0:0')).toEqual({ bold: true, italic: true });
    await expect(control(page, 'Italic')).toBeFocused();
    await expect(control(page, 'Italic')).toHaveAttribute('aria-pressed', 'true');
  });

  test('a control that gets disabled hands the focus to its neighbor', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.clickCell(1, 0);
    await page.keyboard.type('Zed');
    await page.keyboard.press('Enter');

    await control(page, 'Undo').focus();
    await page.keyboard.press('Enter');
    await expect(control(page, 'Undo')).toBeDisabled();
    await expect(control(page, 'Redo')).toBeFocused();

    await control(page, 'Italic').focus();
    await spreadsheet.setReadOnly(true);
    expect(await focusedControl(page)).not.toBeNull();
    await page.keyboard.press('End');
    await expect(control(page, 'Name manager')).toBeFocused();
  });

  test('dialogs opened from the toolbar return the focus to it', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.clickCell(1, 0);

    for (const label of ['Name manager', 'Paste Special']) {
      await control(page, label).focus();
      await page.keyboard.press('Enter');
      await page.waitForTimeout(150);
      expect(await focusedControl(page), `${label} takes the focus`).toBeNull();
      await page.keyboard.press('Escape');
      await expect(control(page, label)).toBeFocused();
    }

    await control(page, 'Conditional formatting').focus();
    await page.keyboard.press('ArrowUp'); // the last item: Manage rules
    await page.keyboard.press('Enter');
    await expect(spreadsheet.shadow('y11n-conditional-format-dialog')).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(control(page, 'Conditional formatting')).toBeFocused();

    // Opened from the grid, they return to the cell as before
    await spreadsheet.clickCell(1, 0);
    await page.keyboard.press('Control+F3');
    await expect(spreadsheet.shadow('y11n-name-manager')).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(spreadsheet.cell(1, 0)).toBeFocused();
  });

  test('a color choice is one undo step', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.clickCell(1, 0);
    await page.evaluate(() => {
      const input = document.querySelector('#tb')!.shadowRoot!.querySelector<HTMLInputElement>('[aria-label="Text color"]')!;
      // A drag in the picker: many input events, then one change
      for (const color of ['#110000', '#220000', '#330000']) {
        input.value = color;
        input.dispatchEvent(new Event('input', { bubbles: true }));
      }
      input.dispatchEvent(new Event('change', { bubbles: true }));
    });
    expect(await spreadsheet.getCellFormat('1:0')).toEqual({ textColor: '#330000' });
    expect(await sheetEval(page, (sheet) => [sheet.undo(), sheet.undo()])).toEqual([true, false]);
  });

  test('freeze refuses what would not fit, like the keyboard shortcut', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await sheetEval(page, (sheet) => sheet.goTo('C40'));
    await control(page, 'Freeze panes').click();
    await menuItem(page, 'Freeze up to the active cell').click();
    expect(await sheetEval(page, (sheet) => [sheet.frozenRows, sheet.frozenCols])).toEqual([0, 0]);
    await expect(spreadsheet.shadow('[role="log"]')).toHaveText('Cannot freeze: the frozen area would not fit on screen');
  });

  test('reapply filter takes in new rows, as Ctrl+Alt+L does', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await sheetEval(page, (sheet) =>
      sheet.setFilter({ start: { row: 0, col: 0 }, end: { row: 3, col: 1 } }, [{ col: 0, values: ['Ann'] }])
    );
    await spreadsheet.clickCell(4, 0);
    await page.keyboard.type('Dee');
    await page.keyboard.press('Enter');
    expect(await sheetEval(page, (sheet) => sheet.isRowHidden(4))).toBe(false);

    await control(page, 'Reapply filter').click();
    expect(await sheetEval(page, (sheet) => sheet.isRowHidden(4))).toBe(true);
    await expect(spreadsheet.shadow('[role="log"]')).toContainText('1 of 4');
  });

  test('follows data that is loaded later and a grid that is replaced', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.clickCell(0, 0);
    await expect(control(page, 'Bold')).toHaveAttribute('aria-pressed', 'false');

    await page.evaluate(() => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      const data = sheet.getData();
      data.get('0:0').format = { bold: true };
      sheet.setData(data);
    });
    await expect(control(page, 'Bold')).toHaveAttribute('aria-pressed', 'true');

    await page.evaluate(() => {
      const old = document.querySelector('y11n-spreadsheet')!;
      const fresh = document.createElement('y11n-spreadsheet') as any;
      fresh.id = old.id;
      fresh.setData(new Map([['0:0', { rawValue: 'x', displayValue: 'x', type: 'text', format: { italic: true } }]]));
      old.replaceWith(fresh);
    });
    await expect(control(page, 'Italic')).toHaveAttribute('aria-pressed', 'true');
    await expect(control(page, 'Bold')).toHaveAttribute('aria-pressed', 'false');
  });

  test('a toolbar that comes before its grid finds it', async ({ page }) => {
    await page.goto('/');
    await page.evaluate(() => {
      document.querySelector('y11n-spreadsheet')!.remove();
      const toolbar = document.createElement('y11n-toolbar');
      toolbar.id = 'tb';
      toolbar.setAttribute('for', 'late');
      document.body.append(toolbar);
    });
    await page.waitForTimeout(200);
    await page.evaluate(() => {
      const sheet = document.createElement('y11n-spreadsheet') as any;
      sheet.id = 'late';
      sheet.readOnly = true;
      sheet.style.height = '300px';
      document.body.append(sheet);
    });
    await expect(control(page, 'Bold')).toBeDisabled();
    await expect(control(page, 'Undo')).toBeDisabled();
  });

  test('a mouse click on a control leaves the keyboard in the grid', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.clickCell(8, 1);

    // Click Bold, go on typing
    await control(page, 'Bold').click();
    expect(await focusedControl(page)).toBeNull();
    await expect(spreadsheet.cell(8, 1)).toBeFocused();
    await page.keyboard.type('abc');
    await page.keyboard.press('Enter');
    await spreadsheet.waitForCellText(8, 1, 'abc');
    expect(await spreadsheet.getCellFormat('8:1')).toEqual({ bold: true });

    // A select needs the focus to open; it hands it back once a value is chosen
    await spreadsheet.clickCell(1, 1);
    await control(page, 'Number format').selectOption('percent');
    await expect(spreadsheet.cell(1, 1)).toBeFocused();
    await page.keyboard.press('ArrowDown');
    expect(await sheetEval(page, (sheet) => sheet.getSelection().activeCell.row)).toBe(2);

    // Also when the focus was in the toolbar before the click
    await control(page, 'Italic').focus();
    await control(page, 'Underline').click();
    await expect(spreadsheet.cell(2, 1)).toBeFocused();
    expect(await spreadsheet.getCellFormat('2:1')).toEqual({ underline: true });

    // Dialogs opened with the mouse go back to the cell
    await control(page, 'Name manager').click();
    await expect(spreadsheet.shadow('y11n-name-manager')).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(spreadsheet.cell(2, 1)).toBeFocused();
  });

  test('a menu opened with the mouse does not take the focus until a key is pressed', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.clickCell(1, 0);

    await control(page, 'Insert rows or columns').click();
    await expect(menuItem(page, 'Insert row above')).toBeVisible();
    await expect(spreadsheet.cell(1, 0)).toBeFocused();
    await menuItem(page, 'Insert row above').click();
    await expect(spreadsheet.cell(1, 0)).toBeFocused();
    await spreadsheet.waitForCellText(2, 0, 'Cy');

    await control(page, 'Insert rows or columns').click();
    await page.keyboard.press('Escape');
    await expect(page.locator('#tb').locator('[role="menu"]')).toHaveCount(0);
    await expect(spreadsheet.cell(1, 0)).toBeFocused();

    await control(page, 'Insert rows or columns').click();
    await page.keyboard.press('ArrowDown');
    await expect(menuItem(page, 'Insert row above')).toBeFocused();
    await page.keyboard.press('ArrowDown');
    await expect(menuItem(page, 'Insert row below')).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(control(page, 'Insert rows or columns')).toBeFocused();
  });

  test('a click during an edit commits the edit and then acts', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.dblClickCell(1, 0);
    await spreadsheet.typeInEditor('Zed');
    await control(page, 'Bold').click();
    await spreadsheet.waitForCellText(1, 0, 'Zed');
    expect(await spreadsheet.getCellFormat('1:0')).toEqual({ bold: true });
    expect(await spreadsheet.isEditorVisible()).toBe(false);
    await expect(spreadsheet.cell(1, 0)).toBeFocused();
  });

  test('with a formula open on another sheet, a control ends the edit and does nothing else', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await sheetEval(page, (sheet) => {
      sheet.addSheet('Second', { data: new Map([['0:0', { rawValue: 'other', displayValue: 'other', type: 'text' }]]) });
    });
    const tabs = page.locator('y11n-sheet-tabs');
    const state = () =>
      sheetEval(page, (sheet) => ({
        rows: sheet.rows,
        format: sheet.getCellFormat('8:0') ?? null,
        filter: sheet.getFilter() !== null,
        rules: sheet.getConditionalFormats().length,
        shown: sheet.activeSheet,
      }));

    const actions: Array<() => Promise<void>> = [
      () => control(page, 'Bold').click(),
      () => control(page, 'Filter').click(),
      async () => {
        await control(page, 'Delete rows or columns').click();
        await menuItem(page, 'Delete row').click();
      },
      async () => {
        await control(page, 'Conditional formatting').click();
        await menuItem(page, 'Data bars').click();
      },
    ];
    for (const act of actions) {
      await spreadsheet.clickCell(8, 0);
      await page.keyboard.type('=');
      await tabs.locator('[role="tab"]', { hasText: 'Second' }).click();
      expect((await state()).shown).toBe('Second');
      await act();
      // The edit ended and its sheet is shown again; nothing else happened
      expect(await state()).toEqual({ rows: 50, format: null, filter: false, rules: 0, shown: 'Sheet1' });
      expect(await spreadsheet.isEditorVisible()).toBe(false);
      expect(await sheetEval(page, (sheet) => sheet.getSheetData('Second').size)).toBe(1);
    }
  });

  test('an edit ended by going to a control outside the grid leaves the focus there', async ({ spreadsheet, page }) => {
    await spreadsheet.setData(DATA);
    await page.evaluate(() => {
      const input = document.createElement('input');
      input.id = 'outside';
      document.body.prepend(input);
    });
    await spreadsheet.dblClickCell(1, 0);
    await spreadsheet.typeInEditor('Zed');
    await page.locator('#outside').click();
    await spreadsheet.waitForCellText(1, 0, 'Zed');
    await page.waitForTimeout(100);
    await expect(page.locator('#outside')).toBeFocused();
  });

  test('disabled while in use: a menu item hands on the focus, Redo hands it to Undo', async ({ spreadsheet, page }) => {
    await addToolbar(page);
    await spreadsheet.setData(DATA);
    await spreadsheet.clickCell(1, 0);
    await page.keyboard.type('Zed');
    await page.keyboard.press('Enter');

    await control(page, 'Undo').focus();
    await page.keyboard.press('Enter');
    await expect(control(page, 'Redo')).toBeFocused();
    await page.keyboard.press('Enter');
    await expect(control(page, 'Redo')).toBeDisabled();
    await expect(control(page, 'Undo')).toBeFocused();

    await control(page, 'Insert rows or columns').focus();
    await page.keyboard.press('ArrowDown');
    await expect(menuItem(page, 'Insert row above')).toBeFocused();
    await spreadsheet.setReadOnly(true);
    await expect(page.locator('#tb').locator('[role="menu"]')).toHaveCount(0);
    expect(await focusedControl(page)).toBe('Fit column width');
  });

  test('for naming something that is not a grid only fires the events', async ({ spreadsheet, page }) => {
    await spreadsheet.setData(DATA);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.evaluate(() => {
      const div = document.createElement('div');
      div.id = 'notgrid';
      const toolbar = document.createElement('y11n-toolbar');
      toolbar.id = 'tb';
      toolbar.setAttribute('for', 'notgrid');
      toolbar.addEventListener('format-action', () => ((window as any).fired = true));
      document.body.prepend(div, toolbar);
    });
    await control(page, 'Bold').click();
    await control(page, 'Undo').click();
    await page.waitForTimeout(100);
    expect(await page.evaluate(() => (window as any).fired)).toBe(true);
    expect(errors).toEqual([]);
  });
});
