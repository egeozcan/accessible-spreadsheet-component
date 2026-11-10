import { test, expect } from './fixtures';
import type { Page } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';

const log = (page: Page) => page.locator('y11n-spreadsheet').locator('[role="log"]');
const dialog = (page: Page) => page.locator('y11n-conditional-format-dialog');
const rules = (page: Page) =>
  page.evaluate(() => (document.querySelector('y11n-spreadsheet') as any).getConditionalFormats());

/**
 *      A        B       C
 *  1   Region   Sales   Growth
 *  2   North    120     12
 *  3   South    80      -5
 *  4   East     200     30
 *  5   West     45      -20
 *  6   Central  150     8
 */
async function seedTable(page: Page) {
  await page.evaluate(() => {
    const sheet = document.querySelector('y11n-spreadsheet') as any;
    const rows = [
      ['Region', 'Sales', 'Growth'],
      ['North', '120', '12'],
      ['South', '80', '-5'],
      ['East', '200', '30'],
      ['West', '45', '-20'],
      ['Central', '150', '8'],
    ];
    const data = new Map();
    rows.forEach((row, r) =>
      row.forEach((value, c) =>
        data.set(`${r}:${c}`, { rawValue: value, displayValue: value, type: /^-?\d+$/.test(value) ? 'number' : 'text' })
      )
    );
    sheet.setData(data);
  });
}

const SALES = { start: { row: 1, col: 1 }, end: { row: 5, col: 1 } };
const GROWTH = { start: { row: 1, col: 2 }, end: { row: 5, col: 2 } };

async function addRule(page: Page, rule: unknown) {
  await page.evaluate((r) => (document.querySelector('y11n-spreadsheet') as any).addConditionalFormat(r), rule);
}

test.describe('Conditional formatting', () => {
  test('keyboard: create a rule in the rule manager and hear it on the cells', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await spreadsheet.clickCell(1, 1);
    for (let i = 0; i < 4; i++) await page.keyboard.press('Shift+ArrowDown');
    await page.keyboard.press('Alt+Shift+C');

    await expect(dialog(page).locator('[role="dialog"]')).toBeVisible();
    await expect(dialog(page).locator('#cf-new')).toBeFocused();
    await page.keyboard.press('Enter');

    // The form starts on the range, which is the selection
    await expect(dialog(page).locator('#cf-range')).toBeFocused();
    await expect(dialog(page).locator('#cf-range')).toHaveValue('B2:B6');
    await page.keyboard.press('Tab');
    await expect(dialog(page).locator('#cf-kind')).toBeFocused();
    await page.keyboard.press('Tab');
    await page.keyboard.press('Tab');
    await expect(dialog(page).locator('#cf-value')).toBeFocused();
    await page.keyboard.type('100');
    await page.keyboard.press('Enter');

    await expect(dialog(page).locator('[role="status"]')).toHaveText('Rule added: greater than 100, B2:B6');
    await expect(dialog(page).locator('#cf-new')).toBeFocused();
    expect(await rules(page)).toMatchObject([
      { range: SALES, condition: { type: 'cellValue', operator: 'greaterThan', value: 100 } },
    ]);

    // Escape returns to the grid, where the cell says why it is colored
    await page.keyboard.press('Escape');
    await expect(dialog(page)).toHaveCount(0);
    await expect(spreadsheet.cell(5, 1)).toBeFocused();
    await expect(spreadsheet.liveRegion).toContainText('B6: 150, greater than 100');
    await expect(spreadsheet.cell(5, 1)).toHaveCSS('background-color', 'rgb(255, 199, 206)');
    await expect(spreadsheet.cell(2, 1)).not.toHaveCSS('background-color', 'rgb(255, 199, 206)');
    // The label is part of the cell's accessible name
    await expect(spreadsheet.cell(5, 1)).toHaveAccessibleName('150 , greater than 100');

    // Undo removes the rule, redo brings it back
    await page.keyboard.press('Control+z');
    expect(await rules(page)).toEqual([]);
    await page.keyboard.press('Control+y');
    expect(await rules(page)).toHaveLength(1);
  });

  test('restyles as values are typed', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await addRule(page, { range: SALES, condition: { type: 'top', rank: 1 }, format: { backgroundColor: '#c6efce' } });
    await expect(spreadsheet.cell(3, 1)).toHaveClass(/cf-match/);

    await spreadsheet.clickCell(4, 1);
    await page.keyboard.type('999');
    await page.keyboard.press('Enter');
    await expect(spreadsheet.cell(4, 1)).toHaveClass(/cf-match/);
    await expect(spreadsheet.cell(3, 1)).not.toHaveClass(/cf-match/);
  });

  test('data bars, icons and color scales are drawn inside their cells', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await addRule(page, { range: GROWTH, condition: { type: 'dataBar' } });
    await addRule(page, { range: SALES, condition: { type: 'iconSet', icons: 'arrows3' } });

    const cellBox = (await spreadsheet.cell(3, 2).boundingBox())!;
    const bar = (await spreadsheet.cell(3, 2).locator('.ls-cf-bar').boundingBox())!;
    // The largest value's bar runs from the zero axis (40% in: -20 to 30) to the right edge
    expect(bar.x - cellBox.x).toBeCloseTo(cellBox.width * 0.4, 0);
    expect(bar.x + bar.width).toBeCloseTo(cellBox.x + cellBox.width - 1, 0);
    expect(bar.y).toBeGreaterThan(cellBox.y);
    // A negative value's bar ends at the axis
    const negative = (await spreadsheet.cell(4, 2).locator('.ls-cf-bar').boundingBox())!;
    expect(negative.x).toBeCloseTo(cellBox.x, 0);
    expect(negative.x + negative.width - cellBox.x).toBeCloseTo(cellBox.width * 0.4, 0);
    // The value stays readable over the bar
    await expect(spreadsheet.cell(3, 2).locator('.cell-text')).toBeVisible();

    await expect(spreadsheet.cell(3, 1).locator('.ls-cf-icon svg')).toBeVisible();
    await expect(spreadsheet.cell(3, 1)).toHaveAccessibleName('200 , up arrow');
    await expect(spreadsheet.cell(4, 1)).toHaveAccessibleName('45 , down arrow');
  });

  test('selection stays visible on cells a rule fills', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await addRule(page, { range: SALES, condition: { type: 'notBlanks' }, format: { backgroundColor: '#ffc7ce' } });
    await spreadsheet.clickCell(1, 1);
    await page.keyboard.press('Shift+ArrowDown');
    const shadowOf = (row: number) =>
      spreadsheet.cell(row, 1).evaluate((el) => getComputedStyle(el).boxShadow);
    expect(await shadowOf(2)).toContain('inset');
    expect(await shadowOf(4)).toBe('none');
  });

  test('forced colors: highlights, scales, bars and icons keep a visible cue', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await addRule(page, { range: SALES, condition: { type: 'aboveAverage' }, format: { backgroundColor: '#c6efce' } });
    await addRule(page, {
      range: GROWTH,
      condition: { type: 'colorScale', min: { type: 'min', color: '#f8696b' }, max: { type: 'max', color: '#63be7b' } },
    });
    await addRule(page, {
      range: { start: { row: 1, col: 0 }, end: { row: 5, col: 0 } },
      condition: { type: 'formula', formula: '=$B2>100' },
      format: { bold: true },
    });
    await addRule(page, { range: SALES, condition: { type: 'iconSet', icons: 'shapes3' } });

    const marker = (row: number, col: number) =>
      spreadsheet.cell(row, col).evaluate((el) => getComputedStyle(el, '::before').content);
    const meter = spreadsheet.cell(3, 2).locator('.ls-cf-meter');

    // Without forced colors there is no marker or meter
    expect(await marker(3, 1)).toBe('none');
    await expect(meter).toBeHidden();

    await page.emulateMedia({ forcedColors: 'active' });
    // Highlighted cells get an edge marker; cells no rule matches don't
    expect(await marker(3, 1)).not.toBe('none');
    expect(await marker(2, 1)).toBe('none');
    // A color scale becomes a meter line as long as the cell's position on the scale
    await expect(meter).toBeVisible();
    const cellBox = (await spreadsheet.cell(3, 2).boundingBox())!;
    expect((await meter.boundingBox())!.width).toBeCloseTo(cellBox.width - 1, 0);
    // Icons take the system text color instead of their own
    const canvasText = await page.evaluate(() => {
      const probe = document.createElement('span');
      probe.style.color = 'CanvasText';
      document.body.appendChild(probe);
      const color = getComputedStyle(probe).color;
      probe.remove();
      return color;
    });
    await expect(spreadsheet.cell(3, 1).locator('.ls-cf-icon svg')).toHaveCSS('color', canvasText);
  });

  test('rule manager: reorder and delete with the keyboard, focus stays on the rule', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await page.evaluate(
      ([sales, growth]) =>
        (document.querySelector('y11n-spreadsheet') as any).setConditionalFormats([
          { id: 'a', range: sales, condition: { type: 'aboveAverage' }, format: { backgroundColor: '#c6efce' } },
          { id: 'b', range: sales, condition: { type: 'duplicates' }, format: { backgroundColor: '#ffc7ce' } },
          { id: 'c', range: growth, condition: { type: 'dataBar' } },
        ]),
      [SALES, GROWTH]
    );
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.press('Alt+Shift+C');
    await expect(dialog(page).locator('.rule-desc')).toHaveText(['Above average', 'Duplicate', 'Data bar']);

    const button = (rule: string, action: string) => dialog(page).locator(`[data-rule="${rule}"][data-action="${action}"]`);
    await button('c', 'up').focus();
    await page.keyboard.press('Enter');
    await expect(dialog(page).locator('.rule-desc')).toHaveText(['Above average', 'Data bar', 'Duplicate']);
    await expect(button('c', 'up')).toBeFocused();
    await expect(dialog(page).locator('[role="status"]')).toHaveText('Rule moved to position 2 of 3');
    await page.keyboard.press('Enter');
    // At the top "Move up" no longer applies, so focus moves to "Move down"
    await expect(button('c', 'down')).toBeFocused();
    expect((await rules(page)).map((r: { id: string }) => r.id)).toEqual(['c', 'a', 'b']);

    await button('a', 'delete').focus();
    await page.keyboard.press('Enter');
    await expect(dialog(page).locator('.rule-desc')).toHaveText(['Data bar', 'Duplicate']);
    await expect(button('b', 'edit')).toBeFocused();

    // Tab wraps inside the modal dialog
    await dialog(page).locator('#cf-close').focus();
    await page.keyboard.press('Tab');
    await expect(button('c', 'edit')).toBeFocused();

    // Clicking outside dismisses it
    await spreadsheet.clickCell(8, 4);
    await expect(dialog(page)).toHaveCount(0);
  });

  test('toolbar menu: presets, clear and manage', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await page.evaluate(() => {
      const toolbar = document.createElement('y11n-format-toolbar');
      toolbar.id = 'fmt-toolbar';
      const container = document.querySelector('.demo-container')!;
      container.parentElement!.insertBefore(toolbar, container);
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      toolbar.addEventListener('conditional-format-action', (e) => {
        const { action } = (e as CustomEvent).detail;
        if (action === 'manage') sheet.openConditionalFormatManager();
        else if (action === 'clear') sheet.clearSelectionConditionalFormats();
        else sheet.applyConditionalFormatPreset(action);
      });
    });
    const toolbar = page.locator('#fmt-toolbar');
    const menuButton = toolbar.locator('[aria-label="Conditional formatting"]');

    await spreadsheet.clickCell(1, 1);
    for (let i = 0; i < 4; i++) await page.keyboard.press('Shift+ArrowDown');

    await menuButton.focus();
    await page.keyboard.press('ArrowDown');
    await expect(toolbar.locator('[role="menu"]')).toBeVisible();
    await expect(toolbar.locator('[role="menuitem"]').first()).toBeFocused();
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    await expect(toolbar.locator('[role="menuitem"]', { hasText: 'Highlight above average' })).toBeFocused();
    await page.keyboard.press('Enter');

    await expect(toolbar.locator('[role="menu"]')).toHaveCount(0);
    await expect(menuButton).toBeFocused();
    await expect(log(page)).toHaveText('Conditional format added to B2:B6: above average');
    await expect(spreadsheet.cell(3, 1)).toHaveCSS('background-color', 'rgb(198, 239, 206)');
    await expect(spreadsheet.cell(4, 1)).not.toHaveClass(/cf-match/);

    await menuButton.click();
    await toolbar.locator('[role="menuitem"]', { hasText: 'Manage rules' }).click();
    await expect(dialog(page).locator('.rule-desc')).toHaveText(['Above average']);
    await expect(dialog(page).locator('#cf-new')).toBeFocused();
    await page.keyboard.press('Escape');

    await menuButton.click();
    await toolbar.locator('[role="menuitem"]', { hasText: 'Clear rules from selection' }).click();
    await expect(log(page)).toHaveText('Conditional formats cleared from the selection: 1 rule removed');
    expect(await rules(page)).toEqual([]);
  });

  test('has no axe violations with rules applied and with the rule manager open', async ({ spreadsheet, page }) => {
    await seedTable(page);
    await addRule(page, { range: SALES, condition: { type: 'aboveAverage' }, format: { backgroundColor: '#c6efce', textColor: '#006100' } });
    await addRule(page, { range: SALES, condition: { type: 'belowAverage' }, format: { backgroundColor: '#ffc7ce', textColor: '#9c0006' } });
    await addRule(page, { range: GROWTH, condition: { type: 'dataBar' } });
    await addRule(page, {
      range: { start: { row: 1, col: 0 }, end: { row: 5, col: 0 } },
      condition: { type: 'iconSet', icons: 'symbols3' },
    });
    await spreadsheet.clickCell(1, 1);

    // Contrast is checked here on purpose: the stock highlight styles must pass it
    const grid = await new AxeBuilder({ page }).include('y11n-spreadsheet').analyze();
    expect(grid.violations.map((v) => `${v.id}: ${v.help}`)).toEqual([]);

    await page.keyboard.press('Alt+Shift+C');
    await expect(dialog(page).locator('[role="dialog"]')).toBeVisible();
    const list = await new AxeBuilder({ page }).include('y11n-spreadsheet').analyze();
    expect(list.violations.map((v) => `${v.id}: ${v.help}`)).toEqual([]);

    await dialog(page).locator('#cf-new').click();
    await dialog(page).locator('#cf-style').selectOption('custom');
    const form = await new AxeBuilder({ page }).include('y11n-spreadsheet').analyze();
    expect(form.violations.map((v) => `${v.id}: ${v.help}`)).toEqual([]);

    await dialog(page).locator('#cf-kind').selectOption('iconSet');
    const icons = await new AxeBuilder({ page }).include('y11n-spreadsheet').analyze();
    expect(icons.violations.map((v) => `${v.id}: ${v.help}`)).toEqual([]);
  });
});
