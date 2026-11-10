import { test, expect } from './fixtures';
import type { Page } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';

const log = (page: Page) => page.locator('y11n-spreadsheet').locator('[role="log"]');
const tabs = (page: Page) => page.locator('y11n-sheet-tabs');
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

test.describe('Dynamic arrays', () => {
  test('a typed array formula spills, and is announced', async ({ spreadsheet, page }) => {
    await seed(page, [['1'], ['2'], ['3']]);
    await spreadsheet.clickCell(0, 2);
    await page.keyboard.type('=A1:A3*2');
    await page.keyboard.press('Enter');

    await expect(spreadsheet.cell(0, 2).locator('.cell-text')).toHaveText('2');
    await expect(spreadsheet.cell(1, 2).locator('.cell-text')).toHaveText('4');
    await expect(spreadsheet.cell(2, 2).locator('.cell-text')).toHaveText('6');

    // The selection moved down onto a spilled value
    await expect(spreadsheet.liveRegion).toHaveText('C2: 4, spilled from C1');
    await expect(spreadsheet.cell(2, 2)).toHaveClass(/in-spill/);
    await expect(page.locator('y11n-formula-bar').locator('input.formula-input')).toHaveAttribute('placeholder', '=A1:A3*2');

    await page.keyboard.press('ArrowUp');
    await expect(spreadsheet.liveRegion).toHaveText('C1: 2, spills to C3');
  });

  test('typing into the spill range blocks it until the value is cleared', async ({ spreadsheet, page }) => {
    await seed(page, [['1', '', '=A1:A3*2'], ['2'], ['3']]);
    await spreadsheet.clickCell(1, 2);
    await page.keyboard.type('x');
    await page.keyboard.press('Enter');
    await expect(spreadsheet.cell(0, 2).locator('.cell-text')).toHaveText('#SPILL!');
    await expect(spreadsheet.cell(2, 2).locator('.cell-text')).toHaveText('');

    await spreadsheet.clickCell(0, 2);
    await expect(spreadsheet.liveRegion).toContainText('the cells the result would fill are not empty');

    await spreadsheet.clickCell(1, 2);
    await page.keyboard.press('Delete');
    await expect(spreadsheet.cell(0, 2).locator('.cell-text')).toHaveText('2');
    await expect(spreadsheet.cell(1, 2).locator('.cell-text')).toHaveText('4');
  });

  test('array functions and whole-column references', async ({ spreadsheet, page }) => {
    await seed(page, [
      ['Fruit', 'Qty', '', '=SORT(FILTER(A2:B5,B2:B5>2),2,-1)', '', '', '=SUM(B:B)'],
      ['Apple', '3'],
      ['Pear', '5'],
      ['Plum', '2'],
      ['Kiwi', '7'],
    ]);
    await expect(spreadsheet.cell(0, 3).locator('.cell-text')).toHaveText('Kiwi');
    await expect(spreadsheet.cell(0, 4).locator('.cell-text')).toHaveText('7');
    await expect(spreadsheet.cell(2, 3).locator('.cell-text')).toHaveText('Apple');
    await expect(spreadsheet.cell(0, 6).locator('.cell-text')).toHaveText('17');

    // Editing an input updates the spilled result and the column total
    await spreadsheet.clickCell(3, 1);
    await page.keyboard.type('9');
    await page.keyboard.press('Enter');
    await expect(spreadsheet.cell(0, 3).locator('.cell-text')).toHaveText('Plum');
    await expect(spreadsheet.cell(0, 6).locator('.cell-text')).toHaveText('24');
  });
});

test.describe('Lists and arrays in cells', () => {
  test('Ctrl+J makes a list, Alt+Down shows its values', async ({ spreadsheet, page }) => {
    await seed(page, [['Ann, Bob, Cy']]);
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.press('Control+j');
    await expect(log(page)).toHaveText('Turned 1 cell into a list');
    await expect(spreadsheet.liveRegion).toHaveText('A1: Ann, Bob, Cy, list of 3 items. Alt+Down to show the values');
    await expect(spreadsheet.cell(0, 0).locator('.ls-array-button')).toBeVisible();
    expect(await sheet(page, (el) => el.getData().get('0:0').rawValue)).toBe('={{"Ann";"Bob";"Cy"}}');

    await page.keyboard.press('Alt+ArrowDown');
    const card = spreadsheet.shadow('.ls-array-card');
    await expect(card).toBeVisible();
    await expect(card.locator('.ls-array-card-body')).toBeFocused();
    await expect(card.locator('li')).toHaveText(['Ann', 'Bob', 'Cy']);
    await expect(spreadsheet.cell(0, 0)).toHaveAttribute('aria-expanded', 'true');

    await page.keyboard.press('Escape');
    await expect(card).toHaveCount(0);
    await expect(spreadsheet.cell(0, 0)).toBeFocused();

    // Editing shows the items, and what comes back is still a list
    await page.keyboard.press('F2');
    await expect(spreadsheet.editor).toHaveValue('Ann, Bob, Cy');
    await spreadsheet.typeInEditor('Ann, Di');
    await spreadsheet.commitWithEnter();
    expect(await sheet(page, (el) => el.getData().get('0:0').rawValue)).toBe('={{"Ann";"Di"}}');

    // Undo, then back to text
    await page.keyboard.press('Control+z');
    await expect(spreadsheet.cell(0, 0).locator('.cell-text')).toHaveText('Ann, Bob, Cy');
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.press('Control+j');
    await expect(log(page)).toHaveText('Turned 1 list into text');
    await expect(spreadsheet.cell(0, 0).locator('.ls-array-button')).toHaveCount(0);
  });

  test('typing a list with Ctrl+J in the editor, and formulas over the items', async ({ spreadsheet, page }) => {
    await seed(page, [['']]);
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.press('Control+j');
    await expect(spreadsheet.editor).toBeVisible();
    await page.keyboard.type('4, 5, 6');
    await page.keyboard.press('Enter');
    expect(await sheet(page, (el) => el.getData().get('0:0').rawValue)).toBe('={{4;5;6}}');

    await spreadsheet.clickCell(0, 1);
    await page.keyboard.type('=SUM(A1)');
    await page.keyboard.press('Tab');
    await page.keyboard.type('=HAS(A1,5)');
    await page.keyboard.press('Tab');
    await page.keyboard.type('={FILTER(A1,A1>4)}');
    await page.keyboard.press('Enter');
    await expect(spreadsheet.cell(0, 1).locator('.cell-text')).toHaveText('15');
    await expect(spreadsheet.cell(0, 2).locator('.cell-text')).toHaveText('TRUE');
    await expect(spreadsheet.cell(0, 3).locator('.cell-text')).toHaveText('5, 6');
    // The braces keep the result in one cell
    await expect(spreadsheet.cell(1, 3).locator('.cell-text')).toHaveText('');
  });

  test('the icon opens the values and nested arrays can be opened', async ({ spreadsheet, page }) => {
    await seed(page, [['={{{1,2};{3,4}}}']]);
    await spreadsheet.cell(0, 0).locator('.ls-array-button').click();
    const card = spreadsheet.shadow('.ls-array-card');
    await expect(card.locator('.ls-array-nested')).toHaveCount(2);
    await card.locator('.ls-array-nested').nth(1).click();
    await expect(card.locator('li')).toHaveText(['3', '4']);
    await card.locator('.ls-array-back').click();
    await expect(card.locator('.ls-array-nested')).toHaveCount(2);
    // A click elsewhere closes it
    await spreadsheet.clickCell(4, 4);
    await expect(card).toHaveCount(0);
  });

  test('the filter menu lists the items of lists', async ({ spreadsheet, page }) => {
    await seed(page, [['Owners'], ['={{"Ann";"Bob"}}'], ['={{"Bob"}}'], ['Cy']]);
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.press('Control+Shift+L');
    await page.keyboard.press('Alt+ArrowDown');
    const menu = spreadsheet.shadow('.ls-filter-menu');
    await expect(menu).toBeVisible();
    await expect(menu).toContainText('Ann');
    await expect(menu).not.toContainText('Ann, Bob');
    await page.keyboard.press('Escape');

    await sheet(page, (el) => el.setColumnFilter(0, ['Bob']));
    expect(await sheet(page, (el) => el.getFilter().hiddenRows)).toEqual([3]);
  });

  test('has no axe violations with a list, a spill and the values popup', async ({ spreadsheet, page }) => {
    await seed(page, [['={{"Ann";"Bob"}}', '=SEQUENCE(3)']]);
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.press('Alt+ArrowDown');
    await expect(spreadsheet.shadow('.ls-array-card')).toBeVisible();
    const results = await new AxeBuilder({ page }).include('y11n-spreadsheet').analyze();
    expect(results.violations).toEqual([]);
  });
});

test.describe('Sheets and names', () => {
  test('sheets are added from the API and switched with the tabs', async ({ spreadsheet, page }) => {
    await expect(tabs(page)).toHaveCount(0);
    await seed(page, [['=Data!A1*2', '=SUM(Data!A:A)']]);
    await expect(spreadsheet.cell(0, 0).locator('.cell-text')).toHaveText('#REF!');

    await sheet(page, (el) => {
      const data = new Map([
        ['0:0', { rawValue: '5', displayValue: '5', type: 'number' }],
        ['1:0', { rawValue: '6', displayValue: '6', type: 'number' }],
      ]);
      el.addSheet('Data', { data });
    });
    await expect(spreadsheet.cell(0, 0).locator('.cell-text')).toHaveText('10');
    await expect(spreadsheet.cell(0, 1).locator('.cell-text')).toHaveText('11');

    await expect(tabs(page).locator('[role="tab"]')).toHaveText(['Sheet1', 'Data']);
    await tabs(page).locator('[role="tab"]', { hasText: 'Data' }).click();
    await expect(tabs(page).locator('[role="tab"][aria-selected="true"]')).toHaveText('Data');
    await expect(spreadsheet.cell(0, 0).locator('.cell-text')).toHaveText('5');

    // An edit here reaches the formulas on the other sheet
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.type('50');
    await page.keyboard.press('Enter');
    await page.keyboard.press('Alt+PageUp');
    await expect(log(page)).toHaveText('Sheet1, sheet 1 of 2');
    await expect(spreadsheet.cell(0, 0).locator('.cell-text')).toHaveText('100');
    await expect(spreadsheet.cell(0, 1).locator('.cell-text')).toHaveText('56');

    // Undo goes back to the sheet the edit was made on
    await page.keyboard.press('Control+z');
    await expect(tabs(page).locator('[role="tab"][aria-selected="true"]')).toHaveText('Data');
    await expect(spreadsheet.cell(0, 0).locator('.cell-text')).toHaveText('5');
  });

  test('keyboard: add, rename and delete a sheet from the tabs', async ({ spreadsheet, page }) => {
    await sheet(page, (el) => {
      el.sheetTabs = true;
    });
    await seed(page, [['=Sheet2!A1']]);
    await tabs(page).getByRole('button', { name: 'Add sheet' }).click();
    await expect(log(page)).toHaveText('Added Sheet2');
    await expect(tabs(page).locator('[role="tab"][aria-selected="true"]')).toHaveText('Sheet2');
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.type('7');
    await page.keyboard.press('Enter');

    // Rename with F2 on the tab
    await tabs(page).locator('[role="tab"][aria-selected="true"]').focus();
    await page.keyboard.press('F2');
    await page.keyboard.type('Costs 2026');
    await page.keyboard.press('Enter');
    await expect(tabs(page).locator('[role="tab"]')).toHaveText(['Sheet1', 'Costs 2026']);
    await expect(log(page)).toHaveText('Renamed Sheet2 to Costs 2026');
    expect(await sheet(page, (el) => el.getSheetData('Sheet1').get('0:0'))).toMatchObject({
      rawValue: "='Costs 2026'!A1",
      displayValue: '7',
    });

    // Arrow keys move between the tabs and show the sheet
    await page.keyboard.press('ArrowLeft');
    expect(await sheet(page, (el) => el.activeSheet)).toBe('Sheet1');
    await page.keyboard.press('ArrowRight');
    expect(await sheet(page, (el) => el.activeSheet)).toBe('Costs 2026');

    // Delete asks first
    await page.keyboard.press('Delete');
    const confirm = tabs(page).locator('[role="alertdialog"]');
    await expect(confirm).toBeVisible();
    await expect(confirm.getByRole('button', { name: 'Cancel' })).toBeFocused();
    await confirm.getByRole('button', { name: 'Delete' }).click();
    await expect(log(page)).toHaveText('Deleted Costs 2026');
    await expect(tabs(page).locator('[role="tab"]')).toHaveText(['Sheet1']);
    await expect(spreadsheet.cell(0, 0).locator('.cell-text')).toHaveText('#REF!');
  });

  test('before-sheets-change can veto a user change', async ({ spreadsheet, page }) => {
    await expect(spreadsheet.grid).toBeVisible();
    await sheet(page, (el) => {
      el.sheetTabs = true;
      el.addEventListener('before-sheets-change', (e: Event) => e.preventDefault());
    });
    await tabs(page).getByRole('button', { name: 'Add sheet' }).click();
    expect(await sheet(page, (el) => el.sheets)).toEqual(['Sheet1']);
  });

  test('defined names and LET / LAMBDA', async ({ spreadsheet, page }) => {
    await seed(page, [['10', '=SUM(sales)*(1+tax)'], ['20', '=LET(total,SUM(sales),total/2)'], ['', '=double(A1)']]);
    await expect(spreadsheet.cell(0, 1).locator('.cell-text')).toHaveText('#NAME?');
    await sheet(page, (el) => {
      el.defineName('sales', '=A1:A2');
      el.defineName('tax', '=0.5');
      el.defineName('double', '=LAMBDA(x,x*2)');
    });
    await expect(spreadsheet.cell(0, 1).locator('.cell-text')).toHaveText('45');
    await expect(spreadsheet.cell(1, 1).locator('.cell-text')).toHaveText('15');
    await expect(spreadsheet.cell(2, 1).locator('.cell-text')).toHaveText('20');

    // Names move with inserted rows
    await sheet(page, (el) => el.insertRows(0, 1));
    expect(await sheet(page, (el) => el.getNames()[0].refersTo)).toBe('=A2:A3');
    await expect(spreadsheet.cell(1, 1).locator('.cell-text')).toHaveText('45');
  });

  test('has no axe violations with sheet tabs', async ({ spreadsheet, page }) => {
    await expect(spreadsheet.grid).toBeVisible();
    await sheet(page, (el) => {
      el.addSheet('Second');
    });
    await expect(tabs(page).locator('[role="tab"]')).toHaveCount(2);
    const results = await new AxeBuilder({ page }).include('y11n-spreadsheet').analyze();
    expect(results.violations).toEqual([]);
  });
});
