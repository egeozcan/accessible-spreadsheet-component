import { test, expect } from './fixtures';

test.describe('API Methods', () => {
  test('getData() returns the current grid data', async ({ spreadsheet }) => {
    const data = await spreadsheet.getData();

    // The demo page sets some data
    expect(data['0:0']).toBeDefined();
    expect(data['0:0'].rawValue).toBe('Item');
  });

  test('setData() replaces the grid data', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'New A1', displayValue: 'New A1', type: 'text' },
      '0:1': { rawValue: '999', displayValue: '999', type: 'number' },
    });

    await spreadsheet.waitForCellText(0, 0, 'New A1');
    await spreadsheet.waitForCellText(0, 1, '999');
  });

  test('setData() clears previous data', async ({ spreadsheet }) => {
    // Initially has demo data in 1:0 (Widget A)
    await spreadsheet.waitForCellText(1, 0, 'Widget A');

    // Set new data that doesn't include 1:0
    await spreadsheet.setData({
      '0:0': { rawValue: 'Only This', displayValue: 'Only This', type: 'text' },
    });

    // Previous data should be gone
    const text = await spreadsheet.getCellText(1, 0);
    expect(text).toBe('');
  });

  test('getData() reflects edits made via UI', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(6, 0);
    await spreadsheet.cell(6, 0).press('Enter');
    await spreadsheet.typeInEditor('UI Edit');
    await spreadsheet.commitWithEnter();

    const data = await spreadsheet.getData();
    expect(data['6:0']).toBeDefined();
    expect(data['6:0'].rawValue).toBe('UI Edit');
  });

  test('registerFunction() makes custom function available', async ({ spreadsheet }) => {
    // Register a DOUBLE function
    await spreadsheet.page.evaluate(() => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      sheet.registerFunction('DOUBLE', (_ctx: any, val: any) => Number(val) * 2);
    });

    await spreadsheet.setData({
      '0:0': { rawValue: '21', displayValue: '21', type: 'number' },
      '0:1': { rawValue: '=DOUBLE(A1)', displayValue: '', type: 'text' },
    });

    await spreadsheet.waitForCellText(0, 1, '42');
  });

  test('setting rows/cols property changes grid size', async ({ spreadsheet }) => {
    await spreadsheet.setProperty('rows', 10);
    await spreadsheet.setProperty('cols', 5);

    // aria-rowcount should reflect the new size
    await expect(spreadsheet.grid).toHaveAttribute('aria-rowcount', '11'); // 10+1
    await expect(spreadsheet.grid).toHaveAttribute('aria-colcount', '6'); // 5+1

    // Only 5 column headers (+ corner)
    const headers = spreadsheet.columnHeaders;
    const count = await headers.count();
    expect(count).toBe(6); // corner + 5 columns
  });

  test('functions property registers formula functions', async ({ spreadsheet }) => {
    await spreadsheet.page.evaluate(() => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      sheet.functions = {
        TRIPLE: (_ctx: any, val: any) => Number(val) * 3,
      };
    });
    await spreadsheet.page.waitForTimeout(100);

    await spreadsheet.setData({
      '0:0': { rawValue: '10', displayValue: '10', type: 'number' },
      '0:1': { rawValue: '=TRIPLE(A1)', displayValue: '', type: 'text' },
    });

    await spreadsheet.waitForCellText(0, 1, '30');
  });

  test('setData triggers formula recalculation', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: '5', displayValue: '5', type: 'number' },
      '0:1': { rawValue: '10', displayValue: '10', type: 'number' },
      '0:2': { rawValue: '=A1+B1', displayValue: '', type: 'text' },
    });

    await spreadsheet.waitForCellText(0, 2, '15');
  });

  test('getData returns computed displayValue for formula cells', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: '7', displayValue: '7', type: 'number' },
      '0:1': { rawValue: '=A1*6', displayValue: '', type: 'text' },
    });

    await spreadsheet.page.waitForTimeout(200);

    const data = await spreadsheet.getData();
    expect(data['0:1'].rawValue).toBe('=A1*6');
    expect(data['0:1'].displayValue).toBe('42');
  });
});
