import { test, expect } from './fixtures.js';

test.describe('Integration: cross-subsystem workflows', () => {

  test('copy formula cell and paste with reference adjustment', async ({ spreadsheet, page }) => {
    // Set up: A1=10, B1=20, C1==A1+B1
    await spreadsheet.setData({
      '0:0': { rawValue: '10', displayValue: '10', type: 'number' },
      '0:1': { rawValue: '20', displayValue: '20', type: 'number' },
      '0:2': { rawValue: '=A1+B1', displayValue: '30', type: 'number' },
    });

    // Select C1 and copy
    await spreadsheet.clickCell(0, 2);
    await page.keyboard.press('Control+c');

    // Move to C3 and paste
    await spreadsheet.clickCell(2, 2);
    await page.keyboard.press('Control+v');

    // Wait for paste to complete
    await spreadsheet.waitForCellText(2, 2, '0');

    // Verify the pasted formula references were adjusted
    const data = await spreadsheet.getData();
    expect(data['2:2']?.rawValue).toBe('=A3+B3');
  });

  test('edit cell, apply bold, undo twice, redo twice', async ({ spreadsheet, page }) => {
    // Type a value into A1
    await spreadsheet.dblClickCell(0, 0);
    await spreadsheet.typeInEditor('Hello');
    await spreadsheet.commitWithEnter();
    await spreadsheet.waitForCellText(0, 0, 'Hello');

    // Move back to A1 and apply bold
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.press('Control+b');

    // Verify bold is applied
    let format = await spreadsheet.getCellFormat('0:0');
    expect(format?.bold).toBe(true);

    // Undo (should remove bold)
    await page.keyboard.press('Control+z');
    format = await spreadsheet.getCellFormat('0:0');
    expect(format?.bold).toBeFalsy();

    // Undo again (should remove value)
    await page.keyboard.press('Control+z');
    await spreadsheet.waitForCellText(0, 0, '');

    // Redo (should restore value)
    await page.keyboard.press('Control+Shift+z');
    await spreadsheet.waitForCellText(0, 0, 'Hello');

    // Redo again (should restore bold)
    await page.keyboard.press('Control+Shift+z');
    format = await spreadsheet.getCellFormat('0:0');
    expect(format?.bold).toBe(true);
  });

  test('cut cells, paste elsewhere, undo restores both', async ({ spreadsheet, page }) => {
    // Set up: A1=100, A2=200
    await spreadsheet.setData({
      '0:0': { rawValue: '100', displayValue: '100', type: 'number' },
      '1:0': { rawValue: '200', displayValue: '200', type: 'number' },
    });

    // Select A1:A2
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.down('Shift');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.up('Shift');

    // Cut
    await page.keyboard.press('Control+x');

    // Move to C1 and paste
    await spreadsheet.clickCell(0, 2);
    await page.keyboard.press('Control+v');

    // Verify source cleared and target populated
    await spreadsheet.waitForCellText(0, 2, '100');
    await spreadsheet.waitForCellText(1, 2, '200');
    const textA1 = await spreadsheet.getCellText(0, 0);
    expect(textA1).toBe('');

    // Undo paste
    await page.keyboard.press('Control+z');
    // Undo cut
    await page.keyboard.press('Control+z');

    // Verify source restored
    await spreadsheet.waitForCellText(0, 0, '100');
    await spreadsheet.waitForCellText(1, 0, '200');
  });

  test('format cells, clear them, undo restores values with format', async ({ spreadsheet, page }) => {
    // Set up: A1=42
    await spreadsheet.setData({
      '0:0': { rawValue: '42', displayValue: '42', type: 'number' },
    });

    // Apply bold to A1
    await spreadsheet.clickCell(0, 0);
    await page.keyboard.press('Control+b');

    let format = await spreadsheet.getCellFormat('0:0');
    expect(format?.bold).toBe(true);

    // Clear A1
    await page.keyboard.press('Delete');
    await spreadsheet.waitForCellText(0, 0, '');

    // Undo clear (should restore value with bold)
    await page.keyboard.press('Control+z');
    await spreadsheet.waitForCellText(0, 0, '42');
    format = await spreadsheet.getCellFormat('0:0');
    expect(format?.bold).toBe(true);
  });
});
