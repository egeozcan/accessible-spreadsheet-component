import { test, expect } from './fixtures';

test.describe('Cell Editing', () => {
  test('Enter key opens editor with existing cell value', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Enter');

    await expect(spreadsheet.editor).toBeFocused();
    // Cell A1 has "Item" in the demo
    await expect(spreadsheet.editor).toHaveValue('Item');
  });

  test('double-click opens editor', async ({ spreadsheet }) => {
    await spreadsheet.dblClickCell(1, 0);

    await expect(spreadsheet.editor).toBeFocused();
    await expect(spreadsheet.editor).toHaveValue('Widget A');
  });

  test('typing a character starts editing with that character', async ({ spreadsheet }) => {
    // Navigate to an empty cell
    await spreadsheet.clickCell(4, 0);
    await spreadsheet.cell(4, 0).press('h');

    await expect(spreadsheet.editor).toBeFocused();
    await expect(spreadsheet.editor).toHaveValue('h');
  });

  test('committing edit with Enter saves value and moves down', async ({ spreadsheet }) => {
    // Edit an empty cell
    await spreadsheet.clickCell(6, 0);
    await spreadsheet.cell(6, 0).press('Enter');
    await spreadsheet.typeInEditor('Hello World');
    await spreadsheet.commitWithEnter();

    // Value should be saved
    await spreadsheet.waitForCellText(6, 0, 'Hello World');

    // Active cell should move down
    await expect(spreadsheet.cell(7, 0)).toHaveClass(/active-cell/);
  });

  test('committing edit with Tab saves value and moves right', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(6, 1);
    await spreadsheet.cell(6, 1).press('Enter');
    await spreadsheet.typeInEditor('Tab Test');
    await spreadsheet.commitWithTab();

    await spreadsheet.waitForCellText(6, 1, 'Tab Test');

    // Active cell should move right
    await expect(spreadsheet.cell(6, 2)).toHaveClass(/active-cell/);
  });

  test('Shift+Tab during editing commits and moves left', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(6, 3);
    await spreadsheet.cell(6, 3).press('Enter');
    await spreadsheet.typeInEditor('Shift Tab');
    await spreadsheet.editor.press('Shift+Tab');

    await spreadsheet.waitForCellText(6, 3, 'Shift Tab');
    await expect(spreadsheet.cell(6, 2)).toHaveClass(/active-cell/);
  });

  test('Escape cancels edit without saving', async ({ spreadsheet }) => {
    const originalText = await spreadsheet.getCellText(0, 0);

    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Enter');
    await spreadsheet.typeInEditor('CHANGED');
    await spreadsheet.cancelEdit();

    // Value should remain unchanged
    const currentText = await spreadsheet.getCellText(0, 0);
    expect(currentText).toBe(originalText);
  });

  test('editor receives focus when editing starts', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Enter');
    await expect(spreadsheet.editor).toBeFocused();
  });

  test('focus returns to grid cell after commit', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(6, 4);
    await spreadsheet.cell(6, 4).press('Enter');
    await spreadsheet.typeInEditor('test');
    await spreadsheet.commitWithEnter();

    // Focus should move to the next cell (down)
    await expect(spreadsheet.cell(7, 4)).toBeFocused();
  });

  test('editing a cell with a formula shows the raw formula', async ({ spreadsheet }) => {
    // Cell D2 (row 1, col 3) has formula =B2*C2
    await spreadsheet.clickCell(1, 3);
    await spreadsheet.cell(1, 3).press('Enter');

    await expect(spreadsheet.editor).toHaveValue('=B2*C2');
  });

  test('number values are correctly stored and displayed', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(8, 0);
    await spreadsheet.cell(8, 0).press('Enter');
    await spreadsheet.typeInEditor('42');
    await spreadsheet.commitWithEnter();

    await spreadsheet.waitForCellText(8, 0, '42');
  });

  test('Delete key clears selected cells', async ({ spreadsheet }) => {
    // Put some data in a cell first
    await spreadsheet.clickCell(7, 0);
    await spreadsheet.cell(7, 0).press('Enter');
    await spreadsheet.typeInEditor('to delete');
    await spreadsheet.commitWithEnter();

    // Navigate back and delete
    await spreadsheet.clickCell(7, 0);
    await spreadsheet.cell(7, 0).press('Delete');

    // Cell should be cleared
    await spreadsheet.waitForCellText(7, 0, '');
  });

  test('Backspace key clears selected cells', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(8, 1);
    await spreadsheet.cell(8, 1).press('Enter');
    await spreadsheet.typeInEditor('backspace test');
    await spreadsheet.commitWithEnter();

    await spreadsheet.clickCell(8, 1);
    await spreadsheet.cell(8, 1).press('Backspace');

    await spreadsheet.waitForCellText(8, 1, '');
  });

  test('editor blur commits the edit', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(9, 0);
    await spreadsheet.cell(9, 0).press('Enter');
    await spreadsheet.typeInEditor('blur test');

    // Click elsewhere to blur
    await spreadsheet.clickCell(9, 2);

    await spreadsheet.waitForCellText(9, 0, 'blur test');
  });
});
