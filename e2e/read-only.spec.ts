import { test, expect } from './fixtures';

test.describe('Read-Only Mode', () => {
  test.beforeEach(async ({ spreadsheet }) => {
    await spreadsheet.setReadOnly(true);
  });

  test('grid has aria-readonly=true when read-only', async ({ spreadsheet }) => {
    await expect(spreadsheet.grid).toHaveAttribute('aria-readonly', 'true');
  });

  test('cells have aria-readonly=true when read-only', async ({ spreadsheet }) => {
    await expect(spreadsheet.cell(0, 0)).toHaveAttribute('aria-readonly', 'true');
  });

  test('Enter does not open editor in read-only mode', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Enter');

    const visible = await spreadsheet.isEditorVisible();
    expect(visible).toBe(false);
  });

  test('double-click does not open editor in read-only mode', async ({ spreadsheet }) => {
    await spreadsheet.dblClickCell(0, 0);

    const visible = await spreadsheet.isEditorVisible();
    expect(visible).toBe(false);
  });

  test('typing does not start editing in read-only mode', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('h');

    const visible = await spreadsheet.isEditorVisible();
    expect(visible).toBe(false);
  });

  test('Delete does not clear cells in read-only mode', async ({ spreadsheet }) => {
    const originalText = await spreadsheet.getCellText(0, 0);

    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Delete');

    const currentText = await spreadsheet.getCellText(0, 0);
    expect(currentText).toBe(originalText);
  });

  test('keyboard navigation still works in read-only mode', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('ArrowDown');
    await expect(spreadsheet.cell(1, 0)).toHaveClass(/active-cell/);

    await spreadsheet.cell(1, 0).press('ArrowRight');
    await expect(spreadsheet.cell(1, 1)).toHaveClass(/active-cell/);
  });

  test('selection still works in read-only mode', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Shift+ArrowRight');

    await expect(spreadsheet.cell(0, 0)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(0, 1)).toHaveAttribute('aria-selected', 'true');
  });

  test('toggling read-only off allows editing again', async ({ spreadsheet }) => {
    await spreadsheet.setReadOnly(false);

    await spreadsheet.clickCell(6, 0);
    await spreadsheet.cell(6, 0).press('Enter');

    const visible = await spreadsheet.isEditorVisible();
    expect(visible).toBe(true);

    await spreadsheet.cancelEdit();
  });
});
