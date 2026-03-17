import { test, expect } from './fixtures';

test.describe('Keyboard Navigation', () => {
  test('arrow down moves active cell down', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('ArrowDown');

    await expect(spreadsheet.cell(1, 0)).toHaveClass(/active-cell/);
    await expect(spreadsheet.cell(1, 0)).toHaveAttribute('tabindex', '0');
  });

  test('arrow right moves active cell right', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('ArrowRight');

    await expect(spreadsheet.cell(0, 1)).toHaveClass(/active-cell/);
  });

  test('arrow up moves active cell up', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(1, 0);
    await spreadsheet.cell(1, 0).press('ArrowUp');

    await expect(spreadsheet.cell(0, 0)).toHaveClass(/active-cell/);
  });

  test('arrow left moves active cell left', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 1);
    await spreadsheet.cell(0, 1).press('ArrowLeft');

    await expect(spreadsheet.cell(0, 0)).toHaveClass(/active-cell/);
  });

  test('arrow up at row 0 stays at row 0 (clamped)', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('ArrowUp');

    await expect(spreadsheet.cell(0, 0)).toHaveClass(/active-cell/);
  });

  test('arrow left at col 0 stays at col 0 (clamped)', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('ArrowLeft');

    await expect(spreadsheet.cell(0, 0)).toHaveClass(/active-cell/);
  });

  test('Tab moves focus out of the grid (no focus trap)', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Tab');

    // Focus should have left the grid — no grid cell should be focused
    await expect(spreadsheet.cell(0, 0)).not.toBeFocused();
    await expect(spreadsheet.cell(0, 1)).not.toBeFocused();
  });

  test('Shift+Tab moves focus out of the grid backwards', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 2);
    await spreadsheet.cell(0, 2).press('Shift+Tab');

    await expect(spreadsheet.cell(0, 2)).not.toBeFocused();
    await expect(spreadsheet.cell(0, 1)).not.toBeFocused();
  });

  test('multiple arrow key presses navigate correctly', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    const cell = spreadsheet.cell(0, 0);

    await cell.press('ArrowDown');
    await spreadsheet.cell(1, 0).press('ArrowDown');
    await spreadsheet.cell(2, 0).press('ArrowRight');
    await spreadsheet.cell(2, 1).press('ArrowRight');

    await expect(spreadsheet.cell(2, 2)).toHaveClass(/active-cell/);
  });

  test('Escape clears the selection range', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);

    // Create a range selection via Shift+Arrow
    await spreadsheet.cell(0, 0).press('Shift+ArrowRight');
    await spreadsheet.cell(0, 1).press('Shift+ArrowDown');

    // Verify multiple cells are selected
    await expect(spreadsheet.cell(0, 0)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(0, 1)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(1, 0)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(1, 1)).toHaveAttribute('aria-selected', 'true');

    // Press Escape - the head is at (1,1), escape collapses anchor to head
    await spreadsheet.cell(1, 1).press('Escape');

    // After escape, selection collapses to just the head cell (1,1)
    await expect(spreadsheet.cell(1, 1)).toHaveAttribute('aria-selected', 'true');
    // Other cells should no longer be selected
    await expect(spreadsheet.cell(0, 0)).toHaveAttribute('aria-selected', 'false');
    await expect(spreadsheet.cell(0, 1)).toHaveAttribute('aria-selected', 'false');
    await expect(spreadsheet.cell(1, 0)).toHaveAttribute('aria-selected', 'false');
  });

  test('focus follows the active cell after navigation', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('ArrowDown');

    // The newly active cell should be focused
    await expect(spreadsheet.cell(1, 0)).toBeFocused();
  });

  test('Home moves to first cell in current row', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(2, 5);
    await spreadsheet.cell(2, 5).press('Home');

    await expect(spreadsheet.cell(2, 0)).toHaveClass(/active-cell/);
  });

  test('End moves to last cell in current row', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(2, 0);
    await spreadsheet.cell(2, 0).press('End');

    await expect(spreadsheet.cell(2, 25)).toHaveClass(/active-cell/);
  });

  test('Ctrl+Home moves to cell A1', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(5, 5);
    await spreadsheet.cell(5, 5).press('Control+Home');

    await expect(spreadsheet.cell(0, 0)).toHaveClass(/active-cell/);
  });

  test('Ctrl+End moves to last cell', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Control+End');

    await expect(spreadsheet.cell(49, 25)).toHaveClass(/active-cell/);
  });

  test('F2 enters edit mode', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('F2');

    await expect(spreadsheet.editor).toBeFocused();
  });

  test('PageDown moves down by approximately one page', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('PageDown');

    const activeCell = spreadsheet.shadow('.active-cell');
    const row = await activeCell.getAttribute('data-row');
    expect(Number(row)).toBeGreaterThan(5);
  });

  test('PageUp moves up by approximately one page', async ({ spreadsheet }) => {
    // First navigate down to row 20 using PageDown (starts at row 0)
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('PageDown');
    await spreadsheet.shadow('.active-cell').press('PageDown');

    // Now we're well past row 5, press PageUp
    const activeBeforeUp = spreadsheet.shadow('.active-cell');
    const rowBefore = Number(await activeBeforeUp.getAttribute('data-row'));

    await activeBeforeUp.press('PageUp');

    const activeAfterUp = spreadsheet.shadow('.active-cell');
    const rowAfter = Number(await activeAfterUp.getAttribute('data-row'));
    expect(rowAfter).toBeLessThan(rowBefore);
  });

  test('Shift+Home extends selection to row start', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(2, 3);
    await spreadsheet.cell(2, 3).press('Shift+Home');

    await expect(spreadsheet.cell(2, 0)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(2, 1)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(2, 2)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(2, 3)).toHaveAttribute('aria-selected', 'true');
  });
});
