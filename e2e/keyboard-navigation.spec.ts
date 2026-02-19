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

  test('Tab moves right', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Tab');

    await expect(spreadsheet.cell(0, 1)).toHaveClass(/active-cell/);
  });

  test('Shift+Tab moves left', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 2);
    await spreadsheet.cell(0, 2).press('Shift+Tab');

    await expect(spreadsheet.cell(0, 1)).toHaveClass(/active-cell/);
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
});
