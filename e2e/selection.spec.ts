import { test, expect } from './fixtures';

test.describe('Selection', () => {
  test('clicking a cell selects it', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(2, 2);

    await expect(spreadsheet.cell(2, 2)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(2, 2)).toHaveClass(/active-cell/);
  });

  test('clicking a different cell deselects the previous one', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await expect(spreadsheet.cell(0, 0)).toHaveClass(/active-cell/);

    await spreadsheet.clickCell(2, 2);
    await expect(spreadsheet.cell(2, 2)).toHaveClass(/active-cell/);

    // Previous cell should no longer be the active cell
    const classes = await spreadsheet.cell(0, 0).getAttribute('class');
    expect(classes).not.toContain('active-cell');
  });

  test('Shift+Arrow extends selection range', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Shift+ArrowRight');
    await spreadsheet.cell(0, 1).press('Shift+ArrowDown');

    // All four cells in the 2x2 range should be selected
    await expect(spreadsheet.cell(0, 0)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(0, 1)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(1, 0)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(1, 1)).toHaveAttribute('aria-selected', 'true');

    // Cell outside range should not be selected
    await expect(spreadsheet.cell(2, 2)).toHaveAttribute('aria-selected', 'false');
  });

  test('Ctrl+A selects all cells', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Control+a');

    // Spot-check that various cells are selected
    await expect(spreadsheet.cell(0, 0)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(5, 5)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(0, 10)).toHaveAttribute('aria-selected', 'true');
  });

  test('Shift+Click creates a range selection', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(2, 2).click({ modifiers: ['Shift'] });

    // The range 0,0 to 2,2 should all be selected
    await expect(spreadsheet.cell(0, 0)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(1, 1)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(2, 2)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(0, 2)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(2, 0)).toHaveAttribute('aria-selected', 'true');
  });

  test('arrow without Shift collapses selection to single cell', async ({ spreadsheet }) => {
    // Create a range
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Shift+ArrowRight');
    await spreadsheet.cell(0, 1).press('Shift+ArrowDown');

    // Now move without shift - should collapse
    await spreadsheet.cell(1, 1).press('ArrowDown');

    // Only the new active cell should be selected
    await expect(spreadsheet.cell(2, 1)).toHaveClass(/active-cell/);
    await expect(spreadsheet.cell(0, 0)).toHaveAttribute('aria-selected', 'false');
  });

  test('Delete clears all cells in selection range', async ({ spreadsheet }) => {
    // Set up data in a range
    await spreadsheet.setData({
      '4:0': { rawValue: 'A', displayValue: 'A', type: 'text' },
      '4:1': { rawValue: 'B', displayValue: 'B', type: 'text' },
      '5:0': { rawValue: 'C', displayValue: 'C', type: 'text' },
      '5:1': { rawValue: 'D', displayValue: 'D', type: 'text' },
    });

    // Select the range
    await spreadsheet.clickCell(4, 0);
    await spreadsheet.cell(4, 0).press('Shift+ArrowRight');
    await spreadsheet.cell(4, 1).press('Shift+ArrowDown');

    // Delete
    await spreadsheet.cell(5, 1).press('Delete');

    // All cells should be cleared
    await spreadsheet.waitForCellText(4, 0, '');
    await spreadsheet.waitForCellText(4, 1, '');
    await spreadsheet.waitForCellText(5, 0, '');
    await spreadsheet.waitForCellText(5, 1, '');
  });

  test('mouse drag creates a selection range', async ({ spreadsheet }) => {
    const startCell = spreadsheet.cell(0, 0);
    const endCell = spreadsheet.cell(1, 1);

    const startBox = await startCell.boundingBox();
    const endBox = await endCell.boundingBox();

    if (!startBox || !endBox) throw new Error('Could not get cell bounding boxes');

    await spreadsheet.page.mouse.move(
      startBox.x + startBox.width / 2,
      startBox.y + startBox.height / 2
    );
    await spreadsheet.page.mouse.down();
    await spreadsheet.page.mouse.move(
      endBox.x + endBox.width / 2,
      endBox.y + endBox.height / 2
    );
    await spreadsheet.page.mouse.up();

    // Cells in the drag range should be selected
    await expect(spreadsheet.cell(0, 0)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(0, 1)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(1, 0)).toHaveAttribute('aria-selected', 'true');
    await expect(spreadsheet.cell(1, 1)).toHaveAttribute('aria-selected', 'true');
  });
});
