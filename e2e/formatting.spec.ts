import { test, expect } from './fixtures';

test.describe('Cell Formatting', () => {
  test('setCellFormat applies inline styles to rendered cells', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Bold', displayValue: 'Bold', type: 'text' },
    });
    await spreadsheet.setCellFormat('0:0', { bold: true });

    const style = await spreadsheet.getCellInlineStyle(0, 0);
    expect(style).toContain('font-weight: bold');
  });

  test('setCellFormat applies text color', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Red', displayValue: 'Red', type: 'text' },
    });
    await spreadsheet.setCellFormat('0:0', { textColor: '#ff0000' });

    // Browser may convert hex to rgb() in inline style, so check via API
    const fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.textColor).toBe('#ff0000');
    // Also verify the style attribute contains a color property
    const style = await spreadsheet.getCellInlineStyle(0, 0);
    expect(style).toContain('color:');
  });

  test('setCellFormat applies background color', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'BG', displayValue: 'BG', type: 'text' },
    });
    await spreadsheet.setCellFormat('0:0', { backgroundColor: '#00ff00' });

    const fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.backgroundColor).toBe('#00ff00');
    const style = await spreadsheet.getCellInlineStyle(0, 0);
    expect(style).toContain('background-color:');
  });

  test('setCellFormat applies font size', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Big', displayValue: 'Big', type: 'text' },
    });
    await spreadsheet.setCellFormat('0:0', { fontSize: 20 });

    const style = await spreadsheet.getCellInlineStyle(0, 0);
    expect(style).toContain('font-size: 20px');
  });

  test('setCellFormat merges with existing format', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Test', displayValue: 'Test', type: 'text' },
    });
    await spreadsheet.setCellFormat('0:0', { bold: true });
    await spreadsheet.setCellFormat('0:0', { italic: true });

    const fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.bold).toBe(true);
    expect(fmt?.italic).toBe(true);
  });

  test('setRangeFormat styles multiple cells', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'A', displayValue: 'A', type: 'text' },
      '0:1': { rawValue: 'B', displayValue: 'B', type: 'text' },
      '1:0': { rawValue: 'C', displayValue: 'C', type: 'text' },
      '1:1': { rawValue: 'D', displayValue: 'D', type: 'text' },
    });
    await spreadsheet.setRangeFormat(0, 0, 1, 1, { bold: true });

    for (const [r, c] of [[0, 0], [0, 1], [1, 0], [1, 1]]) {
      const style = await spreadsheet.getCellInlineStyle(r, c);
      expect(style).toContain('font-weight: bold');
    }
  });

  test('clearCellFormat removes styling', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Styled', displayValue: 'Styled', type: 'text' },
    });
    await spreadsheet.setCellFormat('0:0', { bold: true, textColor: '#ff0000' });
    await spreadsheet.clearCellFormat('0:0');

    const fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt).toBeUndefined();

    const style = await spreadsheet.getCellInlineStyle(0, 0);
    expect(style).not.toContain('font-weight');
    expect(style).not.toContain('color');
  });

  test('getCellFormat returns correct data', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Test', displayValue: 'Test', type: 'text' },
    });
    await spreadsheet.setCellFormat('0:0', {
      bold: true,
      italic: true,
      textColor: '#1565c0',
      fontSize: 16,
    });

    const fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt).toBeDefined();
    expect(fmt!.bold).toBe(true);
    expect(fmt!.italic).toBe(true);
    expect(fmt!.textColor).toBe('#1565c0');
    expect(fmt!.fontSize).toBe(16);
  });

  test('getCellFormat returns undefined for unformatted cell', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Plain', displayValue: 'Plain', type: 'text' },
    });

    const fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt).toBeUndefined();
  });

  test('undo reverts format changes', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Test', displayValue: 'Test', type: 'text' },
    });
    await spreadsheet.clickCell(0, 0);

    await spreadsheet.setCellFormat('0:0', { bold: true });
    let fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.bold).toBe(true);

    // Undo
    await spreadsheet.cell(0, 0).press('Control+z');
    await spreadsheet.page.waitForTimeout(100);

    fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.bold).toBeFalsy();
  });

  test('redo reapplies format changes', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Test', displayValue: 'Test', type: 'text' },
    });
    await spreadsheet.clickCell(0, 0);

    await spreadsheet.setCellFormat('0:0', { bold: true });

    // Undo
    await spreadsheet.cell(0, 0).press('Control+z');
    await spreadsheet.page.waitForTimeout(100);

    // Redo (Ctrl+Shift+Z)
    await spreadsheet.cell(0, 0).press('Control+Shift+z');
    await spreadsheet.page.waitForTimeout(100);

    const fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.bold).toBe(true);
  });

  test('Ctrl+B toggles bold on selection', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Toggle', displayValue: 'Toggle', type: 'text' },
    });
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Control+b');
    await spreadsheet.page.waitForTimeout(100);

    let fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.bold).toBe(true);

    // Toggle off
    await spreadsheet.cell(0, 0).press('Control+b');
    await spreadsheet.page.waitForTimeout(100);

    fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.bold).toBeFalsy();
  });

  test('Ctrl+I toggles italic on selection', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Toggle', displayValue: 'Toggle', type: 'text' },
    });
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Control+i');
    await spreadsheet.page.waitForTimeout(100);

    const fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.italic).toBe(true);
  });

  test('Ctrl+U toggles underline on selection', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Toggle', displayValue: 'Toggle', type: 'text' },
    });
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Control+u');
    await spreadsheet.page.waitForTimeout(100);

    const fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.underline).toBe(true);
  });

  test('combined text-decoration for underline + strikethrough', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Both', displayValue: 'Both', type: 'text' },
    });
    await spreadsheet.setCellFormat('0:0', { underline: true, strikethrough: true });

    const style = await spreadsheet.getCellInlineStyle(0, 0);
    expect(style).toContain('text-decoration');
    expect(style).toContain('underline');
    expect(style).toContain('line-through');
  });
});

test.describe('Clipboard with Formatting', () => {
  test.use({
    permissions: ['clipboard-read', 'clipboard-write'],
  });

  test('copy formatted cell and paste preserves format', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'Styled', displayValue: 'Styled', type: 'text' },
    });
    await spreadsheet.setCellFormat('0:0', { bold: true, textColor: '#ff0000' });

    // Copy
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.cell(0, 0).press('Control+c');

    // Paste to new location
    await spreadsheet.clickCell(2, 0);
    await spreadsheet.cell(2, 0).press('Control+v');
    await spreadsheet.page.waitForTimeout(200);

    // Verify value was pasted
    await spreadsheet.waitForCellText(2, 0, 'Styled');

    // Verify format was preserved
    const fmt = await spreadsheet.getCellFormat('2:0');
    expect(fmt?.bold).toBe(true);
    expect(fmt?.textColor).toBe('#ff0000');
  });
});
