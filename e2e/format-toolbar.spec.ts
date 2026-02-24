import { test, expect } from './fixtures';

test.describe('Format Toolbar', () => {
  /** Inject a <y11n-format-toolbar> and wire it to the spreadsheet */
  async function setupToolbar(page: import('@playwright/test').Page) {
    await page.evaluate(() => {
      // Import the component (already registered via index.ts)
      const toolbar = document.createElement('y11n-format-toolbar') as any;
      toolbar.id = 'fmt-toolbar';
      const container = document.querySelector('.demo-container')!;
      container.insertBefore(toolbar, container.firstChild);

      const sheet = document.querySelector('y11n-spreadsheet') as any;

      // Track the current selection range from events
      let currentRange: any = null;

      // Wire toolbar actions → spreadsheet
      toolbar.addEventListener('format-action', (e: CustomEvent) => {
        const { action, value } = e.detail;

        if (action === 'clearFormat') {
          if (currentRange) sheet.clearRangeFormat(currentRange);
        } else if (['bold', 'italic', 'underline', 'strikethrough'].includes(action)) {
          // toggleFormat uses internal selection, works regardless of focus
          sheet.toggleFormat(action);
        } else {
          if (currentRange) sheet.setRangeFormat(currentRange, { [action]: value });
        }
      });

      // Wire selection-change → update toolbar state
      sheet.addEventListener('selection-change', (e: CustomEvent) => {
        const range = e.detail.range;
        currentRange = range;
        const cellId = `${range.start.row}:${range.start.col}`;
        const fmt = sheet.getCellFormat(cellId) || {};
        toolbar.bold = !!fmt.bold;
        toolbar.italic = !!fmt.italic;
        toolbar.underline = !!fmt.underline;
        toolbar.strikethrough = !!fmt.strikethrough;
        toolbar.textColor = fmt.textColor || '#000000';
        toolbar.backgroundColor = fmt.backgroundColor || '#ffffff';
        toolbar.textAlign = fmt.textAlign || 'left';
        toolbar.fontSize = fmt.fontSize || 13;
      });
    });
    await page.waitForTimeout(200);
  }

  function toolbarButton(page: import('@playwright/test').Page, label: string) {
    return page.locator('#fmt-toolbar').locator(`[aria-label="${label}"]`);
  }

  test('toolbar reflects cell format on selection change', async ({ spreadsheet }) => {
    await setupToolbar(spreadsheet.page);

    // Format cell A1 as bold
    await spreadsheet.setData({
      '0:0': { rawValue: 'Bold', displayValue: 'Bold', type: 'text' },
    });
    await spreadsheet.setCellFormat('0:0', { bold: true });

    // Click the cell to trigger selection-change
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.page.waitForTimeout(200);

    // Bold button should be pressed
    const boldBtn = toolbarButton(spreadsheet.page, 'Bold');
    await expect(boldBtn).toHaveAttribute('aria-pressed', 'true');
  });

  test('toolbar reflects unformatted cell as all off', async ({ spreadsheet }) => {
    await setupToolbar(spreadsheet.page);

    await spreadsheet.setData({
      '0:0': { rawValue: 'Bold', displayValue: 'Bold', type: 'text' },
      '1:0': { rawValue: 'Plain', displayValue: 'Plain', type: 'text' },
    });
    await spreadsheet.setCellFormat('0:0', { bold: true });

    // Select formatted cell first
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.page.waitForTimeout(100);

    // Now select unformatted cell
    await spreadsheet.clickCell(1, 0);
    await spreadsheet.page.waitForTimeout(200);

    const boldBtn = toolbarButton(spreadsheet.page, 'Bold');
    await expect(boldBtn).toHaveAttribute('aria-pressed', 'false');
  });

  test('bold button toggles formatting on selected cell', async ({ spreadsheet }) => {
    await setupToolbar(spreadsheet.page);

    await spreadsheet.setData({
      '0:0': { rawValue: 'Test', displayValue: 'Test', type: 'text' },
    });
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.page.waitForTimeout(100);

    // Click bold button
    const boldBtn = toolbarButton(spreadsheet.page, 'Bold');
    await boldBtn.click();
    await spreadsheet.page.waitForTimeout(200);

    // Cell should be bold
    const fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.bold).toBe(true);

    const style = await spreadsheet.getCellInlineStyle(0, 0);
    expect(style).toContain('font-weight: bold');
  });

  test('italic button toggles formatting', async ({ spreadsheet }) => {
    await setupToolbar(spreadsheet.page);

    await spreadsheet.setData({
      '0:0': { rawValue: 'Test', displayValue: 'Test', type: 'text' },
    });
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.page.waitForTimeout(100);

    const italicBtn = toolbarButton(spreadsheet.page, 'Italic');
    await italicBtn.click();
    await spreadsheet.page.waitForTimeout(200);

    const fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.italic).toBe(true);
  });

  test('underline button toggles formatting', async ({ spreadsheet }) => {
    await setupToolbar(spreadsheet.page);

    await spreadsheet.setData({
      '0:0': { rawValue: 'Test', displayValue: 'Test', type: 'text' },
    });
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.page.waitForTimeout(100);

    const underlineBtn = toolbarButton(spreadsheet.page, 'Underline');
    await underlineBtn.click();
    await spreadsheet.page.waitForTimeout(200);

    const fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.underline).toBe(true);
  });

  test('alignment buttons work as radio group', async ({ spreadsheet }) => {
    await setupToolbar(spreadsheet.page);

    await spreadsheet.setData({
      '0:0': { rawValue: 'Align', displayValue: 'Align', type: 'text' },
    });
    await spreadsheet.clickCell(0, 0);
    await spreadsheet.page.waitForTimeout(100);

    // Click center align
    const centerBtn = toolbarButton(spreadsheet.page, 'Align center');
    await centerBtn.click();
    await spreadsheet.page.waitForTimeout(200);

    let fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.textAlign).toBe('center');

    // Click right align
    const rightBtn = toolbarButton(spreadsheet.page, 'Align right');
    await rightBtn.click();
    await spreadsheet.page.waitForTimeout(200);

    fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.textAlign).toBe('right');
  });

  test('Ctrl+B/I/U keyboard shortcuts toggle formatting', async ({ spreadsheet }) => {
    await setupToolbar(spreadsheet.page);

    await spreadsheet.setData({
      '0:0': { rawValue: 'Keys', displayValue: 'Keys', type: 'text' },
    });
    await spreadsheet.clickCell(0, 0);

    // Ctrl+B
    await spreadsheet.cell(0, 0).press('Control+b');
    await spreadsheet.page.waitForTimeout(100);
    let fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.bold).toBe(true);

    // Ctrl+I
    await spreadsheet.cell(0, 0).press('Control+i');
    await spreadsheet.page.waitForTimeout(100);
    fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.italic).toBe(true);

    // Ctrl+U
    await spreadsheet.cell(0, 0).press('Control+u');
    await spreadsheet.page.waitForTimeout(100);
    fmt = await spreadsheet.getCellFormat('0:0');
    expect(fmt?.underline).toBe(true);
  });
});
