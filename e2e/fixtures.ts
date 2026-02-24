import { test as base, type Page, type Locator } from '@playwright/test';

/**
 * Helper class encapsulating common spreadsheet interactions for E2E tests.
 */
export class SpreadsheetPage {
  readonly page: Page;

  constructor(page: Page) {
    this.page = page;
  }

  /** Navigate to the demo page and wait for the component to render */
  async goto() {
    await this.page.goto('/');
    await this.page.waitForSelector('y11n-spreadsheet');
    // Wait for shadow DOM to be ready
    await this.page.waitForFunction(() => {
      const el = document.querySelector('y11n-spreadsheet');
      return el?.shadowRoot?.querySelector('[role="grid"]') !== null;
    });
  }

  /** Get the spreadsheet host element */
  get host(): Locator {
    return this.page.locator('y11n-spreadsheet');
  }

  /** Get a locator inside the shadow DOM */
  shadow(selector: string): Locator {
    return this.page.locator(`y11n-spreadsheet`).locator(selector);
  }

  /** Get the grid element */
  get grid(): Locator {
    return this.shadow('[role="grid"]');
  }

  /** Get a specific cell by row and col (0-indexed) */
  cell(row: number, col: number): Locator {
    return this.shadow(`[data-row="${row}"][data-col="${col}"]`);
  }

  /** Get the editor input */
  get editor(): Locator {
    return this.shadow('#editor');
  }

  /** Get column headers */
  get columnHeaders(): Locator {
    return this.shadow('[role="columnheader"]');
  }

  /** Get row headers */
  get rowHeaders(): Locator {
    return this.shadow('[role="rowheader"]');
  }

  /** Get the live region for screen reader announcements */
  get liveRegion(): Locator {
    return this.shadow('[role="status"][aria-live="polite"]');
  }

  /** Click on a specific cell */
  async clickCell(row: number, col: number) {
    await this.cell(row, col).click();
  }

  /** Double-click on a specific cell */
  async dblClickCell(row: number, col: number) {
    await this.cell(row, col).dblclick();
  }

  /** Get the displayed text of a cell */
  async getCellText(row: number, col: number): Promise<string> {
    return (await this.cell(row, col).locator('.cell-text').textContent()) ?? '';
  }

  /** Type into the editor (assumes editor is active) */
  async typeInEditor(text: string) {
    await this.editor.fill(text);
  }

  /** Commit an edit by pressing Enter */
  async commitWithEnter() {
    await this.editor.press('Enter');
  }

  /** Commit an edit by pressing Tab */
  async commitWithTab() {
    await this.editor.press('Tab');
  }

  /** Cancel an edit by pressing Escape */
  async cancelEdit() {
    await this.editor.press('Escape');
  }

  /** Set data on the spreadsheet via JS evaluation */
  async setData(data: Record<string, { rawValue: string; displayValue: string; type: string }>) {
    await this.page.evaluate((entries) => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      const map = new Map();
      for (const [key, value] of Object.entries(entries)) {
        map.set(key, value);
      }
      sheet.setData(map);
    }, data);
    // Wait for Lit to update
    await this.page.waitForTimeout(100);
  }

  /** Get data from the spreadsheet */
  async getData(): Promise<Record<string, any>> {
    return await this.page.evaluate(() => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      const data = sheet.getData();
      const result: Record<string, any> = {};
      data.forEach((value: any, key: string) => {
        result[key] = value;
      });
      return result;
    });
  }

  /** Set a property on the spreadsheet element */
  async setProperty(prop: string, value: any) {
    await this.page.evaluate(
      ({ prop, value }) => {
        const sheet = document.querySelector('y11n-spreadsheet') as any;
        (sheet as any)[prop] = value;
      },
      { prop, value }
    );
    await this.page.waitForTimeout(100);
  }

  /** Set the read-only attribute */
  async setReadOnly(readOnly: boolean) {
    await this.page.evaluate((ro) => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      if (ro) {
        sheet.setAttribute('read-only', '');
      } else {
        sheet.removeAttribute('read-only');
      }
    }, readOnly);
    await this.page.waitForTimeout(100);
  }

  /** Set cell format via JS evaluation */
  async setCellFormat(cellId: string, format: Record<string, any>) {
    await this.page.evaluate(
      ({ cellId, format }) => {
        const sheet = document.querySelector('y11n-spreadsheet') as any;
        sheet.setCellFormat(cellId, format);
      },
      { cellId, format }
    );
    await this.page.waitForTimeout(100);
  }

  /** Get cell format via JS evaluation */
  async getCellFormat(cellId: string): Promise<Record<string, any> | undefined> {
    return await this.page.evaluate((cellId) => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      return sheet.getCellFormat(cellId);
    }, cellId);
  }

  /** Get the inline style of a rendered cell */
  async getCellInlineStyle(row: number, col: number): Promise<string> {
    return await this.page.evaluate(
      ({ row, col }) => {
        const sheet = document.querySelector('y11n-spreadsheet');
        const cell = sheet?.shadowRoot?.querySelector(
          `[data-row="${row}"][data-col="${col}"]`
        ) as HTMLElement | null;
        return cell?.getAttribute('style') ?? '';
      },
      { row, col }
    );
  }

  /** Set range format via JS evaluation */
  async setRangeFormat(
    startRow: number, startCol: number,
    endRow: number, endCol: number,
    format: Record<string, any>
  ) {
    await this.page.evaluate(
      ({ startRow, startCol, endRow, endCol, format }) => {
        const sheet = document.querySelector('y11n-spreadsheet') as any;
        sheet.setRangeFormat(
          { start: { row: startRow, col: startCol }, end: { row: endRow, col: endCol } },
          format
        );
      },
      { startRow, startCol, endRow, endCol, format }
    );
    await this.page.waitForTimeout(100);
  }

  /** Clear cell format via JS evaluation */
  async clearCellFormat(cellId: string) {
    await this.page.evaluate((cellId) => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      sheet.clearCellFormat(cellId);
    }, cellId);
    await this.page.waitForTimeout(100);
  }

  /** Clear range format via JS evaluation */
  async clearRangeFormat(
    startRow: number, startCol: number,
    endRow: number, endCol: number
  ) {
    await this.page.evaluate(
      ({ startRow, startCol, endRow, endCol }) => {
        const sheet = document.querySelector('y11n-spreadsheet') as any;
        sheet.clearRangeFormat(
          { start: { row: startRow, col: startCol }, end: { row: endRow, col: endCol } }
        );
      },
      { startRow, startCol, endRow, endCol }
    );
    await this.page.waitForTimeout(100);
  }

  /** Check if the editor is visible */
  async isEditorVisible(): Promise<boolean> {
    const style = await this.editor.getAttribute('style');
    return style !== null && !style.includes('display:none') && !style.includes('display: none');
  }

  /** Wait for cell to have specific text */
  async waitForCellText(row: number, col: number, text: string) {
    await this.page.waitForFunction(
      ({ row, col, text }) => {
        const sheet = document.querySelector('y11n-spreadsheet');
        const cell = sheet?.shadowRoot?.querySelector(
          `[data-row="${row}"][data-col="${col}"] .cell-text`
        );
        return cell?.textContent === text;
      },
      { row, col, text }
    );
  }
}

/** Extended test fixture with SpreadsheetPage */
export const test = base.extend<{ spreadsheet: SpreadsheetPage }>({
  spreadsheet: async ({ page }, use) => {
    const sp = new SpreadsheetPage(page);
    await sp.goto();
    await use(sp);
  },
});

export { expect } from '@playwright/test';
