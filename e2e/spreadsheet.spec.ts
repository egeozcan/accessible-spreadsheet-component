import { test, expect, type Page } from '@playwright/test';

// Helper: get the spreadsheet component
function getSpreadsheet(page: Page) {
  return page.locator('y11n-spreadsheet');
}

// Helper: get a cell by row/col (0-indexed) inside shadow DOM
function getCell(page: Page, row: number, col: number) {
  return page.locator('y11n-spreadsheet').locator(`[data-row="${row}"][data-col="${col}"]`);
}

// Helper: get the editor input inside shadow DOM
function getEditor(page: Page) {
  return page.locator('y11n-spreadsheet').locator('#editor');
}

// Helper: get the formula bar input inside shadow DOM
function getFormulaBarInput(page: Page) {
  return page.locator('y11n-spreadsheet').locator('y11n-formula-bar input');
}

// Helper: get cell text content
async function getCellText(page: Page, row: number, col: number): Promise<string> {
  const cell = getCell(page, row, col);
  const span = cell.locator('.cell-text');
  return (await span.textContent()) ?? '';
}

// Helper: type into a cell and commit
async function editCell(page: Page, row: number, col: number, value: string) {
  const cell = getCell(page, row, col);
  await cell.click();
  await cell.dblclick();
  const editor = getEditor(page);
  await editor.fill(value);
  await editor.press('Enter');
}

async function pressUndo(page: Page) {
  await page.keyboard.press('ControlOrMeta+z');
}

async function pressRedo(page: Page) {
  await page.keyboard.press('ControlOrMeta+Shift+z');
}

async function installClipboardMock(page: Page, initialText = '') {
  await page.evaluate((text) => {
    let clipboardText = text;
    Object.defineProperty(window.navigator, 'clipboard', {
      configurable: true,
      value: {
        readText: async () => clipboardText,
        writeText: async (nextText: string) => {
          clipboardText = nextText;
        },
      },
    });
    (window as unknown as { __setClipboardText?: (nextText: string) => void }).__setClipboardText = (
      nextText: string
    ) => {
      clipboardText = nextText;
    };
  }, initialText);
}

async function setClipboardText(page: Page, text: string) {
  await page.evaluate((nextText) => {
    (window as unknown as { __setClipboardText: (value: string) => void }).__setClipboardText(nextText);
  }, text);
}

test.describe('Spreadsheet Component', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    // Wait for the web component to be defined and rendered
    await page.waitForSelector('y11n-spreadsheet');
    // Wait for shadow DOM to be ready
    await page.locator('y11n-spreadsheet').locator('[data-row="0"]').first().waitFor();
  });

  test.describe('Rendering', () => {
    test('renders the grid with headers', async ({ page }) => {
      const spreadsheet = getSpreadsheet(page);
      await expect(spreadsheet).toBeVisible();

      // Check column headers exist
      const colHeaders = spreadsheet.locator('.ls-col-header');
      const count = await colHeaders.count();
      expect(count).toBeGreaterThan(0);

      // First header should be "A"
      await expect(colHeaders.first()).toHaveText('A');
    });

    test('renders row numbers', async ({ page }) => {
      const spreadsheet = getSpreadsheet(page);
      const rowHeaders = spreadsheet.locator('.ls-row-header');
      const first = rowHeaders.first();
      await expect(first).toHaveText('1');
    });

    test('renders pre-populated data', async ({ page }) => {
      // The demo page populates cells: Item, Qty, Price, Total
      const itemCell = await getCellText(page, 0, 0);
      expect(itemCell).toBe('Item');

      const qtyCell = await getCellText(page, 0, 1);
      expect(qtyCell).toBe('Qty');
    });
  });

  test.describe('Navigation', () => {
    test('navigates with arrow keys', async ({ page }) => {
      const cell00 = getCell(page, 0, 0);
      await cell00.click();
      await expect(cell00).toBeFocused();

      // Move right
      await page.keyboard.press('ArrowRight');
      const cell01 = getCell(page, 0, 1);
      await expect(cell01).toBeFocused();

      // Move down
      await page.keyboard.press('ArrowDown');
      const cell11 = getCell(page, 1, 1);
      await expect(cell11).toBeFocused();

      // Move left
      await page.keyboard.press('ArrowLeft');
      const cell10 = getCell(page, 1, 0);
      await expect(cell10).toBeFocused();

      // Move up
      await page.keyboard.press('ArrowUp');
      await expect(cell00).toBeFocused();
    });

    test('navigates with Tab', async ({ page }) => {
      const cell00 = getCell(page, 0, 0);
      await cell00.click();

      await page.keyboard.press('Tab');
      const cell01 = getCell(page, 0, 1);
      await expect(cell01).toBeFocused();
    });

    test('does not navigate past grid boundaries', async ({ page }) => {
      const cell00 = getCell(page, 0, 0);
      await cell00.click();

      // Try to go up from row 0
      await page.keyboard.press('ArrowUp');
      await expect(cell00).toBeFocused();

      // Try to go left from col 0
      await page.keyboard.press('ArrowLeft');
      await expect(cell00).toBeFocused();
    });
  });

  test.describe('Editing', () => {
    test('enters edit mode on Enter key', async ({ page }) => {
      const cell00 = getCell(page, 0, 0);
      await cell00.click();
      await page.keyboard.press('Enter');

      const editor = getEditor(page);
      await expect(editor).toBeFocused();
    });

    test('enters edit mode on double click', async ({ page }) => {
      const cell00 = getCell(page, 0, 0);
      await cell00.dblclick();

      const editor = getEditor(page);
      await expect(editor).toBeFocused();
    });

    test('enters edit mode on character input', async ({ page }) => {
      // Click an empty cell
      const cell = getCell(page, 10, 0);
      await cell.click();
      await page.keyboard.press('h');

      const editor = getEditor(page);
      await expect(editor).toBeFocused();
    });

    test('commits edit on Enter', async ({ page }) => {
      await editCell(page, 10, 0, 'test value');

      const text = await getCellText(page, 10, 0);
      expect(text).toBe('test value');
    });

    test('cancels edit on Escape', async ({ page }) => {
      const cell = getCell(page, 10, 0);
      await cell.click();
      await cell.dblclick();

      const editor = getEditor(page);
      await editor.fill('should be cancelled');
      await page.keyboard.press('Escape');

      // Cell should remain empty (original value)
      const text = await getCellText(page, 10, 0);
      expect(text).toBe('');
    });

    test('moves down after committing with Enter', async ({ page }) => {
      await editCell(page, 10, 0, 'hello');

      // After Enter, focus should move to the cell below
      const cellBelow = getCell(page, 11, 0);
      await expect(cellBelow).toBeFocused();
    });

    test('formula bar shows raw value for formula cells and can commit edits', async ({ page }) => {
      const formulaCell = getCell(page, 1, 3);
      await formulaCell.click();

      const formulaBar = getFormulaBarInput(page);
      await expect(formulaBar).toHaveValue('=B2*C2');

      await formulaBar.fill('=B2*C2+1');
      await formulaBar.press('Enter');

      const display = await getCellText(page, 1, 3);
      expect(parseFloat(display)).toBeCloseTo(60.9, 1);
    });

    test('formula bar formatted mode shows display value for formula cells', async ({ page }) => {
      const formulaCell = getCell(page, 1, 3);
      await formulaCell.click();

      const formattedBtn = page
        .locator('y11n-spreadsheet')
        .locator('y11n-formula-bar button:has-text("Formatted")');
      await formattedBtn.click();

      const formulaBar = getFormulaBarInput(page);
      const formatted = await formulaBar.inputValue();
      expect(parseFloat(formatted)).toBeCloseTo(59.9, 1);
    });
  });

  test.describe('Formulas', () => {
    test('evaluates arithmetic formulas', async ({ page }) => {
      // The demo has =B2*C2 in D2 (row index 1, col index 3)
      // B2=10, C2=5.99 -> D2 should be 59.9
      const totalText = await getCellText(page, 1, 3);
      expect(parseFloat(totalText)).toBeCloseTo(59.9, 1);
    });

    test('evaluates SUM formula', async ({ page }) => {
      // D6 (row 5, col 3) has =SUM(D2:D4)
      const sumText = await getCellText(page, 5, 3);
      const sumVal = parseFloat(sumText);
      expect(sumVal).toBeGreaterThan(0);
    });

    test('formula recalculates when input changes', async ({ page }) => {
      // Get original total
      const originalTotal = await getCellText(page, 1, 3);

      // Edit B2 (qty) from 10 to 20
      await editCell(page, 1, 1, '20');

      // D2 = B2*C2 should now be 20*5.99 = 119.8
      const newTotal = await getCellText(page, 1, 3);
      expect(parseFloat(newTotal)).toBeCloseTo(119.8, 1);
      expect(newTotal).not.toBe(originalTotal);
    });

    test('string formulas can be entered and evaluated', async ({ page }) => {
      // Enter a value, then a UPPER formula
      await editCell(page, 15, 0, 'hello world');
      await editCell(page, 15, 1, '=UPPER(A16)');

      const upperText = await getCellText(page, 15, 1);
      expect(upperText).toBe('HELLO WORLD');
    });

    test('LEN formula works', async ({ page }) => {
      await editCell(page, 16, 0, 'hello');
      await editCell(page, 16, 1, '=LEN(A17)');

      const lenText = await getCellText(page, 16, 1);
      expect(lenText).toBe('5');
    });

    test('UPPER works on formula cell references', async ({ page }) => {
      // A18 = "hello", B18 = "world", C18 = =CONCAT(A18, " ", B18), D18 = =UPPER(C18)
      await editCell(page, 17, 0, 'hello');
      await editCell(page, 17, 1, 'world');
      await editCell(page, 17, 2, '=CONCAT(A18, " ", B18)');
      await editCell(page, 17, 3, '=UPPER(C18)');

      const concatText = await getCellText(page, 17, 2);
      expect(concatText).toBe('hello world');

      const upperText = await getCellText(page, 17, 3);
      expect(upperText).toBe('HELLO WORLD');
    });

    test('LEN works on formula cell references', async ({ page }) => {
      await editCell(page, 18, 0, 'hello');
      await editCell(page, 18, 1, 'world');
      await editCell(page, 18, 2, '=A19 & " " & B19');
      await editCell(page, 18, 3, '=LEN(C19)');

      const lenText = await getCellText(page, 18, 3);
      expect(lenText).toBe('11');
    });

    test('TRIM formula works', async ({ page }) => {
      await editCell(page, 19, 0, '  spaced  ');
      await editCell(page, 19, 1, '=TRIM(A20)');

      const trimText = await getCellText(page, 19, 1);
      expect(trimText).toBe('spaced');
    });

    test('LOWER formula works', async ({ page }) => {
      await editCell(page, 20, 0, 'HELLO WORLD');
      await editCell(page, 20, 1, '=LOWER(A21)');

      const lowerText = await getCellText(page, 20, 1);
      expect(lowerText).toBe('hello world');
    });

    test('IF formula works', async ({ page }) => {
      await editCell(page, 8, 0, '100');
      await editCell(page, 8, 1, '50');
      await editCell(page, 8, 2, '=IF(A9>B9, "bigger", "smaller")');

      const ifText = await getCellText(page, 8, 2);
      expect(ifText).toBe('bigger');
    });

    test('shows error for invalid formula', async ({ page }) => {
      await editCell(page, 9, 0, '=UNKNOWN_FUNC()');

      const text = await getCellText(page, 9, 0);
      expect(text).toBe('#NAME?');
    });
  });

  test.describe('Selection', () => {
    test('selects a cell on click', async ({ page }) => {
      const cell = getCell(page, 2, 1);
      await cell.click();

      await expect(cell).toHaveAttribute('aria-selected', 'true');
    });

    test('extends selection with Shift+Arrow', async ({ page }) => {
      const cell00 = getCell(page, 0, 0);
      await cell00.click();

      // Shift+ArrowDown to extend
      await page.keyboard.press('Shift+ArrowDown');
      await page.keyboard.press('Shift+ArrowRight');

      // Both original and extended cells should be selected
      const cell01 = getCell(page, 0, 1);
      const cell10 = getCell(page, 1, 0);
      const cell11 = getCell(page, 1, 1);

      await expect(cell00).toHaveAttribute('aria-selected', 'true');
      await expect(cell01).toHaveAttribute('aria-selected', 'true');
      await expect(cell10).toHaveAttribute('aria-selected', 'true');
      await expect(cell11).toHaveAttribute('aria-selected', 'true');
    });

    test('clears selection on Escape', async ({ page }) => {
      const cell00 = getCell(page, 0, 0);
      await cell00.click();

      // Extend selection
      await page.keyboard.press('Shift+ArrowDown');
      await page.keyboard.press('Shift+ArrowRight');

      // Verify extended selection exists
      const cell10 = getCell(page, 1, 0);
      await expect(cell10).toHaveAttribute('aria-selected', 'true');

      // Press Escape to clear
      await page.keyboard.press('Escape');

      // After Escape the range collapses — previously extended cells
      // should no longer be selected
      await expect(cell10).toHaveAttribute('aria-selected', 'false');
      await expect(cell00).toHaveAttribute('aria-selected', 'false');

      // The head cell (1,1) remains as the active cell
      const cell11 = getCell(page, 1, 1);
      await expect(cell11).toHaveAttribute('aria-selected', 'true');
    });
  });

  test.describe('Cell clearing', () => {
    test('clears cell on Delete', async ({ page }) => {
      // Cell A1 has "Item"
      const cell00 = getCell(page, 0, 0);
      await cell00.click();

      const textBefore = await getCellText(page, 0, 0);
      expect(textBefore).toBe('Item');

      await page.keyboard.press('Delete');

      const textAfter = await getCellText(page, 0, 0);
      expect(textAfter).toBe('');
    });

    test('clears cell on Backspace', async ({ page }) => {
      // First set a value in an empty cell
      await editCell(page, 12, 0, 'temp');
      expect(await getCellText(page, 12, 0)).toBe('temp');

      // Navigate back to the cell
      const cell = getCell(page, 12, 0);
      await cell.click();
      await page.keyboard.press('Backspace');

      expect(await getCellText(page, 12, 0)).toBe('');
    });
  });

  test.describe('Undo / Redo', () => {
    test('undoes and redoes committed edit', async ({ page }) => {
      await editCell(page, 12, 1, 'original');
      await editCell(page, 12, 1, 'updated');
      expect(await getCellText(page, 12, 1)).toBe('updated');

      const cell = getCell(page, 12, 1);
      await cell.click();

      await pressUndo(page);
      expect(await getCellText(page, 12, 1)).toBe('original');

      await pressRedo(page);
      expect(await getCellText(page, 12, 1)).toBe('updated');
    });

    test('undo restores value after Delete and restores focus to affected cell', async ({ page }) => {
      const cell00 = getCell(page, 0, 0);
      await cell00.click();
      await page.keyboard.press('Delete');
      expect(await getCellText(page, 0, 0)).toBe('');

      await page.keyboard.press('ArrowRight');
      await expect(getCell(page, 0, 1)).toBeFocused();

      await pressUndo(page);
      expect(await getCellText(page, 0, 0)).toBe('Item');
      await expect(cell00).toBeFocused();
    });

    test('undo restores cut cells', async ({ page }) => {
      await installClipboardMock(page);

      const cell00 = getCell(page, 0, 0);
      await cell00.click();
      await page.keyboard.press('ControlOrMeta+x');
      expect(await getCellText(page, 0, 0)).toBe('');

      await pressUndo(page);
      expect(await getCellText(page, 0, 0)).toBe('Item');
    });

    test('undo and redo pasted range', async ({ page }) => {
      await installClipboardMock(page, 'p11\tp12\np21\tp22');

      const target = getCell(page, 22, 0);
      await target.click();
      await page.keyboard.press('ControlOrMeta+v');

      expect(await getCellText(page, 22, 0)).toBe('p11');
      expect(await getCellText(page, 22, 1)).toBe('p12');
      expect(await getCellText(page, 23, 0)).toBe('p21');
      expect(await getCellText(page, 23, 1)).toBe('p22');

      await pressUndo(page);
      expect(await getCellText(page, 22, 0)).toBe('');
      expect(await getCellText(page, 22, 1)).toBe('');
      expect(await getCellText(page, 23, 0)).toBe('');
      expect(await getCellText(page, 23, 1)).toBe('');

      await pressRedo(page);
      expect(await getCellText(page, 22, 0)).toBe('p11');
      expect(await getCellText(page, 22, 1)).toBe('p12');
      expect(await getCellText(page, 23, 0)).toBe('p21');
      expect(await getCellText(page, 23, 1)).toBe('p22');

      await setClipboardText(page, 'unused');
    });

    test('redo stack clears after new user edit', async ({ page }) => {
      await editCell(page, 13, 0, 'one');
      await editCell(page, 13, 0, 'two');

      const cell = getCell(page, 13, 0);
      await cell.click();
      await pressUndo(page);
      expect(await getCellText(page, 13, 0)).toBe('one');

      await editCell(page, 13, 0, 'three');
      expect(await getCellText(page, 13, 0)).toBe('three');

      await pressRedo(page);
      expect(await getCellText(page, 13, 0)).toBe('three');
    });

    test('read-only mode disables undo and redo', async ({ page }) => {
      await editCell(page, 12, 2, 'locked-value');
      expect(await getCellText(page, 12, 2)).toBe('locked-value');

      await page.evaluate(() => {
        const el = document.querySelector('y11n-spreadsheet') as (HTMLElement & { readOnly: boolean }) | null;
        if (el) {
          el.readOnly = true;
        }
      });

      const cell = getCell(page, 12, 2);
      await cell.click();
      await pressUndo(page);
      expect(await getCellText(page, 12, 2)).toBe('locked-value');

      await pressRedo(page);
      expect(await getCellText(page, 12, 2)).toBe('locked-value');
    });
  });

  test.describe('Accessibility', () => {
    test('grid has correct ARIA role', async ({ page }) => {
      const grid = getSpreadsheet(page).locator('[role="grid"]');
      await expect(grid).toBeVisible();
    });

    test('cells have gridcell role', async ({ page }) => {
      const cell = getCell(page, 0, 0);
      await expect(cell).toHaveAttribute('role', 'gridcell');
    });

    test('column headers have columnheader role', async ({ page }) => {
      const header = getSpreadsheet(page).locator('[role="columnheader"]').first();
      await expect(header).toBeVisible();
    });

    test('row headers have rowheader role', async ({ page }) => {
      const header = getSpreadsheet(page).locator('[role="rowheader"]').first();
      await expect(header).toBeVisible();
    });

    test('active cell has tabindex 0', async ({ page }) => {
      const cell = getCell(page, 0, 0);
      await cell.click();
      await expect(cell).toHaveAttribute('tabindex', '0');
    });

    test('non-active cells have tabindex -1', async ({ page }) => {
      const cell = getCell(page, 0, 0);
      await cell.click();

      const otherCell = getCell(page, 1, 1);
      await expect(otherCell).toHaveAttribute('tabindex', '-1');
    });

    test('has aria-live region for announcements', async ({ page }) => {
      const liveRegion = getSpreadsheet(page).locator('.sr-only[aria-live="polite"]');
      await expect(liveRegion).toBeAttached();
    });
  });

  test.describe('Arrow Key Reference Selection', () => {
    test('arrow key inserts cell reference when editing formula', async ({ page }) => {
      // Navigate to C11 (row 10, col 2)
      const cell = getCell(page, 10, 2);
      await cell.click();

      // Start editing a formula
      await page.keyboard.type('=');

      // Press ArrowDown - should insert reference to cell below (C12)
      await page.keyboard.press('ArrowDown');

      const editor = getEditor(page);
      await expect(editor).toHaveValue('=C12');
    });

    test('subsequent arrow keys move the reference', async ({ page }) => {
      const cell = getCell(page, 10, 2);
      await cell.click();
      await page.keyboard.type('=');

      // ArrowDown → C12, then ArrowRight → D12
      await page.keyboard.press('ArrowDown');
      await page.keyboard.press('ArrowRight');

      const editor = getEditor(page);
      await expect(editor).toHaveValue('=D12');
    });

    test('reference is committed and evaluates correctly', async ({ page }) => {
      // Put a value in B11 (row 10, col 1)
      await editCell(page, 10, 1, '42');

      // Navigate to C11 and create formula referencing B11
      const cell = getCell(page, 10, 2);
      await cell.click();
      await page.keyboard.type('=');

      // ArrowLeft selects B11
      await page.keyboard.press('ArrowLeft');

      const editor = getEditor(page);
      await expect(editor).toHaveValue('=B11');

      // Commit the formula
      await page.keyboard.press('Enter');

      // C11 should display 42 (value of B11)
      const text = await getCellText(page, 10, 2);
      expect(text).toBe('42');
    });

    test('typing an operator exits ref mode and allows a new reference', async ({ page }) => {
      const cell = getCell(page, 10, 2);
      await cell.click();
      await page.keyboard.type('=');

      // First reference: ArrowDown → C12
      await page.keyboard.press('ArrowDown');
      // Type '+' to exit ref mode
      await page.keyboard.type('+');
      // Second reference: ArrowUp → C10 (starts from editing cell C11, goes up)
      await page.keyboard.press('ArrowUp');

      const editor = getEditor(page);
      await expect(editor).toHaveValue('=C12+C10');
    });

    test('arrow keys work normally in non-formula edit', async ({ page }) => {
      const cell = getCell(page, 10, 2);
      await cell.click();
      // Type a non-formula value
      await page.keyboard.type('hello');

      const editor = getEditor(page);
      // ArrowLeft should move cursor, not insert a reference
      await page.keyboard.press('ArrowLeft');
      await expect(editor).toHaveValue('hello');
    });

    test('arrow keys move cursor when not at operator position in formula', async ({ page }) => {
      const cell = getCell(page, 10, 2);
      await cell.click();
      // Type a formula with a complete reference
      await page.keyboard.type('=A1');

      const editor = getEditor(page);
      // Cursor is after '1' which is not an operator - arrow should move cursor
      await page.keyboard.press('ArrowLeft');
      await expect(editor).toHaveValue('=A1');
    });

    test('Escape cancels edit during reference mode', async ({ page }) => {
      const cell = getCell(page, 10, 2);
      await cell.click();
      await page.keyboard.type('=');
      await page.keyboard.press('ArrowDown');

      // Escape should cancel the entire edit
      await page.keyboard.press('Escape');

      const text = await getCellText(page, 10, 2);
      expect(text).toBe('');
    });

    test('referenced cell is highlighted during selection', async ({ page }) => {
      const cell = getCell(page, 10, 2);
      await cell.click();
      await page.keyboard.type('=');
      await page.keyboard.press('ArrowDown');

      // C12 (row 11, col 2) should have ref-highlight class
      const refCell = getCell(page, 11, 2);
      await expect(refCell).toHaveClass(/ref-highlight/);
    });

    test('highlight moves when reference changes', async ({ page }) => {
      const cell = getCell(page, 10, 2);
      await cell.click();
      await page.keyboard.type('=');

      // ArrowDown → C12 highlighted
      await page.keyboard.press('ArrowDown');
      const cellC12 = getCell(page, 11, 2);
      await expect(cellC12).toHaveClass(/ref-highlight/);

      // ArrowRight → D12 highlighted, C12 no longer highlighted
      await page.keyboard.press('ArrowRight');
      const cellD12 = getCell(page, 11, 3);
      await expect(cellD12).toHaveClass(/ref-highlight/);
      await expect(cellC12).not.toHaveClass(/ref-highlight/);
    });

    test('highlight disappears after committing', async ({ page }) => {
      const cell = getCell(page, 10, 2);
      await cell.click();
      await page.keyboard.type('=');
      await page.keyboard.press('ArrowDown');

      const refCell = getCell(page, 11, 2);
      await expect(refCell).toHaveClass(/ref-highlight/);

      // Commit the edit
      await page.keyboard.press('Enter');

      // Highlight should be gone
      await expect(refCell).not.toHaveClass(/ref-highlight/);
    });

    test('builds a complete formula with multiple arrow-key references', async ({ page }) => {
      // Set up values: A8 = 10, B8 = 20
      await editCell(page, 7, 0, '10');
      await editCell(page, 7, 1, '20');

      // Navigate to C8 (row 7, col 2) and build formula =A8+B8
      const cell = getCell(page, 7, 2);
      await cell.click();
      await page.keyboard.type('=');

      // ArrowLeft twice to reach A8: first → B8, second → A8
      await page.keyboard.press('ArrowLeft');
      await page.keyboard.press('ArrowLeft');

      const editor = getEditor(page);
      await expect(editor).toHaveValue('=A8');

      // Type '+' to exit ref mode, then ArrowLeft for B8
      await page.keyboard.type('+');
      await page.keyboard.press('ArrowLeft');
      await expect(editor).toHaveValue('=A8+B8');

      // Commit
      await page.keyboard.press('Enter');

      // C8 should show 30
      const text = await getCellText(page, 7, 2);
      expect(text).toBe('30');
    });

    test('reference works after opening parenthesis in function', async ({ page }) => {
      // Put a value in A14
      await editCell(page, 13, 0, '100');

      // Navigate to B14 and enter edit mode, then type a SUM formula
      const cell = getCell(page, 13, 1);
      await cell.dblclick();
      const editor = getEditor(page);
      await editor.fill('=SUM(');

      // ArrowLeft to reference A14
      await page.keyboard.press('ArrowLeft');
      await expect(editor).toHaveValue('=SUM(A14');

      // Close paren and commit
      await page.keyboard.type(')');
      await page.keyboard.press('Enter');

      const text = await getCellText(page, 13, 1);
      expect(text).toBe('100');
    });

    test('reference works after comma in function arguments', async ({ page }) => {
      // Put values: A15 = 5, C15 = 10
      await editCell(page, 14, 0, '5');
      await editCell(page, 14, 2, '10');

      // Navigate to D15 and enter edit mode, type a SUM formula with first arg
      const cell = getCell(page, 14, 3);
      await cell.dblclick();
      const editor = getEditor(page);
      await editor.fill('=SUM(A15,');

      // ArrowLeft to reference C15
      await page.keyboard.press('ArrowLeft');
      await expect(editor).toHaveValue('=SUM(A15,C15');

      // Close and commit
      await page.keyboard.type(')');
      await page.keyboard.press('Enter');

      const text = await getCellText(page, 14, 3);
      expect(text).toBe('15');
    });
  });
});
