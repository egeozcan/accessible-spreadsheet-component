import { test, expect } from './fixtures';

test.describe('Floating point display', () => {
  test('multiplication formula displays clean number without floating point noise', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: '10', displayValue: '10', type: 'number' },
      '0:1': { rawValue: '5.99', displayValue: '5.99', type: 'number' },
      '0:2': { rawValue: '=A1*B1', displayValue: '', type: 'text' },
    });
    // Should display "59.9", not "59.900000000000006"
    await spreadsheet.waitForCellText(0, 2, '59.9');
  });

  test('AVERAGE displays clean number', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: '10', displayValue: '10', type: 'number' },
      '0:1': { rawValue: '20', displayValue: '20', type: 'number' },
      '0:2': { rawValue: '30', displayValue: '30', type: 'number' },
      '1:0': { rawValue: '=AVERAGE(A1:C1)', displayValue: '', type: 'text' },
    });
    await spreadsheet.waitForCellText(1, 0, '20');
  });

  test('division producing repeating decimal displays cleanly', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: '10', displayValue: '10', type: 'number' },
      '0:1': { rawValue: '3', displayValue: '3', type: 'number' },
      '0:2': { rawValue: '=A1/B1', displayValue: '', type: 'text' },
    });
    const text = await spreadsheet.getCellText(0, 2);
    // Should be a reasonable representation, not trailing garbage
    expect(text.length).toBeLessThanOrEqual(20);
    expect(parseFloat(text)).toBeCloseTo(3.333333333333333, 10);
  });
});

test.describe('VLOOKUP correctness', () => {
  test('VLOOKUP finds value in lookup table', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'ID', displayValue: 'ID', type: 'text' },
      '0:1': { rawValue: 'Name', displayValue: 'Name', type: 'text' },
      '1:0': { rawValue: 'P001', displayValue: 'P001', type: 'text' },
      '1:1': { rawValue: 'Widget', displayValue: 'Widget', type: 'text' },
      '2:0': { rawValue: 'P002', displayValue: 'P002', type: 'text' },
      '2:1': { rawValue: 'Gadget', displayValue: 'Gadget', type: 'text' },
      '4:0': { rawValue: 'P002', displayValue: 'P002', type: 'text' },
      '4:1': { rawValue: '=VLOOKUP(A5, A2:B3, 2, FALSE)', displayValue: '', type: 'text' },
    });
    await spreadsheet.waitForCellText(4, 1, 'Gadget');
  });
});

test.describe('Grid scrollability', () => {
  test('grid with many rows has scrollable container', async ({ spreadsheet }) => {
    // Set many rows on a height-constrained grid
    const data: Record<string, { rawValue: string; displayValue: string; type: string }> = {};
    for (let r = 0; r < 100; r++) {
      data[`${r}:0`] = { rawValue: String(r), displayValue: String(r), type: 'number' };
    }
    await spreadsheet.setData(data);
    await spreadsheet.setProperty('rows', 100);

    // Check that the grid scroll container has scrollHeight > clientHeight
    const isScrollable = await spreadsheet.page.evaluate(() => {
      const el = document.querySelector('y11n-spreadsheet')!;
      const grid = el.shadowRoot!.querySelector('.ls-grid')!;
      return grid.scrollHeight > grid.clientHeight;
    });
    expect(isScrollable).toBe(true);
  });
});

test.describe('COUNTIF correctness', () => {
  test('COUNTIF counts only matching text values', async ({ spreadsheet }) => {
    await spreadsheet.setData({
      '0:0': { rawValue: 'OK', displayValue: 'OK', type: 'text' },
      '1:0': { rawValue: 'REORDER', displayValue: 'REORDER', type: 'text' },
      '2:0': { rawValue: 'OK', displayValue: 'OK', type: 'text' },
      '3:0': { rawValue: 'REORDER', displayValue: 'REORDER', type: 'text' },
      '4:0': { rawValue: 'OK', displayValue: 'OK', type: 'text' },
      '5:0': { rawValue: 'REORDER', displayValue: 'REORDER', type: 'text' },
      '7:0': { rawValue: '=COUNTIF(A1:A6,"REORDER")', displayValue: '', type: 'text' },
    });
    await spreadsheet.waitForCellText(7, 0, '3');
  });
});
