import { test, expect } from './fixtures';

test.describe('Virtual Scrolling', () => {
  test('only renders a subset of rows (not all 50)', async ({ spreadsheet }) => {
    // The grid should not render all 50 rows at once
    const totalRenderedRows = await spreadsheet.shadow('[role="row"]').count();

    // Header row + visible rows + buffer (not all 50 data rows)
    // Exact count depends on viewport, but should be less than 51
    expect(totalRenderedRows).toBeLessThan(51);
    expect(totalRenderedRows).toBeGreaterThan(1); // at least header + some rows
  });

  test('scrolling reveals new rows', async ({ spreadsheet }) => {
    const initialVisibleRows = await spreadsheet.shadow('[data-row]').evaluateAll(
      (elements) => elements.map((el) => parseInt((el as HTMLElement).dataset.row!))
    );
    const initialMaxRow = Math.max(...initialVisibleRows);

    // Scroll well down the grid and wait for virtualization to catch up.
    await spreadsheet.grid.evaluate((grid, top) => {
      grid.scrollTop = top as number;
    }, 1000);
    await spreadsheet.page.waitForTimeout(200);

    const visibleRowsAfterScroll = await spreadsheet.shadow('[data-row]').evaluateAll(
      (elements) => elements.map((el) => parseInt((el as HTMLElement).dataset.row!))
    );
    const maxRowAfterScroll = Math.max(...visibleRowsAfterScroll);

    expect(maxRowAfterScroll).toBeGreaterThan(initialMaxRow);
    expect(maxRowAfterScroll).toBeGreaterThan(30);
  });

  test('scrolled-to cells are accessible and clickable', async ({ spreadsheet }) => {
    // Scroll to make a lower row visible
    await spreadsheet.grid.evaluate((grid) => {
      grid.scrollTop = 400;
    });
    await spreadsheet.page.waitForTimeout(200);

    // Find a row that's now rendered
    const visibleRows = await spreadsheet.shadow('[data-row]').evaluateAll(
      (elements) => elements.map((el) => parseInt((el as HTMLElement).dataset.row!))
    );

    const highRow = Math.max(...visibleRows);

    // Click on the high row cell
    await spreadsheet.clickCell(highRow, 0);
    await expect(spreadsheet.cell(highRow, 0)).toHaveClass(/active-cell/);
  });

  test('navigating with arrow keys scrolls the view', async ({ spreadsheet }) => {
    // Click first cell
    await spreadsheet.clickCell(0, 0);

    // Press arrow down many times to go past visible area
    for (let i = 0; i < 30; i++) {
      const activeRow = await spreadsheet.page.evaluate(() => {
        const sheet = document.querySelector('y11n-spreadsheet');
        const active = sheet?.shadowRoot?.querySelector('.active-cell') as HTMLElement;
        return active ? parseInt(active.dataset.row!) : -1;
      });

      const nextCell = spreadsheet.cell(activeRow, 0);
      await nextCell.press('ArrowDown');
    }

    // The active cell should be at row 30
    await expect(spreadsheet.cell(30, 0)).toHaveClass(/active-cell/);
  });

  test('ensureCellVisible scrolls cell fully into viewport below sticky header', async ({ spreadsheet }) => {
    // Programmatically move the selection to a far-off row via the component's
    // internal API, which calls _ensureCellVisible but does NOT trigger the
    // browser's native focus-scroll. This isolates the scroll calculation.
    await spreadsheet.clickCell(0, 0);

    const visibility = await spreadsheet.page.evaluate(async () => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      // Use the public keyboard-like API: set selection then request update
      // We access the selection manager to jump directly
      const sel = sheet._selection;
      sel.moveTo(30, 0);
      // Trigger the internal ensureCellVisible (same as keyboard nav does)
      sheet._ensureCellVisible(30, 0);
      sheet.requestUpdate();
      await sheet.updateComplete;

      const grid = sheet.shadowRoot!.querySelector('[role="grid"]') as HTMLElement;
      const activeCell = sheet.shadowRoot!.querySelector('[data-row="30"][data-col="0"]') as HTMLElement;
      if (!activeCell || !grid) return { visible: false, reason: 'missing element' };

      const gridRect = grid.getBoundingClientRect();
      const cellRect = activeCell.getBoundingClientRect();

      // Find the sticky header height
      const header = sheet.shadowRoot!.querySelector('.ls-col-header') as HTMLElement;
      const headerHeight = header ? header.getBoundingClientRect().height : 0;

      const dataTop = gridRect.top + headerHeight;
      const dataBottom = gridRect.bottom;

      return {
        visible: cellRect.top >= dataTop - 1 && cellRect.bottom <= dataBottom + 1,
        cellTop: Math.round(cellRect.top),
        cellBottom: Math.round(cellRect.bottom),
        dataTop: Math.round(dataTop),
        dataBottom: Math.round(dataBottom),
      };
    });

    expect(visibility.visible).toBe(true);
  });

  test('large grid renders efficiently with spacers', async ({ spreadsheet }) => {
    // Set the grid to have many rows
    await spreadsheet.setProperty('rows', 500);
    await spreadsheet.page.waitForTimeout(200);

    // Count distinct rendered rows using data-row attributes
    const renderedRowNumbers = await spreadsheet.shadow('[data-row]').evaluateAll(
      (elements) => {
        const rows = new Set(elements.map((el) => (el as HTMLElement).dataset.row));
        return rows.size;
      }
    );
    // Should render far fewer rows than 500 due to virtualization
    expect(renderedRowNumbers).toBeLessThan(100);
    expect(renderedRowNumbers).toBeGreaterThan(0);

    // aria-rowcount should reflect the full size
    await expect(spreadsheet.grid).toHaveAttribute('aria-rowcount', '501');
  });
});

test.describe('Cell sizes in non-px units', () => {
  // 6.3rem x 1.9rem = 100.8px x 30.4px at the default 16px root font size.
  // Fractional on purpose: rounding errors would add up over many columns.
  const WIDTH = 6.3 * 16;
  const HEIGHT = 1.9 * 16;

  test.beforeEach(async ({ spreadsheet }) => {
    await spreadsheet.host.evaluate((el) => {
      el.style.setProperty('--ls-cell-width', '6.3rem');
      el.style.setProperty('--ls-cell-height', '1.9rem');
    });
    // The component re-measures through a ResizeObserver; wait for the re-render
    await expect
      .poll(async () => (await spreadsheet.cell(0, 0).boundingBox())?.width)
      .toBeCloseTo(WIDTH, 1);
  });

  test('scrolled-to region renders the cells under the viewport', async ({ spreadsheet }) => {
    await spreadsheet.grid.evaluate((grid) => {
      grid.scrollLeft = 1200;
      grid.scrollTop = 800;
    });

    await expect
      .poll(() =>
        spreadsheet.page.evaluate(() => {
          const sheet = document.querySelector('y11n-spreadsheet')!;
          const grid = sheet.shadowRoot!.querySelector('[role="grid"]')!;
          const rect = grid.getBoundingClientRect();
          const hit = sheet.shadowRoot!.elementFromPoint(
            rect.left + rect.width / 2,
            rect.top + rect.height / 2
          );
          const cell = hit?.closest<HTMLElement>('[data-row][data-col]');
          return cell ? `${cell.dataset.row}:${cell.dataset.col}` : null;
        })
      )
      .toBe(
        await spreadsheet.grid.evaluate(
          (grid, size) => {
            const rowHeader = 50;
            const row = Math.floor((grid.scrollTop + grid.clientHeight / 2 - size.height) / size.height);
            const col = Math.floor((grid.scrollLeft + grid.clientWidth / 2 - rowHeader) / size.width);
            return `${row}:${col}`;
          },
          { width: WIDTH, height: HEIGHT }
        )
      );
  });

  test('keyboard navigation keeps the active cell inside the data viewport', async ({ spreadsheet }) => {
    await spreadsheet.clickCell(0, 0);

    const result = await spreadsheet.page.evaluate(async () => {
      const sheet = document.querySelector('y11n-spreadsheet') as any;
      // Jump directly, bypassing the browser's native focus scrolling, so only
      // the component's own scroll math is exercised. Row 20 keeps the target
      // scroll position within the rendered rows, whose headers are the
      // grid's scroll-snap points.
      sheet._selection.moveTo(20, 20);
      sheet._ensureCellVisible(20, 20);
      sheet.requestUpdate();
      await sheet.updateComplete;

      const root = sheet.shadowRoot!;
      const grid = root.querySelector('[role="grid"]') as HTMLElement;
      const cell = root.querySelector('[data-row="20"][data-col="20"]') as HTMLElement;
      const header = root.querySelector('.ls-col-header') as HTMLElement;
      const rowHeader = root.querySelector('[role="rowheader"]') as HTMLElement;
      const g = grid.getBoundingClientRect();
      const c = cell.getBoundingClientRect();
      return {
        top: c.top - (g.top + header.getBoundingClientRect().height),
        bottom: g.top + grid.clientHeight - c.bottom,
        left: c.left - (g.left + rowHeader.getBoundingClientRect().width),
        right: g.left + grid.clientWidth - c.right,
      };
    });

    for (const gap of Object.values(result)) expect(gap).toBeGreaterThanOrEqual(-1);
    // The view scrolled just far enough: the cell sits at the bottom-right edge
    expect(Math.abs(result.bottom)).toBeLessThan(2);
    expect(Math.abs(result.right)).toBeLessThan(2);
  });

  test('inline editor lines up with the cell being edited', async ({ spreadsheet }) => {
    await spreadsheet.grid.evaluate((grid) => {
      grid.scrollLeft = 700;
      grid.scrollTop = 500;
    });
    await spreadsheet.page.waitForTimeout(200);

    const target = await spreadsheet.page.evaluate(() => {
      const sheet = document.querySelector('y11n-spreadsheet')!;
      const grid = sheet.shadowRoot!.querySelector('[role="grid"]')!;
      const rect = grid.getBoundingClientRect();
      const cell = sheet
        .shadowRoot!.elementFromPoint(rect.left + rect.width / 2, rect.top + rect.height / 2)
        ?.closest<HTMLElement>('[data-row][data-col]');
      return { row: Number(cell!.dataset.row), col: Number(cell!.dataset.col) };
    });

    await spreadsheet.dblClickCell(target.row, target.col);
    await expect(spreadsheet.editor).toBeVisible();

    const cellBox = (await spreadsheet.cell(target.row, target.col).boundingBox())!;
    const editorBox = (await spreadsheet.editor.boundingBox())!;
    expect(Math.abs(editorBox.x - cellBox.x)).toBeLessThan(1);
    expect(Math.abs(editorBox.y - cellBox.y)).toBeLessThan(1);
    expect(editorBox.width).toBeCloseTo(cellBox.width, 0);
    expect(editorBox.height).toBeCloseTo(cellBox.height, 0);
  });

  test('re-measures when the size variable changes at runtime', async ({ spreadsheet }) => {
    await spreadsheet.host.evaluate((el) => el.style.setProperty('--ls-cell-width', '9rem'));
    await expect
      .poll(async () => (await spreadsheet.cell(0, 0).boundingBox())?.width)
      .toBeCloseTo(144, 1);

    // PageDown moves by the number of rows that fit, which depends on the height
    await spreadsheet.host.evaluate((el) => el.style.setProperty('--ls-cell-height', '3rem'));
    await expect
      .poll(async () => (await spreadsheet.cell(0, 0).boundingBox())?.height)
      .toBeCloseTo(48, 1);
    await spreadsheet.clickCell(0, 0);
    const pageRows = await spreadsheet.grid.evaluate(
      (grid) => Math.max(1, Math.floor(grid.clientHeight / 48) - 1)
    );
    await spreadsheet.cell(0, 0).press('PageDown');
    await expect(spreadsheet.cell(pageRows, 0)).toHaveClass(/active-cell/);
  });
});
