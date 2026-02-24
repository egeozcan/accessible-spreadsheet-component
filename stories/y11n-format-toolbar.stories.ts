import { html } from 'lit';
import type { Meta, StoryObj } from '@storybook/web-components';
import '../src/y11n-spreadsheet.js';
import '../src/components/y11n-format-toolbar.js';
import type { Y11nSpreadsheet } from '../src/y11n-spreadsheet.js';
import type { Y11nFormatToolbar, FormatActionDetail } from '../src/components/y11n-format-toolbar.js';
import { computeSelectionFormat } from '../src/components/y11n-format-toolbar.js';
import type { CellFormat } from '../src/types.js';
import { cellKey } from '../src/types.js';
import { gridFromRows } from './helpers.js';

const meta: Meta = {
  title: 'Components/y11n-format-toolbar',
  component: 'y11n-format-toolbar',
  tags: ['autodocs'],
  parameters: {
    docs: {
      description: {
        component: `
# &lt;y11n-format-toolbar&gt;

A companion toolbar component for cell-level formatting in \`<y11n-spreadsheet>\`.

## Features

- **Bold / Italic / Underline / Strikethrough** toggle buttons
- **Text color** and **background color** pickers
- **Text alignment** (left / center / right)
- **Font size** selector
- **Clear formatting** button
- Full **WAI-ARIA toolbar** pattern with roving tabindex
- All controls disabled in read-only mode

## Usage

\`\`\`html
<y11n-format-toolbar
  .bold=\${format.bold}
  .italic=\${format.italic}
  @format-action=\${handleFormatAction}
></y11n-format-toolbar>
<y11n-spreadsheet
  @selection-change=\${updateToolbar}
></y11n-spreadsheet>
\`\`\`
        `,
      },
    },
  },
};

export default meta;

// ─── 1. Default Toolbar ──────────────────────────────────

export const DefaultToolbar: StoryObj = {
  name: 'Default Toolbar',
  render: () => html`
    <y11n-format-toolbar></y11n-format-toolbar>
  `,
  parameters: {
    docs: {
      description: {
        story: 'The toolbar in its default state. All toggle buttons are off, colors are default black/white, alignment is left, font size is 13.',
      },
    },
  },
};

// ─── 2. Active Formatting State ──────────────────────────

export const ActiveFormattingState: StoryObj = {
  name: 'Active Formatting State',
  render: () => html`
    <y11n-format-toolbar
      .bold=${true}
      .italic=${true}
      .textColor=${'#d32f2f'}
      .backgroundColor=${'#fff9c4'}
      .textAlign=${'center'}
      .fontSize=${16}
    ></y11n-format-toolbar>
  `,
  parameters: {
    docs: {
      description: {
        story: 'Toolbar reflecting an actively formatted cell: bold, italic, red text, yellow background, centered, 16px font.',
      },
    },
  },
};

// ─── 3. Read-Only Toolbar ────────────────────────────────

export const ReadOnlyToolbar: StoryObj = {
  name: 'Read-Only Toolbar',
  render: () => html`
    <y11n-format-toolbar read-only .bold=${true}></y11n-format-toolbar>
  `,
  parameters: {
    docs: {
      description: {
        story: 'When `read-only` is set, all buttons and inputs are disabled. The toolbar still reflects the current state but cannot be interacted with.',
      },
    },
  },
};

// ─── 4. Toolbar + Spreadsheet Integration ────────────────

export const ToolbarIntegration: StoryObj = {
  name: 'Toolbar + Spreadsheet Integration',
  render: () => {
    const wrapper = document.createElement('div');
    wrapper.style.display = 'flex';
    wrapper.style.flexDirection = 'column';
    wrapper.style.height = '100%';

    const toolbar = document.createElement('y11n-format-toolbar') as Y11nFormatToolbar;
    const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
    el.rows = 20;
    el.cols = 10;
    el.style.flex = '1';
    el.style.minHeight = '0';

    const data = gridFromRows([
      ['Name', 'Department', 'Salary'],
      ['Alice', 'Engineering', '95000'],
      ['Bob', 'Design', '82000'],
      ['Charlie', 'Marketing', '78000'],
    ]);
    el.setData(data);

    // Sync toolbar on selection or format change
    const syncToolbar = () => {
      const range = (el as any)._selection.range;
      const fmt = computeSelectionFormat(
        (id) => (el as any)._internalData.get(id),
        range
      );
      toolbar.bold = fmt.bold ?? false;
      toolbar.italic = fmt.italic ?? false;
      toolbar.underline = fmt.underline ?? false;
      toolbar.strikethrough = fmt.strikethrough ?? false;
      toolbar.textColor = fmt.textColor ?? '#000000';
      toolbar.backgroundColor = fmt.backgroundColor ?? '#ffffff';
      toolbar.textAlign = fmt.textAlign ?? 'left';
      toolbar.fontSize = fmt.fontSize ?? 13;
    };

    el.addEventListener('selection-change', syncToolbar);
    el.addEventListener('format-change', syncToolbar);

    // Handle toolbar actions
    toolbar.addEventListener('format-action', ((e: CustomEvent<FormatActionDetail>) => {
      const { action, value } = e.detail;
      switch (action) {
        case 'bold':
        case 'italic':
        case 'underline':
        case 'strikethrough':
          el.toggleFormat(action);
          break;
        case 'textColor':
        case 'backgroundColor':
        case 'textAlign':
        case 'fontSize':
          el.setRangeFormat((el as any)._selection.range, { [action]: value });
          break;
        case 'clearFormat':
          el.clearRangeFormat((el as any)._selection.range);
          break;
      }
      syncToolbar();
    }) as EventListener);

    wrapper.appendChild(toolbar);
    wrapper.appendChild(el);
    return wrapper;
  },
  parameters: {
    docs: {
      description: {
        story: `A fully working integration. Select cells, then use the toolbar to format them. Try:

- Select cells → click **B** to bold
- Use color pickers to change text/background color
- **Ctrl+B**, **Ctrl+I**, **Ctrl+U** keyboard shortcuts work too
- **Ctrl+Z** to undo formatting changes`,
      },
    },
  },
  decorators: [
    (story) => html`
      <div style="height: 500px; border: 1px solid #ccc; border-radius: 4px; overflow: hidden;">
        ${story()}
      </div>
    `,
  ],
};

// ─── 5. Pre-formatted Data ───────────────────────────────

export const PreformattedData: StoryObj = {
  name: 'Pre-formatted Data',
  render: () => {
    const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
    el.rows = 10;
    el.cols = 6;
    el.style.height = '100%';

    const data = gridFromRows([
      ['Name', 'Score', 'Grade', 'Status'],
      ['Alice', '95', 'A', 'Pass'],
      ['Bob', '72', 'C', 'Pass'],
      ['Charlie', '45', 'F', 'Fail'],
      ['Diana', '88', 'B', 'Pass'],
    ]);
    el.setData(data);

    // Apply formatting programmatically
    // Bold headers
    for (let c = 0; c < 4; c++) {
      el.setCellFormat(cellKey(0, c), { bold: true, backgroundColor: '#e3f2fd' });
    }
    // Green for pass, red for fail
    el.setCellFormat(cellKey(1, 3), { textColor: '#2e7d32', bold: true });
    el.setCellFormat(cellKey(2, 3), { textColor: '#2e7d32', bold: true });
    el.setCellFormat(cellKey(3, 3), { textColor: '#c62828', bold: true });
    el.setCellFormat(cellKey(4, 3), { textColor: '#2e7d32', bold: true });
    // Italic scores
    el.setCellFormat(cellKey(1, 1), { italic: true });
    el.setCellFormat(cellKey(2, 1), { italic: true });
    el.setCellFormat(cellKey(3, 1), { italic: true, textColor: '#c62828' });
    el.setCellFormat(cellKey(4, 1), { italic: true });
    // Right-align numbers
    for (let r = 1; r <= 4; r++) {
      el.setCellFormat(cellKey(r, 1), {
        ...el.getCellFormat(cellKey(r, 1)),
        textAlign: 'right',
      });
    }

    const container = document.createElement('div');
    container.style.height = '100%';
    container.appendChild(el);
    return container;
  },
  parameters: {
    docs: {
      description: {
        story: 'A spreadsheet with various formatting applied via the `setCellFormat()` API: bold headers with blue background, green/red status colors, italic scores, and right-aligned numbers.',
      },
    },
  },
  decorators: [
    (story) => html`
      <div style="height: 350px; border: 1px solid #ccc; border-radius: 4px; overflow: hidden;">
        ${story()}
      </div>
    `,
  ],
};

// ─── 6. Formatted Budget Tracker ─────────────────────────

export const FormattedBudgetTracker: StoryObj = {
  name: 'Formatted Budget Tracker',
  render: () => {
    const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
    el.rows = 12;
    el.cols = 6;
    el.style.height = '100%';

    const data = gridFromRows([
      ['Category', 'Budget', 'Actual', 'Difference', 'Status'],
      ['Housing', '2000', '1950', '=B2-C2', '=IF(D2>=0,"Under","Over")'],
      ['Food', '800', '920', '=B3-C3', '=IF(D3>=0,"Under","Over")'],
      ['Transport', '400', '380', '=B4-C4', '=IF(D4>=0,"Under","Over")'],
      ['Entertainment', '300', '450', '=B5-C5', '=IF(D5>=0,"Under","Over")'],
      ['Utilities', '250', '230', '=B6-C6', '=IF(D6>=0,"Under","Over")'],
      ['', '', '', '', ''],
      ['Totals', '=SUM(B2:B6)', '=SUM(C2:C6)', '=SUM(D2:D6)', ''],
    ]);
    el.setData(data);

    // Header formatting
    for (let c = 0; c < 5; c++) {
      el.setCellFormat(cellKey(0, c), {
        bold: true,
        backgroundColor: '#1565c0',
        textColor: '#ffffff',
      });
    }
    // Right-align numbers
    for (let r = 1; r <= 7; r++) {
      for (let c = 1; c <= 3; c++) {
        el.setCellFormat(cellKey(r, c), { textAlign: 'right' });
      }
    }
    // Conditional coloring on difference/status
    const diffCells = [
      { r: 1, positive: true },
      { r: 2, positive: false },
      { r: 3, positive: true },
      { r: 4, positive: false },
      { r: 5, positive: true },
    ];
    for (const { r, positive } of diffCells) {
      const color = positive ? '#2e7d32' : '#c62828';
      el.setCellFormat(cellKey(r, 3), {
        ...el.getCellFormat(cellKey(r, 3)),
        textColor: color,
        bold: true,
      });
      el.setCellFormat(cellKey(r, 4), { textColor: color, bold: true });
    }
    // Totals row
    el.setCellFormat(cellKey(7, 0), { bold: true });
    for (let c = 1; c <= 3; c++) {
      el.setCellFormat(cellKey(7, c), {
        ...el.getCellFormat(cellKey(7, c)),
        bold: true,
        backgroundColor: '#e8eaf6',
      });
    }

    const container = document.createElement('div');
    container.style.height = '100%';
    container.appendChild(el);
    return container;
  },
  parameters: {
    docs: {
      description: {
        story: 'A budget tracker with: bold white-on-blue headers, right-aligned numbers, green text for under-budget items, red for over-budget, and a highlighted totals row.',
      },
    },
  },
  decorators: [
    (story) => html`
      <div style="height: 400px; border: 1px solid #ccc; border-radius: 4px; overflow: hidden;">
        ${story()}
      </div>
    `,
  ],
};

// ─── 7. Color Palette Demo ───────────────────────────────

export const ColorPaletteDemo: StoryObj = {
  name: 'Color Palette Demo',
  render: () => {
    const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
    el.rows = 10;
    el.cols = 10;
    el.style.height = '100%';
    el.setAttribute('style', el.getAttribute('style') + '; --ls-cell-width: 60px; --ls-cell-height: 32px; --ls-border-color: transparent;');

    el.setData(new Map());

    // Generate rainbow gradient
    const colors = [
      ['#f44336', '#e91e63', '#9c27b0', '#673ab7', '#3f51b5', '#2196f3', '#03a9f4', '#00bcd4', '#009688', '#4caf50'],
      ['#ef5350', '#ec407a', '#ab47bc', '#7e57c2', '#5c6bc0', '#42a5f5', '#29b6f6', '#26c6da', '#26a69a', '#66bb6a'],
      ['#e57373', '#f06292', '#ba68c8', '#9575cd', '#7986cb', '#64b5f6', '#4fc3f7', '#4dd0e1', '#4db6ac', '#81c784'],
      ['#ef9a9a', '#f48fb1', '#ce93d8', '#b39ddb', '#9fa8da', '#90caf9', '#81d4fa', '#80deea', '#80cbc4', '#a5d6a7'],
      ['#ffcdd2', '#f8bbd0', '#e1bee7', '#d1c4e9', '#c5cae9', '#bbdefb', '#b3e5fc', '#b2ebf2', '#b2dfdb', '#c8e6c9'],
      ['#fff9c4', '#fff59d', '#fff176', '#ffee58', '#ffeb3b', '#fdd835', '#fbc02d', '#f9a825', '#f57f17', '#e65100'],
    ];

    for (let r = 0; r < colors.length; r++) {
      for (let c = 0; c < colors[r].length; c++) {
        el.setCellFormat(cellKey(r, c), { backgroundColor: colors[r][c] });
      }
    }

    const container = document.createElement('div');
    container.style.height = '100%';
    container.appendChild(el);
    return container;
  },
  parameters: {
    docs: {
      description: {
        story: 'A rainbow/gradient pattern using `backgroundColor` on cells. Demonstrates that cell formatting is purely visual and does not affect cell data.',
      },
    },
  },
  decorators: [
    (story) => html`
      <div style="height: 300px; border: 1px solid #ccc; border-radius: 4px; overflow: hidden;">
        ${story()}
      </div>
    `,
  ],
};

// ─── 8. Duck Pixel Art (Animated) ────────────────────────

export const DuckPixelArt: StoryObj = {
  name: 'Duck Pixel Art (Animated)',
  render: () => {
    const wrapper = document.createElement('div');
    wrapper.style.display = 'flex';
    wrapper.style.height = '100%';
    wrapper.style.gap = '0';

    // Spreadsheet side
    const gridContainer = document.createElement('div');
    gridContainer.style.flex = '1';
    gridContainer.style.minWidth = '0';

    const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
    el.rows = 16;
    el.cols = 16;
    el.readOnly = true;
    el.style.height = '100%';
    el.setAttribute('style', el.getAttribute('style') + '; --ls-cell-width: 28px; --ls-cell-height: 28px; --ls-border-color: transparent; --ls-header-bg: transparent;');
    el.setData(new Map());
    gridContainer.appendChild(el);

    // Code panel side
    const codePanel = document.createElement('div');
    codePanel.style.width = '320px';
    codePanel.style.overflow = 'auto';
    codePanel.style.borderLeft = '2px solid #e0e0e0';
    codePanel.style.background = '#1e1e1e';
    codePanel.style.color = '#d4d4d4';
    codePanel.style.fontFamily = "'JetBrains Mono', 'Fira Code', monospace";
    codePanel.style.fontSize = '11px';
    codePanel.style.padding = '12px';
    codePanel.style.whiteSpace = 'pre';
    codePanel.style.lineHeight = '1.5';

    const header = document.createElement('div');
    header.style.color = '#569cd6';
    header.style.fontWeight = 'bold';
    header.style.marginBottom = '8px';
    header.textContent = '// setCellFormat API calls';
    codePanel.appendChild(header);

    const codeContent = document.createElement('div');
    codePanel.appendChild(codeContent);

    wrapper.appendChild(gridContainer);
    wrapper.appendChild(codePanel);

    // Duck pixel art: _ = sky, Y = yellow, O = orange, B = black, W = water
    const SKY = '#E3F2FD';
    const YELLOW = '#FDD835';
    const ORANGE = '#FF8F00';
    const BLACK = '#333333';
    const WATER = '#81D4FA';

    const duckPixels = [
      '________________',
      '________________',
      '____YY__________',
      '___YYYY_________',
      '___YBYY_________',
      '___YYYY_________',
      '____YOOO________',
      '____YY__________',
      '___YYYY_________',
      '__YYYYYY________',
      '_YYYYYYYY_______',
      '_YYYYYYYY_______',
      'WWYWWWWYWWWWWWWW',
      'WWWWWWWWWWWWWWWW',
      'WWWWWWWWWWWWWWWW',
      'WWWWWWWWWWWWWWWW',
    ];

    const charToColor: Record<string, string> = {
      '_': SKY,
      'Y': YELLOW,
      'O': ORANGE,
      'B': BLACK,
      'W': WATER,
    };

    // Animate row by row
    let currentRow = 0;
    const interval = setInterval(() => {
      if (currentRow >= 16) {
        clearInterval(interval);
        return;
      }

      const row = duckPixels[currentRow];
      const codeLines: string[] = [];

      for (let c = 0; c < 16; c++) {
        const ch = row[c];
        const color = charToColor[ch] ?? SKY;
        el.setCellFormat(cellKey(currentRow, c), { backgroundColor: color });
        if (ch !== '_' || currentRow >= 12) {
          codeLines.push(
            `<span style="color:#ce9178">setCellFormat</span>(<span style="color:#b5cea8">"${currentRow}:${c}"</span>, { bg: <span style="color:#6a9955">"${color}"</span> })`
          );
        }
      }

      if (codeLines.length > 0) {
        const rowLabel = document.createElement('div');
        rowLabel.style.color = '#608b4e';
        rowLabel.style.marginTop = '4px';
        rowLabel.textContent = `// Row ${currentRow}`;
        codeContent.appendChild(rowLabel);

        for (const line of codeLines) {
          const lineEl = document.createElement('div');
          lineEl.innerHTML = line;
          codeContent.appendChild(lineEl);
        }

        codePanel.scrollTop = codePanel.scrollHeight;
      }

      currentRow++;
    }, 150);

    return wrapper;
  },
  parameters: {
    docs: {
      description: {
        story: `A 16x16 duck pixel art drawn cell-by-cell using \`setCellFormat()\` with \`backgroundColor\`. Rows animate in at 150ms intervals (~2.4s total). The code panel on the right shows the API calls in real-time.

Colors used:
- **#FDD835** (yellow body)
- **#FF8F00** (orange beak)
- **#333** (eye)
- **#81D4FA** (water)
- **#E3F2FD** (sky)`,
      },
    },
  },
  decorators: [
    (story) => html`
      <div style="height: 520px; border: 1px solid #ccc; border-radius: 4px; overflow: hidden;">
        ${story()}
      </div>
    `,
  ],
};

// ─── 9. Programmatic Format API ──────────────────────────

export const ProgrammaticFormatAPI: StoryObj = {
  name: 'Programmatic Format API',
  render: () => {
    const wrapper = document.createElement('div');
    wrapper.style.display = 'flex';
    wrapper.style.flexDirection = 'column';
    wrapper.style.height = '100%';

    // Toolbar with API buttons
    const apiBar = document.createElement('div');
    apiBar.style.padding = '8px 12px';
    apiBar.style.display = 'flex';
    apiBar.style.gap = '8px';
    apiBar.style.flexWrap = 'wrap';
    apiBar.style.borderBottom = '1px solid #e0e0e0';
    apiBar.style.background = '#fafafa';
    apiBar.style.flexShrink = '0';

    const btnStyle = 'padding: 6px 12px; border: 1px solid #ccc; border-radius: 4px; background: #fff; cursor: pointer; font-size: 12px;';

    const output = document.createElement('pre');
    output.style.cssText = 'margin: 0; padding: 4px 12px; font-size: 11px; color: #666; flex-shrink: 0; max-height: 80px; overflow: auto; background: #f9f9f9; border-bottom: 1px solid #e0e0e0;';
    output.textContent = '// Click buttons to test the formatting API';

    const gridContainer = document.createElement('div');
    gridContainer.style.flex = '1';
    gridContainer.style.minHeight = '0';

    const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
    el.rows = 10;
    el.cols = 8;
    el.style.height = '100%';

    const data = gridFromRows([
      ['Name', 'Value', 'Notes'],
      ['Alpha', '100', 'First item'],
      ['Beta', '200', 'Second item'],
      ['Gamma', '300', 'Third item'],
    ]);
    el.setData(data);

    // API buttons
    const actions: Array<{ label: string; fn: () => void }> = [
      {
        label: 'Bold A1:C1',
        fn: () => {
          el.setRangeFormat({ start: { row: 0, col: 0 }, end: { row: 0, col: 2 } }, { bold: true, backgroundColor: '#e3f2fd' });
          output.textContent = 'setRangeFormat({0,0}:{0,2}, { bold: true, backgroundColor: "#e3f2fd" })';
        },
      },
      {
        label: 'Red text B2',
        fn: () => {
          el.setCellFormat('1:1', { textColor: '#c62828', bold: true });
          output.textContent = 'setCellFormat("1:1", { textColor: "#c62828", bold: true })';
        },
      },
      {
        label: 'Italic C2:C4',
        fn: () => {
          el.setRangeFormat({ start: { row: 1, col: 2 }, end: { row: 3, col: 2 } }, { italic: true });
          output.textContent = 'setRangeFormat({1,2}:{3,2}, { italic: true })';
        },
      },
      {
        label: 'getCellFormat(A1)',
        fn: () => {
          const fmt = el.getCellFormat('0:0');
          output.textContent = `getCellFormat("0:0") → ${JSON.stringify(fmt ?? 'undefined')}`;
        },
      },
      {
        label: 'Clear all format',
        fn: () => {
          el.clearRangeFormat({ start: { row: 0, col: 0 }, end: { row: 9, col: 7 } });
          output.textContent = 'clearRangeFormat({0,0}:{9,7}) — all formatting cleared';
        },
      },
      {
        label: 'Undo (Ctrl+Z)',
        fn: () => {
          (el as any)._undo();
          output.textContent = '// Undo last format change';
        },
      },
    ];

    for (const action of actions) {
      const btn = document.createElement('button');
      btn.textContent = action.label;
      btn.setAttribute('style', btnStyle);
      btn.addEventListener('click', action.fn);
      apiBar.appendChild(btn);
    }

    gridContainer.appendChild(el);
    wrapper.appendChild(apiBar);
    wrapper.appendChild(output);
    wrapper.appendChild(gridContainer);

    return wrapper;
  },
  parameters: {
    docs: {
      description: {
        story: `Interactive demo of the formatting API methods:

- **setCellFormat(cellId, format)** — applies format to a single cell
- **setRangeFormat(range, format)** — applies format to a range
- **getCellFormat(cellId)** — returns the current format
- **clearRangeFormat(range)** — removes all formatting
- **Undo** — format changes are fully undoable`,
      },
    },
  },
  decorators: [
    (story) => html`
      <div style="height: 450px; border: 1px solid #ccc; border-radius: 4px; overflow: hidden;">
        ${story()}
      </div>
    `,
  ],
};
