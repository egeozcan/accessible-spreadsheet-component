import { html } from 'lit';
import type { Meta, StoryObj } from '@storybook/web-components';
import '../src/y11n-spreadsheet.js';
import type { Y11nSpreadsheet } from '../src/y11n-spreadsheet.js';
import {
  gridFromRows,
  generateProductCatalog,
  generateExpenseData,
  generateLogicData,
  generateStringData,
  generateAbsoluteRefData,
} from './helpers.js';

const meta: Meta = {
  title: 'Components/y11n-spreadsheet',
  component: 'y11n-spreadsheet',
  tags: ['autodocs'],
  argTypes: {
    rows: {
      control: { type: 'number', min: 1, max: 1000 },
      description: 'Total number of rows in the grid',
    },
    cols: {
      control: { type: 'number', min: 1, max: 52 },
      description: 'Total number of columns in the grid',
    },
    readOnly: {
      control: 'boolean',
      description: 'When true, disables all editing',
    },
  },
  parameters: {
    docs: {
      description: {
        component: `
# &lt;y11n-spreadsheet&gt;

A lightweight, fully accessible, framework-agnostic spreadsheet web component built with **Lit 3.0** and **TypeScript**.

## Features

- **WAI-ARIA Grid** pattern with roving tabindex
- Excel-like **keyboard navigation** (Arrow keys, Tab, Enter, Escape)
- **Range selection** via Shift+Arrow or mouse drag
- **Formula engine** with built-in functions (SUM, AVERAGE, IF, etc.)
- **Clipboard** support (Ctrl+C, Ctrl+V, Ctrl+X) with TSV format
- **CSS custom properties** for full style customization
- **Virtual rendering** for large datasets

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| Arrow Keys | Navigate cells |
| Shift + Arrow | Extend selection |
| Enter / Double-click | Edit cell |
| Escape | Cancel edit / clear selection |
| Tab / Shift+Tab | Move right / left |
| Delete / Backspace | Clear selected cells |
| Ctrl+C / Ctrl+V / Ctrl+X | Copy / Paste / Cut |
| Ctrl+Z / Ctrl+Shift+Z | Undo / Redo |
| Ctrl+A | Select all |
| Type any character | Start editing with that character |
        `,
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

export default meta;

// ─── 1. Default (Empty Grid) ──────────────────────────

export const Default: StoryObj = {
  name: 'Default (Empty Grid)',
  args: {
    rows: 50,
    cols: 10,
    readOnly: false,
  },
  render: (args) => html`
    <y11n-spreadsheet
      .rows=${args.rows}
      .cols=${args.cols}
      ?read-only=${args.readOnly}
    ></y11n-spreadsheet>
  `,
  parameters: {
    docs: {
      description: {
        story: 'An empty spreadsheet grid. Click a cell and start typing, or press Enter to edit. Use arrow keys to navigate.',
      },
    },
  },
};

// ─── 2. Pre-populated Data ────────────────────────────

export const PrePopulatedData: StoryObj = {
  name: 'Pre-populated Data',
  render: () => {
    const data = gridFromRows([
      ['Name',     'Department', 'Salary',  'Start Date'],
      ['Alice',    'Engineering','95000',   '2021-03-15'],
      ['Bob',      'Design',    '82000',   '2020-07-01'],
      ['Charlie',  'Marketing', '78000',   '2022-01-10'],
      ['Diana',    'Engineering','105000',  '2019-11-20'],
      ['Eve',      'Product',   '91000',   '2021-08-05'],
    ]);

    return html`
      <y11n-spreadsheet
        .rows=${20}
        .cols=${10}
        .data=${data}
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: 'A spreadsheet pre-populated with employee data. The `.data` property accepts a `Map<string, CellData>` where keys are `"row:col"` format.',
      },
    },
  },
};

// ─── 3. Formulas & Calculations ───────────────────────

export const FormulasAndCalculations: StoryObj = {
  name: 'Formulas & Calculations',
  render: () => {
    const data = gridFromRows([
      ['Item',       'Qty',  'Price',   'Total'],
      ['Widget A',   '10',   '5.99',    '=B2*C2'],
      ['Widget B',   '25',   '3.50',    '=B3*C3'],
      ['Widget C',   '7',    '12.00',   '=B4*C4'],
      ['Gadget X',   '15',   '8.25',    '=B5*C5'],
      ['Gadget Y',   '3',    '45.00',   '=B6*C6'],
      ['',           '',     '',        ''],
      ['Subtotal',   '',     '',        '=SUM(D2:D6)'],
      ['Item Count', '',     '',        '=COUNT(B2:B6)'],
      ['Avg Price',  '',     '',        '=AVERAGE(C2:C6)'],
      ['Min Price',  '',     '',        '=MIN(C2:C6)'],
      ['Max Price',  '',     '',        '=MAX(C2:C6)'],
    ]);

    return html`
      <y11n-spreadsheet
        .rows=${20}
        .cols=${10}
        .data=${data}
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: `Demonstrates the built-in formula engine. Column D uses formulas like \`=B2*C2\` for line totals and \`=SUM(D2:D6)\` for the subtotal. Supported functions include:

- **SUM**, **AVERAGE**, **MIN**, **MAX**, **COUNT**, **COUNTA**
- **IF**(condition, true_val, false_val)
- **ABS**, **ROUND**(value, digits)
- **UPPER**, **LOWER**, **LEN**, **TRIM**, **CONCAT**
- Arithmetic: \`+\`, \`-\`, \`*\`, \`/\`
- Comparisons: \`=\`, \`<>\`, \`<\`, \`>\`, \`<=\`, \`>=\`
- String concatenation: \`&\`

Try editing a quantity or price cell — the totals update automatically.`,
      },
    },
  },
};

// ─── 4. Read-Only Mode ────────────────────────────────

export const ReadOnly: StoryObj = {
  name: 'Read-Only Mode',
  render: () => {
    const data = gridFromRows([
      ['Metric',              'Q1',     'Q2',     'Q3',     'Q4'],
      ['Revenue ($K)',        '1250',   '1480',   '1320',   '1690'],
      ['Costs ($K)',          '890',    '920',    '870',    '950'],
      ['Profit ($K)',         '=B2-B3', '=C2-C3', '=D2-D3', '=E2-E3'],
      ['Margin (%)',          '=B4/B2*100', '=C4/C2*100', '=D4/D2*100', '=E4/E2*100'],
    ]);

    return html`
      <y11n-spreadsheet
        .rows=${10}
        .cols=${8}
        .data=${data}
        read-only
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: 'With `read-only` attribute set, cells cannot be edited, deleted, or pasted into. Navigation and selection still work. Useful for displaying computed dashboards or reports.',
      },
    },
  },
};

// ─── 5. Custom Functions ──────────────────────────────

export const CustomFunctions: StoryObj = {
  name: 'Custom Functions',
  render: () => {
    const data = gridFromRows([
      ['Base Price', 'Tax Rate', 'With Tax',        'Discounted (20%)'],
      ['100',        '0.08',     '=TAX(A2, B2)',    '=DISCOUNT(A2, 20)'],
      ['250',        '0.10',     '=TAX(A3, B3)',    '=DISCOUNT(A3, 20)'],
      ['49.99',      '0.065',    '=TAX(A4, B4)',    '=DISCOUNT(A4, 20)'],
      ['1200',       '0.12',     '=TAX(A5, B5)',    '=DISCOUNT(A5, 20)'],
      ['',           '',         '',                ''],
      ['Total Tax',  '',         '=SUM(C2:C5)-SUM(A2:A5)', ''],
    ]);

    const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
    el.rows = 15;
    el.cols = 8;

    // Register custom functions before setting data
    el.registerFunction('TAX', (_ctx, price, rate) => {
      return Number(price) * (1 + Number(rate));
    });

    el.registerFunction('DISCOUNT', (_ctx, price, pct) => {
      return Number(price) * (1 - Number(pct) / 100);
    });

    el.setData(data);

    const container = document.createElement('div');
    container.style.height = '100%';
    container.appendChild(el);
    return container;
  },
  parameters: {
    docs: {
      description: {
        story: `Demonstrates extensibility via \`registerFunction()\`. Two custom functions are registered:

- **TAX(price, rate)**: Calculates price with tax (\`price * (1 + rate)\`)
- **DISCOUNT(price, pct)**: Applies a percentage discount (\`price * (1 - pct/100)\`)

These are called directly in formulas like \`=TAX(A2, B2)\`. You can register any custom function to extend the formula engine.`,
      },
    },
  },
};

// ─── 6. Compact Grid ─────────────────────────────────

export const CompactGrid: StoryObj = {
  name: 'Compact Grid (Small)',
  render: () => {
    const data = gridFromRows([
      ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
      ['8',   '7',   '9',   '6',   '8'],
      ['7',   '8',   '7',   '8',   '9'],
      ['9',   '6',   '8',   '7',   '7'],
    ]);

    return html`
      <y11n-spreadsheet
        .rows=${5}
        .cols=${5}
        .data=${data}
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: 'A minimal 5x5 grid — useful for small data entry widgets or quick lookups. The component adapts to any row/col count.',
      },
    },
  },
};

// ─── 7. Large Dataset ─────────────────────────────────

export const LargeDataset: StoryObj = {
  name: 'Large Dataset (1000 rows)',
  render: () => {
    const data: Map<string, { rawValue: string; displayValue: string; type: 'text' | 'number' }> = new Map();

    // Headers
    const headers = ['ID', 'First Name', 'Last Name', 'Email', 'Score', 'Status'];
    headers.forEach((h, c) => {
      data.set(`0:${c}`, { rawValue: h, displayValue: h, type: 'text' });
    });

    // Generate 1000 rows of data
    const firstNames = ['Alice', 'Bob', 'Charlie', 'Diana', 'Eve', 'Frank', 'Grace', 'Hank', 'Ivy', 'Jack'];
    const lastNames = ['Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis', 'Wilson', 'Moore'];
    const statuses = ['Active', 'Inactive', 'Pending'];

    for (let r = 1; r <= 1000; r++) {
      const first = firstNames[r % firstNames.length];
      const last = lastNames[r % lastNames.length];
      const score = String(Math.floor(50 + (r * 7 + 13) % 51));
      const status = statuses[r % statuses.length];

      data.set(`${r}:0`, { rawValue: String(r), displayValue: String(r), type: 'number' });
      data.set(`${r}:1`, { rawValue: first, displayValue: first, type: 'text' });
      data.set(`${r}:2`, { rawValue: last, displayValue: last, type: 'text' });
      data.set(`${r}:3`, { rawValue: `${first.toLowerCase()}.${last.toLowerCase()}@example.com`, displayValue: `${first.toLowerCase()}.${last.toLowerCase()}@example.com`, type: 'text' });
      data.set(`${r}:4`, { rawValue: score, displayValue: score, type: 'number' });
      data.set(`${r}:5`, { rawValue: status, displayValue: status, type: 'text' });
    }

    return html`
      <y11n-spreadsheet
        .rows=${1002}
        .cols=${8}
        .data=${data}
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: 'Tests virtual rendering with 1000+ rows of data. The component only renders visible rows plus a buffer, keeping performance smooth even with large datasets. Scroll down to see virtualization in action.',
      },
    },
  },
};

// ─── 8. Event Logging ─────────────────────────────────

export const EventLogging: StoryObj = {
  name: 'Event Logging',
  render: () => {
    const data = gridFromRows([
      ['Name',   'Value'],
      ['Alpha',  '100'],
      ['Beta',   '200'],
      ['Gamma',  '300'],
      ['Total',  '=SUM(B2:B4)'],
    ]);

    const wrapper = document.createElement('div');
    wrapper.style.display = 'flex';
    wrapper.style.flexDirection = 'column';
    wrapper.style.height = '100%';

    const spreadsheetContainer = document.createElement('div');
    spreadsheetContainer.style.flex = '1';
    spreadsheetContainer.style.minHeight = '0';

    const logContainer = document.createElement('div');
    logContainer.style.height = '150px';
    logContainer.style.overflow = 'auto';
    logContainer.style.borderTop = '2px solid #e0e0e0';
    logContainer.style.padding = '8px 12px';
    logContainer.style.fontFamily = 'monospace';
    logContainer.style.fontSize = '12px';
    logContainer.style.background = '#1e1e1e';
    logContainer.style.color = '#d4d4d4';

    const header = document.createElement('div');
    header.style.fontWeight = 'bold';
    header.style.marginBottom = '4px';
    header.style.color = '#569cd6';
    header.textContent = '// Event Log — edit cells to see events';
    logContainer.appendChild(header);

    const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
    el.rows = 10;
    el.cols = 6;
    el.setData(data);
    el.style.height = '100%';

    const log = (name: string, detail: unknown) => {
      const line = document.createElement('div');
      line.style.marginBottom = '2px';
      const ts = new Date().toLocaleTimeString();
      line.innerHTML = `<span style="color:#6a9955">${ts}</span> <span style="color:#ce9178">${name}</span> ${JSON.stringify(detail)}`;
      logContainer.appendChild(line);
      logContainer.scrollTop = logContainer.scrollHeight;
    };

    el.addEventListener('cell-change', (e: Event) => log('cell-change', (e as CustomEvent).detail));
    el.addEventListener('selection-change', (e: Event) => log('selection-change', (e as CustomEvent).detail));
    el.addEventListener('data-change', (e: Event) => log('data-change', (e as CustomEvent).detail));

    spreadsheetContainer.appendChild(el);
    wrapper.appendChild(spreadsheetContainer);
    wrapper.appendChild(logContainer);

    return wrapper;
  },
  parameters: {
    docs: {
      description: {
        story: `Shows all three events in real-time:

- **cell-change**: Fires when a single cell value is committed (after editing)
- **selection-change**: Fires when the selection/active cell moves
- **data-change**: Fires on bulk operations (paste, cut, delete of range)

The dark panel below the grid shows a live event log. Click cells, type values, and select ranges to see events.`,
      },
    },
  },
};

// ─── 9. Custom Styled (Dark Theme) ───────────────────

export const DarkTheme: StoryObj = {
  name: 'Dark Theme (CSS Variables)',
  render: () => {
    const data = gridFromRows([
      ['Planet',   'Diameter (km)', 'Distance from Sun (M km)', 'Moons'],
      ['Mercury',  '4879',          '57.9',                     '0'],
      ['Venus',    '12104',         '108.2',                    '0'],
      ['Earth',    '12756',         '149.6',                    '1'],
      ['Mars',     '6792',          '227.9',                    '2'],
      ['Jupiter',  '142984',        '778.6',                    '95'],
      ['Saturn',   '120536',        '1433.5',                   '146'],
      ['Uranus',   '51118',         '2872.5',                   '28'],
      ['Neptune',  '49528',         '4495.1',                   '16'],
    ]);

    return html`
      <y11n-spreadsheet
        .rows=${15}
        .cols=${8}
        .data=${data}
        style="
          --ls-header-bg: #2d2d2d;
          --ls-border-color: #444;
          --ls-font-family: 'JetBrains Mono', 'Fira Code', monospace;
          --ls-font-size: 13px;
          --ls-text-color: #e0e0e0;
          --ls-selection-border: 2px solid #bb86fc;
          --ls-selection-bg: rgba(187, 134, 252, 0.15);
          --ls-focus-ring: 2px solid #bb86fc;
          --ls-editor-bg: #1e1e1e;
          --ls-editor-shadow: 0 2px 8px rgba(0,0,0,0.5);
          --ls-cell-height: 32px;
          --ls-cell-width: 160px;
          background: #1e1e1e;
        "
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: `The component uses CSS custom properties for full theming. This example shows a dark theme using:

\`\`\`css
--ls-header-bg: #2d2d2d;
--ls-border-color: #444;
--ls-text-color: #e0e0e0;
--ls-selection-border: 2px solid #bb86fc;
--ls-selection-bg: rgba(187, 134, 252, 0.15);
--ls-editor-bg: #1e1e1e;
--ls-cell-height: 32px;
--ls-cell-width: 160px;
\`\`\`

All visual aspects can be customized without breaking Shadow DOM encapsulation.`,
      },
    },
  },
  decorators: [
    (story) => html`
      <div style="height: 500px; border: 1px solid #444; border-radius: 4px; overflow: hidden; background: #1e1e1e;">
        ${story()}
      </div>
    `,
  ],
};

// ─── 10. Pastel Theme ─────────────────────────────────

export const PastelTheme: StoryObj = {
  name: 'Pastel Theme',
  render: () => {
    const data = gridFromRows([
      ['Task',            'Assignee', 'Priority', 'Status'],
      ['Design mockups',  'Alice',    'High',     'Done'],
      ['API integration', 'Bob',      'High',     'In Progress'],
      ['Unit tests',      'Charlie',  'Medium',   'Todo'],
      ['Documentation',   'Diana',    'Low',      'Todo'],
      ['Code review',     'Eve',      'Medium',   'In Progress'],
    ]);

    return html`
      <y11n-spreadsheet
        .rows=${12}
        .cols=${6}
        .data=${data}
        style="
          --ls-header-bg: #fce4ec;
          --ls-border-color: #f8bbd0;
          --ls-text-color: #4a148c;
          --ls-selection-border: 2px solid #e91e63;
          --ls-selection-bg: rgba(233, 30, 99, 0.08);
          --ls-focus-ring: 2px solid #e91e63;
          --ls-editor-bg: #fff;
          --ls-font-family: 'Georgia', serif;
          --ls-cell-height: 34px;
          --ls-cell-width: 140px;
          background: #fce4ec;
        "
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: 'Another theme variant demonstrating the flexibility of CSS custom properties. This pastel/pink theme uses serif fonts and softer colors.',
      },
    },
  },
  decorators: [
    (story) => html`
      <div style="height: 500px; border: 1px solid #f8bbd0; border-radius: 8px; overflow: hidden; background: #fce4ec;">
        ${story()}
      </div>
    `,
  ],
};

// ─── 11. Budget Tracker with Formulas ─────────────────

export const BudgetTracker: StoryObj = {
  name: 'Budget Tracker',
  render: () => {
    const data = gridFromRows([
      ['Category',      'Budget',  'Actual',  'Difference',    'Within Budget?'],
      ['Housing',       '2000',    '1950',    '=B2-C2',        '=IF(D2>=0, "YES", "NO")'],
      ['Food',          '800',     '920',     '=B3-C3',        '=IF(D3>=0, "YES", "NO")'],
      ['Transport',     '400',     '380',     '=B4-C4',        '=IF(D4>=0, "YES", "NO")'],
      ['Entertainment', '300',     '450',     '=B5-C5',        '=IF(D5>=0, "YES", "NO")'],
      ['Utilities',     '250',     '230',     '=B6-C6',        '=IF(D6>=0, "YES", "NO")'],
      ['Healthcare',    '200',     '150',     '=B7-C7',        '=IF(D7>=0, "YES", "NO")'],
      ['Savings',       '500',     '500',     '=B8-C8',        '=IF(D8>=0, "YES", "NO")'],
      ['',              '',        '',        '',              ''],
      ['Totals',        '=SUM(B2:B8)', '=SUM(C2:C8)', '=SUM(D2:D8)', ''],
      ['Avg Category',  '=AVERAGE(B2:B8)', '=AVERAGE(C2:C8)', '', ''],
    ]);

    return html`
      <y11n-spreadsheet
        .rows=${18}
        .cols=${8}
        .data=${data}
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: `A realistic budget tracking example that showcases:

- **Arithmetic formulas**: \`=B2-C2\` for differences
- **IF function**: \`=IF(D2>=0, "YES", "NO")\` for conditional display
- **SUM and AVERAGE**: Aggregations in the totals row

Edit the "Actual" column values to see formulas recalculate in real-time.`,
      },
    },
  },
};

// ─── 12. Grade Book ───────────────────────────────────

export const GradeBook: StoryObj = {
  name: 'Grade Book',
  render: () => {
    const data = gridFromRows([
      ['Student',  'Test 1', 'Test 2', 'Test 3', 'Test 4', 'Average',           'Highest',          'Lowest'],
      ['Alice',    '92',     '88',     '95',     '90',     '=AVERAGE(B2:E2)',   '=MAX(B2:E2)',      '=MIN(B2:E2)'],
      ['Bob',      '78',     '82',     '75',     '88',     '=AVERAGE(B3:E3)',   '=MAX(B3:E3)',      '=MIN(B3:E3)'],
      ['Charlie',  '95',     '91',     '98',     '94',     '=AVERAGE(B4:E4)',   '=MAX(B4:E4)',      '=MIN(B4:E4)'],
      ['Diana',    '85',     '79',     '82',     '86',     '=AVERAGE(B5:E5)',   '=MAX(B5:E5)',      '=MIN(B5:E5)'],
      ['Eve',      '88',     '92',     '87',     '91',     '=AVERAGE(B6:E6)',   '=MAX(B6:E6)',      '=MIN(B6:E6)'],
      ['',         '',       '',       '',       '',       '',                  '',                 ''],
      ['Class Avg','=AVERAGE(B2:B6)', '=AVERAGE(C2:C6)', '=AVERAGE(D2:D6)', '=AVERAGE(E2:E6)', '=AVERAGE(F2:F6)', '', ''],
      ['Highest',  '=MAX(B2:B6)',    '=MAX(C2:C6)',    '=MAX(D2:D6)',    '=MAX(E2:E6)',    '',                '', ''],
    ]);

    return html`
      <y11n-spreadsheet
        .rows=${15}
        .cols=${10}
        .data=${data}
        style="--ls-cell-width: 90px;"
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: `A grade book demonstrating cross-row and cross-column formulas:

- Each student row calculates **AVERAGE**, **MAX**, and **MIN** across their test scores
- Bottom rows calculate per-test class averages and highest scores
- Demonstrates that formulas can reference other formula cells (chained evaluation)`,
      },
    },
  },
};

// ─── 13. String Formulas ──────────────────────────────

export const StringFormulas: StoryObj = {
  name: 'String Formulas',
  render: () => {
    const data = gridFromRows([
      ['First Name', 'Last Name', 'Full Name',           'Uppercase',           'Length',          'Trimmed'],
      ['  Alice ',   'Smith',     '=CONCAT(TRIM(A2), " ", B2)', '=UPPER(C2)',    '=LEN(C2)',        '=TRIM(A2)'],
      ['Bob',        'Johnson',   '=CONCAT(A3, " ", B3)', '=UPPER(C3)',         '=LEN(C3)',        '=TRIM(A3)'],
      ['  Charlie ', 'Williams',  '=CONCAT(TRIM(A4), " ", B4)', '=LOWER(C4)',   '=LEN(C4)',        '=TRIM(A4)'],
      ['Diana',      'Brown',     '=A5 & " " & B5',      '=UPPER(C5)',         '=LEN(C5)',        '=TRIM(A5)'],
    ]);

    return html`
      <y11n-spreadsheet
        .rows=${10}
        .cols=${8}
        .data=${data}
        style="--ls-cell-width: 130px;"
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: `Demonstrates string manipulation formulas:

- **CONCAT(a, b, ...)**: Joins values
- **UPPER(text)** / **LOWER(text)**: Case conversion
- **LEN(text)**: String length
- **TRIM(text)**: Removes leading/trailing whitespace
- **& operator**: String concatenation (\`=A5 & " " & B5\`)

Note how "Alice" and "Charlie" have leading/trailing spaces — TRIM removes them.`,
      },
    },
  },
};

// ─── 14. Narrow Cells ─────────────────────────────────

export const WideCells: StoryObj = {
  name: 'Wide Cells (Custom Dimensions)',
  render: () => {
    const data = gridFromRows([
      ['Description',                        'Amount'],
      ['Monthly subscription fee',           '29.99'],
      ['Annual support contract',            '299.00'],
      ['One-time setup charge',              '150.00'],
      ['Total',                              '=SUM(B2:B4)'],
    ]);

    return html`
      <y11n-spreadsheet
        .rows=${10}
        .cols=${4}
        .data=${data}
        style="
          --ls-cell-width: 250px;
          --ls-cell-height: 36px;
          --ls-font-size: 15px;
        "
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: `Demonstrates using CSS custom properties to adjust cell dimensions:

\`\`\`css
--ls-cell-width: 250px;
--ls-cell-height: 36px;
--ls-font-size: 15px;
\`\`\`

Wider cells are useful when columns contain longer text like descriptions or URLs.`,
      },
    },
  },
};

// ─── 15. Conditional Logic with IF ────────────────────

export const ConditionalLogic: StoryObj = {
  name: 'Conditional Logic (IF)',
  render: () => {
    const data = gridFromRows([
      ['Product',     'Stock', 'Reorder Point', 'Status',                            'Order Qty'],
      ['Bolts M6',    '150',   '100',           '=IF(B2>C2, "OK", "REORDER")',       '=IF(B2>C2, 0, C2-B2+50)'],
      ['Nuts M6',     '80',    '100',            '=IF(B3>C3, "OK", "REORDER")',       '=IF(B3>C3, 0, C3-B3+50)'],
      ['Washers M6',  '200',   '100',            '=IF(B4>C4, "OK", "REORDER")',       '=IF(B4>C4, 0, C4-B4+50)'],
      ['Screws M4',   '45',    '100',            '=IF(B5>C5, "OK", "REORDER")',       '=IF(B5>C5, 0, C5-B5+50)'],
      ['Rivets',      '300',   '200',            '=IF(B6>C6, "OK", "REORDER")',       '=IF(B6>C6, 0, C6-B6+50)'],
      ['Springs',     '90',    '150',            '=IF(B7>C7, "OK", "REORDER")',       '=IF(B7>C7, 0, C7-B7+50)'],
      ['',            '',      '',               '',                                  ''],
      ['Items to reorder', '', '',               '=COUNTA(D2:D7)-COUNT(D2:D7)',       ''],
    ]);

    return html`
      <y11n-spreadsheet
        .rows=${15}
        .cols=${8}
        .data=${data}
        style="--ls-cell-width: 120px;"
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: `An inventory management example showcasing the **IF** function for conditional logic:

- **Status column**: \`=IF(B2>C2, "OK", "REORDER")\` — shows "REORDER" when stock is below the reorder point
- **Order Qty column**: \`=IF(B2>C2, 0, C2-B2+50)\` — calculates how many to order, or 0 if stock is sufficient

Try changing stock values to see statuses update dynamically.`,
      },
    },
  },
};

// ─── 16. API Methods Demo ─────────────────────────────

export const APIMethods: StoryObj = {
  name: 'API Methods (getData / setData)',
  render: () => {
    const wrapper = document.createElement('div');
    wrapper.style.display = 'flex';
    wrapper.style.flexDirection = 'column';
    wrapper.style.height = '100%';
    wrapper.style.gap = '0';

    const toolbar = document.createElement('div');
    toolbar.style.padding = '8px 12px';
    toolbar.style.display = 'flex';
    toolbar.style.gap = '8px';
    toolbar.style.borderBottom = '1px solid #e0e0e0';
    toolbar.style.background = '#fafafa';
    toolbar.style.flexShrink = '0';

    const btnStyle = 'padding: 6px 14px; border: 1px solid #ccc; border-radius: 4px; background: #fff; cursor: pointer; font-size: 13px;';

    const btnGetData = document.createElement('button');
    btnGetData.textContent = 'getData()';
    btnGetData.setAttribute('style', btnStyle);

    const btnSetSample = document.createElement('button');
    btnSetSample.textContent = 'setData(sample)';
    btnSetSample.setAttribute('style', btnStyle);

    const btnClear = document.createElement('button');
    btnClear.textContent = 'setData(empty)';
    btnClear.setAttribute('style', btnStyle);

    const output = document.createElement('pre');
    output.style.cssText = 'margin: 0; padding: 4px 12px; font-size: 11px; color: #666; flex-shrink: 0; max-height: 60px; overflow: auto; background: #f9f9f9; border-bottom: 1px solid #e0e0e0;';
    output.textContent = '// Click getData() to inspect grid state';

    toolbar.appendChild(btnGetData);
    toolbar.appendChild(btnSetSample);
    toolbar.appendChild(btnClear);

    const spreadsheetContainer = document.createElement('div');
    spreadsheetContainer.style.flex = '1';
    spreadsheetContainer.style.minHeight = '0';

    const el = document.createElement('y11n-spreadsheet') as Y11nSpreadsheet;
    el.rows = 10;
    el.cols = 6;
    el.style.height = '100%';

    const sampleData = gridFromRows([
      ['City', 'Population'],
      ['Tokyo', '37400000'],
      ['Delhi', '30290000'],
      ['Shanghai', '27058000'],
    ]);
    el.setData(sampleData);

    btnGetData.addEventListener('click', () => {
      const d = el.getData();
      const entries: string[] = [];
      d.forEach((v, k) => entries.push(`  "${k}": { raw: "${v.rawValue}", display: "${v.displayValue}" }`));
      output.textContent = `GridData (${d.size} cells):\n${entries.join('\n')}`;
    });

    btnSetSample.addEventListener('click', () => {
      const newData = gridFromRows([
        ['Fruit', 'Count', 'Price', 'Total'],
        ['Apple', '10', '1.20', '=B2*C2'],
        ['Banana', '25', '0.50', '=B3*C3'],
        ['Cherry', '100', '0.10', '=B4*C4'],
      ]);
      el.setData(newData);
      output.textContent = '// setData() called with new fruit data';
    });

    btnClear.addEventListener('click', () => {
      el.setData(new Map());
      output.textContent = '// setData() called with empty Map';
    });

    spreadsheetContainer.appendChild(el);
    wrapper.appendChild(toolbar);
    wrapper.appendChild(output);
    wrapper.appendChild(spreadsheetContainer);

    return wrapper;
  },
  parameters: {
    docs: {
      description: {
        story: `Demonstrates the public API methods:

- **getData()**: Returns the current \`GridData\` map — useful for serialization or external processing
- **setData(data)**: Replaces all grid data — useful for loading saved state or switching datasets

Click the buttons above the grid to interact with the API. The output panel shows the result of \`getData()\`.`,
      },
    },
  },
};

// ─── 17. Lookup Functions ───────────────────────────────

export const LookupFunctions: StoryObj = {
  name: 'Lookup Functions',
  render: () => {
    const catalog = generateProductCatalog();
    const rows = [
      ...catalog,
      ['', '', '', '', ''],
      ['Lookup ID', 'Found Name',              'Found Price',             'Found Category'],
      ['P003',      '=VLOOKUP(A11, A2:E9, 2)', '=VLOOKUP(A11, A2:E9, 4)', '=VLOOKUP(A11, A2:E9, 3)'],
      ['P006',      '=VLOOKUP(A12, A2:E9, 2)', '=VLOOKUP(A12, A2:E9, 4)', '=VLOOKUP(A12, A2:E9, 3)'],
      ['',          '',                         '',                        ''],
      ['INDEX/MATCH Demo', '',                  '',                        ''],
      ['Row 3 Col 2',      '=INDEX(A2:E9, 3, 2)', '',                    ''],
      ['MATCH "Office"',   '=MATCH("Office", C2:C9, 0)', '',             ''],
    ];

    const data = gridFromRows(rows);

    return html`
      <y11n-spreadsheet
        .rows=${25}
        .cols=${8}
        .data=${data}
        style="--ls-cell-width: 140px;"
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: `Demonstrates lookup functions with a product catalog:

- **VLOOKUP(lookup, range, col_index)**: Searches the first column of a range and returns a value from a specified column
- **INDEX(range, row, col)**: Returns the value at a given row/column position in a range
- **MATCH(value, range, type)**: Returns the position of a value in a range

Try changing the Lookup ID values (A11, A12) to look up different products.

*Note: These functions require Agent 2's formula additions to render correctly.*`,
      },
    },
  },
};

// ─── 18. Conditional Aggregation ─────────────────────────

export const ConditionalAggregation: StoryObj = {
  name: 'Conditional Aggregation',
  render: () => {
    const expenses = generateExpenseData();
    const rows = [
      ...expenses,
      ['', '', '', ''],
      ['Summary', '', '', ''],
      ['Category',    'Total',                                  'Count',                                   'Average'],
      ['Technology',  '=SUMIF(C2:C11, "Technology", D2:D11)',   '=COUNTIF(C2:C11, "Technology")',           '=AVERAGEIF(C2:C11, "Technology", D2:D11)'],
      ['Office',      '=SUMIF(C2:C11, "Office", D2:D11)',       '=COUNTIF(C2:C11, "Office")',               '=AVERAGEIF(C2:C11, "Office", D2:D11)'],
      ['Food',        '=SUMIF(C2:C11, "Food", D2:D11)',         '=COUNTIF(C2:C11, "Food")',                 '=AVERAGEIF(C2:C11, "Food", D2:D11)'],
      ['', '', '', ''],
      ['Grand Total', '=SUM(D2:D11)', '=COUNTA(D2:D11)', '=AVERAGE(D2:D11)'],
    ];

    const data = gridFromRows(rows);

    return html`
      <y11n-spreadsheet
        .rows=${25}
        .cols=${6}
        .data=${data}
        style="--ls-cell-width: 140px;"
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: `Demonstrates conditional aggregation with categorized expense data:

- **SUMIF(range, criteria, sum_range)**: Sums values where the criteria matches
- **COUNTIF(range, criteria)**: Counts cells matching a criteria
- **AVERAGEIF(range, criteria, avg_range)**: Averages values where the criteria matches

The summary table below the data shows totals, counts, and averages per category.

*Note: SUMIF/COUNTIF/AVERAGEIF require Agent 2's formula additions to render correctly.*`,
      },
    },
  },
};

// ─── 19. Logic Functions ─────────────────────────────────

export const LogicFunctions: StoryObj = {
  name: 'Logic Functions',
  render: () => {
    const logic = generateLogicData();
    const rows = [
      // Original columns A-E, then computed columns F-I
      [...logic[0], 'Pass All (>=70)?',                        'Pass Any (>=70)?',                     'Good Attendance?',                      'Safe Division'],
      [...logic[1], '=AND(B2>=70, C2>=70, D2>=70)',            '=OR(B2>=70, C2>=70, D2>=70)',          '=NOT(E2<90)',                           '=IFERROR(B2/0, "N/A")'],
      [...logic[2], '=AND(B3>=70, C3>=70, D3>=70)',            '=OR(B3>=70, C3>=70, D3>=70)',          '=NOT(E3<90)',                           '=IFERROR(B3/0, "N/A")'],
      [...logic[3], '=AND(B4>=70, C4>=70, D4>=70)',            '=OR(B4>=70, C4>=70, D4>=70)',          '=NOT(E4<90)',                           '=IFERROR(B4/0, "N/A")'],
      [...logic[4], '=AND(B5>=70, C5>=70, D5>=70)',            '=OR(B5>=70, C5>=70, D5>=70)',          '=NOT(E5<90)',                           '=IFERROR(B5/0, "N/A")'],
      [...logic[5], '=AND(B6>=70, C6>=70, D6>=70)',            '=OR(B6>=70, C6>=70, D6>=70)',          '=NOT(E6<90)',                           '=IFERROR(B6/0, "N/A")'],
      [...logic[6], '=AND(B7>=70, C7>=70, D7>=70)',            '=OR(B7>=70, C7>=70, D7>=70)',          '=NOT(E7<90)',                           '=IFERROR(B7/0, "N/A")'],
    ];

    const data = gridFromRows(rows);

    return html`
      <y11n-spreadsheet
        .rows=${15}
        .cols=${12}
        .data=${data}
        style="--ls-cell-width: 130px;"
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: `Demonstrates logical functions with student grade data:

- **AND(cond1, cond2, ...)**: Returns TRUE only if ALL conditions are true
- **OR(cond1, cond2, ...)**: Returns TRUE if ANY condition is true
- **NOT(condition)**: Inverts a boolean value
- **IFERROR(value, fallback)**: Returns fallback if value produces an error

Column F checks if a student passes all subjects (>=70). Column G checks if they pass at least one. Column H checks attendance. Column I shows IFERROR catching a division-by-zero.

*Note: AND/OR/NOT/IFERROR require Agent 2's formula additions to render correctly.*`,
      },
    },
  },
};

// ─── 20. Extended String Functions ──────────────────────

export const ExtendedStringFunctions: StoryObj = {
  name: 'Extended String Functions',
  render: () => {
    const strings = generateStringData();
    const rows = [
      // Original columns A-C, then computed columns D-I
      [...strings[0], 'First Name',        'Last Name',             'Domain',                                   'Area Code',           'Cleaned Phone',                          'Name Swap'],
      [...strings[1], '=LEFT(A2, FIND(" ", A2)-1)',  '=MID(A2, FIND(" ", A2)+1, 100)',  '=MID(B2, FIND("@", B2)+1, 100)',  '=MID(C2, 2, 3)',  '=SUBSTITUTE(SUBSTITUTE(C2, "(", ""), ")", "")',  '=SUBSTITUTE(A2, LEFT(A2, FIND(" ", A2)-1), MID(A2, FIND(" ", A2)+1, 100))'],
      [...strings[2], '=LEFT(A3, FIND(" ", A3)-1)',  '=MID(A3, FIND(" ", A3)+1, 100)',  '=MID(B3, FIND("@", B3)+1, 100)',  '=MID(C3, 2, 3)',  '=SUBSTITUTE(SUBSTITUTE(C3, "(", ""), ")", "")',  '=SUBSTITUTE(A3, LEFT(A3, FIND(" ", A3)-1), MID(A3, FIND(" ", A3)+1, 100))'],
      [...strings[3], '=LEFT(A4, FIND(" ", A4)-1)',  '=MID(A4, FIND(" ", A4)+1, 100)',  '=MID(B4, FIND("@", B4)+1, 100)',  '=MID(C4, 2, 3)',  '=SUBSTITUTE(SUBSTITUTE(C4, "(", ""), ")", "")',  '=SUBSTITUTE(A4, LEFT(A4, FIND(" ", A4)-1), MID(A4, FIND(" ", A4)+1, 100))'],
      [...strings[4], '=LEFT(A5, FIND(" ", A5)-1)',  '=MID(A5, FIND(" ", A5)+1, 100)',  '=MID(B5, FIND("@", B5)+1, 100)',  '=MID(C5, 2, 3)',  '=SUBSTITUTE(SUBSTITUTE(C5, "(", ""), ")", "")',  '=SUBSTITUTE(A5, LEFT(A5, FIND(" ", A5)-1), MID(A5, FIND(" ", A5)+1, 100))'],
      [...strings[5], '=LEFT(A6, FIND(" ", A6)-1)',  '=MID(A6, FIND(" ", A6)+1, 100)',  '=MID(B6, FIND("@", B6)+1, 100)',  '=MID(C6, 2, 3)',  '=SUBSTITUTE(SUBSTITUTE(C6, "(", ""), ")", "")',  '=SUBSTITUTE(A6, LEFT(A6, FIND(" ", A6)-1), MID(A6, FIND(" ", A6)+1, 100))'],
    ];

    const data = gridFromRows(rows);

    return html`
      <y11n-spreadsheet
        .rows=${12}
        .cols=${12}
        .data=${data}
        style="--ls-cell-width: 160px;"
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: `Demonstrates extended string functions for text manipulation:

- **LEFT(text, num)**: Extracts characters from the start
- **RIGHT(text, num)**: Extracts characters from the end
- **MID(text, start, length)**: Extracts characters from the middle
- **FIND(search, text)**: Returns the position of a substring
- **SUBSTITUTE(text, old, new)**: Replaces occurrences of a substring

Examples include extracting first/last names from full names, parsing email domains, extracting area codes from phone numbers, and swapping name parts.

*Note: LEFT/RIGHT/MID/FIND/SUBSTITUTE require Agent 2's formula additions to render correctly.*`,
      },
    },
  },
};

// ─── 21. Absolute References ─────────────────────────────

export const AbsoluteReferences: StoryObj = {
  name: 'Absolute References',
  render: () => {
    const base = generateAbsoluteRefData();
    const rows = [
      ...base,
      ['', '', '', '', ''],
      ['Tax Rate:', '0.08', '', '', ''],
      ['', '', '', '', ''],
      ['', 'Jan', 'Feb', 'Mar', 'Apr'],
      ['Revenue After Tax', '=B2*(1-$B$5)', '=C2*(1-$B$5)', '=D2*(1-$B$5)', '=E2*(1-$B$5)'],
      ['Profit',            '=B2-B3',       '=C2-C3',       '=D2-D3',       '=E2-E3'],
      ['Profit After Tax',  '=B9*(1-$B$5)', '=C9*(1-$B$5)', '=D9*(1-$B$5)', '=E9*(1-$B$5)'],
      ['',                  '',             '',             '',             ''],
      ['Pct of Jan Sales',  '=B2/B2',       '=C2/$B$2',    '=D2/$B$2',    '=E2/$B$2'],
    ];

    const data = gridFromRows(rows);

    return html`
      <y11n-spreadsheet
        .rows=${18}
        .cols=${8}
        .data=${data}
        style="--ls-cell-width: 130px;"
      ></y11n-spreadsheet>
    `;
  },
  parameters: {
    docs: {
      description: {
        story: `Demonstrates absolute vs relative cell references:

- **Relative reference** (\`B2\`): Adjusts when copied to other cells
- **Absolute reference** (\`$B$5\`): Always refers to the same cell regardless of position

In this example:
- Row 8 uses \`$B$5\` (absolute) to reference the tax rate in every column
- Row 12 uses \`$B$2\` (absolute) to always compare against January sales
- Row 9 uses relative references for column-specific profit calculations

Try changing the Tax Rate in B5 to see all tax-dependent formulas update.

*Note: Absolute references (\`$\` notation) require Agent 2's formula additions to render correctly.*`,
      },
    },
  },
};
