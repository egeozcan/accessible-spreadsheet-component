import { html } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import type { Meta, StoryObj } from '@storybook/web-components-vite';
import 'y11n-spreadsheet';

interface PlaygroundArgs {
  rows: number;
  cols: number;
  readOnly: boolean;
  frozenRows: number;
  frozenCols: number;
  sheetTabs: boolean;
  hideStatusBar: boolean;
  rawInput: boolean;
  gridLabel: string;
}

const meta: Meta<PlaygroundArgs> = {
  title: 'Examples/Playground',
  parameters: {
    controls: { disable: false, expanded: true },
    docs: {
      // The snippet is the markup as rendered, without the frame around it
      source: {
        excludeDecorators: true,
        transform: (code: string) => code.replace(/=""/g, ''),
      },
    },
  },
  argTypes: {
    rows: { control: { type: 'number', min: 1, max: 5000 }, description: 'Number of rows. Attribute: `rows`' },
    cols: { control: { type: 'number', min: 1, max: 200 }, description: 'Number of columns. Attribute: `cols`' },
    readOnly: { control: 'boolean', description: 'Turns off editing. Attribute: `read-only`' },
    frozenRows: {
      control: { type: 'number', min: 0, max: 10 },
      description: 'Rows at the top that stay in view while scrolling. Attribute: `frozen-rows`',
    },
    frozenCols: {
      control: { type: 'number', min: 0, max: 10 },
      description: 'Columns at the left that stay in view while scrolling. Attribute: `frozen-cols`',
    },
    sheetTabs: {
      control: 'boolean',
      description: 'Shows the sheet tabs even with one sheet, so users can add more. Attribute: `sheet-tabs`',
    },
    hideStatusBar: { control: 'boolean', description: 'Hides the status bar under the grid. Attribute: `hide-status-bar`' },
    rawInput: {
      control: 'boolean',
      description: 'Stores typed values as typed, without recognising `15%`, `$1,200` or dates. Attribute: `raw-input`',
    },
    gridLabel: { control: 'text', description: 'Accessible name of the grid. Attribute: `grid-label`' },
  },
};

export default meta;

/**
 * Every setting that can be written as an HTML attribute. Change them in the
 * **Controls** panel; the **Code** panel shows the markup for the current
 * values.
 */
export const Playground: StoryObj<PlaygroundArgs> = {
  args: {
    rows: 50,
    cols: 12,
    readOnly: false,
    frozenRows: 0,
    frozenCols: 0,
    sheetTabs: false,
    hideStatusBar: false,
    rawInput: false,
    gridLabel: 'Spreadsheet',
  },
  render: (args) => html`
    <y11n-spreadsheet
      rows=${args.rows}
      cols=${args.cols}
      ?read-only=${args.readOnly}
      frozen-rows=${ifDefined(args.frozenRows || undefined)}
      frozen-cols=${ifDefined(args.frozenCols || undefined)}
      ?sheet-tabs=${args.sheetTabs}
      ?hide-status-bar=${args.hideStatusBar}
      ?raw-input=${args.rawInput}
      grid-label=${ifDefined(args.gridLabel === 'Spreadsheet' ? undefined : args.gridLabel)}
    ></y11n-spreadsheet>
  `,
};
