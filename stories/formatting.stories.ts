import type { Meta, StoryObj } from '@storybook/web-components-vite';
import { example } from './support/example.js';
import cellFormats from './examples/cell-formats.js';
import formatApi from './examples/format-api.js';
import conditionalFormatting from './examples/conditional-formatting.js';
import formatToolbar from './examples/format-toolbar.js';
import toolbarStates from './examples/toolbar-states.js';

const meta: Meta = {
  title: 'Examples/Formatting',
  tags: ['autodocs'],
  parameters: {
    docs: {
      description: {
        component: `How cells look: formats stored with the cells, formats that follow the values (conditional formatting), and the optional toolbar element.`,
      },
    },
  },
};

export default meta;

export const CellFormats: StoryObj = {
  name: 'Cell formats',
  render: cellFormats,
  parameters: example('cell-formats', { height: 300 }),
};

export const FormatApi: StoryObj = {
  name: 'Format API',
  render: formatApi,
  parameters: example('format-api', { height: 360 }),
};

export const ConditionalFormatting: StoryObj = {
  name: 'Conditional formatting',
  render: conditionalFormatting,
  parameters: example('conditional-formatting', { height: 360 }),
};

export const FormatToolbar: StoryObj = {
  name: 'Format toolbar',
  render: formatToolbar,
  parameters: example('format-toolbar'),
};

export const ToolbarStates: StoryObj = {
  name: 'Toolbar states',
  render: toolbarStates,
  parameters: example('toolbar-states', { height: 220 }),
};
