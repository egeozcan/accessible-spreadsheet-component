import type { Meta, StoryObj } from '@storybook/web-components-vite';
import { example } from './support/example.js';
import cellFormats from './examples/cell-formats.js';
import formatApi from './examples/format-api.js';
import conditionalFormatting from './examples/conditional-formatting.js';

const meta: Meta = {
  title: 'Examples/Formatting',
  tags: ['autodocs'],
  parameters: {
    docs: {
      description: {
        component: `How cells look: formats stored with the cells, and formats that follow the values (conditional formatting). The toolbar in these examples is the optional \`<y11n-toolbar>\` element (see Toolbar).`,
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
