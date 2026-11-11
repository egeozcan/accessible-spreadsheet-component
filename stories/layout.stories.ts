import type { Meta, StoryObj } from '@storybook/web-components-vite';
import { example } from './support/example.js';
import columnWidths from './examples/column-widths.js';
import frozenPanes from './examples/frozen-panes.js';
import largeDataset from './examples/large-dataset.js';

const meta: Meta = {
  title: 'Examples/Layout',
  tags: ['autodocs'],
  parameters: {
    docs: {
      description: {
        component: `Column widths, frozen rows and columns, and what happens with a lot of data.`,
      },
    },
  },
};

export default meta;

export const ColumnWidths: StoryObj = {
  name: 'Column widths',
  render: columnWidths,
  parameters: example('column-widths', { height: 300 }),
};

export const FrozenPanes: StoryObj = {
  name: 'Frozen panes',
  render: frozenPanes,
  parameters: example('frozen-panes'),
};

export const LargeDataset: StoryObj = {
  name: 'Large dataset',
  render: largeDataset,
  parameters: example('large-dataset'),
};
