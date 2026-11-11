import type { Meta, StoryObj } from '@storybook/web-components-vite';
import { example } from './support/example.js';
import emptyGrid from './examples/empty-grid.js';
import loadingData from './examples/loading-data.js';
import formulas from './examples/formulas.js';
import readOnly from './examples/read-only.js';
import events from './examples/events.js';
import savingAndRestoring from './examples/saving-and-restoring.js';

const meta: Meta = {
  title: 'Examples/Basics',
  tags: ['autodocs'],
  parameters: {
    docs: {
      description: {
        component: `The first things to know: how to put a grid on the page, how to give it data, and how to hear about changes. Read these in order.

Every example on this page is one file. **Show code** under an example shows that file, and so does the **Code** panel when you open the example from the sidebar.`,
      },
    },
  },
};

export default meta;

export const EmptyGrid: StoryObj = {
  name: 'Empty grid',
  render: emptyGrid,
  parameters: example('empty-grid'),
};

export const LoadingData: StoryObj = {
  name: 'Loading data',
  render: loadingData,
  parameters: example('loading-data'),
};

export const Formulas: StoryObj = {
  name: 'Formulas',
  render: formulas,
  parameters: example('formulas'),
};

export const ReadOnly: StoryObj = {
  name: 'Read-only',
  render: readOnly,
  parameters: example('read-only', { height: 300 }),
};

export const Events: StoryObj = {
  name: 'Events',
  render: events,
  parameters: example('events'),
};

export const SavingAndRestoring: StoryObj = {
  name: 'Saving and restoring',
  render: savingAndRestoring,
  parameters: example('saving-and-restoring'),
};
