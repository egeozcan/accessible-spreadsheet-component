import type { Meta, StoryObj } from '@storybook/web-components-vite';
import { example } from './support/example.js';
import validation from './examples/validation.js';
import findAndFill from './examples/find-and-fill.js';
import sorting from './examples/sorting.js';
import filter from './examples/filter.js';

const meta: Meta = {
  title: 'Examples/Editing and Tables',
  tags: ['autodocs'],
  parameters: {
    docs: {
      description: {
        component: `Controlling what users can enter, and the tools they get for working with a table of data: find, fill, sort and filter. Most of this needs no code at all; the examples show the properties, methods and events that let your application take part.`,
      },
    },
  },
};

export default meta;

export const Validation: StoryObj = {
  name: 'Validation and locked cells',
  render: validation,
  parameters: example('validation', { height: 320 }),
};

export const FindAndFill: StoryObj = {
  name: 'Find, fill and row insertion',
  render: findAndFill,
  parameters: example('find-and-fill'),
};

export const Sorting: StoryObj = {
  name: 'Sorting',
  render: sorting,
  parameters: example('sorting', { height: 360 }),
};

export const Filter: StoryObj = {
  name: 'Filter',
  render: filter,
  parameters: example('filter'),
};
