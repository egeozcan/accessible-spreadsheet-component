import type { Meta, StoryObj } from '@storybook/web-components-vite';
import { example } from './support/example.js';
import budgetTracker from './examples/budget-tracker.js';
import gradeBook from './examples/grade-book.js';
import inventory from './examples/inventory.js';
import pixelArt from './examples/pixel-art.js';

const meta: Meta = {
  title: 'Examples/Recipes',
  tags: ['autodocs'],
  parameters: {
    docs: {
      description: {
        component: `Small complete sheets that combine several features. Good starting points to copy from.`,
      },
    },
  },
};

export default meta;

export const BudgetTracker: StoryObj = {
  name: 'Budget tracker',
  render: budgetTracker,
  parameters: example('budget-tracker'),
};

export const GradeBook: StoryObj = {
  name: 'Grade book',
  render: gradeBook,
  parameters: example('grade-book', { height: 360 }),
};

export const Inventory: StoryObj = {
  name: 'Inventory',
  render: inventory,
  parameters: example('inventory', { height: 380 }),
};

export const PixelArt: StoryObj = {
  name: 'Pixel art',
  render: pixelArt,
  parameters: example('pixel-art', { height: 540 }),
};
