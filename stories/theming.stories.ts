import type { Meta, StoryObj } from '@storybook/web-components-vite';
import { example } from './support/example.js';
import darkTheme from './examples/dark-theme.js';
import pastelTheme from './examples/pastel-theme.js';
import cellSize from './examples/cell-size.js';

const meta: Meta = {
  title: 'Examples/Theming',
  tags: ['autodocs'],
  parameters: {
    docs: {
      description: {
        component: `The grid is styled with CSS custom properties. Every property and its default is listed under **Reference / Theming**.`,
      },
    },
  },
};

export default meta;

export const DarkTheme: StoryObj = {
  name: 'Dark theme',
  render: darkTheme,
  parameters: example('dark-theme'),
};

export const PastelTheme: StoryObj = {
  name: 'Pastel theme',
  render: pastelTheme,
  parameters: example('pastel-theme', { height: 360 }),
};

export const CellSize: StoryObj = {
  name: 'Cell size',
  render: cellSize,
  parameters: example('cell-size', { height: 320 }),
};
