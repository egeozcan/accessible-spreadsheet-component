import type { Meta, StoryObj } from '@storybook/web-components-vite';
import { example } from './support/example.js';
import toolbar from './examples/toolbar.js';
import toolbarItems from './examples/toolbar-items.js';
import toolbarEvents from './examples/toolbar-events.js';
import toolbarStates from './examples/toolbar-states.js';

const meta: Meta = {
  title: 'Examples/Toolbar',
  tags: ['autodocs'],
  parameters: {
    docs: {
      description: {
        component: `The optional toolbar element, \`<y11n-toolbar>\`: buttons for what the grid can do. Point it at a grid and it works; choose which controls it shows, and step in before any of them acts. Many examples in the other groups use it with a few controls each.`,
      },
    },
  },
};

export default meta;

export const Toolbar: StoryObj = {
  name: 'Toolbar',
  render: toolbar,
  parameters: example('toolbar'),
};

export const ToolbarItems: StoryObj = {
  name: 'Choosing the controls',
  render: toolbarItems,
  parameters: example('toolbar-items'),
};

export const ToolbarEvents: StoryObj = {
  name: 'Toolbar events',
  render: toolbarEvents,
  parameters: example('toolbar-events'),
};

export const ToolbarStates: StoryObj = {
  name: 'Toolbar states',
  render: toolbarStates,
  parameters: example('toolbar-states', { height: 260 }),
};
