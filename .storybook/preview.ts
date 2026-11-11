import { html } from 'lit';
import type { Preview } from '@storybook/web-components-vite';
import ExamplesPage from './ExamplesPage.mdx';

const preview: Preview = {
  decorators: [
    // The grid fills its parent, so every example gets a parent with a height:
    // the whole canvas for a single story, a fixed frame on a docs page.
    (story, context) => {
      const height = context.viewMode === 'docs' ? `${context.parameters.frameHeight ?? 420}px` : '100vh';
      return html`<div class="example-frame" style="height: ${height}">${story()}</div>`;
    },
  ],

  parameters: {
    layout: 'fullscreen',

    // Most examples have nothing to control; the Playground turns this back on
    controls: { disable: true },
    actions: { disable: true },

    options: {
      storySort: {
        order: [
          'Getting Started',
          ['Introduction', 'Installation', 'Your First Spreadsheet'],
          'Examples',
          ['Playground', 'Basics', 'Formulas', 'Editing and Tables', 'Layout', 'Formatting', 'Theming', 'Recipes'],
          'Reference',
          [
            'Properties and Methods',
            'Events',
            'Cell Formats',
            'Formula Engine',
            'Feature Details',
            'Theming',
            'Keyboard Shortcuts',
            'Types and Exports',
          ],
        ],
      },
    },

    docs: {
      codePanel: true,
      toc: { headingSelector: 'h2, h3' },
      page: ExamplesPage,
    },
  },
};

export default preview;
