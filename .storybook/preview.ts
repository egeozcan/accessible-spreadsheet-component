import type { Preview } from '@storybook/web-components-vite'

const preview: Preview = {
  parameters: {
    controls: {
      matchers: {
        color: /(background|color)$/i,
        date: /Date$/i,
      },
    },

    layout: 'fullscreen',

    options: {
      storySort: {
        order: ['Components', ['y11n-spreadsheet', '*']],
      },
    },

    docs: {
      codePanel: true
    }
  },
};

export default preview;