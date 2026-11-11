import { fileURLToPath } from 'node:url';
import remarkGfm from 'remark-gfm';
import type { StorybookConfig } from '@storybook/web-components-vite';

const config: StorybookConfig = {
  stories: ['../stories/**/*.mdx', '../stories/**/*.stories.@(js|ts)'],
  addons: [
    '@storybook/addon-a11y',
    {
      name: '@storybook/addon-docs',
      // Tables in the MDX pages are GitHub-flavored Markdown, which MDX leaves out by default
      options: { mdxPluginOptions: { mdxCompileOptions: { remarkPlugins: [remarkGfm] } } },
    },
  ],
  framework: {
    name: '@storybook/web-components-vite',
    options: {},
  },
  features: {
    sidebarOnboardingChecklist: false,
  },
  async viteFinal(viteConfig) {
    const { mergeConfig } = await import('vite');
    return mergeConfig(viteConfig, {
      resolve: {
        // The examples import the package by its name, exactly as an app does
        alias: [
          { find: /^y11n-spreadsheet$/, replacement: fileURLToPath(new URL('../src/index.ts', import.meta.url)) },
        ],
      },
    });
  },
};

export default config;
