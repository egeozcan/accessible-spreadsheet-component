import { defineConfig } from '@playwright/test';

const cwd = '/Users/egecan/Code/accessible-spreadsheet-component';

export default defineConfig({
  testDir: `${cwd}/e2e`,
  fullyParallel: true,
  retries: 0,
  reporter: 'list',
  use: {
    baseURL: 'http://localhost:4174',
    headless: true,
  },
  projects: [
    {
      name: 'chromium',
      use: {
        browserName: 'chromium',
        launchOptions: {
          executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH || undefined,
        },
      },
    },
  ],
  webServer: {
    command: `cd ${cwd} && npx vite --port 4174`,
    port: 4174,
    reuseExistingServer: false,
  },
});
