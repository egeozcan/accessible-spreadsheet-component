import { defineConfig, devices } from '@playwright/test';

export default defineConfig({
  testDir: './e2e',
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  workers: process.env.CI ? 1 : undefined,
  reporter: 'list',
  use: {
    baseURL: 'http://localhost:5173',
    headless: true,
    trace: 'on-first-retry',
  },
  projects: [
    {
      name: 'chromium',
      use: { ...devices['Desktop Chrome'] },
      testIgnore: 'storybook-a11y.spec.ts',
    },
    {
      name: 'storybook-a11y',
      use: { ...devices['Desktop Chrome'], baseURL: 'http://localhost:6006' },
      testMatch: 'storybook-a11y.spec.ts',
    },
  ],
  webServer: [
    {
      command: 'npx vite --port 5173',
      port: 5173,
      reuseExistingServer: !process.env.CI,
    },
    {
      command: 'npx storybook dev -p 6006 --no-open',
      port: 6006,
      reuseExistingServer: !process.env.CI,
      timeout: 60_000,
    },
  ],
});
