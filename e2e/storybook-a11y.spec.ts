import { test, expect } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';

/**
 * Story metadata for data-driven axe audits.
 * Each entry navigates to the Storybook iframe URL for that story.
 */
interface StoryMeta {
  id: string;
  name: string;
  /** CSS selector to wait for before scanning (inside shadow DOM or light DOM) */
  waitFor: string;
  /** Extra delay in ms for stories with animations or large data */
  extraWait?: number;
}

const spreadsheetStories: StoryMeta[] = [
  { id: 'components-y11n-spreadsheet--default', name: 'Default', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--pre-populated-data', name: 'Pre-populated Data', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--formulas-and-calculations', name: 'Formulas & Calculations', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--read-only', name: 'Read-Only', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--custom-functions', name: 'Custom Functions', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--compact-grid', name: 'Compact Grid', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--large-dataset', name: 'Large Dataset', waitFor: 'y11n-spreadsheet', extraWait: 5000 },
  { id: 'components-y11n-spreadsheet--event-logging', name: 'Event Logging', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--dark-theme', name: 'Dark Theme', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--pastel-theme', name: 'Pastel Theme', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--budget-tracker', name: 'Budget Tracker', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--grade-book', name: 'Grade Book', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--string-formulas', name: 'String Formulas', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--wide-cells', name: 'Wide Cells', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--conditional-logic', name: 'Conditional Logic', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--api-methods', name: 'API Methods', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--lookup-functions', name: 'Lookup Functions', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--conditional-aggregation', name: 'Conditional Aggregation', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--logic-functions', name: 'Logic Functions', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--extended-string-functions', name: 'Extended String Functions', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-spreadsheet--absolute-references', name: 'Absolute References', waitFor: 'y11n-spreadsheet' },
];

const toolbarStories: StoryMeta[] = [
  { id: 'components-y11n-format-toolbar--default-toolbar', name: 'Default Toolbar', waitFor: 'y11n-format-toolbar' },
  { id: 'components-y11n-format-toolbar--active-formatting-state', name: 'Active Formatting State', waitFor: 'y11n-format-toolbar' },
  { id: 'components-y11n-format-toolbar--read-only-toolbar', name: 'Read-Only Toolbar', waitFor: 'y11n-format-toolbar' },
  { id: 'components-y11n-format-toolbar--toolbar-integration', name: 'Toolbar Integration', waitFor: 'y11n-format-toolbar' },
  { id: 'components-y11n-format-toolbar--preformatted-data', name: 'Preformatted Data', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-format-toolbar--formatted-budget-tracker', name: 'Formatted Budget Tracker', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-format-toolbar--color-palette-demo', name: 'Color Palette Demo', waitFor: 'y11n-spreadsheet' },
  { id: 'components-y11n-format-toolbar--duck-pixel-art', name: 'Duck Pixel Art', waitFor: 'y11n-spreadsheet', extraWait: 3500 },
  { id: 'components-y11n-format-toolbar--programmatic-format-api', name: 'Programmatic Format API', waitFor: 'y11n-spreadsheet' },
];

const allStories = [...spreadsheetStories, ...toolbarStories];

/**
 * Rules disabled globally:
 * - color-contrast: themed stories use intentional non-standard contrast
 * - page-has-heading-one: Storybook iframe has no headings (not a component concern)
 * - landmark-one-main: Storybook iframe has no landmarks (not a component concern)
 * - region: Storybook iframe content isn't wrapped in landmarks (not a component concern)
 */
const DISABLED_RULES = [
  'color-contrast',
  'page-has-heading-one',
  'landmark-one-main',
  'region',
];

test.describe('Storybook A11y Audit', () => {
  for (const story of allStories) {
    test(`${story.name} (${story.id}) has no axe violations`, async ({ page }) => {
      await page.goto(`/iframe.html?id=${story.id}&viewMode=story`);

      // Wait for the web component to render
      await page.waitForSelector(story.waitFor, { timeout: 15_000 });

      // Extra wait for animations or large datasets
      if (story.extraWait) {
        await page.waitForTimeout(story.extraWait);
      }

      const results = await new AxeBuilder({ page })
        .disableRules(DISABLED_RULES)
        .analyze();

      const violations = results.violations.map((v) => ({
        id: v.id,
        impact: v.impact,
        description: v.description,
        nodes: v.nodes.length,
        help: v.help,
        helpUrl: v.helpUrl,
      }));

      expect(violations, `Axe violations found:\n${JSON.stringify(violations, null, 2)}`).toEqual([]);
    });
  }
});
