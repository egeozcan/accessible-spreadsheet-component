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

const allStories: StoryMeta[] = [
  { id: 'examples-playground--playground', name: 'Playground', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-basics--empty-grid', name: 'Empty grid', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-basics--loading-data', name: 'Loading data', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-basics--formulas', name: 'Formulas', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-basics--read-only', name: 'Read-only', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-basics--events', name: 'Events', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-basics--saving-and-restoring', name: 'Saving and restoring', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-formulas--lookup-functions', name: 'Lookup functions', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-formulas--conditional-aggregation', name: 'Conditional aggregation', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-formulas--logic-functions', name: 'Logic functions', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-formulas--text-functions', name: 'Text functions', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-formulas--absolute-references', name: 'Absolute references', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-formulas--custom-functions', name: 'Custom functions', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-formulas--dynamic-arrays', name: 'Dynamic arrays', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-formulas--lists-in-cells', name: 'Lists and arrays in cells', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-formulas--sheets-and-names', name: 'Sheets and names', waitFor: 'y11n-sheet-tabs' },
  { id: 'examples-editing-and-tables--validation', name: 'Validation and locked cells', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-editing-and-tables--find-and-fill', name: 'Find, fill and row insertion', waitFor: '#find-input' },
  { id: 'examples-editing-and-tables--sorting', name: 'Sorting', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-editing-and-tables--filter', name: 'Filter', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-layout--column-widths', name: 'Column widths', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-layout--frozen-panes', name: 'Frozen panes', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-layout--large-dataset', name: 'Large dataset', waitFor: 'y11n-spreadsheet', extraWait: 5000 },
  { id: 'examples-formatting--cell-formats', name: 'Cell formats', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-formatting--format-api', name: 'Format API', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-formatting--conditional-formatting', name: 'Conditional formatting', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-toolbar--toolbar', name: 'Toolbar', waitFor: 'y11n-toolbar' },
  { id: 'examples-toolbar--toolbar-items', name: 'Choosing the controls', waitFor: 'y11n-toolbar' },
  { id: 'examples-toolbar--toolbar-events', name: 'Toolbar events', waitFor: 'y11n-toolbar' },
  { id: 'examples-toolbar--toolbar-states', name: 'Toolbar states', waitFor: 'y11n-toolbar' },
  { id: 'examples-theming--dark-theme', name: 'Dark theme', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-theming--pastel-theme', name: 'Pastel theme', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-theming--cell-size', name: 'Cell size', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-recipes--budget-tracker', name: 'Budget tracker', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-recipes--grade-book', name: 'Grade book', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-recipes--inventory', name: 'Inventory', waitFor: 'y11n-spreadsheet' },
  { id: 'examples-recipes--pixel-art', name: 'Pixel art', waitFor: 'y11n-spreadsheet', extraWait: 3500 },
];

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
