import type { Meta, StoryObj } from '@storybook/web-components-vite';
import { example } from './support/example.js';
import lookupFunctions from './examples/lookup-functions.js';
import conditionalAggregation from './examples/conditional-aggregation.js';
import logicFunctions from './examples/logic-functions.js';
import textFunctions from './examples/text-functions.js';
import absoluteReferences from './examples/absolute-references.js';
import customFunctions from './examples/custom-functions.js';
import dynamicArrays from './examples/dynamic-arrays.js';
import listsInCells from './examples/lists-in-cells.js';
import sheetsAndNames from './examples/sheets-and-names.js';

const meta: Meta = {
  title: 'Examples/Formulas',
  tags: ['autodocs'],
  parameters: {
    docs: {
      description: {
        component: `The formula language follows Excel: the same function names, argument order and reference syntax. These examples go from everyday functions to the parts that are specific to a modern spreadsheet: dynamic arrays, several values in one cell, sheets and names. The full list of functions is under **Reference / Formula Engine**.`,
      },
    },
  },
};

export default meta;

export const LookupFunctions: StoryObj = {
  name: 'Lookup functions',
  render: lookupFunctions,
  parameters: example('lookup-functions', { height: 480 }),
};

export const ConditionalAggregation: StoryObj = {
  name: 'Conditional aggregation',
  render: conditionalAggregation,
  parameters: example('conditional-aggregation', { height: 560 }),
};

export const LogicFunctions: StoryObj = {
  name: 'Logic functions',
  render: logicFunctions,
  parameters: example('logic-functions', { height: 300 }),
};

export const TextFunctions: StoryObj = {
  name: 'Text functions',
  render: textFunctions,
  parameters: example('text-functions', { height: 300 }),
};

export const AbsoluteReferences: StoryObj = {
  name: 'Absolute references',
  render: absoluteReferences,
  parameters: example('absolute-references'),
};

export const CustomFunctions: StoryObj = {
  name: 'Custom functions',
  render: customFunctions,
  parameters: example('custom-functions', { height: 300 }),
};

export const DynamicArrays: StoryObj = {
  name: 'Dynamic arrays',
  render: dynamicArrays,
  parameters: example('dynamic-arrays'),
};

export const ListsInCells: StoryObj = {
  name: 'Lists and arrays in cells',
  render: listsInCells,
  parameters: example('lists-in-cells', { height: 340 }),
};

export const SheetsAndNames: StoryObj = {
  name: 'Sheets and names',
  render: sheetsAndNames,
  parameters: example('sheets-and-names', { height: 360 }),
};
