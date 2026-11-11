import { addons } from 'storybook/manager-api';

// Open on the Code panel, beside the example it belongs to and wide enough to read
addons.setConfig({
  selectedPanel: 'storybook/docs/panel',
  layout: { panelPosition: 'right', rightPanelWidth: 560 },
});
