import { LitElement, html, css, nothing, svg, type TemplateResult } from 'lit';
import { customElement, property, state } from 'lit/decorators.js';
import { live } from 'lit/directives/live.js';
import type { CellFormat, ConditionalFormatPreset, NumberFormatOptions, SelectionRange } from '../types.js';
import type { Y11nSpreadsheet } from '../y11n-spreadsheet.js';

/**
 * What the Conditional formatting menu asks for: a one-step rule for the
 * selection, clearing the selection's rules, or opening the rule manager.
 */
export type ConditionalFormatMenuAction = ConditionalFormatPreset | 'clear' | 'manage';

/** Detail for the `conditional-format-action` event dispatched by the toolbar. */
export interface ConditionalFormatActionDetail {
  action: ConditionalFormatMenuAction;
}

/** Detail for the `format-action` custom event dispatched by the toolbar. */
export interface FormatActionDetail {
  /** The formatting action to perform */
  action:
    | 'bold'
    | 'italic'
    | 'underline'
    | 'strikethrough'
    | 'textColor'
    | 'backgroundColor'
    | 'textAlign'
    | 'verticalAlign'
    | 'wrap'
    | 'fontSize'
    | 'numberFormat'
    | 'clearFormat';
  /**
   * The new value for the action: true/false for toggles, a color, an
   * alignment, a font size, or for `numberFormat` the format (undefined for
   * General). Except for the four text-style toggles and `clearFormat`,
   * `{ [action]: value }` is a `CellFormat`.
   */
  value?: string | number | boolean | NumberFormatOptions;
}

/**
 * What the toolbar's other controls ask for. Each one is something
 * `<y11n-spreadsheet>` has a method for (see `toolbar-action` in the README).
 */
export type ToolbarAction =
  | 'undo'
  | 'redo'
  | 'cut'
  | 'copy'
  | 'paste'
  | 'pasteSpecial'
  | 'decreaseDecimals'
  | 'increaseDecimals'
  | 'insertRowAbove'
  | 'insertRowBelow'
  | 'insertColumnLeft'
  | 'insertColumnRight'
  | 'deleteRow'
  | 'deleteColumn'
  | 'autoFit'
  | 'sortAscending'
  | 'sortDescending'
  | 'filter'
  | 'reapplyFilter'
  | 'list'
  | 'autoSum'
  | 'freeze'
  | 'freezeFirstRow'
  | 'freezeFirstColumn'
  | 'unfreeze'
  | 'find'
  | 'names';

/** Detail for the `toolbar-action` event dispatched by the toolbar. */
export interface ToolbarActionDetail {
  action: ToolbarAction;
}

/** A control of the toolbar: a button, a group of buttons, a select or a menu. */
export type ToolbarItem =
  | 'undo'
  | 'redo'
  | 'cut'
  | 'copy'
  | 'paste'
  | 'pasteSpecial'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strikethrough'
  | 'fontSize'
  | 'textColor'
  | 'backgroundColor'
  | 'textAlign'
  | 'verticalAlign'
  | 'wrap'
  | 'numberFormat'
  | 'decreaseDecimals'
  | 'increaseDecimals'
  | 'conditionalFormat'
  | 'clearFormat'
  | 'insert'
  | 'delete'
  | 'autoFit'
  | 'sortAscending'
  | 'sortDescending'
  | 'filter'
  | 'reapplyFilter'
  | 'list'
  | 'autoSum'
  | 'freeze'
  | 'find'
  | 'names';

/** A named set of toolbar items, usable wherever an item is. */
export type ToolbarGroup =
  | 'history'
  | 'clipboard'
  | 'font'
  | 'alignment'
  | 'number'
  | 'styles'
  | 'cells'
  | 'data'
  | 'view';

/** The toolbar's groups and their items, in the order the toolbar shows them by default. */
export const TOOLBAR_GROUPS: Readonly<Record<ToolbarGroup, readonly ToolbarItem[]>> = {
  history: ['undo', 'redo'],
  clipboard: ['cut', 'copy', 'paste', 'pasteSpecial'],
  font: ['bold', 'italic', 'underline', 'strikethrough', 'fontSize', 'textColor', 'backgroundColor'],
  alignment: ['textAlign', 'verticalAlign', 'wrap'],
  number: ['numberFormat', 'decreaseDecimals', 'increaseDecimals'],
  styles: ['conditionalFormat', 'clearFormat'],
  cells: ['insert', 'delete', 'autoFit'],
  data: ['sortAscending', 'sortDescending', 'filter', 'reapplyFilter', 'list', 'autoSum'],
  view: ['freeze', 'find', 'names'],
};

/** Every toolbar item, in default order. */
export const TOOLBAR_ITEMS: readonly ToolbarItem[] = Object.values(TOOLBAR_GROUPS).flat();

/** What `items` accepts: items, group names, and `'|'` for a separator. */
export type ToolbarItemSpec = ToolbarItem | ToolbarGroup | '|';

const GROUP_LABELS: Record<ToolbarGroup, string> = {
  history: 'History',
  clipboard: 'Clipboard',
  font: 'Font',
  alignment: 'Alignment',
  number: 'Numbers',
  styles: 'Styles',
  cells: 'Rows and columns',
  data: 'Data',
  view: 'View',
};

type MenuEntry<T extends string> = { action: T; label: string; disabled?: boolean } | null;

const CONDITIONAL_MENU: Array<MenuEntry<ConditionalFormatMenuAction>> = [
  { action: 'duplicates', label: 'Highlight duplicate values' },
  { action: 'unique', label: 'Highlight unique values' },
  { action: 'aboveAverage', label: 'Highlight above average' },
  { action: 'belowAverage', label: 'Highlight below average' },
  { action: 'top10', label: 'Highlight top 10' },
  { action: 'bottom10', label: 'Highlight bottom 10' },
  null,
  { action: 'colorScale', label: 'Color scale' },
  { action: 'dataBar', label: 'Data bars' },
  { action: 'iconSet', label: 'Icon set' },
  null,
  { action: 'clear', label: 'Clear rules from selection' },
  { action: 'manage', label: 'Manage rules…' },
];

const INSERT_MENU: Array<MenuEntry<ToolbarAction>> = [
  { action: 'insertRowAbove', label: 'Insert row above' },
  { action: 'insertRowBelow', label: 'Insert row below' },
  null,
  { action: 'insertColumnLeft', label: 'Insert column left' },
  { action: 'insertColumnRight', label: 'Insert column right' },
];

const DELETE_MENU: Array<MenuEntry<ToolbarAction>> = [
  { action: 'deleteRow', label: 'Delete row' },
  { action: 'deleteColumn', label: 'Delete column' },
];

/** The Number format select: [value, label, format]. The formats are those of Ctrl+Shift+1/3/4/5. */
const NUMBER_FORMATS: Array<[string, string, NumberFormatOptions | undefined]> = [
  ['general', 'General', undefined],
  ['number', 'Number', { type: 'number', decimals: 2, thousandsSep: true }],
  ['currency', 'Currency', { type: 'currency', decimals: 2 }],
  ['percent', 'Percent', { type: 'percent', decimals: 0 }],
  ['scientific', 'Scientific', { type: 'scientific', decimals: 2 }],
  ['date', 'Date', { type: 'date', datePattern: 'YYYY-MM-DD' }],
];

const FONT_SIZES = [8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 32, 36];

/** Events of the grid after which the toolbar's state may be out of date. */
const SYNC_EVENTS = [
  'selection-change',
  'format-change',
  'cell-change',
  'data-change',
  'history-change',
  'filter-change',
  'freeze-change',
  'structure-change',
  'sheet-change',
];

// Letters and digits in the icons are drawn as paths: text in a button would
// have to be part of its accessible name
const LETTER_A_TOP = 'M1.5 7L3.75 1.5 6 7M2.4 5h2.7';
const LETTER_A_BOTTOM = 'M1.5 14.5L3.75 9 6 14.5M2.4 12.5h2.7';
const LETTER_Z_TOP = 'M1.75 1.75h4l-4 5h4';
const LETTER_Z_BOTTOM = 'M1.75 9.5h4l-4 5h4';
const SORT_ARROW = 'M11.5 2.5v11M9 11l2.5 2.5L14 11';

const ICONS = {
  undo: svg`<path d="M5.5 3.5l-3 3 3 3"/><path d="M2.5 6.5h7a4 4 0 0 1 0 8H6"/>`,
  redo: svg`<path d="M10.5 3.5l3 3-3 3"/><path d="M13.5 6.5h-7a4 4 0 0 0 0 8H10"/>`,
  cut: svg`<circle cx="4.5" cy="12" r="2"/><circle cx="11.5" cy="12" r="2"/><path d="M5.7 10.4L11.5 2M10.3 10.4L4.5 2"/>`,
  copy: svg`<rect x="5.5" y="5.5" width="8" height="9" rx="1"/><path d="M3 11V3a1 1 0 0 1 1-1h6.5"/>`,
  paste: svg`<path d="M5.5 3H4a1 1 0 0 0-1 1v9.5a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V4a1 1 0 0 0-1-1h-1.5"/><rect x="5.5" y="1.5" width="5" height="3" rx="0.75"/>`,
  pasteSpecial: svg`<path d="M5.5 3H4a1 1 0 0 0-1 1v9.5a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V4a1 1 0 0 0-1-1h-1.5"/><rect x="5.5" y="1.5" width="5" height="3" rx="0.75"/><path d="M5.5 8h5M5.5 11h3"/>`,
  wrap: svg`<path d="M2 3.5h12M2 8h9.75a2.25 2.25 0 0 1 0 4.5H8.5M2 12.5h3.5"/><path d="M10 11l-1.5 1.5L10 14"/>`,
  top: svg`<path d="M2 2.5h12M8 13.5v-8M5.5 8L8 5.5 10.5 8"/>`,
  middle: svg`<path d="M2 8h12M8 1.5v4M6.5 4L8 5.5 9.5 4M8 14.5v-4M6.5 12L8 10.5 9.5 12"/>`,
  bottom: svg`<path d="M2 13.5h12M8 2.5v8M5.5 8L8 10.5 10.5 8"/>`,
  decreaseDecimals: svg`<path d="M4 7.25h.01" stroke-width="2"/><ellipse cx="8.25" cy="4.5" rx="1.75" ry="2.75" stroke-width="1.25"/><path d="M13 12.5H5.5M8 10l-2.5 2.5L8 15"/>`,
  increaseDecimals: svg`<path d="M2 7.25h.01" stroke-width="2"/><ellipse cx="6" cy="4.5" rx="1.75" ry="2.75" stroke-width="1.25"/><ellipse cx="11.25" cy="4.5" rx="1.75" ry="2.75" stroke-width="1.25"/><path d="M3 12.5h7.5M8 10l2.5 2.5L8 15"/>`,
  conditionalFormat: svg`<path d="M2 3h12v2.5H2zM2 6.75h8v2.5H2zM2 10.5h4.5V13H2z" fill="currentColor" stroke="none"/>`,
  clearFormat: svg`<path d="M3 3h8M7 3v9.5M10.5 10.5l3.5 3.5M14 10.5L10.5 14"/>`,
  insert: svg`<rect x="2" y="2" width="12" height="5" rx="1"/><path d="M8 9.5v5M5.5 12h5"/>`,
  delete: svg`<rect x="2" y="2" width="12" height="5" rx="1"/><path d="M5.75 9.75l4.5 4.5M10.25 9.75l-4.5 4.5"/>`,
  autoFit: svg`<path d="M2 2v12M14 2v12M4.5 8h7M6.5 6l-2 2 2 2M9.5 6l2 2-2 2"/>`,
  sortAscending: svg`<path d="${LETTER_A_TOP}${LETTER_Z_BOTTOM}" stroke-width="1.25"/><path d="${SORT_ARROW}"/>`,
  sortDescending: svg`<path d="${LETTER_Z_TOP}${LETTER_A_BOTTOM}" stroke-width="1.25"/><path d="${SORT_ARROW}"/>`,
  filter: svg`<path d="M2 3h12L9.5 8.5v4l-3 1.5V8.5z"/>`,
  reapplyFilter: svg`<path d="M13.5 8a5.5 5.5 0 1 1-1.6-3.9"/><path d="M13.5 2v3.5H10"/>`,
  list: svg`<path d="M6 4h8M6 8h8M6 12h8"/><path d="M2.5 4h.01M2.5 8h.01M2.5 12h.01" stroke-width="2.25"/>`,
  freeze: svg`<path d="M2.75 2.75h10.5v2.5H2.75zM2.75 6.75h2.5v6.5h-2.5z" fill="currentColor" stroke="none" opacity="0.45"/><rect x="2" y="2" width="12" height="12" rx="1"/><path d="M2 6h12M6 2v12"/>`,
  find: svg`<circle cx="7" cy="7" r="4.5"/><path d="M10.5 10.5L14 14"/>`,
  names: svg`<path d="M2.5 2.5h5l6 6-5 5-6-6z"/><circle cx="5.5" cy="5.5" r="0.75" fill="currentColor"/>`,
};

const icon = (name: keyof typeof ICONS) => html`<svg
  class="icon"
  viewBox="0 0 16 16"
  width="16"
  height="16"
  aria-hidden="true"
  focusable="false"
  fill="none"
  stroke="currentColor"
  stroke-width="1.5"
  stroke-linecap="round"
  stroke-linejoin="round"
>${ICONS[name]}</svg>`;

const caret = html`<svg class="caret" viewBox="0 0 10 10" width="8" height="8" aria-hidden="true" focusable="false">
  <path d="M1.5 3.5 5 7l3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.5" />
</svg>`;

/** `<input type="color">` only takes `#rrggbb`. */
const colorInputValue = (color: string, fallback: string) => (/^#[0-9a-f]{6}$/i.test(color) ? color : fallback);

/**
 * `<y11n-toolbar>` -- A toolbar for `<y11n-spreadsheet>`: undo and
 * redo, clipboard, cell formats, number formats, conditional formatting,
 * rows and columns, sorting, filtering, frozen panes, find and names.
 *
 * Point it at a grid and it wires itself: its controls show the state of
 * the selection and act on it.
 *
 * ```html
 * <y11n-toolbar for="sheet"></y11n-toolbar>
 * <y11n-spreadsheet id="sheet"></y11n-spreadsheet>
 * ```
 *
 * (From script: `toolbar.target = sheet`.) `items` chooses the controls and
 * their order: items, group names and `|` for a separator.
 *
 * ```html
 * <y11n-toolbar for="sheet" items="history | bold italic | sortAscending filter"></y11n-toolbar>
 * ```
 *
 * Every control fires an event before it acts: `format-action`,
 * `conditional-format-action` or `toolbar-action`. Call `preventDefault()`
 * to do something else instead. Without a target the toolbar only fires
 * the events and shows what its properties say, so it can be wired by hand.
 *
 * @fires format-action - A formatting control was used: `{ action, value }`
 * @fires conditional-format-action - An item of the Conditional formatting menu was picked: `{ action }`
 * @fires toolbar-action - Any other control was used: `{ action }`
 */
@customElement('y11n-toolbar')
export class Y11nToolbar extends LitElement {
  // ─── Configuration ─────────────────────────────────────

  /**
   * The controls to show, in order: item names, group names (see
   * `TOOLBAR_GROUPS`) and `'|'` for a separator, as an array or as text with
   * spaces or commas between them (`'history | bold italic'`), which is also
   * how the attribute is written. Not set (or null): everything.
   */
  @property({ type: String }) items: readonly ToolbarItemSpec[] | string | null = null;

  /** The grid to act on and to follow. Takes precedence over `for`. */
  @property({ attribute: false }) target: Y11nSpreadsheet | null = null;

  /** The id of the grid to act on and to follow, looked up next to the toolbar. */
  @property({ type: String, attribute: 'for' }) htmlFor = '';

  /** The toolbar's accessible name. Give each toolbar its own when a page has several. */
  @property({ type: String }) label = 'Spreadsheet toolbar';

  // ─── State shown by the controls (set from the target, if there is one) ───

  @property({ type: Boolean }) bold = false;
  @property({ type: Boolean }) italic = false;
  @property({ type: Boolean }) underline = false;
  @property({ type: Boolean }) strikethrough = false;
  /** The selection's text color; empty for none (the swatch then shows the theme's text color) */
  @property({ type: String }) textColor = '#000000';
  /** The selection's fill color; empty for none */
  @property({ type: String }) backgroundColor = '#ffffff';
  @property({ type: String }) textAlign: 'left' | 'center' | 'right' = 'left';
  @property({ type: String }) verticalAlign: 'top' | 'middle' | 'bottom' = 'middle';
  @property({ type: Boolean }) wrap = false;
  @property({ type: Number }) fontSize = 13;
  /** The selection's number format; undefined for General */
  @property({ attribute: false }) numberFormat: NumberFormatOptions | undefined = undefined;
  /** Disables the controls that change the grid; those that only change the view stay available */
  @property({ type: Boolean, attribute: 'read-only', reflect: true }) readOnly = false;
  /** Whether Undo is available */
  @property({ attribute: false }) canUndo = true;
  /** Whether Redo is available */
  @property({ attribute: false }) canRedo = true;
  /** Whether the sheet has a filter: shown by the Filter button, and needed by Reapply filter */
  @property({ type: Boolean }) filtered = false;
  /** Whether the sheet has frozen rows or columns: needed by Unfreeze panes */
  @property({ type: Boolean }) frozen = false;

  /** The menu that is open */
  @state() private _menu: string | null = null;
  /** Whether the open menu is aligned to the right edge of its button (no room on the right) */
  @state() private _menuAtEnd = false;
  /** Whether the open menu is above its button (no room below, as for a toolbar under the grid) */
  @state() private _menuAbove = false;
  private _menuCleanup: (() => void) | null = null;

  /** The control that Tab enters the toolbar on */
  private _tabStop = '';
  private _connected: Y11nSpreadsheet | null = null;
  private _disconnectTarget: (() => void) | null = null;
  private _syncQueued = false;
  private _warned = new Set<string>();

  connectedCallback(): void {
    super.connectedCallback();
    this._attach();
    // An element named by `for` may come later in the document
    if (!this._connected && this.htmlFor) setTimeout(() => this._attach());
  }

  disconnectedCallback(): void {
    super.disconnectedCallback();
    this._closeMenu(false);
    this._attach();
  }

  protected willUpdate(changed: Map<string, unknown>): void {
    if (changed.has('target') || changed.has('htmlFor')) this._attach();
  }

  protected updated(): void {
    // One control is in the Tab order (the last one used, or the first);
    // the arrow keys move between them
    const controls = this._controls();
    const stop = controls.find((el) => el.dataset.item === this._tabStop) ?? controls[0];
    for (const el of this.renderRoot.querySelectorAll<HTMLElement>('[data-item]')) {
      el.tabIndex = el === stop ? 0 : -1;
    }
  }

  // ─── Styles ────────────────────────────────────────────

  static styles = css`
    :host {
      display: block;
      font-family: var(--ls-font-family, system-ui, -apple-system, sans-serif);
      font-size: var(--ls-font-size, 13px);
      color: var(--ls-text-color, #333);
    }

    .toolbar {
      display: flex;
      align-items: center;
      gap: 2px;
      padding: 4px 8px;
      border-bottom: 1px solid var(--ls-border-color, #e0e0e0);
      background: var(--ls-toolbar-bg, #f8fafc);
      flex-wrap: wrap;
    }

    .group {
      display: inline-flex;
      align-items: center;
      gap: 2px;
    }

    .separator {
      width: 1px;
      height: 20px;
      background: var(--ls-border-color, #e0e0e0);
      margin: 0 4px;
    }

    button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 28px;
      height: 28px;
      padding: 0 4px;
      border: 1px solid transparent;
      border-radius: 4px;
      background: transparent;
      cursor: pointer;
      font: inherit;
      font-size: 13px;
      color: inherit;
    }

    button:hover:not(:disabled) {
      background: var(--ls-toolbar-hover-bg, #e8eaed);
    }

    button[aria-pressed='true'] {
      background: var(--ls-toolbar-active-bg, #d3e3fd);
      border-color: var(--ls-toolbar-active-border, #a8c7fa);
    }

    button:disabled {
      opacity: 0.4;
      cursor: not-allowed;
    }

    button.bold-btn {
      font-weight: bold;
    }

    button.italic-btn {
      font-style: italic;
    }

    button.underline-btn {
      text-decoration: underline;
    }

    button.strikethrough-btn {
      text-decoration: line-through;
    }

    button.sum-btn {
      font-size: 15px;
    }

    .icon {
      display: block;
    }

    .color-wrapper {
      position: relative;
      display: inline-flex;
      align-items: center;
      border-radius: 4px;
    }

    .color-wrapper:hover:not(.disabled) {
      background: var(--ls-toolbar-hover-bg, #e8eaed);
    }

    .color-wrapper.disabled {
      opacity: 0.4;
    }

    .color-wrapper:focus-within {
      outline: var(--ls-focus-ring, 2px solid #1a73e8);
      outline-offset: 1px;
    }

    .color-wrapper input[type='color'] {
      position: absolute;
      inset: 0;
      opacity: 0;
      width: 100%;
      height: 100%;
      cursor: pointer;
      border: none;
    }

    .color-wrapper input[type='color']:disabled {
      cursor: not-allowed;
    }

    .color-swatch {
      display: inline-flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      min-width: 28px;
      height: 28px;
      padding: 0 4px;
      box-sizing: border-box;
      border: 1px solid transparent;
      border-radius: 4px;
      background: transparent;
      font-size: 13px;
      line-height: 1;
      pointer-events: none;
    }

    .color-swatch .indicator {
      display: block;
      width: 14px;
      height: 3px;
      margin-top: 1px;
      border-radius: 1px;
    }

    .color-swatch .indicator.fill {
      width: 16px;
      height: 16px;
      margin: 0;
      border: 1px solid var(--ls-border-color, #ccc);
      border-radius: 2px;
      box-sizing: border-box;
    }

    .button-group {
      display: inline-flex;
      gap: 0;
    }

    .align-icon {
      display: flex;
      flex-direction: column;
      gap: 2px;
      width: 14px;
    }

    .align-icon span {
      display: block;
      height: 2px;
      background: currentColor;
      border-radius: 1px;
    }

    .align-left span:nth-child(1) { width: 100%; }
    .align-left span:nth-child(2) { width: 70%; }
    .align-left span:nth-child(3) { width: 85%; }

    .align-center span:nth-child(1) { width: 100%; margin: 0 auto; }
    .align-center span:nth-child(2) { width: 70%; margin: 0 auto; }
    .align-center span:nth-child(3) { width: 85%; margin: 0 auto; }

    .align-right span:nth-child(1) { width: 100%; margin-left: auto; }
    .align-right span:nth-child(2) { width: 70%; margin-left: auto; }
    .align-right span:nth-child(3) { width: 85%; margin-left: auto; }

    select {
      height: 28px;
      padding: 0 4px;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 4px;
      background: var(--ls-toolbar-input-bg, #fff);
      font: inherit;
      font-size: 12px;
      color: inherit;
      cursor: pointer;
    }

    .font-size-select {
      width: 52px;
    }

    .number-format-select {
      width: 96px;
    }

    select:disabled {
      opacity: 0.4;
      cursor: not-allowed;
    }

    .menu-wrapper {
      position: relative;
      display: inline-flex;
    }

    .menu-button {
      gap: 3px;
    }

    .menu {
      position: absolute;
      top: 100%;
      left: 0;
      z-index: 30;
      display: flex;
      flex-direction: column;
      min-width: 200px;
      margin-top: 2px;
      padding: 4px;
      background: var(--ls-filter-menu-bg, #fff);
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 4px;
      box-shadow: var(--ls-editor-shadow, 0 2px 6px rgba(0, 0, 0, 0.2));
    }

    .menu.at-end {
      left: auto;
      right: 0;
    }

    .menu.above {
      top: auto;
      bottom: 100%;
      margin: 0 0 2px;
    }

    .menu button {
      justify-content: flex-start;
      width: 100%;
      padding: 0 8px;
      white-space: nowrap;
    }

    .menu [role='separator'] {
      height: 1px;
      margin: 4px 0;
      background: var(--ls-border-color, #e0e0e0);
    }

    button:focus-visible,
    select:focus-visible {
      outline: var(--ls-focus-ring, 2px solid #1a73e8);
      outline-offset: 1px;
    }

    @media (forced-colors: active) {
      button[aria-pressed='true'] {
        forced-color-adjust: none;
        background: Highlight;
        border-color: Highlight;
        color: HighlightText;
      }

      button:disabled,
      select:disabled {
        opacity: 1;
        color: GrayText;
      }

      .separator,
      .menu [role='separator'] {
        background: CanvasText;
      }
    }
  `;

  // ─── The Target Grid ───────────────────────────────────

  private _findTarget(): Y11nSpreadsheet | null {
    if (this.target) return this.target;
    if (!this.htmlFor) return null;
    const root = this.getRootNode() as Document | ShadowRoot;
    return (root.getElementById?.(this.htmlFor) as Y11nSpreadsheet | null) ?? null;
  }

  /** Follow the target grid (or stop, when there is none or the toolbar left the document). */
  private _attach(): Y11nSpreadsheet | null {
    const next = this.isConnected ? this._findTarget() : null;
    if (next === this._connected) return next;
    this._disconnectTarget?.();
    this._disconnectTarget = null;
    this._connected = next;
    if (!next) return null;

    const sync = () => this._queueSync();
    for (const type of SYNC_EVENTS) next.addEventListener(type, sync);
    // Properties that are set without an event; all three are reflected as attributes
    const observer =
      typeof MutationObserver === 'undefined' ? null : new MutationObserver(sync);
    observer?.observe(next, { attributes: true, attributeFilter: ['read-only', 'frozen-rows', 'frozen-cols'] });
    this._disconnectTarget = () => {
      for (const type of SYNC_EVENTS) next.removeEventListener(type, sync);
      observer?.disconnect();
    };
    this._queueSync();
    return next;
  }

  private _queueSync(): void {
    if (this._syncQueued) return;
    this._syncQueued = true;
    queueMicrotask(() => {
      this._syncQueued = false;
      this._sync();
    });
  }

  /** Show the state of the target grid. */
  private _sync(): void {
    const grid = this._connected;
    if (!grid) return;
    if (typeof grid.getSelectionFormat !== 'function') {
      // Named by `for` before <y11n-spreadsheet> was defined
      void customElements.whenDefined(grid.localName).then(() => this._queueSync());
      return;
    }
    const format = grid.getSelectionFormat();
    this.bold = format.bold ?? false;
    this.italic = format.italic ?? false;
    this.underline = format.underline ?? false;
    this.strikethrough = format.strikethrough ?? false;
    // No color of their own: the swatches show the theme's text color and no fill
    this.textColor = format.textColor ?? '';
    this.backgroundColor = format.backgroundColor ?? '';
    this.textAlign = format.textAlign ?? 'left';
    this.verticalAlign = format.verticalAlign ?? 'middle';
    this.wrap = format.wrap ?? false;
    this.fontSize = format.fontSize ?? 13;
    this.numberFormat = format.numberFormat;
    this.readOnly = grid.readOnly;
    this.canUndo = grid.canUndo;
    this.canRedo = grid.canRedo;
    this.filtered = grid.getFilter() !== null;
    this.frozen = grid.frozenRows > 0 || grid.frozenCols > 0;
  }

  // ─── Actions ───────────────────────────────────────────

  /** Fire a cancelable event; true unless a listener called preventDefault(). */
  private _fire<T>(type: string, detail: T): boolean {
    return this.dispatchEvent(new CustomEvent<T>(type, { detail, bubbles: true, composed: true, cancelable: true }));
  }

  private _format(action: FormatActionDetail['action'], value?: FormatActionDetail['value']): void {
    const proceed = this._fire<FormatActionDetail>('format-action', { action, value });
    // The selects show the selection's format again if nothing came of the choice
    this.requestUpdate();
    const grid = proceed ? this._attach() : null;
    if (!grid) return;
    if (action === 'bold' || action === 'italic' || action === 'underline' || action === 'strikethrough') {
      grid.toggleFormat(action);
    } else if (action === 'clearFormat') {
      grid.clearSelectionFormat();
    } else {
      grid.formatSelection({ [action]: value });
    }
    this._queueSync();
  }

  private _conditional(action: ConditionalFormatMenuAction): void {
    const proceed = this._fire<ConditionalFormatActionDetail>('conditional-format-action', { action });
    const grid = proceed ? this._attach() : null;
    if (!grid) return;
    if (action === 'manage') void grid.openConditionalFormatManager();
    else if (action === 'clear') grid.clearSelectionConditionalFormats();
    else grid.applyConditionalFormatPreset(action);
  }

  private _act(action: ToolbarAction): void {
    const proceed = this._fire<ToolbarActionDetail>('toolbar-action', { action });
    const grid = proceed ? this._attach() : null;
    if (!grid) return;
    switch (action) {
      case 'undo':
        grid.undo();
        break;
      case 'redo':
        grid.redo();
        break;
      case 'cut':
        void grid.cut();
        break;
      case 'copy':
        void grid.copy();
        break;
      case 'paste':
        void grid.paste();
        break;
      case 'pasteSpecial':
        void grid.openPasteSpecial();
        break;
      case 'decreaseDecimals':
        grid.changeDecimals(-1);
        break;
      case 'increaseDecimals':
        grid.changeDecimals(1);
        break;
      case 'insertRowAbove':
        grid.insertRowsAtSelection('above');
        break;
      case 'insertRowBelow':
        grid.insertRowsAtSelection('below');
        break;
      case 'insertColumnLeft':
        grid.insertColumnsAtSelection('left');
        break;
      case 'insertColumnRight':
        grid.insertColumnsAtSelection('right');
        break;
      case 'deleteRow':
        grid.deleteSelectedRows();
        break;
      case 'deleteColumn':
        grid.deleteSelectedColumns();
        break;
      case 'autoFit':
        grid.autoFitSelectedColumns();
        break;
      case 'sortAscending':
        grid.sortSelection('ascending');
        break;
      case 'sortDescending':
        grid.sortSelection('descending');
        break;
      case 'filter':
        grid.toggleFilter();
        break;
      case 'reapplyFilter':
        grid.reapplyFilter();
        grid.announce('Filter reapplied');
        break;
      case 'list':
        grid.toggleLists();
        break;
      case 'autoSum':
        grid.autoSum();
        break;
      case 'freeze': {
        const { row, col } = grid.getSelection().activeCell;
        if (row === 0 && col === 0) {
          grid.announce('Nothing to freeze: move below or right of the rows and columns to keep in view');
        } else {
          grid.freezePanes(row, col);
          this._announceFrozen(grid);
        }
        break;
      }
      case 'freezeFirstRow':
        grid.freezePanes(1, 0);
        this._announceFrozen(grid);
        break;
      case 'freezeFirstColumn':
        grid.freezePanes(0, 1);
        this._announceFrozen(grid);
        break;
      case 'unfreeze':
        grid.unfreezePanes();
        grid.announce('Panes unfrozen');
        break;
      case 'find':
        void grid.openFind();
        break;
      case 'names':
        void grid.openNameManager();
        break;
    }
    this._queueSync();
  }

  private _announceFrozen(grid: Y11nSpreadsheet): void {
    const { frozenRows: rows, frozenCols: cols } = grid;
    const parts: string[] = [];
    if (rows > 0) parts.push(rows === 1 ? '1 row' : `${rows} rows`);
    if (cols > 0) parts.push(cols === 1 ? '1 column' : `${cols} columns`);
    grid.announce(parts.length > 0 ? `Froze ${parts.join(' and ')}` : 'Nothing was frozen');
  }

  // ─── Keyboard Navigation ───────────────────────────────

  /** The controls the arrow keys move between: enabled, and not inside a menu. */
  private _controls(): HTMLElement[] {
    return [...this.renderRoot.querySelectorAll<HTMLElement>('[data-item]:not(:disabled)')];
  }

  private _handleToolbarFocusin(e: FocusEvent): void {
    const item = (e.target as HTMLElement).dataset?.item;
    if (!item || item === this._tabStop) return;
    this._tabStop = item;
    this.requestUpdate();
  }

  private _handleToolbarKeydown(e: KeyboardEvent): void {
    const controls = this._controls();
    const current = controls.indexOf(e.target as HTMLElement);
    if (current === -1) return;

    let next = -1;
    if (e.key === 'ArrowRight') {
      next = (current + 1) % controls.length;
    } else if (e.key === 'ArrowLeft') {
      next = (current - 1 + controls.length) % controls.length;
    } else if (e.key === 'Home') {
      next = 0;
    } else if (e.key === 'End') {
      next = controls.length - 1;
    }

    if (next >= 0) {
      e.preventDefault();
      controls[next].focus();
    }
  }

  // ─── Menus ─────────────────────────────────────────────

  private _menuItems(): HTMLElement[] {
    return [...this.renderRoot.querySelectorAll<HTMLElement>('.menu [role="menuitem"]:not(:disabled)')];
  }

  private async _openMenu(id: string, focus: 'first' | 'last' = 'first'): Promise<void> {
    this._menu = id;
    this._menuAtEnd = false;
    this._menuAbove = false;
    if (!this._menuCleanup) {
      // Clicking anywhere else closes the menu, like a native dropdown
      const onPointerDown = (e: PointerEvent) => {
        const wrapper = this.renderRoot.querySelector('.menu-wrapper.open');
        if (wrapper && e.composedPath().includes(wrapper)) return;
        this._closeMenu(false);
      };
      document.addEventListener('pointerdown', onPointerDown);
      this._menuCleanup = () => document.removeEventListener('pointerdown', onPointerDown);
    }
    await this.updateComplete;
    const menu = this.renderRoot.querySelector<HTMLElement>('.menu');
    if (menu) {
      const rect = menu.getBoundingClientRect();
      const { clientWidth, clientHeight } = document.documentElement;
      this._menuAtEnd = rect.right > clientWidth;
      // Above the button when it does not fit below and does fit above
      const buttonTop = menu.parentElement!.getBoundingClientRect().top;
      this._menuAbove = rect.bottom > clientHeight && rect.height + 4 <= buttonTop;
      await this.updateComplete;
    }
    const items = this._menuItems();
    (focus === 'first' ? items[0] : items[items.length - 1])?.focus();
  }

  private _closeMenu(focusButton: boolean): void {
    this._menuCleanup?.();
    this._menuCleanup = null;
    if (this._menu === null) return;
    const button = this.renderRoot.querySelector<HTMLElement>('.menu-wrapper.open .menu-button');
    this._menu = null;
    if (focusButton) button?.focus();
  }

  private _handleMenuButtonKeydown(e: KeyboardEvent, id: string): void {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      e.stopPropagation();
      void this._openMenu(id, e.key === 'ArrowDown' ? 'first' : 'last');
    }
  }

  private _handleMenuKeydown(e: KeyboardEvent): void {
    // Arrow keys move inside the menu here, not along the toolbar
    e.stopPropagation();
    const items = this._menuItems();
    const current = items.indexOf((this.renderRoot as ShadowRoot).activeElement as HTMLElement);
    let next = -1;
    if (e.key === 'ArrowDown') next = (current + 1) % items.length;
    else if (e.key === 'ArrowUp') next = (current - 1 + items.length) % items.length;
    else if (e.key === 'Home') next = 0;
    else if (e.key === 'End') next = items.length - 1;
    else if (e.key === 'Escape') {
      e.preventDefault();
      this._closeMenu(true);
    } else if (e.key === 'Tab') {
      this._closeMenu(false);
    }
    if (next >= 0) {
      e.preventDefault();
      items[next]?.focus();
    }
  }

  private _renderMenu<T extends string>(
    item: ToolbarItem,
    label: string,
    iconName: keyof typeof ICONS,
    entries: Array<MenuEntry<T>>,
    choose: (action: T) => void,
    disabled: boolean
  ) {
    const open = this._menu === item && !disabled;
    return html`
      <div class="menu-wrapper ${open ? 'open' : ''}">
        <button
          class="menu-button"
          type="button"
          data-item="${item}"
          aria-label="${label}"
          aria-haspopup="menu"
          aria-expanded="${open}"
          title="${label}"
          ?disabled=${disabled}
          @click=${() => (this._menu === item ? this._closeMenu(false) : void this._openMenu(item))}
          @keydown=${(e: KeyboardEvent) => this._handleMenuButtonKeydown(e, item)}
        >${icon(iconName)}${caret}</button>
        ${open
          ? html`<div
              class="menu ${this._menuAtEnd ? 'at-end' : ''} ${this._menuAbove ? 'above' : ''}"
              role="menu"
              aria-label="${label}"
              @keydown=${this._handleMenuKeydown}
            >
              ${entries.map((entry) =>
                entry === null
                  ? html`<div role="separator"></div>`
                  : html`<button
                      type="button"
                      role="menuitem"
                      tabindex="-1"
                      data-action="${entry.action}"
                      ?disabled=${entry.disabled ?? false}
                      @click=${() => {
                        this._closeMenu(true);
                        choose(entry.action);
                      }}
                    >${entry.label}</button>`
              )}
            </div>`
          : nothing}
      </div>
    `;
  }

  // ─── Render ────────────────────────────────────────────

  /**
   * The items to show, as groups with a separator between them. A group
   * name is a group of its own; items between two `|` are one.
   */
  private _layout(): Array<{ label?: string; items: ToolbarItem[] }> {
    const specs: readonly string[] =
      typeof this.items === 'string'
        ? this.items.split(/[\s,]+/).filter(Boolean)
        : this.items ?? (Object.keys(TOOLBAR_GROUPS) as ToolbarGroup[]);
    const groups: Array<{ label?: string; items: ToolbarItem[] }> = [];
    const seen = new Set<string>();
    let open: ToolbarItem[] | null = null;
    const take = (item: ToolbarItem) => {
      if (seen.has(item)) return false;
      seen.add(item);
      return true;
    };
    for (const spec of specs) {
      if (spec === '|') {
        open = null;
      } else if (Object.prototype.hasOwnProperty.call(TOOLBAR_GROUPS, spec)) {
        const items = TOOLBAR_GROUPS[spec as ToolbarGroup].filter(take);
        if (items.length > 0) groups.push({ label: GROUP_LABELS[spec as ToolbarGroup], items });
        open = null;
      } else if ((TOOLBAR_ITEMS as readonly string[]).includes(spec)) {
        if (!take(spec as ToolbarItem)) continue;
        if (!open) groups.push({ items: (open = []) });
        open.push(spec as ToolbarItem);
      } else if (!this._warned.has(spec)) {
        this._warned.add(spec);
        console.warn(`<y11n-toolbar>: "${spec}" is not a toolbar item or group`);
      }
    }
    return groups;
  }

  private _button(
    item: ToolbarItem,
    label: string,
    content: unknown,
    onClick: () => void,
    options: { disabled?: boolean; pressed?: boolean; shortcut?: string; className?: string } = {}
  ) {
    return html`<button
      class="${options.className ?? ''}"
      type="button"
      data-item="${item}"
      aria-pressed="${options.pressed ?? nothing}"
      aria-label="${label}"
      title="${options.shortcut ? `${label} (${options.shortcut})` : label}"
      ?disabled=${options.disabled ?? false}
      @click=${onClick}
    >${content}</button>`;
  }

  private _renderItem(item: ToolbarItem): TemplateResult {
    const d = this.readOnly;
    switch (item) {
      case 'undo':
        return this._button(item, 'Undo', icon('undo'), () => this._act('undo'), {
          disabled: d || !this.canUndo,
          shortcut: 'Ctrl+Z',
        });
      case 'redo':
        return this._button(item, 'Redo', icon('redo'), () => this._act('redo'), {
          disabled: d || !this.canRedo,
          shortcut: 'Ctrl+Y',
        });
      case 'cut':
        return this._button(item, 'Cut', icon('cut'), () => this._act('cut'), { disabled: d, shortcut: 'Ctrl+X' });
      case 'copy':
        return this._button(item, 'Copy', icon('copy'), () => this._act('copy'), { shortcut: 'Ctrl+C' });
      case 'paste':
        return this._button(item, 'Paste', icon('paste'), () => this._act('paste'), { disabled: d, shortcut: 'Ctrl+V' });
      case 'pasteSpecial':
        return this._button(item, 'Paste Special', icon('pasteSpecial'), () => this._act('pasteSpecial'), {
          disabled: d,
          shortcut: 'Ctrl+Alt+V',
        });

      case 'bold':
        return this._button(item, 'Bold', 'B', () => this._format('bold', !this.bold), {
          disabled: d,
          pressed: this.bold,
          shortcut: 'Ctrl+B',
          className: 'bold-btn',
        });
      case 'italic':
        return this._button(item, 'Italic', 'I', () => this._format('italic', !this.italic), {
          disabled: d,
          pressed: this.italic,
          shortcut: 'Ctrl+I',
          className: 'italic-btn',
        });
      case 'underline':
        return this._button(item, 'Underline', 'U', () => this._format('underline', !this.underline), {
          disabled: d,
          pressed: this.underline,
          shortcut: 'Ctrl+U',
          className: 'underline-btn',
        });
      case 'strikethrough':
        return this._button(item, 'Strikethrough', 'S', () => this._format('strikethrough', !this.strikethrough), {
          disabled: d,
          pressed: this.strikethrough,
          className: 'strikethrough-btn',
        });
      case 'fontSize':
        return html`<select
          class="font-size-select"
          data-item="${item}"
          .value=${live(String(this.fontSize))}
          ?disabled=${d}
          aria-label="Font size"
          title="Font size"
          @change=${(e: Event) => this._format('fontSize', Number((e.target as HTMLSelectElement).value))}
        >
          ${FONT_SIZES.map(
            (size) => html`<option value=${size} ?selected=${this.fontSize === size}>${size}</option>`
          )}
        </select>`;
      case 'textColor':
        return html`<div class="color-wrapper ${d ? 'disabled' : ''}">
          <div class="color-swatch" aria-hidden="true">
            A<span class="indicator" style="background-color: ${this.textColor || 'currentColor'}"></span>
          </div>
          <input
            type="color"
            data-item="${item}"
            .value=${live(colorInputValue(this.textColor, '#000000'))}
            ?disabled=${d}
            aria-label="Text color"
            title="Text color"
            @input=${(e: Event) => this._format('textColor', (e.target as HTMLInputElement).value)}
          />
        </div>`;
      case 'backgroundColor':
        return html`<div class="color-wrapper ${d ? 'disabled' : ''}">
          <div class="color-swatch" aria-hidden="true">
            <span class="indicator fill" style="background-color: ${this.backgroundColor || 'transparent'}"></span>
          </div>
          <input
            type="color"
            data-item="${item}"
            .value=${live(colorInputValue(this.backgroundColor, '#ffffff'))}
            ?disabled=${d}
            aria-label="Background color"
            title="Background color"
            @input=${(e: Event) => this._format('backgroundColor', (e.target as HTMLInputElement).value)}
          />
        </div>`;

      case 'textAlign':
        return html`<div class="button-group" role="group" aria-label="Text alignment">
          ${(['left', 'center', 'right'] as const).map(
            (align) => html`<button
              type="button"
              data-item="${item}-${align}"
              aria-pressed="${this.textAlign === align}"
              aria-label="Align ${align}"
              title="Align ${align}"
              ?disabled=${d}
              @click=${() => this._format('textAlign', align)}
            ><span class="align-icon align-${align}" aria-hidden="true"><span></span><span></span><span></span></span></button>`
          )}
        </div>`;
      case 'verticalAlign':
        return html`<div class="button-group" role="group" aria-label="Vertical alignment">
          ${(['top', 'middle', 'bottom'] as const).map(
            (align) => html`<button
              type="button"
              data-item="${item}-${align}"
              aria-pressed="${this.verticalAlign === align}"
              aria-label="Align ${align}"
              title="Align ${align}"
              ?disabled=${d}
              @click=${() => this._format('verticalAlign', align)}
            >${icon(align)}</button>`
          )}
        </div>`;
      case 'wrap':
        return this._button(item, 'Wrap text', icon('wrap'), () => this._format('wrap', !this.wrap), {
          disabled: d,
          pressed: this.wrap,
        });

      case 'numberFormat':
        return html`<select
          class="number-format-select"
          data-item="${item}"
          .value=${live(this.numberFormat?.type ?? 'general')}
          ?disabled=${d}
          aria-label="Number format"
          title="Number format"
          @change=${(e: Event) => {
            const chosen = NUMBER_FORMATS.find(([value]) => value === (e.target as HTMLSelectElement).value);
            if (chosen) this._format('numberFormat', chosen[2]);
          }}
        >
          ${NUMBER_FORMATS.map(
            ([value, label]) =>
              html`<option value=${value} ?selected=${(this.numberFormat?.type ?? 'general') === value}>${label}</option>`
          )}
        </select>`;
      case 'decreaseDecimals':
        return this._button(item, 'Decrease decimals', icon('decreaseDecimals'), () => this._act('decreaseDecimals'), {
          disabled: d,
        });
      case 'increaseDecimals':
        return this._button(item, 'Increase decimals', icon('increaseDecimals'), () => this._act('increaseDecimals'), {
          disabled: d,
        });

      case 'conditionalFormat':
        // One-step rules, clear, and the rule manager
        return this._renderMenu(
          item,
          'Conditional formatting',
          'conditionalFormat',
          CONDITIONAL_MENU,
          (action) => this._conditional(action),
          d
        );
      case 'clearFormat':
        return this._button(item, 'Clear formatting', icon('clearFormat'), () => this._format('clearFormat'), {
          disabled: d,
        });

      case 'insert':
        return this._renderMenu(item, 'Insert rows or columns', 'insert', INSERT_MENU, (action) => this._act(action), d);
      case 'delete':
        return this._renderMenu(item, 'Delete rows or columns', 'delete', DELETE_MENU, (action) => this._act(action), d);
      case 'autoFit':
        return this._button(item, 'Fit column width', icon('autoFit'), () => this._act('autoFit'), {
          shortcut: 'Alt+Shift+0',
        });

      case 'sortAscending':
        return this._button(item, 'Sort ascending', icon('sortAscending'), () => this._act('sortAscending'), {
          disabled: d,
          shortcut: 'Alt+Shift+Up',
        });
      case 'sortDescending':
        return this._button(item, 'Sort descending', icon('sortDescending'), () => this._act('sortDescending'), {
          disabled: d,
          shortcut: 'Alt+Shift+Down',
        });
      case 'filter':
        return this._button(item, 'Filter', icon('filter'), () => this._act('filter'), {
          pressed: this.filtered,
          shortcut: 'Ctrl+Shift+L',
        });
      case 'reapplyFilter':
        return this._button(item, 'Reapply filter', icon('reapplyFilter'), () => this._act('reapplyFilter'), {
          disabled: !this.filtered,
          shortcut: 'Ctrl+Alt+L',
        });
      case 'list':
        return this._button(item, 'List in cell', icon('list'), () => this._act('list'), {
          disabled: d,
          shortcut: 'Ctrl+J',
        });
      case 'autoSum':
        return this._button(item, 'AutoSum', 'Σ', () => this._act('autoSum'), {
          disabled: d,
          shortcut: 'Alt+=',
          className: 'sum-btn',
        });

      case 'freeze':
        return this._renderMenu<ToolbarAction>(
          item,
          'Freeze panes',
          'freeze',
          [
            { action: 'freeze', label: 'Freeze up to the active cell' },
            { action: 'freezeFirstRow', label: 'Freeze first row' },
            { action: 'freezeFirstColumn', label: 'Freeze first column' },
            null,
            { action: 'unfreeze', label: 'Unfreeze panes', disabled: !this.frozen },
          ],
          (action) => this._act(action),
          false
        );
      case 'find':
        return this._button(item, 'Find', icon('find'), () => this._act('find'), { shortcut: 'Ctrl+F' });
      case 'names':
        return this._button(item, 'Name manager', icon('names'), () => this._act('names'), { shortcut: 'Ctrl+F3' });
    }
  }

  render() {
    const groups = this._layout();
    return html`
      <div
        class="toolbar"
        role="toolbar"
        aria-label="${this.label}"
        @keydown=${this._handleToolbarKeydown}
        @focusin=${this._handleToolbarFocusin}
      >
        ${groups.map(
          (group, i) => html`
            ${i > 0 ? html`<div class="separator" role="separator" aria-orientation="vertical"></div>` : nothing}
            <div class="group" role="${group.label ? 'group' : nothing}" aria-label="${group.label ?? nothing}">
              ${group.items.map((item) => this._renderItem(item))}
            </div>
          `
        )}
      </div>
    `;
  }
}

/**
 * Compute the consensus format from all cells in a selection range.
 * For each property, returns the common value if all cells agree, or undefined if mixed.
 *
 * With a grid at hand, `sheet.getSelectionFormat()` does this for its
 * selection (and a toolbar with a `target` needs neither).
 *
 * @param getData - Callback to retrieve cell data by key
 * @param range - The selection range to inspect
 * @returns A CellFormat with only the unanimously shared properties set
 */
export function computeSelectionFormat(
  getData: (cellId: string) => { format?: CellFormat } | undefined,
  range: SelectionRange
): CellFormat {
  const props = [
    'bold',
    'italic',
    'underline',
    'strikethrough',
    'textColor',
    'backgroundColor',
    'textAlign',
    'verticalAlign',
    'wrap',
    'fontSize',
    'numberFormat',
  ] as const;
  const same = (a: unknown, b: unknown) => a === b || JSON.stringify(a) === JSON.stringify(b);
  // A toggle that is off is the same as one that is not set
  const valueOf = (format: CellFormat | undefined, prop: (typeof props)[number]) => format?.[prop] || undefined;

  const shared: Record<string, unknown> = {};
  let first = true;
  for (let r = range.start.row; r <= range.end.row; r++) {
    for (let c = range.start.col; c <= range.end.col; c++) {
      const format = getData(`${r}:${c}`)?.format;
      for (const prop of props) {
        const value = valueOf(format, prop);
        if (first) shared[prop] = value;
        else if (!same(shared[prop], value)) shared[prop] = undefined;
      }
      first = false;
    }
  }

  const result: Record<string, unknown> = {};
  for (const prop of props) {
    if (shared[prop] !== undefined) result[prop] = shared[prop];
  }
  return result as CellFormat;
}

declare global {
  interface HTMLElementTagNameMap {
    'y11n-toolbar': Y11nToolbar;
  }
}
