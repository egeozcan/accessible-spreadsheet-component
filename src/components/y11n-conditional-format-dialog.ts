import { LitElement, html, css, nothing } from 'lit';
import { customElement, property, state } from 'lit/decorators.js';
import { repeat } from 'lit/directives/repeat.js';
import { styleMap } from 'lit/directives/style-map.js';
import type {
  ConditionalComparison,
  ConditionalDatePeriod,
  ConditionalFormatCondition,
  ConditionalFormatRule,
  ConditionalFormatRuleInput,
  ConditionalFormatStyle,
  ConditionalIconSetName,
  ConditionalOperand,
  ConditionalTextOperator,
  SelectionRange,
} from '../types.js';
import {
  DATE_PERIOD_TEXT,
  HIGHLIGHT_STYLES,
  ICON_SET_LABELS,
  ICON_SET_NAMES,
  describeCondition,
  isGraphicCondition,
  normalizeConditionalRule,
  parseA1Range,
  parseHexColor,
  rangeToA1,
} from '../utils/conditional-format.js';
import { parseInput } from '../utils/input-parse.js';
import { conditionalIcon, conditionalIconStyles } from './conditional-icons.js';

/** Detail for the dialog's `cf-save` event: a new rule (`id: null`) or a changed one. */
export interface ConditionalFormatSaveDetail {
  id: string | null;
  rule: ConditionalFormatRuleInput;
}

/** Detail for the dialog's `cf-move` event: the rule and its new position (0 = first). */
export interface ConditionalFormatMoveDetail {
  id: string;
  index: number;
}

type RuleKind = ConditionalFormatCondition['type'];
type StylePreset = keyof typeof HIGHLIGHT_STYLES | 'custom';

/** The edit form's fields. Everything is kept as typed until the rule is saved. */
interface Draft {
  /** Rule being edited; null for a new rule */
  id: string | null;
  range: string;
  kind: RuleKind;
  comparison: ConditionalComparison;
  value: string;
  value2: string;
  textOperator: ConditionalTextOperator;
  text: string;
  period: ConditionalDatePeriod;
  rank: string;
  percent: boolean;
  orEqual: boolean;
  formula: string;
  stylePreset: StylePreset;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  useTextColor: boolean;
  textColor: string;
  useFill: boolean;
  fillColor: string;
  stopIfTrue: boolean;
  threeColors: boolean;
  minColor: string;
  midColor: string;
  maxColor: string;
  customBarColor: boolean;
  barColor: string;
  showValue: boolean;
  icons: ConditionalIconSetName;
  reverse: boolean;
  label: string;
}

const KIND_GROUPS: Array<[string, Array<[RuleKind, string]>]> = [
  [
    'Highlight cells',
    [
      ['cellValue', 'Cell value'],
      ['text', 'Text'],
      ['date', 'Date'],
      ['blanks', 'Blank cells'],
      ['notBlanks', 'Cells that are not blank'],
      ['errors', 'Errors'],
      ['notErrors', 'Cells without errors'],
      ['duplicates', 'Duplicate values'],
      ['unique', 'Unique values'],
    ],
  ],
  [
    'Top and bottom',
    [
      ['top', 'Top values'],
      ['bottom', 'Bottom values'],
      ['aboveAverage', 'Above average'],
      ['belowAverage', 'Below average'],
    ],
  ],
  ['Formula', [['formula', 'Formula is true']]],
  [
    'Graphics in the cell',
    [
      ['colorScale', 'Color scale'],
      ['dataBar', 'Data bar'],
      ['iconSet', 'Icon set'],
    ],
  ],
];

const COMPARISONS: Array<[ConditionalComparison, string]> = [
  ['greaterThan', 'greater than'],
  ['greaterThanOrEqual', 'greater than or equal to'],
  ['lessThan', 'less than'],
  ['lessThanOrEqual', 'less than or equal to'],
  ['equal', 'equal to'],
  ['notEqual', 'not equal to'],
  ['between', 'between'],
  ['notBetween', 'not between'],
];

const TEXT_OPERATORS: Array<[ConditionalTextOperator, string]> = [
  ['contains', 'contains'],
  ['notContains', 'does not contain'],
  ['beginsWith', 'begins with'],
  ['endsWith', 'ends with'],
];

const STYLE_PRESETS: Array<[StylePreset, string]> = [
  ['red', 'Light red fill with dark red text'],
  ['yellow', 'Yellow fill with dark yellow text'],
  ['green', 'Green fill with dark green text'],
  ['custom', 'Custom'],
];

const DEFAULT_BAR_COLOR = '#8fb8ea';

/** A color for `<input type="color">`, which only takes `#rrggbb`. */
function colorInputValue(color: string, fallback: string): string {
  const rgb = parseHexColor(color);
  return rgb ? `#${rgb.map((n) => n.toString(16).padStart(2, '0')).join('')}` : fallback;
}

/** Inline styles that show a rule's format, for previews. */
function styleCss(format: ConditionalFormatStyle | undefined): Record<string, string> {
  const css: Record<string, string> = {};
  if (!format) return css;
  if (format.bold) css['font-weight'] = 'bold';
  if (format.italic) css['font-style'] = 'italic';
  const lines = [format.underline ? 'underline' : '', format.strikethrough ? 'line-through' : ''].filter(Boolean);
  if (lines.length > 0) css['text-decoration'] = lines.join(' ');
  if (format.textColor) css.color = format.textColor;
  if (format.backgroundColor) css['background-color'] = format.backgroundColor;
  return css;
}

function capitalize(text: string): string {
  return text.charAt(0).toUpperCase() + text.slice(1);
}

function presetOf(format: ConditionalFormatStyle | undefined): StylePreset {
  for (const [name, style] of Object.entries(HIGHLIGHT_STYLES)) {
    if (
      format &&
      Object.keys(format).length === 2 &&
      format.backgroundColor === style.backgroundColor &&
      format.textColor === style.textColor
    ) {
      return name as StylePreset;
    }
  }
  return 'custom';
}

function draftFor(rule: ConditionalFormatRule | null, selection: SelectionRange): Draft {
  const draft: Draft = {
    id: rule?.id ?? null,
    range: rangeToA1(rule?.range ?? selection),
    kind: rule?.condition.type ?? 'cellValue',
    comparison: 'greaterThan',
    value: '',
    value2: '',
    textOperator: 'contains',
    text: '',
    period: 'today',
    rank: '10',
    percent: false,
    orEqual: false,
    formula: '',
    stylePreset: rule ? presetOf(rule.format) : 'red',
    bold: !!rule?.format?.bold,
    italic: !!rule?.format?.italic,
    underline: !!rule?.format?.underline,
    strikethrough: !!rule?.format?.strikethrough,
    useTextColor: !!rule?.format?.textColor,
    textColor: rule?.format?.textColor ?? HIGHLIGHT_STYLES.red.textColor,
    useFill: !!rule?.format?.backgroundColor,
    fillColor: rule?.format?.backgroundColor ?? HIGHLIGHT_STYLES.red.backgroundColor,
    stopIfTrue: !!rule?.stopIfTrue,
    threeColors: true,
    minColor: '#f8696b',
    midColor: '#ffeb84',
    maxColor: '#63be7b',
    customBarColor: false,
    barColor: DEFAULT_BAR_COLOR,
    showValue: true,
    icons: 'arrows3',
    reverse: false,
    label: rule?.label ?? '',
  };
  const c = rule?.condition;
  if (!c) return draft;
  switch (c.type) {
    case 'cellValue':
      draft.comparison = c.operator;
      draft.value = String(c.value);
      draft.value2 = c.value2 === undefined ? '' : String(c.value2);
      break;
    case 'text':
      draft.textOperator = c.operator;
      draft.text = c.text;
      break;
    case 'date':
      draft.period = c.period;
      break;
    case 'top':
    case 'bottom':
      draft.rank = String(c.rank);
      draft.percent = !!c.percent;
      break;
    case 'aboveAverage':
    case 'belowAverage':
      draft.orEqual = !!c.orEqual;
      break;
    case 'formula':
      draft.formula = c.formula;
      break;
    case 'colorScale':
      draft.threeColors = !!c.mid;
      draft.minColor = c.min.color;
      draft.midColor = c.mid?.color ?? draft.midColor;
      draft.maxColor = c.max.color;
      break;
    case 'dataBar':
      draft.customBarColor = c.color !== undefined;
      draft.barColor = c.color ?? DEFAULT_BAR_COLOR;
      draft.showValue = c.showValue !== false;
      break;
    case 'iconSet':
      draft.icons = c.icons;
      draft.reverse = !!c.reverse;
      draft.showValue = c.showValue !== false;
      break;
  }
  return draft;
}

/** What was typed for a `cellValue` rule: a formula, a number (`15%`, `$1,200`, a date) or text. */
function operandFrom(text: string, name: string): ConditionalOperand {
  const trimmed = text.trim();
  if (trimmed === '') throw new TypeError(`Enter ${name}`);
  if (trimmed.startsWith('=')) return trimmed;
  const num = Number(trimmed);
  if (Number.isFinite(num)) return num;
  const parsed = parseInput(trimmed);
  return parsed ? Number(parsed.value) : trimmed;
}

function styleFrom(draft: Draft): ConditionalFormatStyle {
  if (draft.stylePreset !== 'custom') return { ...HIGHLIGHT_STYLES[draft.stylePreset] };
  const style: ConditionalFormatStyle = {};
  if (draft.bold) style.bold = true;
  if (draft.italic) style.italic = true;
  if (draft.underline) style.underline = true;
  if (draft.strikethrough) style.strikethrough = true;
  if (draft.useTextColor) style.textColor = draft.textColor;
  if (draft.useFill) style.backgroundColor = draft.fillColor;
  return style;
}

/**
 * The condition the form describes. Settings the form has no fields for
 * (custom thresholds, icon labels) are carried over from `original` when the
 * rule keeps its type.
 */
function conditionFrom(draft: Draft, original: ConditionalFormatCondition | undefined): ConditionalFormatCondition {
  switch (draft.kind) {
    case 'cellValue': {
      const value = operandFrom(draft.value, 'a value to compare with');
      if (draft.comparison === 'between' || draft.comparison === 'notBetween') {
        return {
          type: 'cellValue',
          operator: draft.comparison,
          value,
          value2: operandFrom(draft.value2, 'the second value'),
        };
      }
      return { type: 'cellValue', operator: draft.comparison, value };
    }
    case 'text':
      if (draft.text === '') throw new TypeError('Enter the text to look for');
      return { type: 'text', operator: draft.textOperator, text: draft.text };
    case 'date':
      return { type: 'date', period: draft.period };
    case 'top':
    case 'bottom': {
      const rank = Number(draft.rank);
      if (!Number.isInteger(rank) || rank < 1 || (draft.percent && rank > 100)) {
        throw new TypeError(draft.percent ? 'Enter a whole percentage from 1 to 100' : 'Enter a whole number of 1 or more');
      }
      return { type: draft.kind, rank, percent: draft.percent };
    }
    case 'aboveAverage':
    case 'belowAverage':
      return { type: draft.kind, orEqual: draft.orEqual };
    case 'formula':
      if (draft.formula.trim() === '' || draft.formula.trim() === '=') throw new TypeError('Enter a formula');
      return { type: 'formula', formula: draft.formula };
    case 'colorScale': {
      const kept = original?.type === 'colorScale' ? original : undefined;
      const scale: ConditionalFormatCondition = {
        type: 'colorScale',
        min: { ...(kept?.min ?? { type: 'min' }), color: draft.minColor },
        max: { ...(kept?.max ?? { type: 'max' }), color: draft.maxColor },
      };
      if (draft.threeColors) {
        scale.mid = { ...(kept?.mid ?? { type: 'percentile', value: 50 }), color: draft.midColor };
      }
      return scale;
    }
    case 'dataBar': {
      const bar: ConditionalFormatCondition = { ...(original?.type === 'dataBar' ? original : {}), type: 'dataBar' };
      if (draft.customBarColor) bar.color = draft.barColor;
      else delete bar.color;
      bar.showValue = draft.showValue;
      return bar;
    }
    case 'iconSet': {
      // Thresholds and labels belong to one icon set: drop them when it changes
      const kept = original?.type === 'iconSet' && original.icons === draft.icons ? original : {};
      return { ...kept, type: 'iconSet', icons: draft.icons, reverse: draft.reverse, showValue: draft.showValue };
    }
    default:
      return { type: draft.kind };
  }
}

/**
 * `<y11n-conditional-format-dialog>` -- The rule manager of `<y11n-spreadsheet>`:
 * a modal dialog that lists the conditional-format rules in priority order
 * (edit, reorder, delete, stop-if-true) and has a form to add or change one.
 *
 * It holds no rules itself. The spreadsheet passes them in and applies the
 * changes it asks for:
 *
 * @fires cf-save - A rule was added (`id: null`) or changed
 * @fires cf-remove - Delete a rule (`detail.id`)
 * @fires cf-move - Move a rule to another position
 * @fires cf-close - The dialog was dismissed
 */
@customElement('y11n-conditional-format-dialog')
export class Y11nConditionalFormatDialog extends LitElement {
  /** The rules in priority order */
  @property({ attribute: false }) rules: readonly ConditionalFormatRule[] = [];
  /** The grid's selection: the range a new rule starts with */
  @property({ attribute: false }) selection: SelectionRange = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };
  /** Grid size, to check ranges against */
  @property({ type: Number }) rows = 100;
  @property({ type: Number }) cols = 26;
  /** When true, rules can be looked at but not changed */
  @property({ type: Boolean, attribute: 'read-only', reflect: true }) readOnly = false;
  /** Outcome of the last change, announced from inside the dialog */
  @property({ type: String }) status = '';

  @state() private _draft: Draft | null = null;
  @state() private _error = '';
  private _pendingFocus: { selector: string; fallback: string } | null = null;

  static styles = [
    conditionalIconStyles,
    css`
    :host {
      position: absolute;
      top: 8px;
      left: 50%;
      transform: translateX(-50%);
      z-index: 20;
      display: block;
      box-sizing: border-box;
      width: min(540px, calc(100% - 16px));
      max-height: calc(100% - 16px);
      overflow: auto;
      background: var(--ls-filter-menu-bg, #fff);
      color: var(--ls-text-color, #333);
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 4px;
      box-shadow: var(--ls-editor-shadow, 0 2px 6px rgba(0, 0, 0, 0.2));
      font-family: var(--ls-font-family, system-ui, -apple-system, sans-serif);
      font-size: 13px;
    }

    *,
    *::before,
    *::after {
      box-sizing: border-box;
    }

    .dialog,
    form {
      display: flex;
      flex-direction: column;
      gap: 8px;
    }

    .dialog {
      padding: 10px;
    }

    h2 {
      margin: 0;
      font-size: 14px;
      font-weight: 600;
    }

    ol {
      display: flex;
      flex-direction: column;
      gap: 6px;
      margin: 0;
      padding: 0;
      list-style: none;
    }

    li {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 6px 8px;
      padding: 6px;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 3px;
    }

    .rule-main {
      display: flex;
      flex: 1 1 180px;
      align-items: center;
      gap: 8px;
      min-width: 0;
    }

    .rule-text {
      display: flex;
      flex-direction: column;
      min-width: 0;
    }

    .rule-desc {
      overflow-wrap: anywhere;
    }

    .rule-range,
    .hint,
    .empty {
      font-size: 12px;
      color: var(--ls-find-status-color, #555);
    }

    .rule-actions {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 4px;
    }

    .swatch {
      position: relative;
      flex: 0 0 auto;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 44px;
      height: 22px;
      overflow: hidden;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 2px;
    }

    .swatch-bar {
      position: absolute;
      inset: 3px 30% 3px 3px;
      background: var(--ls-data-bar-color, #8fb8ea);
    }

    .preview {
      align-self: flex-start;
      padding: 3px 12px;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 2px;
    }

    .field {
      display: flex;
      flex-direction: column;
      gap: 2px;
    }

    .row {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 6px 10px;
    }

    .check {
      display: inline-flex;
      align-items: center;
      gap: 5px;
      cursor: pointer;
    }

    fieldset {
      display: flex;
      flex-direction: column;
      gap: 6px;
      margin: 0;
      padding: 6px 8px;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 3px;
    }

    input[type='text'],
    input[type='number'],
    select {
      min-height: 26px;
      padding: 2px 6px;
      font: inherit;
      color: inherit;
      background: var(--ls-editor-bg, #fff);
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 3px;
    }

    input[type='number'] {
      width: 6em;
    }

    input[type='color'] {
      width: 36px;
      height: 26px;
      padding: 1px;
      background: var(--ls-editor-bg, #fff);
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 3px;
    }

    button {
      min-height: 26px;
      padding: 2px 8px;
      font: inherit;
      color: inherit;
      background: transparent;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 3px;
      cursor: pointer;
    }

    button:hover:not(:disabled):not([aria-disabled='true']) {
      background: var(--ls-header-bg, #f3f3f3);
    }

    button.primary {
      color: #fff;
      background: var(--ls-filter-active-color, #1a73e8);
      border-color: var(--ls-filter-active-color, #1a73e8);
    }

    button.primary:hover:not(:disabled) {
      background: var(--ls-filter-active-color, #1a73e8);
      filter: brightness(0.92);
    }

    button:disabled,
    button[aria-disabled='true'] {
      opacity: 0.5;
      cursor: default;
    }

    button:focus-visible,
    input:focus-visible,
    select:focus-visible {
      outline: var(--ls-focus-ring, 2px solid #1a73e8);
      outline-offset: 1px;
    }

    .buttons {
      display: flex;
      justify-content: flex-end;
      gap: 6px;
    }

    .status {
      min-height: 1.2em;
      font-size: 12px;
      color: var(--ls-find-status-color, #555);
    }

    .error {
      padding: 4px 8px;
      color: var(--ls-validation-color, #8a1c12);
      background: var(--ls-validation-bg, #fdecea);
      border: 1px solid var(--ls-validation-border, #d93025);
      border-radius: 3px;
    }

    .icons {
      display: inline-flex;
      align-items: center;
      gap: 4px;
    }

    @media (forced-colors: active) {
      .swatch-bar {
        forced-color-adjust: none;
        background: CanvasText;
      }
    }
  `,
  ];

  /** Move focus into the dialog: the first field of the form, or the list's New rule button. */
  focusInitial(): void {
    const root = this.renderRoot;
    const target = this._draft
      ? root.querySelector<HTMLElement>('#cf-range')
      : (root.querySelector<HTMLElement>('#cf-new:not(:disabled)') ?? root.querySelector<HTMLElement>('#cf-close'));
    target?.focus();
  }

  private _emit<T>(name: string, detail?: T): void {
    this.dispatchEvent(new CustomEvent<T>(name, { detail, bubbles: true, composed: true }));
  }

  /**
   * Focus an element once the dialog has re-rendered. A change is applied by
   * the spreadsheet, which then passes new rules in: only after that update
   * are the rows in their final place (a row that moves loses focus).
   */
  private _focusAfterUpdate(selector: string, fallback = '#cf-new'): void {
    this._pendingFocus = { selector, fallback };
    this.requestUpdate();
  }

  protected updated(): void {
    const pending = this._pendingFocus;
    if (!pending) return;
    this._pendingFocus = null;
    const root = this.renderRoot;
    (root.querySelector<HTMLElement>(pending.selector) ?? root.querySelector<HTMLElement>(pending.fallback))?.focus();
  }

  private _handleKeydown(e: KeyboardEvent): void {
    // Keys typed in the dialog are not grid shortcuts
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      if (this._draft) this._closeForm();
      else this._emit('cf-close');
    } else if (e.key === 'Tab') {
      // Modal: keep focus inside
      const focusables = [
        ...this.renderRoot.querySelectorAll<HTMLElement>('button:not(:disabled), input:not(:disabled), select:not(:disabled)'),
      ];
      const first = focusables[0];
      const last = focusables[focusables.length - 1];
      const current = (this.renderRoot as ShadowRoot).activeElement;
      if (e.shiftKey && current === first) {
        e.preventDefault();
        last?.focus();
      } else if (!e.shiftKey && current === last) {
        e.preventDefault();
        first?.focus();
      }
    }
  }

  // ─── List ──────────────────────────────────────────────

  private _openForm(rule: ConditionalFormatRule | null): void {
    if (this.readOnly) return;
    this._draft = draftFor(rule, this.selection);
    this._error = '';
    this._focusAfterUpdate('#cf-range');
  }

  private _closeForm(): void {
    const id = this._draft?.id;
    this._draft = null;
    this._error = '';
    this._focusAfterUpdate(id ? `[data-rule="${id}"][data-action="edit"]` : '#cf-new');
  }

  private _move(rule: ConditionalFormatRule, index: number, delta: number): void {
    const target = index + delta;
    if (this.readOnly || target < 0 || target >= this.rules.length) return;
    this._emit<ConditionalFormatMoveDetail>('cf-move', { id: rule.id, index: target });
    // The button moved with its rule; at either end, the opposite button still works
    const action = target === 0 ? 'down' : target === this.rules.length - 1 ? 'up' : delta < 0 ? 'up' : 'down';
    this._focusAfterUpdate(`[data-rule="${rule.id}"][data-action="${action}"]`);
  }

  private _remove(rule: ConditionalFormatRule, index: number): void {
    if (this.readOnly) return;
    const neighbor = this.rules[index + 1] ?? this.rules[index - 1];
    this._emit('cf-remove', { id: rule.id });
    this._focusAfterUpdate(neighbor ? `[data-rule="${neighbor.id}"][data-action="edit"]` : '#cf-new');
  }

  private _toggleStop(rule: ConditionalFormatRule, stopIfTrue: boolean): void {
    if (this.readOnly) return;
    const { id, ...rest } = rule;
    this._emit<ConditionalFormatSaveDetail>('cf-save', { id, rule: { ...rest, stopIfTrue } });
  }

  private _renderSwatch(rule: ConditionalFormatRule) {
    const c = rule.condition;
    if (c.type === 'colorScale') {
      const colors = [c.min.color, ...(c.mid ? [c.mid.color] : []), c.max.color].join(', ');
      return html`<span class="swatch" aria-hidden="true" style="background: linear-gradient(90deg, ${colors})"></span>`;
    }
    if (c.type === 'dataBar') {
      return html`<span class="swatch" aria-hidden="true"
        ><span class="swatch-bar" style=${styleMap(c.color ? { background: c.color } : {})}></span
      ></span>`;
    }
    if (c.type === 'iconSet') {
      const top = ICON_SET_LABELS[c.icons].length - 1;
      return html`<span class="swatch" aria-hidden="true">${conditionalIcon(c.icons, c.reverse ? 0 : top)}</span>`;
    }
    return html`<span class="swatch" aria-hidden="true" style=${styleMap(styleCss(rule.format))}>Aa</span>`;
  }

  private _renderList() {
    const count = this.rules.length;
    return html`
      <h2 id="cf-title">Conditional formatting</h2>
      ${count === 0
        ? html`<p class="empty">No rules yet.</p>`
        : html`<p class="hint" id="cf-order">Rules apply from the top: where two rules set the same thing, the higher one wins.</p>
            <ol aria-label="Rules, highest priority first" aria-describedby="cf-order">
              ${repeat(
                this.rules,
                (rule) => rule.id,
                (rule, i) => {
                  const descId = `cf-desc-${i}`;
                  return html`
                    <li>
                      <div class="rule-main" id="${descId}">
                        ${this._renderSwatch(rule)}
                        <span class="rule-text">
                          <span class="rule-desc">${capitalize(describeCondition(rule.condition))}${
                            rule.label ? html` (${rule.label})` : nothing
                          }</span>
                          <span class="rule-range">${rangeToA1(rule.range)}</span>
                        </span>
                      </div>
                      <div class="rule-actions">
                        ${isGraphicCondition(rule.condition)
                          ? nothing
                          : html`<label class="check">
                              <input
                                type="checkbox"
                                data-rule="${rule.id}"
                                data-action="stop"
                                aria-describedby="${descId}"
                                .checked=${!!rule.stopIfTrue}
                                ?disabled=${this.readOnly}
                                @change="${(e: Event) => this._toggleStop(rule, (e.target as HTMLInputElement).checked)}"
                              />
                              <span>Stop if true</span>
                            </label>`}
                        <button
                          type="button"
                          data-rule="${rule.id}"
                          data-action="edit"
                          aria-describedby="${descId}"
                          ?disabled=${this.readOnly}
                          @click="${() => this._openForm(rule)}"
                        >Edit</button>
                        <button
                          type="button"
                          data-rule="${rule.id}"
                          data-action="up"
                          aria-label="Move up"
                          aria-describedby="${descId}"
                          title="Move up"
                          aria-disabled="${this.readOnly || i === 0}"
                          @click="${() => this._move(rule, i, -1)}"
                        ><span aria-hidden="true">&#x2191;</span></button>
                        <button
                          type="button"
                          data-rule="${rule.id}"
                          data-action="down"
                          aria-label="Move down"
                          aria-describedby="${descId}"
                          title="Move down"
                          aria-disabled="${this.readOnly || i === count - 1}"
                          @click="${() => this._move(rule, i, 1)}"
                        ><span aria-hidden="true">&#x2193;</span></button>
                        <button
                          type="button"
                          data-rule="${rule.id}"
                          data-action="delete"
                          aria-describedby="${descId}"
                          ?disabled=${this.readOnly}
                          @click="${() => this._remove(rule, i)}"
                        >Delete</button>
                      </div>
                    </li>
                  `;
                }
              )}
            </ol>`}
      <div class="status" role="status">${this.status}</div>
      <div class="buttons">
        <button
          id="cf-new"
          type="button"
          class="primary"
          ?disabled=${this.readOnly}
          @click="${() => this._openForm(null)}"
        >New rule</button>
        <button id="cf-close" type="button" @click="${() => this._emit('cf-close')}">Close</button>
      </div>
    `;
  }

  // ─── Form ──────────────────────────────────────────────

  private _set(changes: Partial<Draft>): void {
    if (this._draft) this._draft = { ...this._draft, ...changes };
  }

  private _text(key: keyof Draft) {
    return (e: Event) => this._set({ [key]: (e.target as HTMLInputElement).value } as Partial<Draft>);
  }

  private _checked(key: keyof Draft) {
    return (e: Event) => this._set({ [key]: (e.target as HTMLInputElement).checked } as Partial<Draft>);
  }

  private _handleSubmit(e: Event): void {
    e.preventDefault();
    const draft = this._draft;
    if (!draft || this.readOnly) return;
    try {
      const range = parseA1Range(draft.range);
      if (!range) throw new TypeError('Enter a range such as A1:C10');
      if (range.end.row >= this.rows || range.end.col >= this.cols) {
        throw new TypeError('The range is outside the grid');
      }
      const original = this.rules.find((r) => r.id === draft.id);
      const condition = conditionFrom(draft, original?.condition);
      const rule: ConditionalFormatRuleInput = { range, condition };
      if (!isGraphicCondition(condition)) {
        rule.format = styleFrom(draft);
        rule.stopIfTrue = draft.stopIfTrue;
      }
      if (draft.label.trim() !== '') rule.label = draft.label;
      // Report anything else that is wrong with it here, not after the dialog closed
      normalizeConditionalRule(rule, draft.id ?? 'new');
      this._emit<ConditionalFormatSaveDetail>('cf-save', { id: draft.id, rule });
      this._closeForm();
    } catch (err) {
      this._error = (err as Error).message.replace(/^conditional format: /, '');
    }
  }

  private _renderCondition(draft: Draft) {
    switch (draft.kind) {
      case 'cellValue': {
        const two = draft.comparison === 'between' || draft.comparison === 'notBetween';
        return html`
          <div class="row">
            <label class="field">
              <span>Cell value is</span>
              <select id="cf-comparison" @change="${this._text('comparison')}">
                ${COMPARISONS.map(
                  ([value, label]) => html`<option value="${value}" ?selected=${draft.comparison === value}>${label}</option>`
                )}
              </select>
            </label>
            <label class="field">
              <span>${two ? 'First value' : 'Value'}</span>
              <input id="cf-value" type="text" autocomplete="off" .value="${draft.value}" @input="${this._text('value')}" />
            </label>
            ${two
              ? html`<label class="field">
                  <span>Second value</span>
                  <input id="cf-value2" type="text" autocomplete="off" .value="${draft.value2}" @input="${this._text('value2')}" />
                </label>`
              : nothing}
          </div>
          <span class="hint">A number, text, or a formula such as =$B$1.</span>
        `;
      }
      case 'text':
        return html`
          <div class="row">
            <label class="field">
              <span>Cell text</span>
              <select id="cf-text-operator" @change="${this._text('textOperator')}">
                ${TEXT_OPERATORS.map(
                  ([value, label]) => html`<option value="${value}" ?selected=${draft.textOperator === value}>${label}</option>`
                )}
              </select>
            </label>
            <label class="field">
              <span>Text</span>
              <input id="cf-text" type="text" autocomplete="off" .value="${draft.text}" @input="${this._text('text')}" />
            </label>
          </div>
        `;
      case 'date':
        return html`
          <label class="field">
            <span>Date is</span>
            <select id="cf-period" @change="${this._text('period')}">
              ${(Object.keys(DATE_PERIOD_TEXT) as ConditionalDatePeriod[]).map(
                (value) => html`<option value="${value}" ?selected=${draft.period === value}>${DATE_PERIOD_TEXT[value]}</option>`
              )}
            </select>
          </label>
        `;
      case 'top':
      case 'bottom':
        return html`
          <div class="row">
            <label class="field">
              <span>${draft.kind === 'top' ? 'How many of the highest values' : 'How many of the lowest values'}</span>
              <input id="cf-rank" type="number" min="1" step="1" .value="${draft.rank}" @input="${this._text('rank')}" />
            </label>
            <label class="check">
              <input id="cf-percent" type="checkbox" .checked=${draft.percent} @change="${this._checked('percent')}" />
              <span>Percent of the range</span>
            </label>
          </div>
        `;
      case 'aboveAverage':
      case 'belowAverage':
        return html`
          <label class="check">
            <input id="cf-or-equal" type="checkbox" .checked=${draft.orEqual} @change="${this._checked('orEqual')}" />
            <span>Include values equal to the average</span>
          </label>
        `;
      case 'formula': {
        const first = parseA1Range(draft.range);
        return html`
          <label class="field">
            <span>Format cells where this formula is true</span>
            <input
              id="cf-formula"
              type="text"
              autocomplete="off"
              spellcheck="false"
              aria-describedby="cf-formula-hint"
              .value="${draft.formula}"
              @input="${this._text('formula')}"
            />
          </label>
          <span class="hint" id="cf-formula-hint">Write it for the first cell of the range${
            first ? ` (${rangeToA1({ start: first.start, end: first.start })})` : ''
          }. It is adjusted for the other cells like a filled formula.</span>
        `;
      }
      case 'colorScale':
        return html`
          <label class="check">
            <input id="cf-three" type="checkbox" .checked=${draft.threeColors} @change="${this._checked('threeColors')}" />
            <span>Use a midpoint color</span>
          </label>
          <div class="row">
            <label class="check">
              <input
                id="cf-min-color"
                type="color"
                .value="${colorInputValue(draft.minColor, '#f8696b')}"
                @input="${this._text('minColor')}"
              />
              <span>Lowest</span>
            </label>
            ${draft.threeColors
              ? html`<label class="check">
                  <input
                    id="cf-mid-color"
                    type="color"
                    .value="${colorInputValue(draft.midColor, '#ffeb84')}"
                    @input="${this._text('midColor')}"
                  />
                  <span>Midpoint</span>
                </label>`
              : nothing}
            <label class="check">
              <input
                id="cf-max-color"
                type="color"
                .value="${colorInputValue(draft.maxColor, '#63be7b')}"
                @input="${this._text('maxColor')}"
              />
              <span>Highest</span>
            </label>
          </div>
          <span class="hint">Each cell also says where it is on the scale, as a percentage.</span>
        `;
      case 'dataBar':
        return html`
          <div class="row">
            <label class="check">
              <input
                id="cf-bar-custom"
                type="checkbox"
                .checked=${draft.customBarColor}
                @change="${this._checked('customBarColor')}"
              />
              <span>Custom bar color</span>
            </label>
            ${draft.customBarColor
              ? html`<label class="check">
                  <input
                    id="cf-bar-color"
                    type="color"
                    .value="${colorInputValue(draft.barColor, DEFAULT_BAR_COLOR)}"
                    @input="${this._text('barColor')}"
                  />
                  <span>Bar color</span>
                </label>`
              : nothing}
          </div>
          <label class="check">
            <input id="cf-show-value" type="checkbox" .checked=${draft.showValue} @change="${this._checked('showValue')}" />
            <span>Show the value beside the bar</span>
          </label>
        `;
      case 'iconSet': {
        const labels = ICON_SET_LABELS[draft.icons];
        const order = labels.map((_, i) => (draft.reverse ? labels.length - 1 - i : i));
        return html`
          <div class="row">
            <label class="field">
              <span>Icons</span>
              <select id="cf-icons" @change="${this._text('icons')}">
                ${(Object.keys(ICON_SET_NAMES) as ConditionalIconSetName[]).map(
                  (value) => html`<option value="${value}" ?selected=${draft.icons === value}>${ICON_SET_NAMES[value]}</option>`
                )}
              </select>
            </label>
            <span class="icons" role="img" aria-label="Lowest to highest: ${order.map((i) => labels[i]).join(', ')}">
              ${order.map((i) => conditionalIcon(draft.icons, i))}
            </span>
          </div>
          <label class="check">
            <input id="cf-reverse" type="checkbox" .checked=${draft.reverse} @change="${this._checked('reverse')}" />
            <span>Reverse the icon order</span>
          </label>
          <label class="check">
            <input id="cf-show-value" type="checkbox" .checked=${draft.showValue} @change="${this._checked('showValue')}" />
            <span>Show the value beside the icon</span>
          </label>
        `;
      }
      default:
        return nothing;
    }
  }

  private _renderStyle(draft: Draft) {
    const custom = draft.stylePreset === 'custom';
    return html`
      <fieldset>
        <legend>Format</legend>
        <div class="row">
          <label class="field">
            <span>Style</span>
            <select id="cf-style" @change="${this._text('stylePreset')}">
              ${STYLE_PRESETS.map(
                ([value, label]) => html`<option value="${value}" ?selected=${draft.stylePreset === value}>${label}</option>`
              )}
            </select>
          </label>
          <span class="preview" aria-hidden="true" style=${styleMap(styleCss(styleFrom(draft)))}>AaBbCc 123</span>
        </div>
        ${custom
          ? html`
              <div class="row">
                <label class="check">
                  <input id="cf-bold" type="checkbox" .checked=${draft.bold} @change="${this._checked('bold')}" />
                  <span>Bold</span>
                </label>
                <label class="check">
                  <input id="cf-italic" type="checkbox" .checked=${draft.italic} @change="${this._checked('italic')}" />
                  <span>Italic</span>
                </label>
                <label class="check">
                  <input id="cf-underline" type="checkbox" .checked=${draft.underline} @change="${this._checked('underline')}" />
                  <span>Underline</span>
                </label>
                <label class="check">
                  <input
                    id="cf-strikethrough"
                    type="checkbox"
                    .checked=${draft.strikethrough}
                    @change="${this._checked('strikethrough')}"
                  />
                  <span>Strikethrough</span>
                </label>
              </div>
              <div class="row">
                <label class="check">
                  <input
                    id="cf-use-text-color"
                    type="checkbox"
                    .checked=${draft.useTextColor}
                    @change="${this._checked('useTextColor')}"
                  />
                  <span>Text color</span>
                </label>
                <input
                  id="cf-text-color"
                  type="color"
                  aria-label="Text color"
                  ?disabled=${!draft.useTextColor}
                  .value="${colorInputValue(draft.textColor, '#9c0006')}"
                  @input="${this._text('textColor')}"
                />
                <label class="check">
                  <input id="cf-use-fill" type="checkbox" .checked=${draft.useFill} @change="${this._checked('useFill')}" />
                  <span>Fill color</span>
                </label>
                <input
                  id="cf-fill-color"
                  type="color"
                  aria-label="Fill color"
                  ?disabled=${!draft.useFill}
                  .value="${colorInputValue(draft.fillColor, '#ffc7ce')}"
                  @input="${this._text('fillColor')}"
                />
              </div>
            `
          : nothing}
        <label class="check">
          <input id="cf-stop" type="checkbox" .checked=${draft.stopIfTrue} @change="${this._checked('stopIfTrue')}" />
          <span>Stop if true (skip lower rules for cells that match)</span>
        </label>
      </fieldset>
    `;
  }

  /** What is read with a cell when the label field is left empty, as a sentence for the hint. */
  private _defaultLabelHint(draft: Draft): string {
    try {
      const condition = conditionFrom(draft, undefined);
      if (condition.type === 'colorScale') {
        return ' Left empty, it is “color scale” and the cell’s position on it, such as “color scale 40%”.';
      }
      if (condition.type === 'dataBar') {
        return ' Left empty, it is “data bar” and the bar’s position, such as “data bar 40%”.';
      }
      if (condition.type === 'iconSet') {
        const labels = ICON_SET_LABELS[condition.icons];
        return ` Left empty, the icon’s name is read, such as “${labels[labels.length - 1]}”.`;
      }
      return ` Left empty, it is “${describeCondition(condition)}”.`;
    } catch {
      // The form isn't filled in far enough to say
      return '';
    }
  }

  private _renderForm(draft: Draft) {
    const graphic = draft.kind === 'colorScale' || draft.kind === 'dataBar' || draft.kind === 'iconSet';
    const fallback = this._defaultLabelHint(draft);
    return html`
      <h2 id="cf-title">${draft.id ? 'Edit rule' : 'New rule'}</h2>
      <form novalidate @submit="${this._handleSubmit}">
        <label class="field">
          <span>Applies to</span>
          <input
            id="cf-range"
            type="text"
            autocomplete="off"
            spellcheck="false"
            .value="${draft.range}"
            @input="${this._text('range')}"
          />
        </label>
        <label class="field">
          <span>Rule type</span>
          <select id="cf-kind" @change="${this._text('kind')}">
            ${KIND_GROUPS.map(
              ([group, kinds]) => html`<optgroup label="${group}">
                ${kinds.map(([value, label]) => html`<option value="${value}" ?selected=${draft.kind === value}>${label}</option>`)}
              </optgroup>`
            )}
          </select>
        </label>
        ${this._renderCondition(draft)}
        ${graphic ? nothing : this._renderStyle(draft)}
        <label class="field">
          <span>Label for screen readers (optional)</span>
          <input
            id="cf-label"
            type="text"
            autocomplete="off"
            aria-describedby="cf-label-hint"
            .value="${draft.label}"
            @input="${this._text('label')}"
          />
        </label>
        <span class="hint" id="cf-label-hint">Read out with each cell the rule applies to, so its meaning does not depend on
          color.${fallback}</span>
        ${this._error ? html`<div class="error" role="alert">${this._error}</div>` : nothing}
        <div class="buttons">
          <button id="cf-save" type="submit" class="primary">${draft.id ? 'Save' : 'Add rule'}</button>
          <button id="cf-cancel" type="button" @click="${this._closeForm}">Cancel</button>
        </div>
      </form>
    `;
  }

  protected render() {
    return html`
      <div class="dialog" role="dialog" aria-modal="true" aria-labelledby="cf-title" @keydown="${this._handleKeydown}">
        ${this._draft ? this._renderForm(this._draft) : this._renderList()}
      </div>
    `;
  }
}

declare global {
  interface HTMLElementTagNameMap {
    'y11n-conditional-format-dialog': Y11nConditionalFormatDialog;
  }
}
