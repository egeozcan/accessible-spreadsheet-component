import { LitElement, html, css, nothing } from 'lit';
import { customElement, property, state } from 'lit/decorators.js';
import type { ConditionalFormatPreset } from '../types.js';

/**
 * What the Conditional formatting menu asks for: a one-step rule for the
 * selection, clearing the selection's rules, or opening the rule manager.
 */
export type ConditionalFormatMenuAction = ConditionalFormatPreset | 'clear' | 'manage';

/** Detail for the `conditional-format-action` event dispatched by the format toolbar. */
export interface ConditionalFormatActionDetail {
  action: ConditionalFormatMenuAction;
}

const CONDITIONAL_MENU: Array<[ConditionalFormatMenuAction, string] | null> = [
  ['duplicates', 'Highlight duplicate values'],
  ['unique', 'Highlight unique values'],
  ['aboveAverage', 'Highlight above average'],
  ['belowAverage', 'Highlight below average'],
  ['top10', 'Highlight top 10'],
  ['bottom10', 'Highlight bottom 10'],
  null,
  ['colorScale', 'Color scale'],
  ['dataBar', 'Data bars'],
  ['iconSet', 'Icon set'],
  null,
  ['clear', 'Clear rules from selection'],
  ['manage', 'Manage rules\u2026'],
];

/** Detail for the `format-action` custom event dispatched by the format toolbar. */
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
    | 'fontSize'
    | 'clearFormat';
  /** The new value for the action (e.g. true/false for toggles, color string, font size number) */
  value?: string | number | boolean;
}

/**
 * `<y11n-format-toolbar>` -- A companion toolbar for cell formatting.
 *
 * This component does not manage state internally. The consumer wires it
 * to `<y11n-spreadsheet>` by binding selection-derived properties and
 * handling the `format-action` event:
 *
 * ```html
 * <y11n-format-toolbar
 *   .bold=${format.bold}
 *   @format-action=${handle}
 * ></y11n-format-toolbar>
 * ```
 *
 * The Conditional formatting menu fires its own event, which maps onto three
 * spreadsheet methods:
 *
 * ```js
 * toolbar.addEventListener('conditional-format-action', (e) => {
 *   const { action } = e.detail;
 *   if (action === 'manage') sheet.openConditionalFormatManager();
 *   else if (action === 'clear') sheet.clearSelectionConditionalFormats();
 *   else sheet.applyConditionalFormatPreset(action);
 * });
 * ```
 *
 * @fires format-action - Dispatched when the user clicks a formatting button or changes a value
 * @fires conditional-format-action - Dispatched when the user picks an item of the Conditional formatting menu
 */
@customElement('y11n-format-toolbar')
export class Y11nFormatToolbar extends LitElement {
  // ─── Properties reflecting current selection state ─────

  @property({ type: Boolean }) bold = false;
  @property({ type: Boolean }) italic = false;
  @property({ type: Boolean }) underline = false;
  @property({ type: Boolean }) strikethrough = false;
  @property({ type: String }) textColor = '#000000';
  @property({ type: String }) backgroundColor = '#ffffff';
  @property({ type: String }) textAlign: 'left' | 'center' | 'right' = 'left';
  @property({ type: Number }) fontSize = 13;
  @property({ type: Boolean, attribute: 'read-only', reflect: true }) readOnly = false;

  /** Whether the Conditional formatting menu is open */
  @state() private _menuOpen = false;
  private _menuCleanup: (() => void) | null = null;

  disconnectedCallback(): void {
    super.disconnectedCallback();
    this._closeMenu(false);
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

    .color-wrapper {
      position: relative;
      display: inline-flex;
      align-items: center;
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
      align-items: center;
      justify-content: center;
      min-width: 28px;
      height: 28px;
      padding: 0 4px;
      border: 1px solid transparent;
      border-radius: 4px;
      background: transparent;
      font-size: 13px;
      pointer-events: none;
    }

    .color-swatch .indicator {
      display: block;
      width: 14px;
      height: 3px;
      margin-top: 1px;
      border-radius: 1px;
    }

    .align-group {
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

    .font-size-select {
      width: 52px;
      height: 28px;
      padding: 0 4px;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 4px;
      background: #fff;
      font: inherit;
      font-size: 12px;
      color: inherit;
      cursor: pointer;
    }

    .font-size-select:disabled {
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
      min-width: 220px;
      margin-top: 2px;
      padding: 4px;
      background: var(--ls-filter-menu-bg, #fff);
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 4px;
      box-shadow: var(--ls-editor-shadow, 0 2px 6px rgba(0, 0, 0, 0.2));
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
  `;

  // ─── Event Dispatching ─────────────────────────────────

  private _dispatch(action: FormatActionDetail['action'], value?: FormatActionDetail['value']): void {
    this.dispatchEvent(
      new CustomEvent<FormatActionDetail>('format-action', {
        detail: { action, value },
        bubbles: true,
        composed: true,
      })
    );
  }

  // ─── Keyboard Navigation ───────────────────────────────

  private _handleToolbarKeydown(e: KeyboardEvent): void {
    const toolbar = e.currentTarget as HTMLElement;
    const focusable = Array.from(
      toolbar.querySelectorAll('button:not(:disabled):not([role="menuitem"]), select:not(:disabled), .color-wrapper')
    ) as HTMLElement[];

    const current = focusable.findIndex(
      (el) => el === e.target || el.contains(e.target as Node)
    );
    if (current === -1) return;

    let next = -1;
    if (e.key === 'ArrowRight') {
      next = (current + 1) % focusable.length;
    } else if (e.key === 'ArrowLeft') {
      next = (current - 1 + focusable.length) % focusable.length;
    } else if (e.key === 'Home') {
      next = 0;
    } else if (e.key === 'End') {
      next = focusable.length - 1;
    }

    if (next >= 0) {
      e.preventDefault();
      const el = focusable[next];
      const inner = el.querySelector('input, select, button') as HTMLElement | null;
      (inner ?? el).focus();
    }
  }

  // ─── Conditional Formatting Menu ───────────────────────

  private _menuItems(): HTMLElement[] {
    return [...this.renderRoot.querySelectorAll<HTMLElement>('.menu [role="menuitem"]:not(:disabled)')];
  }

  private async _openMenu(focus: 'first' | 'last' = 'first'): Promise<void> {
    this._menuOpen = true;
    if (!this._menuCleanup) {
      // Clicking anywhere else closes the menu, like a native dropdown
      const onPointerDown = (e: PointerEvent) => {
        const wrapper = this.renderRoot.querySelector('.menu-wrapper');
        if (wrapper && e.composedPath().includes(wrapper)) return;
        this._closeMenu(false);
      };
      document.addEventListener('pointerdown', onPointerDown);
      this._menuCleanup = () => document.removeEventListener('pointerdown', onPointerDown);
    }
    await this.updateComplete;
    const items = this._menuItems();
    (focus === 'first' ? items[0] : items[items.length - 1])?.focus();
  }

  private _closeMenu(focusButton: boolean): void {
    this._menuCleanup?.();
    this._menuCleanup = null;
    if (!this._menuOpen) return;
    this._menuOpen = false;
    if (focusButton) this.renderRoot.querySelector<HTMLElement>('.menu-button')?.focus();
  }

  private _handleMenuButtonKeydown(e: KeyboardEvent): void {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      e.stopPropagation();
      void this._openMenu(e.key === 'ArrowDown' ? 'first' : 'last');
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

  private _chooseMenuItem(action: ConditionalFormatMenuAction): void {
    this._closeMenu(true);
    this.dispatchEvent(
      new CustomEvent<ConditionalFormatActionDetail>('conditional-format-action', {
        detail: { action },
        bubbles: true,
        composed: true,
      })
    );
  }

  private _renderConditionalMenu() {
    return html`
      <div class="menu-wrapper">
        <button
          class="menu-button"
          type="button"
          aria-label="Conditional formatting"
          aria-haspopup="menu"
          aria-expanded="${this._menuOpen}"
          title="Conditional formatting"
          ?disabled=${this.readOnly}
          tabindex="-1"
          @click=${() => (this._menuOpen ? this._closeMenu(false) : void this._openMenu())}
          @keydown=${this._handleMenuButtonKeydown}
        >
          <svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true" focusable="false">
            <path d="M2 3h12v2.5H2zM2 6.75h8v2.5H2zM2 10.5h4.5V13H2z" fill="currentColor" />
          </svg>
          <svg viewBox="0 0 10 10" width="8" height="8" aria-hidden="true" focusable="false">
            <path d="M1.5 3.5 5 7l3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.5" />
          </svg>
        </button>
        ${this._menuOpen && !this.readOnly
          ? html`<div class="menu" role="menu" aria-label="Conditional formatting" @keydown=${this._handleMenuKeydown}>
              ${CONDITIONAL_MENU.map((item) =>
                item === null
                  ? html`<div role="separator"></div>`
                  : html`<button
                      type="button"
                      role="menuitem"
                      tabindex="-1"
                      data-action="${item[0]}"
                      @click=${() => this._chooseMenuItem(item[0])}
                    >${item[1]}</button>`
              )}
            </div>`
          : nothing}
      </div>
    `;
  }

  // ─── Render ────────────────────────────────────────────

  render() {
    const d = this.readOnly;

    return html`
      <div
        class="toolbar"
        role="toolbar"
        aria-label="Cell formatting"
        @keydown=${this._handleToolbarKeydown}
      >
        <!-- Text style toggles -->
        <button
          class="bold-btn"
          type="button"
          aria-pressed="${this.bold}"
          aria-label="Bold"
          title="Bold (Ctrl+B)"
          ?disabled=${d}
          tabindex="0"
          @click=${() => this._dispatch('bold', !this.bold)}
        >B</button>
        <button
          class="italic-btn"
          type="button"
          aria-pressed="${this.italic}"
          aria-label="Italic"
          title="Italic (Ctrl+I)"
          ?disabled=${d}
          tabindex="-1"
          @click=${() => this._dispatch('italic', !this.italic)}
        >I</button>
        <button
          class="underline-btn"
          type="button"
          aria-pressed="${this.underline}"
          aria-label="Underline"
          title="Underline (Ctrl+U)"
          ?disabled=${d}
          tabindex="-1"
          @click=${() => this._dispatch('underline', !this.underline)}
        >U</button>
        <button
          class="strikethrough-btn"
          type="button"
          aria-pressed="${this.strikethrough}"
          aria-label="Strikethrough"
          title="Strikethrough"
          ?disabled=${d}
          tabindex="-1"
          @click=${() => this._dispatch('strikethrough', !this.strikethrough)}
        >S</button>

        <div class="separator" role="separator"></div>

        <!-- Text color -->
        <div class="color-wrapper" tabindex="-1">
          <div class="color-swatch" aria-hidden="true">
            A<span class="indicator" style="background-color: ${this.textColor}"></span>
          </div>
          <input
            type="color"
            .value=${this.textColor}
            ?disabled=${d}
            aria-label="Text color"
            title="Text color"
            @input=${(e: Event) =>
              this._dispatch('textColor', (e.target as HTMLInputElement).value)}
          />
        </div>

        <!-- Background color -->
        <div class="color-wrapper" tabindex="-1">
          <div class="color-swatch" aria-hidden="true">
            <span class="indicator" style="background-color: ${this.backgroundColor}; width: 16px; height: 16px; border-radius: 2px; border: 1px solid #ccc;"></span>
          </div>
          <input
            type="color"
            .value=${this.backgroundColor}
            ?disabled=${d}
            aria-label="Background color"
            title="Background color"
            @input=${(e: Event) =>
              this._dispatch('backgroundColor', (e.target as HTMLInputElement).value)}
          />
        </div>

        <div class="separator" role="separator"></div>

        <!-- Alignment group -->
        <div class="align-group" role="group" aria-label="Text alignment">
          <button
            type="button"
            aria-pressed="${this.textAlign === 'left'}"
            aria-label="Align left"
            title="Align left"
            ?disabled=${d}
            tabindex="-1"
            @click=${() => this._dispatch('textAlign', 'left')}
          ><span class="align-icon align-left" aria-hidden="true"><span></span><span></span><span></span></span></button>
          <button
            type="button"
            aria-pressed="${this.textAlign === 'center'}"
            aria-label="Align center"
            title="Align center"
            ?disabled=${d}
            tabindex="-1"
            @click=${() => this._dispatch('textAlign', 'center')}
          ><span class="align-icon align-center" aria-hidden="true"><span></span><span></span><span></span></span></button>
          <button
            type="button"
            aria-pressed="${this.textAlign === 'right'}"
            aria-label="Align right"
            title="Align right"
            ?disabled=${d}
            tabindex="-1"
            @click=${() => this._dispatch('textAlign', 'right')}
          ><span class="align-icon align-right" aria-hidden="true"><span></span><span></span><span></span></span></button>
        </div>

        <div class="separator" role="separator"></div>

        <!-- Font size -->
        <select
          class="font-size-select"
          .value=${String(this.fontSize)}
          ?disabled=${d}
          aria-label="Font size"
          title="Font size"
          tabindex="-1"
          @change=${(e: Event) =>
            this._dispatch('fontSize', Number((e.target as HTMLSelectElement).value))}
        >
          ${[8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 32, 36].map(
            (size) => html`<option value=${size} ?selected=${this.fontSize === size}>${size}</option>`
          )}
        </select>

        <div class="separator" role="separator"></div>

        <!-- Conditional formatting: one-step rules, clear, and the rule manager -->
        ${this._renderConditionalMenu()}

        <div class="separator" role="separator"></div>

        <!-- Clear formatting -->
        <button
          type="button"
          aria-label="Clear formatting"
          title="Clear formatting"
          ?disabled=${d}
          tabindex="-1"
          @click=${() => this._dispatch('clearFormat')}
        >${'\u2718'}</button>
      </div>
    `;
  }
}

/**
 * Compute the consensus format from all cells in a selection range.
 * For each property, returns the common value if all cells agree, or undefined if mixed.
 *
 * @param getData - Callback to retrieve cell data by key
 * @param range - The selection range to inspect
 * @returns A CellFormat with only the unanimously shared properties set
 */
export function computeSelectionFormat(
  getData: (cellId: string) => { format?: import('../types.js').CellFormat } | undefined,
  range: import('../types.js').SelectionRange
): import('../types.js').CellFormat {
  let bold: boolean | undefined;
  let italic: boolean | undefined;
  let underline: boolean | undefined;
  let strikethrough: boolean | undefined;
  let textColor: string | undefined;
  let backgroundColor: string | undefined;
  let textAlign: 'left' | 'center' | 'right' | undefined;
  let fontSize: number | undefined;
  let first = true;

  for (let r = range.start.row; r <= range.end.row; r++) {
    for (let c = range.start.col; c <= range.end.col; c++) {
      const cellId = `${r}:${c}`;
      const cell = getData(cellId);
      const fmt = cell?.format;

      if (first) {
        bold = fmt?.bold ?? false;
        italic = fmt?.italic ?? false;
        underline = fmt?.underline ?? false;
        strikethrough = fmt?.strikethrough ?? false;
        textColor = fmt?.textColor;
        backgroundColor = fmt?.backgroundColor;
        textAlign = fmt?.textAlign;
        fontSize = fmt?.fontSize;
        first = false;
      } else {
        if (bold !== (fmt?.bold ?? false)) bold = undefined;
        if (italic !== (fmt?.italic ?? false)) italic = undefined;
        if (underline !== (fmt?.underline ?? false)) underline = undefined;
        if (strikethrough !== (fmt?.strikethrough ?? false)) strikethrough = undefined;
        if (textColor !== fmt?.textColor) textColor = undefined;
        if (backgroundColor !== fmt?.backgroundColor) backgroundColor = undefined;
        if (textAlign !== fmt?.textAlign) textAlign = undefined;
        if (fontSize !== fmt?.fontSize) fontSize = undefined;
      }
    }
  }

  const result: import('../types.js').CellFormat = {};
  if (bold) result.bold = true;
  if (italic) result.italic = true;
  if (underline) result.underline = true;
  if (strikethrough) result.strikethrough = true;
  if (textColor !== undefined) result.textColor = textColor;
  if (backgroundColor !== undefined) result.backgroundColor = backgroundColor;
  if (textAlign !== undefined) result.textAlign = textAlign;
  if (fontSize !== undefined) result.fontSize = fontSize;
  return result;
}

declare global {
  interface HTMLElementTagNameMap {
    'y11n-format-toolbar': Y11nFormatToolbar;
  }
}
