import { LitElement, html, css } from 'lit';
import { customElement, property } from 'lit/decorators.js';

export interface FormatActionDetail {
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
  value?: string | number | boolean;
}

/**
 * <y11n-format-toolbar> - A companion toolbar for cell formatting.
 *
 * Connected to <y11n-spreadsheet> via properties and events.
 * Consumer wires them together:
 *
 * ```html
 * <y11n-format-toolbar
 *   .bold=${format.bold}
 *   @format-action=${handle}
 * ></y11n-format-toolbar>
 * ```
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
      toolbar.querySelectorAll('button:not(:disabled), select:not(:disabled), .color-wrapper')
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
          >${'\u2261'}</button>
          <button
            type="button"
            aria-pressed="${this.textAlign === 'center'}"
            aria-label="Align center"
            title="Align center"
            ?disabled=${d}
            tabindex="-1"
            @click=${() => this._dispatch('textAlign', 'center')}
          >${'\u2261'}</button>
          <button
            type="button"
            aria-pressed="${this.textAlign === 'right'}"
            aria-label="Align right"
            title="Align right"
            ?disabled=${d}
            tabindex="-1"
            @click=${() => this._dispatch('textAlign', 'right')}
          >${'\u2261'}</button>
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

/** Compute consensus format from a selection range */
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
