import { LitElement, html, css } from 'lit';
import { customElement, property, state } from 'lit/decorators.js';

export type FormulaBarMode = 'raw' | 'formatted';

interface FormulaBarCommitDetail {
  value: string;
}

interface FormulaBarModeChangeDetail {
  mode: FormulaBarMode;
}

/**
 * Reusable formula bar that supports raw/formatted value display.
 * Emits commit + mode-change events and leaves data persistence to parent components.
 */
@customElement('y11n-formula-bar')
export class Y11nFormulaBar extends LitElement {
  @property({ type: String, attribute: 'cell-ref' }) cellRef = 'A1';
  @property({ type: String, attribute: 'raw-value' }) rawValue = '';
  @property({ type: String, attribute: 'display-value' }) displayValue = '';
  @property({ type: String }) mode: FormulaBarMode = 'raw';
  @property({ type: Boolean, attribute: 'read-only', reflect: true }) readOnly = false;

  @state() private _draft = '';

  static styles = css`
    :host {
      display: grid;
      grid-template-columns: minmax(48px, auto) auto minmax(0, 1fr);
      align-items: stretch;
      gap: 8px;
      padding: 8px;
      border-bottom: 1px solid var(--ls-border-color, #e0e0e0);
      background: var(--ls-formula-bar-bg, #f8fafc);
      color: var(--ls-text-color, #333);
      font-family: var(--ls-font-family, system-ui, -apple-system, sans-serif);
      font-size: var(--ls-font-size, 13px);
    }

    .cell-ref {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 48px;
      padding: 0 8px;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 4px;
      font-weight: 600;
      background: var(--ls-formula-ref-bg, #fff);
    }

    .mode-toggle {
      display: inline-flex;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 4px;
      overflow: hidden;
      background: #fff;
    }

    .mode-toggle button {
      border: 0;
      background: transparent;
      padding: 0 10px;
      min-height: 32px;
      cursor: pointer;
      font: inherit;
      color: inherit;
    }

    .mode-toggle button[aria-pressed='true'] {
      background: var(--ls-formula-mode-active-bg, #dbeafe);
      font-weight: 600;
    }

    input {
      width: 100%;
      min-height: 32px;
      padding: 0 10px;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 4px;
      font: inherit;
      color: inherit;
      outline: none;
      background: #fff;
    }

    input:focus {
      border-color: #1a73e8;
      box-shadow: 0 0 0 1px rgba(26, 115, 232, 0.2);
    }

    button:disabled,
    input:disabled {
      opacity: 0.7;
      cursor: not-allowed;
    }
  `;

  willUpdate(changedProps: Map<string, unknown>): void {
    if (
      changedProps.has('mode') ||
      changedProps.has('rawValue') ||
      changedProps.has('displayValue')
    ) {
      this._draft = this._sourceValue();
    }
  }

  private _sourceValue(): string {
    return this.mode === 'raw' ? this.rawValue : this.displayValue;
  }

  private _setMode(mode: FormulaBarMode): void {
    if (this.mode === mode) return;
    this.dispatchEvent(
      new CustomEvent<FormulaBarModeChangeDetail>('formula-bar-mode-change', {
        detail: { mode },
        bubbles: true,
        composed: true,
      })
    );
  }

  private _onInput(e: Event): void {
    this._draft = (e.target as HTMLInputElement).value;
  }

  private _commit(): void {
    this.dispatchEvent(
      new CustomEvent<FormulaBarCommitDetail>('formula-bar-commit', {
        detail: { value: this._draft },
        bubbles: true,
        composed: true,
      })
    );
  }

  private _onKeydown(e: KeyboardEvent): void {
    if (e.key === 'Enter') {
      e.preventDefault();
      this._commit();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      this._draft = this._sourceValue();
    }
  }

  render() {
    return html`
      <output class="cell-ref" aria-live="polite" aria-atomic="true">${this.cellRef}</output>
      <div class="mode-toggle" role="group" aria-label="Formula bar mode">
        <button
          type="button"
          aria-pressed="${this.mode === 'raw'}"
          @click="${() => this._setMode('raw')}"
        >
          Raw
        </button>
        <button
          type="button"
          aria-pressed="${this.mode === 'formatted'}"
          @click="${() => this._setMode('formatted')}"
        >
          Formatted
        </button>
      </div>
      <input
        type="text"
        .value="${this._draft}"
        ?disabled="${this.readOnly}"
        @input="${this._onInput}"
        @keydown="${this._onKeydown}"
        @blur="${this._commit}"
        aria-label="Formula bar"
      />
    `;
  }
}

declare global {
  interface HTMLElementTagNameMap {
    'y11n-formula-bar': Y11nFormulaBar;
  }
}
