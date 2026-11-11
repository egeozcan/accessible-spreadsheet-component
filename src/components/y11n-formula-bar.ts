import { LitElement, html, css } from 'lit';
import { customElement, property, state } from 'lit/decorators.js';
import { live } from 'lit/directives/live.js';

/** Display mode for the formula bar input. */
export type FormulaBarMode = 'raw' | 'formatted';

/** Detail for the `formula-bar-commit` event. */
interface FormulaBarCommitDetail {
  /** The committed value from the input field */
  value: string;
  /** The cell reference the value applies to (e.g. "A1") */
  cellRef: string;
}

/** Detail for the `name-box-commit` event. */
export interface NameBoxCommitDetail {
  /** What was typed into the Name Box: a cell reference, a range or a name */
  value: string;
}

/** Detail for the `formula-bar-mode-change` event. */
interface FormulaBarModeChangeDetail {
  /** The newly selected display mode */
  mode: FormulaBarMode;
}

/**
 * `<y11n-formula-bar>` -- Reusable formula bar sub-component.
 *
 * Has the Name Box (the active cell's reference, or the name of the
 * selection; type a reference or a name and press Enter to go there, or a new
 * name to name the selection) and an input for editing cell values.
 * Supports toggling between raw (formula) and formatted (display) modes.
 * All data persistence is handled by the parent component; this component
 * communicates via custom events:
 *
 * @fires formula-bar-commit - Fired on Enter or blur when the draft differs from the source value
 * @fires formula-bar-mode-change - Fired when the user toggles between raw and formatted mode
 * @fires name-box-commit - Fired on Enter in the Name Box, with what was typed
 * @fires name-box-cancel - Fired on Escape in the Name Box
 */
@customElement('y11n-formula-bar')
export class Y11nFormulaBar extends LitElement {
  @property({ type: String, attribute: 'cell-ref' }) cellRef = 'A1';
  @property({ type: String, attribute: 'raw-value' }) rawValue = '';
  @property({ type: String, attribute: 'display-value' }) displayValue = '';
  @property({ type: String }) mode: FormulaBarMode = 'raw';
  @property({ type: Boolean, attribute: 'read-only', reflect: true }) readOnly = false;
  /**
   * Text shown greyed out while the input is empty. The spreadsheet uses it
   * for the formula behind a spilled value, which can't be edited in that cell.
   */
  @property({ type: String }) placeholder = '';

  /** The defined name the selection matches, shown in the Name Box instead of the cell reference */
  @property({ type: String, attribute: 'selection-label' }) selectionLabel = '';
  /** Defined names, offered as suggestions in the Name Box */
  @property({ attribute: false }) names: string[] = [];

  @state() private _draft = '';
  @state() private _nameDraft = '';

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

    input.cell-ref {
      box-sizing: border-box;
      width: var(--ls-name-box-width, 9em);
      min-width: 48px;
      padding: 0 8px;
      text-align: center;
      font-weight: 600;
      background: var(--ls-formula-ref-bg, #fff);
    }

    .mode-toggle {
      display: inline-flex;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 4px;
      overflow: hidden;
      background: var(--ls-formula-input-bg, #fff);
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
      background: var(--ls-formula-input-bg, #fff);
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
      changedProps.has('cellRef') ||
      changedProps.has('rawValue') ||
      changedProps.has('displayValue')
    ) {
      this._draft = this._sourceValue();
    }
    if (changedProps.has('cellRef') || changedProps.has('selectionLabel')) {
      this._nameDraft = this._nameSource();
    }
  }

  /**
   * Discard the draft and show the cell's value again. The spreadsheet calls
   * this when it rejects a commit (read-only cell, failed validation, or a
   * cancelled `before-change`), so the rejected text isn't committed later.
   */
  resetDraft(): void {
    this._draft = this._sourceValue();
  }

  /** Discard what was typed into the Name Box and show the selection again. */
  resetNameBox(): void {
    this._nameDraft = this._nameSource();
  }

  /** Focus the Name Box with its text selected, ready to type a reference or a name. */
  focusNameBox(): void {
    const box = this.shadowRoot?.querySelector<HTMLInputElement>('.name-box');
    box?.focus();
    box?.select();
  }

  private _nameSource(): string {
    return this.selectionLabel || this.cellRef;
  }

  private _onNameInput(e: Event): void {
    this._nameDraft = (e.target as HTMLInputElement).value;
  }

  private _onNameKeydown(e: KeyboardEvent): void {
    if (e.key === 'Enter') {
      e.preventDefault();
      const value = this._nameDraft.trim();
      if (value === '') {
        this._nameDraft = this._nameSource();
        return;
      }
      this.dispatchEvent(
        new CustomEvent<NameBoxCommitDetail>('name-box-commit', { detail: { value }, bubbles: true, composed: true })
      );
    } else if (e.key === 'Escape') {
      e.preventDefault();
      this._nameDraft = this._nameSource();
      this.dispatchEvent(new CustomEvent('name-box-cancel', { bubbles: true, composed: true }));
    }
  }

  private _onNameFocus(e: FocusEvent): void {
    (e.target as HTMLInputElement).select();
  }

  private _onNameBlur(): void {
    this._nameDraft = this._nameSource();
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
        detail: { value: this._draft, cellRef: this.cellRef },
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
      // Reverting the draft is enough: blur only commits when the draft
      // differs from the source, so no extra suppression flag is needed
      // (a sticky flag would silently drop edits typed after Escape).
      this._draft = this._sourceValue();
    }
  }

  private _onBlur(): void {
    if (this._draft !== this._sourceValue()) {
      this._commit();
    }
  }

  render() {
    return html`
      <input
        class="cell-ref name-box"
        type="text"
        list="name-list"
        autocomplete="off"
        spellcheck="false"
        aria-label="Name box"
        title="Name box: type a cell, a range or a name and press Enter"
        .value="${live(this._nameDraft)}"
        @input="${this._onNameInput}"
        @keydown="${this._onNameKeydown}"
        @focus="${this._onNameFocus}"
        @blur="${this._onNameBlur}"
      />
      <datalist id="name-list">
        ${this.names.map((name) => html`<option value="${name}"></option>`)}
      </datalist>
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
        class="formula-input"
        type="text"
        .value="${live(this._draft)}"
        ?disabled="${this.readOnly}"
        placeholder="${this.placeholder}"
        @input="${this._onInput}"
        @keydown="${this._onKeydown}"
        @blur="${this._onBlur}"
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
