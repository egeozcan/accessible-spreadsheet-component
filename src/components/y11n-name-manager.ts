import { LitElement, html, css, nothing } from 'lit';
import type { PropertyValues } from 'lit';
import { customElement, property, state } from 'lit/decorators.js';
import { keyed } from 'lit/directives/keyed.js';
import type { DefinedName } from '../types.js';
import { isValidName } from '../engine/formula-engine.js';
import { parseFormula } from '../engine/parser.js';

/** Detail for the name manager's `name-save` event: a new name, or a changed one. */
export interface NameSaveDetail {
  name: string;
  /** What the name stands for, as a formula that starts with `=` */
  refersTo: string;
  /** Sheet that references without a sheet name point to */
  sheet: string;
  /** When an existing name was edited: its name before the edit */
  previous?: string;
}

/** Detail for the name manager's `name-remove` event: the name whose deletion was confirmed. */
export interface NameRemoveDetail {
  name: string;
}

/** Detail for the name manager's `name-go` event: the name whose cells the user wants to select. */
export interface NameGoDetail {
  name: string;
}

type RowAction = 'go' | 'edit' | 'delete';

/** The form's fields, kept as typed until the name is saved. */
interface Draft {
  /** Name being edited; null for a new name */
  previous: string | null;
  name: string;
  refersTo: string;
  sheet: string;
}

interface FormErrors {
  name?: string;
  refersTo?: string;
}

const INVALID_NAME =
  'A name starts with a letter or underscore, has only letters, digits, underscores and periods, and is not a cell reference';
const INVALID_REFERS_TO = 'Refers to is not a valid formula';

/** What was typed in "Refers to" as a formula with its `=`, or null when it is empty or does not parse. */
function normalizeRefersTo(text: string): string | null {
  const trimmed = text.trim();
  const body = (trimmed.startsWith('=') ? trimmed.slice(1) : trimmed).trim();
  if (body === '') return null;
  try {
    parseFormula(body);
  } catch {
    return null;
  }
  return `=${body}`;
}

function sameName(a: string, b: string): boolean {
  return a.toLowerCase() === b.toLowerCase();
}

/**
 * `<y11n-name-manager>` -- The name manager of `<y11n-spreadsheet>`: a
 * non-modal dialog that lists the defined names (go to, edit, delete) and has
 * a form to add or change one.
 *
 * It holds no names itself. The spreadsheet passes them in, applies the
 * changes it asks for and reports the outcome through `status`:
 *
 * @fires name-save - A name was added, or changed (`detail.previous` is its name before the edit)
 * @fires name-remove - Delete a name (`detail.name`), after the user confirmed
 * @fires name-go - Select what a name refers to (`detail.name`)
 * @fires name-close - The dialog was dismissed
 */
@customElement('y11n-name-manager')
export class Y11nNameManager extends LitElement {
  /** All defined names */
  @property({ attribute: false }) names: DefinedName[] = [];
  /** Sheet names, for the "sheet" choice of a name (what references without a sheet name point to) */
  @property({ attribute: false }) sheets: string[] = [];
  /** The sheet being shown: default sheet of a new name */
  @property({ type: String, attribute: 'active-sheet' }) activeSheet = '';
  /** Default "refers to" of a new name: the current selection as an absolute, sheet-qualified reference, e.g. `=Sheet1!$A$1:$B$3` */
  @property({ type: String, attribute: 'selection-ref' }) selectionRef = '';
  /** Names can be looked at and gone to, but not added, edited or deleted */
  @property({ type: Boolean, attribute: 'read-only', reflect: true }) readOnly = false;
  /** Outcome of the last change, set by the parent; announced in the dialog's status region */
  @property({ type: String }) status = '';

  @state() private _draft: Draft | null = null;
  @state() private _errors: FormErrors = {};
  /** Name the delete confirmation asks about */
  @state() private _confirmName: string | null = null;
  /** Counts failed saves, so each one announces its messages again */
  private _attempt = 0;
  private _pendingFocus: (() => HTMLElement | null | undefined) | null = null;

  static styles = css`
    :host {
      position: absolute;
      top: 8px;
      left: 50%;
      transform: translateX(-50%);
      z-index: 20;
      display: flex;
      flex-direction: column;
      box-sizing: border-box;
      width: min(600px, calc(100% - 16px));
      max-height: calc(100% - 16px);
      overflow: hidden;
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

    /* Scrolls as a whole only when the list can't shrink any further */
    .dialog {
      flex: 1 1 auto;
      min-height: 0;
      padding: 10px;
      overflow: auto;
    }

    h2 {
      margin: 0;
      font-size: 14px;
      font-weight: 600;
    }

    p {
      margin: 0;
    }

    /* The names scroll in here, so the buttons below stay in view */
    .list {
      flex: 0 1 auto;
      min-height: 64px;
      overflow: auto;
      /* The sticky header covers the top */
      scroll-padding-top: 32px;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 3px;
    }

    /* In a narrow grid the list scrolls sideways rather than squeeze "Refers to" away */
    table {
      width: 100%;
      min-width: 460px;
      table-layout: fixed;
      border-collapse: separate;
      border-spacing: 0;
    }

    /* The padding is also the room a button's focus ring needs */
    th,
    td {
      padding: 4px 6px;
      font-weight: normal;
      text-align: start;
      vertical-align: middle;
      border-bottom: 1px solid var(--ls-border-color, #e0e0e0);
    }

    tbody tr:last-child > * {
      border-bottom: 0;
    }

    thead th {
      position: sticky;
      top: 0;
      z-index: 1;
      font-weight: 600;
      background: var(--ls-header-bg, #f3f3f3);
    }

    .col-name {
      width: 24%;
    }

    .col-sheet {
      width: 16%;
    }

    .col-actions {
      width: 184px;
    }

    :host([read-only]) .col-actions {
      width: 76px;
    }

    tbody th {
      font-weight: 600;
    }

    tr.confirming > * {
      background: var(--ls-header-bg, #f3f3f3);
    }

    /* Long names and formulas are cut off; the full text is in the title and in the edit form */
    .clip {
      display: block;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .actions {
      display: flex;
      flex-wrap: wrap;
      gap: 4px;
    }

    .hint,
    .empty {
      font-size: 12px;
      color: var(--ls-find-status-color, #555);
    }

    .field {
      display: flex;
      flex-direction: column;
      gap: 2px;
    }

    input[type='text'],
    select {
      min-height: 26px;
      padding: 2px 6px;
      font: inherit;
      color: inherit;
      background: var(--ls-editor-bg, #fff);
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 3px;
    }

    select {
      align-self: flex-start;
      max-width: 100%;
    }

    input[aria-invalid='true'] {
      border-color: var(--ls-validation-border, #d93025);
    }

    button {
      min-width: 24px;
      min-height: 26px;
      padding: 2px 8px;
      font: inherit;
      color: inherit;
      white-space: nowrap;
      background: transparent;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 3px;
      cursor: pointer;
    }

    button:hover {
      background: var(--ls-header-bg, #f3f3f3);
    }

    button.primary {
      color: #fff;
      background: var(--ls-filter-active-color, #1a73e8);
      border-color: var(--ls-filter-active-color, #1a73e8);
    }

    button.danger {
      color: var(--ls-name-delete-color, var(--ls-sheet-delete-color, #fff));
      background: var(--ls-name-delete-bg, var(--ls-sheet-delete-bg, #c5221f));
      border-color: var(--ls-name-delete-bg, var(--ls-sheet-delete-bg, #c5221f));
    }

    button.primary:hover,
    button.danger:hover {
      filter: brightness(0.92);
    }

    button.primary:hover {
      background: var(--ls-filter-active-color, #1a73e8);
    }

    button.danger:hover {
      background: var(--ls-name-delete-bg, var(--ls-sheet-delete-bg, #c5221f));
    }

    button:focus-visible,
    input:focus-visible,
    select:focus-visible {
      outline: var(--ls-focus-ring, 2px solid #1a73e8);
      outline-offset: 1px;
    }

    .buttons {
      display: flex;
      flex: 0 0 auto;
      justify-content: flex-end;
      gap: 6px;
    }

    .confirm {
      display: flex;
      flex: 0 0 auto;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 6px 10px;
      padding: 6px 8px;
      border: 1px solid var(--ls-border-color, #e0e0e0);
      border-radius: 3px;
    }

    .confirm p {
      flex: 1 1 200px;
      overflow-wrap: anywhere;
    }

    .status {
      flex: 0 0 auto;
      min-height: 1.2em;
      font-size: 12px;
      color: var(--ls-find-status-color, #555);
      overflow-wrap: anywhere;
    }

    /* In the form the last outcome is old news: kept for screen readers, where taking it out and putting it back would announce it again */
    .status.offscreen {
      position: absolute;
      width: 1px;
      height: 1px;
      min-height: 0;
      padding: 0;
      margin: -1px;
      overflow: hidden;
      clip: rect(0, 0, 0, 0);
      white-space: nowrap;
      border: 0;
    }

    .error {
      padding: 2px 8px;
      font-size: 12px;
      color: var(--ls-validation-color, #8a1c12);
      background: var(--ls-validation-bg, #fdecea);
      border: 1px solid var(--ls-validation-border, #d93025);
      border-radius: 3px;
    }

    /*
     * Forced colors drop the backgrounds: the main and the delete button, the
     * row being deleted and a field with an error are told apart by their
     * lines instead, and focus is a ring in a system color.
     */
    @media (forced-colors: active) {
      :host {
        border-color: CanvasText;
      }

      button:focus-visible,
      input:focus-visible,
      select:focus-visible {
        outline: 2px solid Highlight;
        outline-offset: 1px;
      }

      button.primary,
      button.danger,
      input[aria-invalid='true'] {
        border-width: 2px;
      }

      thead th {
        border-bottom: 2px solid CanvasText;
      }

      tr.confirming > * {
        border-top: 2px solid Highlight;
        border-bottom: 2px solid Highlight;
      }

      .error {
        border-width: 2px;
      }
    }
  `;

  /**
   * Focus the control a keyboard user should start on: the "New name" button
   * (the Name field when the form is open); in read-only mode the first name's
   * "Go to" button, or the Close button when there are no names.
   */
  focusInitial(): void {
    this._initialTarget()?.focus();
  }

  private _initialTarget(): HTMLElement | null {
    if (this._draft) return this._query('#name-name');
    return this._query('#name-new') ?? this._query('button[data-action="go"]') ?? this._query('#name-close');
  }

  private _query(selector: string): HTMLElement | null {
    return this.renderRoot.querySelector<HTMLElement>(selector);
  }

  /** A row's button. Looked up by comparing names, which need not be safe to put in a selector. */
  private _rowButton(name: string, action: RowAction): HTMLElement | undefined {
    return [...this.renderRoot.querySelectorAll<HTMLElement>(`button[data-action="${action}"]`)].find(
      (button) => button.dataset.name === name
    );
  }

  private _emit<T>(name: string, detail?: T): void {
    this.dispatchEvent(new CustomEvent<T>(name, { detail, bubbles: true, composed: true }));
  }

  /**
   * Focus an element once the dialog has re-rendered: the view changed, or the
   * spreadsheet applied a change and passed new names in.
   */
  private _focusAfterUpdate(target: () => HTMLElement | null | undefined): void {
    this._pendingFocus = target;
    this.requestUpdate();
  }

  protected willUpdate(changed: PropertyValues<this>): void {
    const readOnlyNow = this.readOnly && (this._draft !== null || this._confirmName !== null);
    const confirmedNameGone =
      this._confirmName !== null && changed.has('names') && !this.names.some((n) => n.name === this._confirmName);
    if (!readOnlyNow && !confirmedNameGone) return;
    // What was open can't be finished any more: back to the list. Focus stays in the dialog if it was there.
    this._draft = null;
    this._errors = {};
    this._confirmName = null;
    if ((this.renderRoot as ShadowRoot).activeElement) this._pendingFocus = () => this._initialTarget();
  }

  protected updated(): void {
    const pending = this._pendingFocus;
    if (!pending) return;
    this._pendingFocus = null;
    (pending() ?? this._initialTarget())?.focus();
  }

  private _handleKeydown(e: KeyboardEvent): void {
    // Keys typed in the dialog are not grid shortcuts
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      if (this._confirmName !== null) this._closeConfirm();
      else if (this._draft) this._closeForm();
      else this._emit('name-close');
    } else if (e.key === 'Enter') {
      // Not left to the browser's implicit submission, which a form with several fields may not do
      if (this._draft && !e.isComposing && e.target instanceof HTMLInputElement) {
        e.preventDefault();
        this._save();
      }
    } else if (e.key === 'Tab' && this._confirmName !== null) {
      // The confirmation is modal: focus stays on its two buttons
      const buttons = [...this.renderRoot.querySelectorAll<HTMLElement>('.confirm button')];
      const current = buttons.indexOf((this.renderRoot as ShadowRoot).activeElement as HTMLElement);
      e.preventDefault();
      buttons[(current + (e.shiftKey ? -1 : 1) + buttons.length) % buttons.length]?.focus();
    } else if (e.key === 'Tab') {
      // Modal: keep focus inside
      const focusables = [
        ...this.renderRoot.querySelectorAll<HTMLElement>(
          'button:not(:disabled), input:not(:disabled), select:not(:disabled)'
        ),
      ];
      const first = focusables[0];
      const last = focusables[focusables.length - 1];
      const current = (this.renderRoot as ShadowRoot).activeElement;
      if (e.shiftKey && (current === first || current === null)) {
        e.preventDefault();
        last?.focus();
      } else if (!e.shiftKey && (current === last || current === null)) {
        e.preventDefault();
        first?.focus();
      }
    }
  }

  // ─── List ──────────────────────────────────────────────

  /** The names in the order they are listed: by name, ignoring case. */
  private _sorted(): DefinedName[] {
    return [...this.names].sort((a, b) => {
      const x = a.name.toLowerCase();
      const y = b.name.toLowerCase();
      if (x !== y) return x < y ? -1 : 1;
      return a.name < b.name ? -1 : a.name > b.name ? 1 : 0;
    });
  }

  private _go(name: string): void {
    this._confirmName = null;
    this._emit<NameGoDetail>('name-go', { name });
  }

  /** The sheet the form starts on: the one asked for when there is such a sheet, else the sheet being shown. */
  private _sheetFor(wanted: string | undefined): string {
    const find = (name: string) => this.sheets.find((s) => sameName(s, name));
    return (
      (wanted === undefined ? undefined : find(wanted)) ??
      find(this.activeSheet) ??
      this.sheets[0] ??
      wanted ??
      this.activeSheet
    );
  }

  private _openForm(existing: DefinedName | null): void {
    if (this.readOnly) return;
    this._confirmName = null;
    this._draft = existing
      ? {
          previous: existing.name,
          name: existing.name,
          refersTo: existing.refersTo,
          sheet: this._sheetFor(existing.sheet),
        }
      : { previous: null, name: '', refersTo: this.selectionRef, sheet: this._sheetFor(undefined) };
    this._errors = {};
    this._focusAfterUpdate(() => this._query('#name-name'));
  }

  /** Leave the form without saving: focus returns to the button that opened it. */
  private _closeForm(): void {
    const previous = this._draft?.previous ?? null;
    this._draft = null;
    this._errors = {};
    this._focusAfterUpdate(() => (previous === null ? null : this._rowButton(previous, 'edit')) ?? this._query('#name-new'));
  }

  // ─── Delete ────────────────────────────────────────────

  private _requestDelete(name: string): void {
    if (this.readOnly) return;
    this._confirmName = name;
    this._focusAfterUpdate(() => {
      // The confirmation takes room from the list: keep the row it is about in view
      this._rowButton(name, 'delete')?.closest('tr')?.scrollIntoView?.({ block: 'nearest' });
      return this._query('#name-delete-cancel');
    });
  }

  private _closeConfirm(): void {
    const name = this._confirmName;
    if (name === null) return;
    this._confirmName = null;
    this._focusAfterUpdate(() => this._rowButton(name, 'delete'));
  }

  private _confirmDelete(): void {
    const name = this._confirmName;
    if (name === null || this.readOnly) return;
    this._confirmName = null;
    this._emit<NameRemoveDetail>('name-remove', { name });
    // The name's row goes away once the spreadsheet has removed it
    this._focusAfterUpdate(() => this._query('#name-new'));
  }

  private _renderRow(entry: DefinedName) {
    const { name, refersTo, sheet } = entry;
    return html`
      <tr data-name="${name}" class="${this._confirmName === name ? 'confirming' : ''}">
        <th scope="row"><span class="clip" title="${name}">${name}</span></th>
        <td><span class="clip" title="${refersTo}">${refersTo}</span></td>
        <td><span class="clip" title="${sheet}">${sheet}</span></td>
        <td>
          <div class="actions">
            <button
              type="button"
              data-name="${name}"
              data-action="go"
              aria-label="Go to ${name}"
              @click="${() => this._go(name)}"
            >Go to</button>
            ${this.readOnly
              ? nothing
              : html`<button
                    type="button"
                    data-name="${name}"
                    data-action="edit"
                    aria-label="Edit ${name}"
                    @click="${() => this._openForm(entry)}"
                  >Edit</button>
                  <button
                    type="button"
                    data-name="${name}"
                    data-action="delete"
                    aria-label="Delete ${name}"
                    @click="${() => this._requestDelete(name)}"
                  >Delete</button>`}
          </div>
        </td>
      </tr>
    `;
  }

  private _renderConfirm(name: string) {
    return html`
      <div class="confirm" role="alertdialog" aria-modal="true" aria-label="Delete name" aria-describedby="name-delete-text">
        <p id="name-delete-text">Delete the name "${name}"? Formulas that use it will show a #NAME? error.</p>
        <div class="buttons">
          <button id="name-delete-confirm" type="button" class="danger" @click="${this._confirmDelete}">Delete</button>
          <button id="name-delete-cancel" type="button" @click="${this._closeConfirm}">Cancel</button>
        </div>
      </div>
    `;
  }

  private _renderList() {
    const names = this._sorted();
    return html`
      <h2 id="name-title">Name manager</h2>
      ${names.length === 0
        ? html`<p class="empty">No names yet.</p>`
        : html`<div class="list">
            <table aria-label="Defined names">
              <colgroup>
                <col class="col-name" />
                <col />
                <col class="col-sheet" />
                <col class="col-actions" />
              </colgroup>
              <thead>
                <tr>
                  <th scope="col">Name</th>
                  <th scope="col">Refers to</th>
                  <th scope="col">Sheet</th>
                  <th scope="col">Actions</th>
                </tr>
              </thead>
              <tbody>
                ${names.map((entry) => this._renderRow(entry))}
              </tbody>
            </table>
          </div>`}
      ${this._confirmName === null ? nothing : this._renderConfirm(this._confirmName)}
    `;
  }

  // ─── Form ──────────────────────────────────────────────

  private _set(changes: Partial<Draft>): void {
    if (this._draft) this._draft = { ...this._draft, ...changes };
  }

  private _text(key: 'name' | 'refersTo' | 'sheet') {
    return (e: Event) => this._set({ [key]: (e.target as HTMLInputElement | HTMLSelectElement).value });
  }

  private _handleSubmit(e: Event): void {
    e.preventDefault();
    this._save();
  }

  /** Check the form and ask the spreadsheet to save the name, or show what is wrong with it. */
  private _save(): void {
    const draft = this._draft;
    if (!draft || this.readOnly) return;
    const errors: FormErrors = {};

    const name = draft.name.trim();
    if (!isValidName(name)) {
      errors.name = INVALID_NAME;
    } else {
      const other = this.names.find((n) => sameName(n.name, name));
      // Its own name is fine, in any case: "sales" can become "Sales"
      if (other && !(draft.previous !== null && sameName(other.name, draft.previous))) {
        errors.name = `There is already a name ${other.name}`;
      }
    }

    const refersTo = normalizeRefersTo(draft.refersTo);
    if (refersTo === null) errors.refersTo = INVALID_REFERS_TO;

    if (errors.name !== undefined || refersTo === null) {
      this._errors = errors;
      this._attempt++;
      this._focusAfterUpdate(() => this._query(errors.name !== undefined ? '#name-name' : '#name-refers'));
      return;
    }

    const detail: NameSaveDetail = { name, refersTo, sheet: draft.sheet };
    if (draft.previous !== null) detail.previous = draft.previous;
    this._emit<NameSaveDetail>('name-save', detail);
    this._draft = null;
    this._errors = {};
    this._focusAfterUpdate(() => this._query('#name-new'));
  }

  /** A field's message. Keyed by the attempt, so saving again with the same mistake announces it again. */
  private _renderError(id: string, message: string | undefined) {
    return message === undefined
      ? nothing
      : keyed(this._attempt, html`<div class="error" id="${id}" role="alert">${message}</div>`);
  }

  private _renderForm(draft: Draft) {
    const { name: nameError, refersTo: refersError } = this._errors;
    // A sheet the list doesn't have (no list was passed in) can still be kept
    const sheets = draft.sheet === '' || this.sheets.includes(draft.sheet) ? this.sheets : [...this.sheets, draft.sheet];
    return html`
      <h2 id="name-title">${draft.previous === null ? 'New name' : 'Edit name'}</h2>
      <form novalidate @submit="${this._handleSubmit}">
        <div class="field">
          <label for="name-name">Name</label>
          <input
            id="name-name"
            type="text"
            autocomplete="off"
            autocapitalize="off"
            spellcheck="false"
            aria-invalid="${nameError === undefined ? 'false' : 'true'}"
            aria-describedby="${nameError === undefined ? nothing : 'name-name-error'}"
            .value="${draft.name}"
            @input="${this._text('name')}"
          />
          ${this._renderError('name-name-error', nameError)}
        </div>
        <div class="field">
          <label for="name-refers">Refers to</label>
          <input
            id="name-refers"
            type="text"
            autocomplete="off"
            autocapitalize="off"
            spellcheck="false"
            aria-invalid="${refersError === undefined ? 'false' : 'true'}"
            aria-describedby="${refersError === undefined ? 'name-refers-hint' : 'name-refers-error name-refers-hint'}"
            .value="${draft.refersTo}"
            @input="${this._text('refersTo')}"
          />
          ${this._renderError('name-refers-error', refersError)}
          <span class="hint" id="name-refers-hint">A range, a constant or a function, such as =Sheet1!$A$1:$A$10, =0.19 or
            =LAMBDA(x, x * 2)</span>
        </div>
        <div class="field">
          <label for="name-sheet">Sheet</label>
          <select id="name-sheet" aria-describedby="name-sheet-hint" @change="${this._text('sheet')}">
            ${sheets.map((sheet) => html`<option value="${sheet}" ?selected=${draft.sheet === sheet}>${sheet}</option>`)}
          </select>
          <span class="hint" id="name-sheet-hint">Sheet that references without a sheet name point to</span>
        </div>
        <div class="buttons">
          <button id="name-save" type="submit" class="primary">Save</button>
          <button id="name-cancel" type="button" @click="${this._closeForm}">Cancel</button>
        </div>
      </form>
    `;
  }

  protected render() {
    const draft = this._draft;
    return html`
      <div class="dialog" role="dialog" aria-modal="true" aria-label="Name manager" @keydown="${this._handleKeydown}">
        ${draft ? this._renderForm(draft) : this._renderList()}
        <div class="status ${draft ? 'offscreen' : ''}" id="name-status" role="status">${this.status}</div>
        ${draft
          ? nothing
          : html`<div class="buttons">
              ${this.readOnly
                ? nothing
                : html`<button
                    id="name-new"
                    type="button"
                    class="primary"
                    @click="${() => this._openForm(null)}"
                  >New name</button>`}
              <button id="name-close" type="button" @click="${() => this._emit('name-close')}">Close</button>
            </div>`}
      </div>
    `;
  }
}

declare global {
  interface HTMLElementTagNameMap {
    'y11n-name-manager': Y11nNameManager;
  }
}
