import { LitElement, html, css, nothing } from 'lit';
import type { PropertyValues } from 'lit';
import { customElement, property, state } from 'lit/decorators.js';
import { repeat } from 'lit/directives/repeat.js';
import { MAX_SHEET_NAME_LENGTH, validateSheetName } from '../utils/sheet-names.js';

/** Detail for the `sheet-select` event: the sheet to switch to. */
export interface SheetSelectDetail {
  id: number;
}

/** Detail for the `sheet-rename` event: a sheet and its new, already validated name. */
export interface SheetRenameDetail {
  id: number;
  name: string;
}

/** Detail for the `sheet-remove` event: the sheet whose deletion was confirmed. */
export interface SheetRemoveDetail {
  id: number;
}

/** Detail for the `sheet-move` event: a sheet and its new position (0 = first). */
export interface SheetMoveDetail {
  id: number;
  index: number;
}

type MenuAction = 'rename' | 'delete' | 'left' | 'right';

/** A press on a tab that may become, or is, a drag that reorders it. */
interface TabDrag {
  id: number;
  pointerId: number;
  tab: HTMLElement;
  startX: number;
  /** Where the pointer is now */
  x: number;
  /** Touch and pen drag only after a hold: until then, moving is a swipe that scrolls the list */
  armed: boolean;
  /** Moved past the threshold: the tab is being dragged */
  active: boolean;
  /** Given up with Escape; the rest of the press does nothing */
  cancelled: boolean;
  /** Position the sheet would get if it were dropped now */
  index: number;
  abort: AbortController;
  holdTimer: ReturnType<typeof setTimeout> | undefined;
  frame: number | undefined;
}

/** How far (px) a pressed tab has to move sideways before it is dragged */
const DRAG_THRESHOLD = 5;
/** How long (ms) a finger or pen has to rest on a tab before it can drag it */
const DRAG_HOLD_MS = 300;
/** Width (px) of the zones at the ends of the tab list that scroll it during a drag */
const DRAG_SCROLL_EDGE = 24;
/** Fastest scroll during a drag, in px per frame */
const DRAG_SCROLL_SPEED = 12;

const MENU: Array<[MenuAction, string]> = [
  ['rename', 'Rename'],
  ['delete', 'Delete'],
  ['left', 'Move left'],
  ['right', 'Move right'],
];

/**
 * `<y11n-sheet-tabs>` -- The sheet bar of `<y11n-spreadsheet>`: one tab per
 * sheet, a button to add a sheet and a menu to rename, delete or move the
 * active one.
 *
 * It holds no sheets itself. The host passes them in and applies the changes
 * it asks for:
 *
 * ```html
 * <y11n-sheet-tabs
 *   .sheets=${sheets}
 *   .activeId=${activeId}
 *   @sheet-select=${(e) => activate(e.detail.id)}
 * ></y11n-sheet-tabs>
 * ```
 *
 * Keys on a tab: Left/Right (and Home/End) switch sheets, F2 renames, Delete
 * asks to delete, Ctrl/Cmd+Shift+Left/Right move the sheet. A tab can also be
 * dragged to another position (with a finger or pen: after holding it briefly).
 *
 * @fires sheet-select - Switch to a sheet (`detail.id`)
 * @fires sheet-add - Add a sheet
 * @fires sheet-rename - Rename a sheet (`detail.id`, `detail.name`); the name has passed `validateSheetName`
 * @fires sheet-remove - Delete a sheet (`detail.id`), after the user confirmed
 * @fires sheet-move - Move a sheet to another position (`detail.id`, `detail.index`)
 */
@customElement('y11n-sheet-tabs')
export class Y11nSheetTabs extends LitElement {
  /** The sheets, in tab order */
  @property({ attribute: false }) sheets: Array<{ id: number; name: string }> = [];
  /** Id of the sheet that is shown */
  @property({ type: Number, attribute: 'active-id' }) activeId = 0;
  /** No adding, renaming, deleting or moving; switching sheets still works */
  @property({ type: Boolean, attribute: 'read-only', reflect: true }) readOnly = false;
  /** Accessible name of the tab list */
  @property({ type: String }) label = 'Sheets';

  /** Sheet whose tab is a name field at the moment */
  @state() private _renamingId: number | null = null;
  @state() private _renameError = '';
  /** Sheet the delete confirmation asks about */
  @state() private _confirmId: number | null = null;
  @state() private _menuOpen = false;
  /** Text of the polite live region */
  @state() private _status = '';

  /** Runs once after the next render, to focus something that render creates */
  private _afterRender: (() => void) | null = null;
  /** The tab that had focus when an update started: a tab that moves or goes away drops it */
  private _focusedTabBefore: number | null = null;
  private _outsideCleanup: (() => void) | null = null;

  /** The press on a tab that is being followed, from pointerdown until it ends */
  private _drag: TabDrag | null = null;
  /** Sheet a move was just asked for: its tab is the one to keep in view when the host reorders */
  private _movedId: number | null = null;
  /** Sheet whose tab is shown as picked up */
  @state() private _dragId: number | null = null;
  /** Where the drop line is, in px from the bar's left edge; null when a drop would change nothing */
  @state() private _dropX: number | null = null;

  /**
   * A drag by finger must keep the browser from panning the list instead.
   * Not passive, and always there: one added mid-touch could come too late.
   */
  private _touchMoveListener = {
    handleEvent: (e: TouchEvent): void => {
      if (this._dragId !== null && e.cancelable) e.preventDefault();
    },
    passive: false,
  };

  disconnectedCallback(): void {
    super.disconnectedCallback();
    this._endDrag();
    this._outsideCleanup?.();
    this._outsideCleanup = null;
    this._afterRender = null;
    this._menuOpen = false;
    this._confirmId = null;
    this._renamingId = null;
    this._renameError = '';
  }

  // ─── Styles ────────────────────────────────────────────

  static styles = css`
    :host {
      display: block;
      font-family: var(--ls-font-family, system-ui, -apple-system, sans-serif);
      font-size: var(--ls-font-size, 13px);
      color: var(--ls-text-color, #333);
      --_border-color: var(--ls-border-color, #e0e0e0);
      --_focus-ring: var(--ls-focus-ring, 2px solid #1a73e8);
      --_accent: var(--ls-sheet-tab-accent, #1a73e8);
      --_hover-bg: var(--ls-sheet-tab-hover-bg, color-mix(in srgb, currentColor 10%, transparent));
    }

    *,
    *::before,
    *::after {
      box-sizing: border-box;
    }

    .bar {
      position: relative;
      display: flex;
      align-items: center;
      gap: 2px;
      min-height: var(--ls-sheet-bar-height, 30px);
      padding-inline-end: 4px;
      background: var(--ls-sheet-bar-bg, var(--ls-header-bg, #f3f3f3));
      border-top: 1px solid var(--_border-color);
    }

    /* Scrolls sideways when the tabs don't fit; the scrollbar only shows then */
    .tabs {
      display: flex;
      flex: 0 1 auto;
      align-self: stretch;
      min-width: 0;
      overflow-x: auto;
      overflow-y: hidden;
      scrollbar-width: thin;
    }

    button,
    input {
      margin: 0;
      font: inherit;
      color: inherit;
    }

    button {
      cursor: pointer;
    }

    .tab {
      flex: 0 0 auto;
      display: inline-flex;
      align-items: center;
      min-width: 40px;
      max-width: var(--ls-sheet-tab-max-width, 160px);
      min-height: 28px;
      padding: 0 12px 2px;
      color: var(--ls-sheet-tab-color, var(--ls-text-color, #333));
      background: transparent;
      border: 0 solid var(--_border-color);
      border-top: 2px solid transparent;
      border-inline-end-width: 1px;
      border-radius: 0;
      user-select: none;
      -webkit-user-select: none;
      -webkit-touch-callout: none;
    }

    .tab.dragging {
      opacity: 0.5;
    }

    .bar.reordering,
    .bar.reordering button {
      cursor: grabbing;
    }

    /* Where a dragged tab would land */
    .drop-indicator {
      position: absolute;
      top: 0;
      bottom: 0;
      z-index: 1;
      width: 2px;
      background: var(--_accent);
      pointer-events: none;
    }

    .tab:hover {
      background: var(--_hover-bg);
    }

    .tab[aria-selected='true'] {
      font-weight: 700;
      color: var(--ls-sheet-tab-active-color, var(--ls-text-color, #333));
      background: var(--ls-sheet-tab-active-bg, var(--ls-cell-bg, #fff));
      border-top-color: var(--_accent);
    }

    .tab-name {
      display: block;
      min-width: 0;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    /* Reserves the width of the bold name, so tabs don't shift when another one becomes active */
    .tab-name::after {
      content: attr(data-name);
      display: block;
      height: 0;
      overflow: hidden;
      font-weight: 700;
      visibility: hidden;
    }

    .rename-input {
      flex: 0 0 auto;
      align-self: center;
      field-sizing: content;
      min-width: 8ch;
      max-width: 34ch;
      min-height: 24px;
      margin: 0 2px;
      padding: 1px 6px;
      background: var(--ls-editor-bg, #fff);
      border: 1px solid var(--_accent);
      border-radius: 3px;
    }

    .rename-input[aria-invalid='true'] {
      border-color: var(--ls-validation-border, #d93025);
    }

    .icon-button {
      flex: 0 0 auto;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 26px;
      height: 26px;
      padding: 0;
      font-size: 16px;
      line-height: 1;
      background: transparent;
      border: 1px solid transparent;
      border-radius: 4px;
    }

    .icon-button:hover,
    .icon-button[aria-expanded='true'] {
      background: var(--_hover-bg);
    }

    /* Menu, delete confirmation and name error open over the grid, above the bar */
    .popover {
      position: absolute;
      bottom: 100%;
      left: 0;
      z-index: 30;
      max-width: 100%;
      margin: 0 0 2px;
      color: var(--ls-text-color, #333);
      background: var(--ls-filter-menu-bg, #fff);
      border: 1px solid var(--_border-color);
      border-radius: 4px;
      box-shadow: var(--ls-editor-shadow, 0 2px 6px rgba(0, 0, 0, 0.2));
    }

    .popover.below {
      top: 100%;
      bottom: auto;
      margin: 2px 0 0;
    }

    .menu {
      display: flex;
      flex-direction: column;
      min-width: 140px;
      padding: 4px;
    }

    .menu button {
      min-height: 28px;
      padding: 0 8px;
      text-align: start;
      white-space: nowrap;
      background: transparent;
      border: 1px solid transparent;
      border-radius: 4px;
    }

    .menu button:hover:not([aria-disabled='true']) {
      background: var(--_hover-bg);
    }

    .menu button[aria-disabled='true'] {
      opacity: 0.5;
      cursor: default;
    }

    .confirm {
      display: flex;
      flex-direction: column;
      gap: 8px;
      width: 260px;
      padding: 10px;
    }

    .confirm p {
      margin: 0;
      overflow-wrap: anywhere;
    }

    .confirm-buttons {
      display: flex;
      justify-content: flex-end;
      gap: 6px;
    }

    .confirm button {
      min-height: 26px;
      padding: 2px 10px;
      background: transparent;
      border: 1px solid var(--_border-color);
      border-radius: 3px;
    }

    .confirm button:hover {
      background: var(--_hover-bg);
    }

    .confirm button.danger {
      color: var(--ls-sheet-delete-color, #fff);
      background: var(--ls-sheet-delete-bg, #c5221f);
      border-color: var(--ls-sheet-delete-bg, #c5221f);
    }

    .confirm button.danger:hover {
      filter: brightness(0.92);
    }

    .error {
      width: max-content;
      padding: 4px 8px;
      font-size: 12px;
      color: var(--ls-validation-color, #8a1c12);
      background: var(--ls-validation-bg, #fdecea);
      border-color: var(--ls-validation-border, #d93025);
    }

    /* Inside the ring: the tab list clips what is drawn outside a tab */
    .tab:focus-visible,
    .rename-input:focus-visible {
      outline: var(--_focus-ring);
      outline-offset: -2px;
    }

    .icon-button:focus-visible,
    .menu:focus-visible,
    .menu button:focus-visible,
    .confirm button:focus-visible {
      outline: var(--_focus-ring);
      outline-offset: 1px;
    }

    .sr-only {
      position: absolute;
      width: 1px;
      height: 1px;
      padding: 0;
      margin: -1px;
      overflow: hidden;
      clip: rect(0, 0, 0, 0);
      border: 0;
    }

    /*
     * Forced colors drop the backgrounds and paint transparent borders: only
     * the active tab keeps its top line (thicker, in the highlight color), and
     * focus is a ring in the text color so it stands apart from that line.
     */
    @media (forced-colors: active) {
      .tab {
        border-top-width: 3px;
        border-top-color: ButtonFace;
      }

      .tab[aria-selected='true'] {
        border-top-color: Highlight;
      }

      .tab:focus-visible,
      .rename-input:focus-visible,
      .icon-button:focus-visible,
      .menu:focus-visible,
      .menu button:focus-visible,
      .confirm button:focus-visible {
        outline: 2px solid ButtonText;
      }

      /* Inside the top line, which still has to tell the active tab apart */
      .tab:focus-visible {
        outline-offset: -5px;
      }

      .icon-button[aria-expanded='true'] {
        border-color: Highlight;
      }

      /* Wider, to stand apart from the lines between the tabs */
      .drop-indicator {
        width: 4px;
        margin-left: -1px;
        forced-color-adjust: none;
        background: Highlight;
      }

      .menu button[aria-disabled='true'] {
        color: GrayText;
        opacity: 1;
      }

      .confirm button.danger {
        border-width: 2px;
      }
    }
  `;

  // ─── Public API ────────────────────────────────────────

  /** Focus the active sheet's tab */
  focusActiveTab(): void {
    const id = this._tabStopId();
    if (id !== undefined) this._focusTab(id);
  }

  /** Start renaming a sheet in place (default: the active one) */
  startRename(id: number = this.activeId): void {
    if (this.readOnly || !this._sheet(id)) return;
    this._closeMenu(false);
    this._confirmId = null;
    this._renamingId = id;
    this._renameError = '';
    this._runAfterRender(() => {
      const input = this._renameInput();
      input?.focus();
      input?.select();
    });
  }

  // ─── Helpers ───────────────────────────────────────────

  private _emit<T>(name: string, detail?: T): void {
    this.dispatchEvent(new CustomEvent<T>(name, { detail, bubbles: true, composed: true }));
  }

  private _sheet(id: number): { id: number; name: string } | undefined {
    return this.sheets.find((s) => s.id === id);
  }

  /** The one tab in the page's tab order: the active sheet's, or the first when no sheet is active. */
  private _tabStopId(): number | undefined {
    return (this._sheet(this.activeId) ?? this.sheets[0])?.id;
  }

  /** An element of the rendered bar; null before the first render. */
  private _query<T extends HTMLElement = HTMLElement>(selector: string): T | null {
    return this.hasUpdated ? this.renderRoot.querySelector<T>(selector) : null;
  }

  private _tab(id: number): HTMLElement | null {
    return this._query(`#sheet-tab-${id}`);
  }

  private _renameInput(): HTMLInputElement | null {
    return this._query<HTMLInputElement>('.rename-input');
  }

  private _focusedElement(): Element | null {
    return this.hasUpdated ? (this.renderRoot as ShadowRoot).activeElement : null;
  }

  /**
   * Do something once the bar has rendered again. What a change asks for is
   * applied by the host, which then passes new sheets in: only after that
   * update is everything in its final place.
   */
  private _runAfterRender(action: () => void): void {
    this._afterRender = action;
    this.requestUpdate();
  }

  /** Focus a sheet's tab: now if it is there, after the next render otherwise. */
  private _focusTab(id: number): void {
    const tab = this._tab(id);
    if (tab) tab.focus();
    else this._runAfterRender(() => (this._tab(id) ?? this._renameInput())?.focus());
  }

  private _announce(text: string): void {
    // A live region only speaks when its text changes: make a repeat differ
    this._status = text === this._status ? `${text} ` : text;
  }

  // ─── Lifecycle ─────────────────────────────────────────

  protected willUpdate(): void {
    // Whatever was open for a sheet that is gone, or can't be edited any more, closes
    if (this._renamingId !== null && (this.readOnly || !this._sheet(this._renamingId))) {
      this._renamingId = null;
      this._renameError = '';
    }
    if (this._confirmId !== null && (this.readOnly || this.sheets.length <= 1 || !this._sheet(this._confirmId))) {
      this._confirmId = null;
    }
    if (this.readOnly) this._menuOpen = false;
    if (this._drag && (this.readOnly || this._renamingId !== null || !this._sheet(this._drag.id))) this._endDrag();

    const focused = this._focusedElement();
    this._focusedTabBefore =
      focused instanceof HTMLElement && focused.getAttribute('role') === 'tab' ? Number(focused.dataset.sheetId) : null;
  }

  protected updated(changed: PropertyValues<this>): void {
    this._syncOutsideListener();
    this._placePopovers();

    const action = this._afterRender;
    this._afterRender = null;
    if (action) {
      action();
    } else if (this._focusedTabBefore !== null && !this._focusedElement()) {
      // The focused tab moved (a moved element loses focus) or its sheet was removed
      const id = this._sheet(this._focusedTabBefore) ? this._focusedTabBefore : this._tabStopId();
      if (id !== undefined) this._tab(id)?.focus();
    }

    if (changed.has('activeId') || changed.has('sheets')) {
      // After a move the list stays with the tab that was moved; otherwise it shows the active one
      const moved = changed.has('activeId') || this._movedId === null ? undefined : this._sheet(this._movedId);
      this._movedId = null;
      this._revealTab(moved ? moved.id : this._tabStopId());
    }
  }

  /**
   * Scroll the tab list so a tab is in view. Not `scrollIntoView()`: that
   * would also scroll the page to the bar whenever the sheet changes.
   */
  private _revealTab(id: number | undefined): void {
    const list = this._query('.tabs');
    const target = id === undefined ? null : (this._tab(id) ?? this._renameInput());
    if (!list || !target) return;
    const outer = list.getBoundingClientRect();
    const inner = target.getBoundingClientRect();
    if (inner.left < outer.left) list.scrollLeft -= outer.left - inner.left;
    else if (inner.right > outer.right) list.scrollLeft += inner.right - outer.right;
  }

  /**
   * Line each popover up with what it belongs to (`data-anchor`), kept inside
   * the bar's width. They open upwards, over the grid, unless only the space
   * below the bar is big enough.
   */
  private _placePopovers(): void {
    const bar = this._query('.bar');
    if (!bar) return;
    for (const popover of this.renderRoot.querySelectorAll<HTMLElement>('.popover')) {
      const selector = popover.dataset.anchor;
      const anchor = (selector ? this._query(selector) : null) ?? bar;
      // Measured from the bar's start, where its width doesn't depend on the last position
      popover.style.left = '0px';
      const barRect = bar.getBoundingClientRect();
      const size = popover.getBoundingClientRect();
      const left = Math.min(anchor.getBoundingClientRect().left - barRect.left, barRect.width - size.width);
      popover.style.left = `${Math.max(0, left)}px`;
      const fitsAbove = barRect.top >= size.height;
      const fitsBelow = window.innerHeight - barRect.bottom >= size.height;
      popover.classList.toggle('below', !fitsAbove && fitsBelow);
    }
  }

  /** While the menu or the confirmation is open, pressing anywhere else closes it, like a native dropdown. */
  private _syncOutsideListener(): void {
    const needed = this._menuOpen || this._confirmId !== null;
    if (needed && !this._outsideCleanup) {
      const onPointerDown = (e: Event) => {
        const path = e.composedPath();
        const inside = (selector: string) => {
          const el = this._query(selector);
          return el !== null && path.includes(el);
        };
        if (this._menuOpen && !inside('.menu') && !inside('.menu-button')) this._closeMenu(false);
        if (this._confirmId !== null && !inside('.confirm')) this._confirmId = null;
      };
      document.addEventListener('pointerdown', onPointerDown);
      this._outsideCleanup = () => document.removeEventListener('pointerdown', onPointerDown);
    } else if (!needed && this._outsideCleanup) {
      this._outsideCleanup();
      this._outsideCleanup = null;
    }
  }

  // ─── Tabs ──────────────────────────────────────────────

  private _select(id: number): void {
    this._movedId = null;
    // A name left invalid stays open until the user goes on to something else
    this._endRename(false);
    if (id !== this.activeId) this._emit<SheetSelectDetail>('sheet-select', { id });
  }

  private _add(): void {
    if (this.readOnly) return;
    this._endRename(false);
    this._emit('sheet-add');
  }

  /** Move a sheet one position; nothing happens at either end. */
  private _move(id: number, delta: -1 | 1): void {
    const index = this.sheets.findIndex((s) => s.id === id);
    const target = index + delta;
    if (this.readOnly || index < 0 || target < 0 || target >= this.sheets.length) return;
    const name = this.sheets[index].name;
    this._movedId = id;
    this._emit<SheetMoveDetail>('sheet-move', { id, index: target });
    this._announce(`Moved ${name} ${delta < 0 ? 'left' : 'right'}`);
  }

  private _handleTabKeydown(e: KeyboardEvent): void {
    const id = Number((e.currentTarget as HTMLElement).dataset.sheetId);
    const index = this.sheets.findIndex((s) => s.id === id);
    if (index < 0 || e.altKey) return;
    const count = this.sheets.length;
    const mod = e.ctrlKey || e.metaKey;
    const plain = !mod && !e.shiftKey;
    const arrow = e.key === 'ArrowLeft' ? -1 : e.key === 'ArrowRight' ? 1 : 0;

    if (arrow !== 0 && mod && e.shiftKey) {
      if (this.readOnly) return;
      // The tab moves with its sheet; updated() gives it focus back if the move drops it
      this._move(id, arrow);
    } else if (arrow !== 0 && plain) {
      this._goTo((index + arrow + count) % count);
    } else if (e.key === 'Home' && plain) {
      this._goTo(0);
    } else if (e.key === 'End' && plain) {
      this._goTo(count - 1);
    } else if ((e.key === 'Enter' || e.key === ' ') && plain) {
      this._select(id);
    } else if (e.key === 'F2' && plain) {
      if (this.readOnly) return;
      this.startRename(id);
    } else if ((e.key === 'Delete' || e.key === 'Backspace') && plain) {
      if (this.readOnly) return;
      this._requestDelete(id);
    } else {
      return;
    }
    // Handled here: not a grid shortcut as well
    e.preventDefault();
    e.stopPropagation();
  }

  /** Focus a tab and switch to its sheet (tabs activate as focus moves along them). */
  private _goTo(index: number): void {
    const sheet = this.sheets[index];
    if (!sheet) return;
    this._select(sheet.id);
    this._focusTab(sheet.id);
  }

  // ─── Drag to reorder ───────────────────────────────────

  /**
   * Follow a press on a tab. A mouse drags the tab once it has moved a little
   * sideways; a finger or pen first has to rest on it, because moving straight
   * away is how they scroll the list. Anything less stays a click.
   */
  private _handleTabPointerDown(e: PointerEvent): void {
    if (this._drag?.active && !this._drag.cancelled) return;
    this._endDrag();
    if (this.readOnly || this._renamingId !== null || e.button !== 0 || this.sheets.length < 2) return;
    const tab = e.currentTarget as HTMLElement;
    const id = Number(tab.dataset.sheetId);
    const index = this.sheets.findIndex((s) => s.id === id);
    if (index < 0) return;

    const hold = e.pointerType === 'touch' || e.pointerType === 'pen';
    const abort = new AbortController();
    const drag: TabDrag = {
      id,
      pointerId: e.pointerId,
      tab,
      startX: e.clientX,
      x: e.clientX,
      armed: !hold,
      active: false,
      cancelled: false,
      index,
      abort,
      holdTimer: undefined,
      frame: undefined,
    };
    this._drag = drag;
    if (hold) {
      drag.holdTimer = setTimeout(() => {
        drag.armed = true;
        // Picked up: shown before it moves, as there is no cursor to tell
        this._dragId = id;
      }, DRAG_HOLD_MS);
    }

    const { signal } = abort;
    const onMove = (ev: PointerEvent) => {
      if (ev.pointerId !== drag.pointerId || drag.cancelled) return;
      drag.x = ev.clientX;
      if (!drag.active) {
        if (Math.abs(ev.clientX - drag.startX) <= DRAG_THRESHOLD) return;
        if (!drag.armed) {
          // A swipe: the browser scrolls the list
          this._endDrag();
          return;
        }
        this._startDrag(drag);
      }
      this._updateDrop();
    };
    const onUp = (ev: PointerEvent) => {
      if (ev.pointerId !== drag.pointerId) return;
      drag.x = ev.clientX;
      this._drop(drag);
    };
    const onKeydown = (ev: KeyboardEvent) => {
      if (ev.key !== 'Escape' || this._dragId === null) return;
      ev.preventDefault();
      ev.stopPropagation();
      this._cancelDrag(drag);
    };
    document.addEventListener('pointermove', onMove, { signal });
    document.addEventListener('pointerup', onUp, { signal });
    const onCancel = (ev: PointerEvent) => {
      if (ev.pointerId === drag.pointerId) this._endDrag();
    };
    document.addEventListener('pointercancel', onCancel, { signal });
    document.addEventListener('keydown', onKeydown, { signal, capture: true });
    window.addEventListener('blur', () => this._endDrag(), { signal });
  }

  private _handleTabContextMenu(e: Event): void {
    // A long press on a touch screen asks for a context menu as well: by then it has picked the tab up
    if (this._dragId !== null) e.preventDefault();
    else this._endDrag();
  }

  private _startDrag(drag: TabDrag): void {
    drag.active = true;
    this._dragId = drag.id;
    try {
      // Keeps the moves coming (and the cursor) wherever the pointer goes
      drag.tab.setPointerCapture?.(drag.pointerId);
    } catch {
      // Not an active pointer: the document listeners still follow it
    }
    drag.frame = requestAnimationFrame(() => this._autoScroll(drag));
  }

  /**
   * Work out where the dragged tab would land: after every other tab whose
   * middle the pointer has passed. The result is the sheet's new position, as
   * in `sheet-move`.
   */
  private _updateDrop(): void {
    const drag = this._drag;
    const bar = this._query('.bar');
    const list = this._query('.tabs');
    if (!drag?.active || drag.cancelled || !bar || !list) return;
    const rtl = getComputedStyle(list).direction === 'rtl';
    const rects: DOMRect[] = [];
    for (const sheet of this.sheets) {
      const tab = sheet.id === drag.id ? null : this._tab(sheet.id);
      if (tab) rects.push(tab.getBoundingClientRect());
    }
    let index = 0;
    for (const rect of rects) {
      const middle = rect.left + rect.width / 2;
      if (rtl ? drag.x < middle : drag.x > middle) index++;
    }
    drag.index = index;

    if (rects.length === 0 || index === this.sheets.findIndex((s) => s.id === drag.id)) {
      this._dropX = null;
      return;
    }
    // The line goes before the tab that will follow the dropped one, or after the last tab
    const next = rects[index];
    const last = rects[rects.length - 1];
    let x = next ? (rtl ? next.right : next.left) : rtl ? last.left : last.right;
    // Kept on the part of the list that is in view
    const visible = list.getBoundingClientRect();
    if (visible.width > 2) x = Math.min(Math.max(x, visible.left + 1), visible.right - 1);
    this._dropX = Math.round(x - bar.getBoundingClientRect().left - 1);
  }

  /** While a tab is dragged near (or past) an end of the list, scroll the list that way. */
  private _autoScroll(drag: TabDrag): void {
    if (this._drag !== drag || !drag.active || drag.cancelled) return;
    const list = this._query('.tabs');
    if (list && list.scrollWidth > list.clientWidth) {
      const rect = list.getBoundingClientRect();
      const intoStart = drag.x - (rect.left + DRAG_SCROLL_EDGE);
      const intoEnd = drag.x - (rect.right - DRAG_SCROLL_EDGE);
      const depth = intoStart < 0 ? intoStart : intoEnd > 0 ? intoEnd : 0;
      if (depth !== 0) {
        // Faster the further into the zone
        const speed = Math.ceil(Math.min(Math.abs(depth) / DRAG_SCROLL_EDGE, 1) * DRAG_SCROLL_SPEED);
        const before = list.scrollLeft;
        list.scrollLeft += Math.sign(depth) * speed;
        if (list.scrollLeft !== before) this._updateDrop();
      }
    }
    drag.frame = requestAnimationFrame(() => this._autoScroll(drag));
  }

  /** The press ended: move the sheet if it was dragged somewhere else. */
  private _drop(drag: TabDrag): void {
    const dragged = drag.active;
    if (dragged && !drag.cancelled) this._updateDrop();
    const { id, index } = drag;
    const cancelled = drag.cancelled;
    this._endDrag();
    if (!dragged) return;
    // The release is not a click on whatever is under the pointer now
    this._suppressClick();
    const origin = this.sheets.findIndex((s) => s.id === id);
    if (cancelled || origin < 0 || index === origin) return;
    const { name } = this.sheets[origin];
    const total = this.sheets.length;
    this._movedId = id;
    this._emit<SheetMoveDetail>('sheet-move', { id, index });
    this._announce(`Moved ${name} to position ${index + 1} of ${total}`);
  }

  /** Escape: put the tab back. The press is still followed, so that its release does nothing. */
  private _cancelDrag(drag: TabDrag): void {
    drag.cancelled = true;
    clearTimeout(drag.holdTimer);
    this._stopDragging(drag);
  }

  /** Stop following the press, whatever state it is in. */
  private _endDrag(): void {
    const drag = this._drag;
    if (!drag) return;
    this._drag = null;
    drag.abort.abort();
    clearTimeout(drag.holdTimer);
    this._stopDragging(drag);
  }

  private _stopDragging(drag: TabDrag): void {
    if (drag.frame !== undefined) cancelAnimationFrame(drag.frame);
    drag.frame = undefined;
    try {
      if (drag.tab.hasPointerCapture?.(drag.pointerId)) drag.tab.releasePointerCapture(drag.pointerId);
    } catch {
      // Nothing was captured
    }
    this._dragId = null;
    this._dropX = null;
  }

  /** Swallow the click that the browser sends right after the pointer is released. */
  private _suppressClick(): void {
    const root = this.renderRoot;
    const swallow = (e: Event) => {
      e.stopPropagation();
      e.preventDefault();
    };
    root.addEventListener('click', swallow, { capture: true, once: true });
    setTimeout(() => root.removeEventListener('click', swallow, { capture: true }), 0);
  }

  // ─── Rename ────────────────────────────────────────────

  private _endRename(refocus: boolean): void {
    const id = this._renamingId;
    if (id === null) return;
    this._renamingId = null;
    this._renameError = '';
    if (refocus) this._runAfterRender(() => this._tab(id)?.focus());
  }

  private _commitRename(refocus: boolean): void {
    const id = this._renamingId;
    if (id === null) return;
    const sheet = this._sheet(id);
    const input = this._renameInput();
    if (!sheet || !input) {
      this._endRename(refocus);
      return;
    }
    // Spaces around a name are never wanted (and not allowed)
    const name = input.value.trim();
    if (name === sheet.name) {
      this._endRename(refocus);
      return;
    }
    const error = validateSheetName(
      name,
      this.sheets.filter((s) => s.id !== id).map((s) => s.name)
    );
    if (error) {
      this._renameError = error;
      return;
    }
    this._endRename(refocus);
    this._emit<SheetRenameDetail>('sheet-rename', { id, name });
  }

  private _handleRenameKeydown(e: KeyboardEvent): void {
    // Keys typed in the name are not tab or grid shortcuts
    e.stopPropagation();
    if (e.isComposing) return;
    if (e.key === 'Enter') {
      e.preventDefault();
      this._commitRename(true);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      this._endRename(true);
    }
  }

  private _handleRenameBlur(e: FocusEvent): void {
    // Still the focused element: the window lost focus, not the field. Keep editing
    if (this._focusedElement() === e.target) return;
    // Focus is already where the user sent it: commit without taking it back
    this._commitRename(false);
  }

  private _handleRenameInput(): void {
    if (this._renameError) this._renameError = '';
  }

  // ─── Delete ────────────────────────────────────────────

  /** Open the confirmation for a sheet. The last sheet can't be deleted. */
  private _requestDelete(id: number): void {
    if (this.readOnly || !this._sheet(id)) return;
    if (this.sheets.length <= 1) {
      this._announce('The only sheet cannot be deleted');
      return;
    }
    this._closeMenu(false);
    this._endRename(false);
    this._confirmId = id;
    this._runAfterRender(() => this._query('#sheet-delete-cancel')?.focus());
  }

  private _closeConfirm(): void {
    const id = this._confirmId;
    if (id === null) return;
    this._confirmId = null;
    this._runAfterRender(() => (this._tab(id) ?? this._tab(this.activeId))?.focus());
  }

  private _confirmDelete(): void {
    const id = this._confirmId;
    if (id === null) return;
    this._confirmId = null;
    this._emit<SheetRemoveDetail>('sheet-remove', { id });
    // The sheet's tab goes away: focus the tab of whichever sheet the host shows now
    this._runAfterRender(() => {
      const next = this._tabStopId();
      if (next !== undefined) this._tab(next)?.focus();
    });
  }

  private _handleConfirmKeydown(e: KeyboardEvent): void {
    // Keys pressed in the confirmation are not grid shortcuts
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      this._closeConfirm();
    } else if (e.key === 'Tab') {
      // Modal: focus stays on its two buttons
      const buttons = [...this.renderRoot.querySelectorAll<HTMLElement>('.confirm button')];
      const current = buttons.indexOf(this._focusedElement() as HTMLElement);
      e.preventDefault();
      buttons[(current + (e.shiftKey ? -1 : 1) + buttons.length) % buttons.length]?.focus();
    }
  }

  // ─── Menu ──────────────────────────────────────────────

  private _menuDisabled(action: MenuAction): boolean {
    const index = this.sheets.findIndex((s) => s.id === this.activeId);
    if (index < 0) return true;
    if (action === 'delete') return this.sheets.length <= 1;
    if (action === 'left') return index === 0;
    if (action === 'right') return index === this.sheets.length - 1;
    return false;
  }

  private _menuItems(): HTMLElement[] {
    return [...this.renderRoot.querySelectorAll<HTMLElement>('.menu [role="menuitem"]:not([aria-disabled="true"])')];
  }

  private _openMenu(focus: 'first' | 'last' = 'first'): void {
    if (this.readOnly) return;
    this._endRename(false);
    this._confirmId = null;
    this._menuOpen = true;
    this._runAfterRender(() => {
      const items = this._menuItems();
      const target = focus === 'first' ? items[0] : items[items.length - 1];
      // With nothing to choose, the menu itself takes focus so Escape still closes it
      (target ?? this._query('.menu'))?.focus();
    });
  }

  private _closeMenu(focusButton: boolean): void {
    if (!this._menuOpen) return;
    this._menuOpen = false;
    if (focusButton) this._query('.menu-button')?.focus();
  }

  private _handleMenuButtonKeydown(e: KeyboardEvent): void {
    if (e.key === 'Enter' || e.key === ' ' || e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      // Enter and Space are handled here, so they don't click the button as well
      e.preventDefault();
      e.stopPropagation();
      this._openMenu(e.key === 'ArrowUp' ? 'last' : 'first');
    } else if (e.key === 'Escape' && this._menuOpen) {
      e.preventDefault();
      e.stopPropagation();
      this._closeMenu(false);
    }
  }

  private _handleMenuKeydown(e: KeyboardEvent): void {
    // Keys pressed in the menu are not grid shortcuts
    e.stopPropagation();
    const items = this._menuItems();
    const current = items.indexOf(this._focusedElement() as HTMLElement);
    let next = -1;
    if (e.key === 'ArrowDown') next = (current + 1) % items.length;
    else if (e.key === 'ArrowUp') next = (Math.max(current, 0) - 1 + items.length) % items.length;
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

  private _chooseMenuItem(action: MenuAction): void {
    if (this._menuDisabled(action)) return;
    const id = this.activeId;
    if (action === 'rename') {
      this._closeMenu(false);
      this.startRename(id);
    } else if (action === 'delete') {
      this._closeMenu(false);
      this._requestDelete(id);
    } else {
      // Back on the button, the sheet can be moved again straight away
      this._closeMenu(true);
      this._move(id, action === 'left' ? -1 : 1);
    }
  }

  // ─── Render ────────────────────────────────────────────

  private _renderTab(sheet: { id: number; name: string }, tabStop: number | undefined) {
    if (sheet.id === this._renamingId) {
      const invalid = this._renameError !== '';
      return html`<input
        class="rename-input"
        type="text"
        size="16"
        maxlength="${MAX_SHEET_NAME_LENGTH}"
        autocomplete="off"
        spellcheck="false"
        data-sheet-id="${sheet.id}"
        aria-label="Rename sheet ${sheet.name}"
        aria-invalid="${invalid ? 'true' : nothing}"
        aria-describedby="${invalid ? 'sheet-rename-error' : nothing}"
        .value="${sheet.name}"
        @keydown="${this._handleRenameKeydown}"
        @input="${this._handleRenameInput}"
        @blur="${this._handleRenameBlur}"
      />`;
    }
    const selected = sheet.id === this.activeId;
    return html`<button
      type="button"
      role="tab"
      class="tab ${sheet.id === this._dragId ? 'dragging' : ''}"
      id="sheet-tab-${sheet.id}"
      data-sheet-id="${sheet.id}"
      aria-selected="${selected}"
      tabindex="${sheet.id === tabStop ? 0 : -1}"
      title="${sheet.name}"
      @click="${() => this._select(sheet.id)}"
      @dblclick="${() => this.startRename(sheet.id)}"
      @keydown="${this._handleTabKeydown}"
      @pointerdown="${this._handleTabPointerDown}"
      @contextmenu="${this._handleTabContextMenu}"
    ><span class="tab-name" data-name="${sheet.name}">${sheet.name}</span></button>`;
  }

  private _renderMenu() {
    const active = this._sheet(this.activeId);
    return html`<div
      id="sheet-menu"
      class="popover menu"
      role="menu"
      tabindex="-1"
      aria-label="${active ? `Options for ${active.name}` : 'Sheet options'}"
      data-anchor=".menu-button"
      @keydown="${this._handleMenuKeydown}"
    >
      ${MENU.map(
        ([action, text]) => html`<button
          type="button"
          role="menuitem"
          tabindex="-1"
          data-action="${action}"
          aria-disabled="${this._menuDisabled(action)}"
          @click="${() => this._chooseMenuItem(action)}"
        >${text}</button>`
      )}
    </div>`;
  }

  private _renderConfirm(sheet: { id: number; name: string }) {
    return html`<div
      class="popover confirm"
      role="alertdialog"
      aria-modal="true"
      aria-label="Delete sheet"
      aria-describedby="sheet-delete-text"
      data-anchor="#sheet-tab-${sheet.id}"
      @keydown="${this._handleConfirmKeydown}"
    >
      <p id="sheet-delete-text">Delete "${sheet.name}"? This cannot be undone.</p>
      <div class="confirm-buttons">
        <button id="sheet-delete-confirm" type="button" class="danger" @click="${this._confirmDelete}">Delete</button>
        <button id="sheet-delete-cancel" type="button" @click="${this._closeConfirm}">Cancel</button>
      </div>
    </div>`;
  }

  protected render() {
    const tabStop = this._tabStopId();
    const confirming = this._confirmId === null ? undefined : this._sheet(this._confirmId);
    return html`
      <div class="bar ${this._dragId !== null ? 'reordering' : ''}">
        <div
          class="tabs"
          role="tablist"
          aria-label="${this.label}"
          aria-orientation="horizontal"
          @touchmove="${this._touchMoveListener}"
        >
          ${repeat(
            this.sheets,
            (sheet) => sheet.id,
            (sheet) => this._renderTab(sheet, tabStop)
          )}
        </div>
        ${this.readOnly
          ? nothing
          : html`
              <button class="icon-button add-button" type="button" aria-label="Add sheet" title="Add sheet" @click="${this._add}">
                <span aria-hidden="true">+</span>
              </button>
              <button
                class="icon-button menu-button"
                type="button"
                aria-label="Sheet options"
                aria-haspopup="menu"
                aria-expanded="${this._menuOpen}"
                title="Sheet options"
                @click="${() => (this._menuOpen ? this._closeMenu(false) : this._openMenu())}"
                @keydown="${this._handleMenuButtonKeydown}"
              >
                <svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true" focusable="false">
                  <circle cx="3" cy="8" r="1.4" fill="currentColor" />
                  <circle cx="8" cy="8" r="1.4" fill="currentColor" />
                  <circle cx="13" cy="8" r="1.4" fill="currentColor" />
                </svg>
              </button>
            `}
        ${this._menuOpen ? this._renderMenu() : nothing}
        ${confirming ? this._renderConfirm(confirming) : nothing}
        ${this._renameError
          ? html`<div id="sheet-rename-error" class="popover error" role="alert" data-anchor=".rename-input">
              ${this._renameError}
            </div>`
          : nothing}
        ${this._dropX === null
          ? nothing
          : html`<div class="drop-indicator" aria-hidden="true" style="left: ${this._dropX}px"></div>`}
        <div class="sr-only" role="status">${this._status}</div>
      </div>
    `;
  }
}

declare global {
  interface HTMLElementTagNameMap {
    'y11n-sheet-tabs': Y11nSheetTabs;
  }
}
