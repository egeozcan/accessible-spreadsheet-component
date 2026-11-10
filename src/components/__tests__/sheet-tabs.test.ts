// @vitest-environment jsdom
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import '../y11n-sheet-tabs.js';
import type { Y11nSheetTabs } from '../y11n-sheet-tabs.js';

const THREE = [
  { id: 1, name: 'Sheet1' },
  { id: 2, name: 'Budget' },
  { id: 3, name: 'Notes' },
];

describe('Y11nSheetTabs', () => {
  let el: Y11nSheetTabs;

  const root = () => el.shadowRoot!;
  const tabs = () => [...root().querySelectorAll<HTMLElement>('[role="tab"]')];
  const tab = (id: number) => root().querySelector<HTMLElement>(`#sheet-tab-${id}`)!;
  const addButton = () => root().querySelector<HTMLButtonElement>('[aria-label="Add sheet"]');
  const menuButton = () => root().querySelector<HTMLButtonElement>('[aria-label="Sheet options"]');
  const menu = () => root().querySelector<HTMLElement>('[role="menu"]');
  const menuItem = (action: string) => root().querySelector<HTMLElement>(`[role="menuitem"][data-action="${action}"]`)!;
  const renameInput = () => root().querySelector<HTMLInputElement>('input');
  const alert = () => root().querySelector<HTMLElement>('[role="alert"]');
  const confirmDialog = () => root().querySelector<HTMLElement>('[role="alertdialog"]');
  const dialogButton = (text: string) =>
    [...confirmDialog()!.querySelectorAll('button')].find((b) => b.textContent === text) as HTMLButtonElement;
  const status = () => root().querySelector<HTMLElement>('[role="status"]')!.textContent!.trim();

  const key = (target: Element, name: string, init: KeyboardEventInit = {}) => {
    const event = new KeyboardEvent('keydown', { key: name, bubbles: true, composed: true, cancelable: true, ...init });
    target.dispatchEvent(event);
    return event;
  };

  /** Let the component render, and run what it does once it has. */
  const settle = async () => {
    await el.updateComplete;
    await el.updateComplete;
  };

  const listen = (name: string) => {
    const handler = vi.fn();
    el.addEventListener(name, handler);
    return handler;
  };
  const details = (handler: ReturnType<typeof vi.fn>) => handler.mock.calls.map((call) => call[0].detail);

  async function setSheets(sheets: Array<{ id: number; name: string }>, activeId = el.activeId) {
    el.sheets = sheets;
    el.activeId = activeId;
    await settle();
  }

  async function startRename(id: number, text?: string) {
    el.startRename(id);
    await settle();
    const input = renameInput()!;
    if (text !== undefined) {
      input.value = text;
      input.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
    }
    return input;
  }

  async function openMenu() {
    menuButton()!.click();
    await settle();
  }

  beforeEach(async () => {
    el = document.createElement('y11n-sheet-tabs') as Y11nSheetTabs;
    el.sheets = THREE;
    el.activeId = 2;
    document.body.appendChild(el);
    await settle();
  });

  afterEach(() => {
    el.remove();
  });

  describe('rendering', () => {
    it('renders a labelled horizontal tab list with a tab per sheet', () => {
      const list = root().querySelector('[role="tablist"]')!;
      expect(list.getAttribute('aria-label')).toBe('Sheets');
      expect(list.getAttribute('aria-orientation')).toBe('horizontal');
      expect(tabs().map((t) => t.textContent!.trim())).toEqual(['Sheet1', 'Budget', 'Notes']);
      expect(tabs().every((t) => t.tagName === 'BUTTON')).toBe(true);
    });

    it('identifies each tab and gives it a title for truncated names', () => {
      expect(tabs().map((t) => t.id)).toEqual(['sheet-tab-1', 'sheet-tab-2', 'sheet-tab-3']);
      expect(tabs().map((t) => t.dataset.sheetId)).toEqual(['1', '2', '3']);
      expect(tab(2).title).toBe('Budget');
    });

    it('marks only the active tab as selected and keeps only it in the tab order', async () => {
      expect(tabs().map((t) => t.getAttribute('aria-selected'))).toEqual(['false', 'true', 'false']);
      expect(tabs().map((t) => t.tabIndex)).toEqual([-1, 0, -1]);

      el.activeId = 3;
      await settle();
      expect(tabs().map((t) => t.getAttribute('aria-selected'))).toEqual(['false', 'false', 'true']);
      expect(tabs().map((t) => t.tabIndex)).toEqual([-1, -1, 0]);
    });

    it('keeps the first tab reachable when no sheet is active', async () => {
      el.activeId = 99;
      await settle();
      expect(tabs().map((t) => t.getAttribute('aria-selected'))).toEqual(['false', 'false', 'false']);
      expect(tabs().map((t) => t.tabIndex)).toEqual([0, -1, -1]);
    });

    it('uses the label property as the name of the tab list', async () => {
      el.label = 'Worksheets';
      await settle();
      expect(root().querySelector('[role="tablist"]')!.getAttribute('aria-label')).toBe('Worksheets');
    });

    it('renders the add button and a closed options menu button', () => {
      expect(addButton()!.textContent!.trim()).toBe('+');
      expect(menuButton()!.getAttribute('aria-haspopup')).toBe('menu');
      expect(menuButton()!.getAttribute('aria-expanded')).toBe('false');
      expect(menu()).toBeNull();
    });

    it('focusActiveTab() focuses the active tab', () => {
      el.focusActiveTab();
      expect(root().activeElement).toBe(tab(2));
    });
  });

  describe('selecting', () => {
    it('fires sheet-select when another tab is clicked, without changing anything itself', async () => {
      const select = listen('sheet-select');
      tab(3).click();
      await settle();
      expect(details(select)).toEqual([{ id: 3 }]);
      const event = select.mock.calls[0][0] as CustomEvent;
      expect(event.bubbles).toBe(true);
      expect(event.composed).toBe(true);
      expect(el.activeId).toBe(2);
      expect(tab(2).getAttribute('aria-selected')).toBe('true');
    });

    it('fires nothing when the active tab is clicked', () => {
      const select = listen('sheet-select');
      tab(2).click();
      expect(select).not.toHaveBeenCalled();
    });

    it('moves focus and selects with ArrowRight and ArrowLeft', () => {
      const select = listen('sheet-select');
      tab(2).focus();
      key(tab(2), 'ArrowRight');
      expect(root().activeElement).toBe(tab(3));
      key(tab(3), 'ArrowLeft');
      expect(root().activeElement).toBe(tab(2));
      key(tab(2), 'ArrowLeft');
      expect(root().activeElement).toBe(tab(1));
      // Back on the active sheet in between: nothing to select
      expect(details(select)).toEqual([{ id: 3 }, { id: 1 }]);
    });

    it('wraps around at both ends', () => {
      const select = listen('sheet-select');
      key(tab(3), 'ArrowRight');
      expect(root().activeElement).toBe(tab(1));
      key(tab(1), 'ArrowLeft');
      expect(root().activeElement).toBe(tab(3));
      expect(details(select)).toEqual([{ id: 1 }, { id: 3 }]);
    });

    it('goes to the first and last tab with Home and End', () => {
      const select = listen('sheet-select');
      key(tab(2), 'Home');
      expect(root().activeElement).toBe(tab(1));
      key(tab(1), 'End');
      expect(root().activeElement).toBe(tab(3));
      expect(details(select)).toEqual([{ id: 1 }, { id: 3 }]);
    });

    it('selects the focused tab with Enter and Space', () => {
      const select = listen('sheet-select');
      key(tab(1), 'Enter');
      key(tab(3), ' ');
      key(tab(2), 'Enter');
      expect(details(select)).toEqual([{ id: 1 }, { id: 3 }]);
    });

    it('keeps handled keys away from the grid and lets the others through', () => {
      const outer = vi.fn();
      el.addEventListener('keydown', outer);
      for (const name of ['ArrowLeft', 'ArrowRight', 'Home', 'End', 'Enter', ' ', 'F2', 'Delete']) {
        expect(key(tab(2), name).defaultPrevented, name).toBe(true);
      }
      expect(key(tab(2), 'ArrowLeft', { ctrlKey: true, shiftKey: true }).defaultPrevented).toBe(true);
      expect(outer).not.toHaveBeenCalled();

      expect(key(tab(2), 'a').defaultPrevented).toBe(false);
      expect(key(tab(2), 'ArrowDown').defaultPrevented).toBe(false);
      expect(outer).toHaveBeenCalledTimes(2);
    });
  });

  describe('adding', () => {
    it('fires sheet-add without a detail', () => {
      const add = listen('sheet-add');
      addButton()!.click();
      expect(add).toHaveBeenCalledOnce();
      expect(add.mock.calls[0][0].detail).toBeNull();
    });
  });

  describe('renaming', () => {
    it('replaces the tab with a filled-in, selected name field on double-click', async () => {
      tab(2).dispatchEvent(new MouseEvent('dblclick', { bubbles: true, composed: true }));
      await settle();
      const input = renameInput()!;
      expect(root().querySelector('#sheet-tab-2')).toBeNull();
      expect(tabs()).toHaveLength(2);
      expect(input.getAttribute('aria-label')).toBe('Rename sheet Budget');
      expect(input.maxLength).toBe(31);
      expect(input.value).toBe('Budget');
      expect(root().activeElement).toBe(input);
      expect([input.selectionStart, input.selectionEnd]).toEqual([0, 6]);
    });

    it('starts on F2, and from startRename() for the active sheet by default', async () => {
      key(tab(3), 'F2');
      await settle();
      expect(renameInput()!.getAttribute('aria-label')).toBe('Rename sheet Notes');
      key(renameInput()!, 'Escape');
      await settle();

      el.startRename();
      await settle();
      expect(renameInput()!.getAttribute('aria-label')).toBe('Rename sheet Budget');
    });

    it('commits on Enter: fires sheet-rename and returns focus to the tab', async () => {
      const rename = listen('sheet-rename');
      const input = await startRename(2, 'Costs');
      expect(key(input, 'Enter').defaultPrevented).toBe(true);
      await settle();
      expect(details(rename)).toEqual([{ id: 2, name: 'Costs' }]);
      expect(renameInput()).toBeNull();
      expect(root().activeElement).toBe(tab(2));
      // The host renames the sheet; until then the tab shows what it was given
      expect(tab(2).textContent!.trim()).toBe('Budget');
    });

    it('commits when the field loses focus, leaving focus where it went', async () => {
      const rename = listen('sheet-rename');
      await startRename(2, 'Costs');
      addButton()!.focus();
      await settle();
      expect(details(rename)).toEqual([{ id: 2, name: 'Costs' }]);
      expect(renameInput()).toBeNull();
      expect(root().activeElement).toBe(addButton());
    });

    it('cancels on Escape: fires nothing and returns focus to the tab', async () => {
      const rename = listen('sheet-rename');
      const input = await startRename(2, 'Costs');
      key(input, 'Escape');
      await settle();
      expect(rename).not.toHaveBeenCalled();
      expect(renameInput()).toBeNull();
      expect(root().activeElement).toBe(tab(2));
    });

    it('fires nothing for an unchanged name', async () => {
      const rename = listen('sheet-rename');
      key(await startRename(2), 'Enter');
      await settle();
      key(await startRename(2, '  Budget '), 'Enter');
      await settle();
      expect(rename).not.toHaveBeenCalled();
      expect(renameInput()).toBeNull();
    });

    it('trims spaces around the new name', async () => {
      const rename = listen('sheet-rename');
      key(await startRename(2, '  Costs  '), 'Enter');
      expect(details(rename)).toEqual([{ id: 2, name: 'Costs' }]);
    });

    it('keeps the field open and explains an invalid name', async () => {
      const rename = listen('sheet-rename');
      const input = await startRename(2, 'Q1/Q2');
      key(input, 'Enter');
      await settle();

      expect(rename).not.toHaveBeenCalled();
      expect(renameInput()).toBe(input);
      expect(root().activeElement).toBe(input);
      expect(input.getAttribute('aria-invalid')).toBe('true');
      const message = alert()!;
      expect(message.textContent!.trim()).toBe('A sheet name cannot contain : \\ / ? * [ or ]');
      expect(input.getAttribute('aria-describedby')).toBe(message.id);
    });

    it('rejects an empty name and the name of another sheet', async () => {
      const rename = listen('sheet-rename');
      const input = await startRename(2, '   ');
      key(input, 'Enter');
      await settle();
      expect(alert()!.textContent!.trim()).toBe('A sheet name cannot be empty');

      input.value = 'notes';
      key(input, 'Enter');
      await settle();
      expect(alert()!.textContent!.trim()).toBe('There is already a sheet named "notes"');
      expect(rename).not.toHaveBeenCalled();
    });

    it('allows changing only the case of a name', async () => {
      const rename = listen('sheet-rename');
      key(await startRename(2, 'BUDGET'), 'Enter');
      expect(details(rename)).toEqual([{ id: 2, name: 'BUDGET' }]);
    });

    it('clears the error on typing and accepts the corrected name', async () => {
      const rename = listen('sheet-rename');
      const input = await startRename(2, 'Notes');
      key(input, 'Enter');
      await settle();
      expect(alert()).not.toBeNull();

      input.value = 'Notes 2';
      input.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
      await settle();
      expect(alert()).toBeNull();
      expect(input.hasAttribute('aria-invalid')).toBe(false);
      expect(input.hasAttribute('aria-describedby')).toBe(false);

      key(input, 'Enter');
      expect(details(rename)).toEqual([{ id: 2, name: 'Notes 2' }]);
    });

    it('does not pass keys typed in the name on to the tabs or the grid', async () => {
      const outer = vi.fn();
      const select = listen('sheet-select');
      el.addEventListener('keydown', outer);
      const input = await startRename(2);
      for (const name of ['a', 'ArrowLeft', 'Home', 'Delete', 'Enter']) key(input, name);
      expect(outer).not.toHaveBeenCalled();
      expect(select).not.toHaveBeenCalled();
    });
  });

  describe('deleting', () => {
    async function pressDelete(id: number) {
      tab(id).focus();
      key(tab(id), 'Delete');
      await settle();
    }

    it('asks for confirmation on Delete, with focus on Cancel', async () => {
      const remove = listen('sheet-remove');
      await pressDelete(3);
      const dialog = confirmDialog()!;
      expect(dialog.getAttribute('aria-label')).toBe('Delete sheet');
      const text = root().getElementById(dialog.getAttribute('aria-describedby')!)!;
      expect(text.textContent).toBe('Delete "Notes"? This cannot be undone.');
      expect(root().activeElement).toBe(dialogButton('Cancel'));
      expect(remove).not.toHaveBeenCalled();
    });

    it('fires sheet-remove and closes when Delete is confirmed', async () => {
      const remove = listen('sheet-remove');
      await pressDelete(3);
      dialogButton('Delete').click();
      await settle();
      expect(details(remove)).toEqual([{ id: 3 }]);
      expect(confirmDialog()).toBeNull();
    });

    it('focuses the tab of the sheet the host shows after the delete', async () => {
      el.addEventListener('sheet-remove', () => {
        el.sheets = THREE.slice(0, 2);
        el.activeId = 1;
      });
      await pressDelete(3);
      dialogButton('Delete').click();
      await settle();
      expect(tabs()).toHaveLength(2);
      expect(root().activeElement).toBe(tab(1));
    });

    it('closes on Cancel and on Escape without firing, refocusing the tab', async () => {
      const remove = listen('sheet-remove');
      await pressDelete(3);
      dialogButton('Cancel').click();
      await settle();
      expect(confirmDialog()).toBeNull();
      expect(root().activeElement).toBe(tab(3));

      await pressDelete(1);
      expect(key(dialogButton('Cancel'), 'Escape').defaultPrevented).toBe(true);
      await settle();
      expect(confirmDialog()).toBeNull();
      expect(root().activeElement).toBe(tab(1));
      expect(remove).not.toHaveBeenCalled();
    });

    it('keeps Tab and Shift+Tab on its two buttons', async () => {
      const outer = vi.fn();
      el.addEventListener('keydown', outer);
      await pressDelete(3);
      expect(key(dialogButton('Cancel'), 'Tab').defaultPrevented).toBe(true);
      expect(root().activeElement).toBe(dialogButton('Delete'));
      key(dialogButton('Delete'), 'Tab');
      expect(root().activeElement).toBe(dialogButton('Cancel'));
      key(dialogButton('Cancel'), 'Tab', { shiftKey: true });
      expect(root().activeElement).toBe(dialogButton('Delete'));
      key(dialogButton('Delete'), 'Tab', { shiftKey: true });
      expect(root().activeElement).toBe(dialogButton('Cancel'));
      expect(outer).not.toHaveBeenCalled();
    });

    it('closes when something outside is pressed', async () => {
      await pressDelete(3);
      confirmDialog()!.dispatchEvent(new Event('pointerdown', { bubbles: true, composed: true }));
      await settle();
      expect(confirmDialog()).not.toBeNull();
      document.body.dispatchEvent(new Event('pointerdown', { bubbles: true, composed: true }));
      await settle();
      expect(confirmDialog()).toBeNull();
    });

    it('never offers to delete the only sheet', async () => {
      const remove = listen('sheet-remove');
      await setSheets([{ id: 1, name: 'Sheet1' }], 1);
      await pressDelete(1);
      expect(confirmDialog()).toBeNull();
      expect(status()).toBe('The only sheet cannot be deleted');

      await openMenu();
      expect(menuItem('delete').getAttribute('aria-disabled')).toBe('true');
      menuItem('delete').click();
      await settle();
      expect(confirmDialog()).toBeNull();
      expect(remove).not.toHaveBeenCalled();
    });
  });

  describe('options menu', () => {
    it('opens on click, focuses the first item and closes on a second click', async () => {
      await openMenu();
      expect(menuButton()!.getAttribute('aria-expanded')).toBe('true');
      const items = [...menu()!.querySelectorAll('[role="menuitem"]')];
      expect(items.map((i) => i.textContent)).toEqual(['Rename', 'Delete', 'Move left', 'Move right']);
      expect(root().activeElement).toBe(items[0]);

      await openMenu();
      expect(menu()).toBeNull();
      expect(menuButton()!.getAttribute('aria-expanded')).toBe('false');
    });

    it('opens from the keyboard with Enter, Space and ArrowDown', async () => {
      const outer = vi.fn();
      el.addEventListener('keydown', outer);
      for (const name of ['Enter', ' ', 'ArrowDown']) {
        expect(key(menuButton()!, name).defaultPrevented, name).toBe(true);
        await settle();
        expect(root().activeElement, name).toBe(menuItem('rename'));
        key(menuItem('rename'), 'Escape');
        await settle();
        expect(menu(), name).toBeNull();
      }
      expect(outer).not.toHaveBeenCalled();
    });

    it('moves through the items with the arrow keys, Home and End, wrapping', async () => {
      await openMenu();
      expect(key(menuItem('rename'), 'ArrowDown').defaultPrevented).toBe(true);
      expect(root().activeElement).toBe(menuItem('delete'));
      key(menuItem('delete'), 'End');
      expect(root().activeElement).toBe(menuItem('right'));
      key(menuItem('right'), 'ArrowDown');
      expect(root().activeElement).toBe(menuItem('rename'));
      key(menuItem('rename'), 'ArrowUp');
      expect(root().activeElement).toBe(menuItem('right'));
      key(menuItem('right'), 'Home');
      expect(root().activeElement).toBe(menuItem('rename'));
    });

    it('disables and skips what the active sheet cannot do', async () => {
      const disabled = () =>
        ['rename', 'delete', 'left', 'right'].filter((a) => menuItem(a).getAttribute('aria-disabled') === 'true');
      await openMenu();
      expect(disabled()).toEqual([]);

      el.activeId = 1;
      await settle();
      expect(disabled()).toEqual(['left']);
      menuItem('delete').focus();
      key(menuItem('delete'), 'ArrowDown');
      expect(root().activeElement).toBe(menuItem('right'));

      el.activeId = 3;
      await settle();
      expect(disabled()).toEqual(['right']);
      key(menuItem('rename'), 'End');
      expect(root().activeElement).toBe(menuItem('left'));
    });

    it('closes on Escape and returns focus to the button', async () => {
      await openMenu();
      expect(key(menuItem('rename'), 'Escape').defaultPrevented).toBe(true);
      await settle();
      expect(menu()).toBeNull();
      expect(root().activeElement).toBe(menuButton());
    });

    it('closes on Tab without holding focus back', async () => {
      await openMenu();
      expect(key(menuItem('rename'), 'Tab').defaultPrevented).toBe(false);
      await settle();
      expect(menu()).toBeNull();
    });

    it('closes when something outside is pressed and stops listening', async () => {
      const removeListener = vi.spyOn(document, 'removeEventListener');
      await openMenu();
      menu()!.dispatchEvent(new Event('pointerdown', { bubbles: true, composed: true }));
      await settle();
      expect(menu()).not.toBeNull();

      document.body.dispatchEvent(new Event('pointerdown', { bubbles: true, composed: true }));
      await settle();
      expect(menu()).toBeNull();
      expect(removeListener).toHaveBeenCalledWith('pointerdown', expect.any(Function));
      removeListener.mockRestore();
    });

    it('stops listening for outside presses when it is disconnected', async () => {
      await openMenu();
      const removeListener = vi.spyOn(document, 'removeEventListener');
      el.remove();
      expect(removeListener).toHaveBeenCalledWith('pointerdown', expect.any(Function));
      removeListener.mockRestore();
    });

    it('Rename closes the menu and renames the active sheet in place', async () => {
      await openMenu();
      menuItem('rename').click();
      await settle();
      expect(menu()).toBeNull();
      expect(renameInput()!.getAttribute('aria-label')).toBe('Rename sheet Budget');
      expect(root().activeElement).toBe(renameInput());
    });

    it('Delete closes the menu and asks about the active sheet', async () => {
      await openMenu();
      menuItem('delete').click();
      await settle();
      expect(menu()).toBeNull();
      expect(confirmDialog()!.textContent).toContain('Delete "Budget"?');
      expect(root().activeElement).toBe(dialogButton('Cancel'));
    });

    it('Move left and Move right fire sheet-move for the active sheet', async () => {
      const move = listen('sheet-move');
      await openMenu();
      menuItem('left').click();
      await settle();
      expect(menu()).toBeNull();
      expect(root().activeElement).toBe(menuButton());
      expect(status()).toBe('Moved Budget left');

      await openMenu();
      menuItem('right').click();
      await settle();
      expect(details(move)).toEqual([
        { id: 2, index: 0 },
        { id: 2, index: 2 },
      ]);
      expect(status()).toBe('Moved Budget right');
    });
  });

  describe('moving from the keyboard', () => {
    it('fires sheet-move for the focused sheet on Ctrl/Cmd+Shift+Arrow', () => {
      const move = listen('sheet-move');
      const select = listen('sheet-select');
      key(tab(2), 'ArrowLeft', { ctrlKey: true, shiftKey: true });
      key(tab(2), 'ArrowRight', { metaKey: true, shiftKey: true });
      // The focused tab, not necessarily the active one
      key(tab(1), 'ArrowRight', { ctrlKey: true, shiftKey: true });
      expect(details(move)).toEqual([
        { id: 2, index: 0 },
        { id: 2, index: 2 },
        { id: 1, index: 1 },
      ]);
      expect(select).not.toHaveBeenCalled();
    });

    it('does nothing at either end', () => {
      const move = listen('sheet-move');
      expect(key(tab(1), 'ArrowLeft', { ctrlKey: true, shiftKey: true }).defaultPrevented).toBe(true);
      key(tab(3), 'ArrowRight', { ctrlKey: true, shiftKey: true });
      expect(move).not.toHaveBeenCalled();
      expect(status()).toBe('');
    });

    it('keeps focus on the tab when the host reorders the sheets', async () => {
      el.addEventListener('sheet-move', (e) => {
        const { id, index } = (e as CustomEvent<{ id: number; index: number }>).detail;
        const next = el.sheets.filter((s) => s.id !== id);
        next.splice(index, 0, el.sheets.find((s) => s.id === id)!);
        el.sheets = next;
      });
      tab(2).focus();
      key(tab(2), 'ArrowLeft', { ctrlKey: true, shiftKey: true });
      await settle();
      expect(tabs().map((t) => t.dataset.sheetId)).toEqual(['2', '1', '3']);
      expect(root().activeElement).toBe(tab(2));
      expect(status()).toBe('Moved Budget left');
    });
  });

  describe('dragging to reorder', () => {
    /** jsdom has no layout: lay the tabs out side by side, each 100px wide, from x = 0. */
    const layOut = () =>
      tabs().forEach((t, i) => {
        const left = i * 100;
        t.getBoundingClientRect = () =>
          ({ left, right: left + 100, top: 0, bottom: 28, width: 100, height: 28, x: left, y: 0, toJSON: () => ({}) }) as DOMRect;
      });
    const pointer = (target: EventTarget, type: string, clientX: number, init: PointerEventInit = {}) => {
      const event = new PointerEvent(type, {
        bubbles: true,
        composed: true,
        cancelable: true,
        clientX,
        pointerId: 1,
        button: 0,
        pointerType: 'mouse',
        isPrimary: true,
        ...init,
      });
      target.dispatchEvent(event);
      return event;
    };
    /** Press a tab at one x and move to another, without letting go. */
    async function dragTo(id: number, from: number, to: number, init: PointerEventInit = {}) {
      pointer(tab(id), 'pointerdown', from, init);
      pointer(document, 'pointermove', to, init);
      await settle();
    }
    async function drop(x: number, init: PointerEventInit = {}) {
      pointer(document, 'pointerup', x, init);
      await settle();
    }
    const indicator = () => root().querySelector<HTMLElement>('.drop-indicator');
    const dragging = () => tabs().filter((t) => t.classList.contains('dragging')).map((t) => t.dataset.sheetId);
    const nextTask = () => new Promise((resolve) => setTimeout(resolve, 0));

    beforeEach(layOut);

    it('moves a tab right once the pointer passes the middle of the next tab', async () => {
      const move = listen('sheet-move');
      const select = listen('sheet-select');
      await dragTo(1, 50, 151);
      await drop(151);
      expect(details(move)).toEqual([{ id: 1, index: 1 }]);
      expect(status()).toBe('Moved Sheet1 to position 2 of 3');
      const event = move.mock.calls[0][0] as CustomEvent;
      expect(event.bubbles).toBe(true);
      expect(event.composed).toBe(true);
      expect(select).not.toHaveBeenCalled();
    });

    it('moves a tab left once the pointer passes the middle of the previous tab', async () => {
      const move = listen('sheet-move');
      await dragTo(3, 250, 149);
      await drop(149);
      expect(details(move)).toEqual([{ id: 3, index: 1 }]);
      expect(status()).toBe('Moved Notes to position 2 of 3');
    });

    it('moves a tab to the last and to the first position', async () => {
      const move = listen('sheet-move');
      await dragTo(1, 50, 400);
      await drop(400);
      await dragTo(3, 250, -20);
      await drop(-20);
      expect(details(move)).toEqual([
        { id: 1, index: 2 },
        { id: 3, index: 0 },
      ]);
      expect(status()).toBe('Moved Notes to position 1 of 3');
    });

    it('uses where the pointer is released, not only where it last moved', async () => {
      const move = listen('sheet-move');
      await dragTo(2, 150, 160);
      await drop(260);
      expect(details(move)).toEqual([{ id: 2, index: 2 }]);
    });

    it('marks the dragged tab and shows where it would land', async () => {
      await dragTo(1, 50, 160);
      expect(dragging()).toEqual(['1']);
      expect(root().querySelector('.bar')!.classList.contains('reordering')).toBe(true);
      // Between Budget and Notes
      expect(indicator()!.style.left).toBe('199px');
      expect(indicator()!.getAttribute('aria-hidden')).toBe('true');

      pointer(document, 'pointermove', 290);
      await settle();
      expect(indicator()!.style.left).toBe('299px');

      await drop(290);
      expect(dragging()).toEqual([]);
      expect(indicator()).toBeNull();
      expect(root().querySelector('.bar')!.classList.contains('reordering')).toBe(false);
    });

    it('fires nothing when the tab is dropped where it was', async () => {
      const move = listen('sheet-move');
      await dragTo(2, 150, 260);
      expect(dragging()).toEqual(['2']);
      expect(indicator()!.style.left).toBe('299px');
      // Back short of the middle of either neighbor: no line, nothing to do
      pointer(document, 'pointermove', 60);
      await settle();
      expect(indicator()).toBeNull();
      await drop(60);
      expect(move).not.toHaveBeenCalled();
      expect(status()).toBe('');
    });

    it('leaves a press that hardly moves a plain click', async () => {
      const move = listen('sheet-move');
      const select = listen('sheet-select');
      await dragTo(3, 250, 255);
      expect(dragging()).toEqual([]);
      await drop(255);
      tab(3).click();
      expect(move).not.toHaveBeenCalled();
      expect(details(select)).toEqual([{ id: 3 }]);
    });

    it('still renames on a double-click', async () => {
      for (let i = 0; i < 2; i++) {
        pointer(tab(2), 'pointerdown', 150);
        pointer(document, 'pointerup', 150);
        tab(2).click();
      }
      tab(2).dispatchEvent(new MouseEvent('dblclick', { bubbles: true, composed: true }));
      await settle();
      expect(renameInput()!.getAttribute('aria-label')).toBe('Rename sheet Budget');
    });

    it('swallows the click that follows a drop, and only that one', async () => {
      const select = listen('sheet-select');
      await dragTo(1, 50, 260);
      await drop(260);
      tab(1).click();
      expect(select).not.toHaveBeenCalled();

      await nextTask();
      tab(1).click();
      expect(details(select)).toEqual([{ id: 1 }]);
    });

    it('swallows the click after a drop in the original position too', async () => {
      const select = listen('sheet-select');
      await dragTo(1, 50, 120);
      await drop(120);
      tab(1).click();
      expect(select).not.toHaveBeenCalled();
    });

    it('cancels on Escape: no event, and the rest of the press does nothing', async () => {
      const move = listen('sheet-move');
      const select = listen('sheet-select');
      const outer = vi.fn();
      el.addEventListener('keydown', outer);
      await dragTo(1, 50, 260);
      expect(key(tab(1), 'Escape').defaultPrevented).toBe(true);
      await settle();
      expect(outer).not.toHaveBeenCalled();
      expect(dragging()).toEqual([]);
      expect(indicator()).toBeNull();

      pointer(document, 'pointermove', 280);
      await settle();
      expect(dragging()).toEqual([]);
      await drop(280);
      tab(1).click();
      expect(move).not.toHaveBeenCalled();
      expect(select).not.toHaveBeenCalled();
    });

    it('leaves Escape alone when nothing is being dragged', () => {
      pointer(tab(1), 'pointerdown', 50);
      expect(key(tab(1), 'Escape').defaultPrevented).toBe(false);
      pointer(document, 'pointerup', 50);
    });

    it('cancels on pointercancel and when the window loses focus', async () => {
      const move = listen('sheet-move');
      await dragTo(1, 50, 260);
      pointer(document, 'pointercancel', 260);
      await settle();
      expect(dragging()).toEqual([]);
      await drop(260);

      await dragTo(1, 50, 260);
      window.dispatchEvent(new Event('blur'));
      await settle();
      expect(dragging()).toEqual([]);
      expect(indicator()).toBeNull();
      await drop(260);
      expect(move).not.toHaveBeenCalled();
    });

    it('does not start in read-only mode, while renaming, or with another button', async () => {
      const move = listen('sheet-move');
      await dragTo(1, 50, 260, { button: 2 });
      expect(dragging()).toEqual([]);
      await drop(260, { button: 2 });

      await startRename(2, 'Notes');
      key(renameInput()!, 'Enter');
      await settle();
      expect(alert()).not.toBeNull();
      await dragTo(1, 50, 260);
      expect(dragging()).toEqual([]);
      await drop(260);
      key(renameInput()!, 'Escape');

      el.readOnly = true;
      await settle();
      layOut();
      await dragTo(1, 50, 260);
      expect(dragging()).toEqual([]);
      await drop(260);
      expect(move).not.toHaveBeenCalled();
    });

    it('ignores other pointers during a drag', async () => {
      const move = listen('sheet-move');
      await dragTo(1, 50, 160);
      pointer(tab(3), 'pointerdown', 250, { pointerId: 2 });
      pointer(document, 'pointermove', 20, { pointerId: 2 });
      pointer(document, 'pointerup', 20, { pointerId: 2 });
      await settle();
      expect(dragging()).toEqual(['1']);
      await drop(160);
      expect(details(move)).toEqual([{ id: 1, index: 1 }]);
    });

    it('stops following the pointer when it is disconnected or the sheet goes away', async () => {
      const move = listen('sheet-move');
      await dragTo(1, 50, 260);
      el.sheets = THREE.slice(1);
      await settle();
      expect(dragging()).toEqual([]);
      expect(indicator()).toBeNull();
      await drop(260);

      el.sheets = THREE;
      await settle();
      layOut();
      const removeListener = vi.spyOn(AbortController.prototype, 'abort');
      await dragTo(1, 50, 260);
      el.remove();
      expect(removeListener).toHaveBeenCalled();
      removeListener.mockRestore();
      pointer(document, 'pointerup', 260);
      expect(move).not.toHaveBeenCalled();
    });

    it('keeps the dropped tab in place and focused when the host reorders the sheets', async () => {
      el.addEventListener('sheet-move', (e) => {
        const { id, index } = (e as CustomEvent<{ id: number; index: number }>).detail;
        const next = el.sheets.filter((s) => s.id !== id);
        next.splice(index, 0, el.sheets.find((s) => s.id === id)!);
        el.sheets = next;
      });
      tab(1).focus();
      await dragTo(1, 50, 400);
      await drop(400);
      expect(tabs().map((t) => t.dataset.sheetId)).toEqual(['2', '3', '1']);
      expect(root().activeElement).toBe(tab(1));
      // Still showing the sheet it showed before
      expect(tab(2).getAttribute('aria-selected')).toBe('true');
    });

    it('keeps the dropped tab in view rather than jumping back to the active one', async () => {
      const list = root().querySelector<HTMLElement>('.tabs')!;
      const rect = (left: number, width: number) =>
        ({ left, right: left + width, top: 0, bottom: 28, width, height: 28, x: left, y: 0, toJSON: () => ({}) }) as DOMRect;
      // A list 200px wide that scrolls over three tabs of 100px, wherever each one is at the time
      list.getBoundingClientRect = () => rect(0, 200);
      tabs().forEach((t) => {
        t.getBoundingClientRect = () => rect(tabs().indexOf(t) * 100 - list.scrollLeft, 100);
      });
      el.addEventListener('sheet-move', () => {
        el.sheets = [THREE[1], THREE[2], THREE[0]];
      });
      await dragTo(1, 50, 400);
      await drop(400);
      expect(tabs().map((t) => t.dataset.sheetId)).toEqual(['2', '3', '1']);
      // Sheet1 is now the third tab; the active Budget tab would not have needed a scroll
      expect(list.scrollLeft).toBe(100);
    });

    describe('with a finger or pen', () => {
      beforeEach(() => {
        vi.useFakeTimers();
      });

      afterEach(() => {
        vi.useRealTimers();
      });

      it('leaves a swipe to scroll the list', async () => {
        const move = listen('sheet-move');
        await dragTo(1, 50, 260, { pointerType: 'touch' });
        expect(dragging()).toEqual([]);
        await drop(260, { pointerType: 'touch' });
        expect(move).not.toHaveBeenCalled();
      });

      it('picks the tab up after a hold, then drags it', async () => {
        const move = listen('sheet-move');
        pointer(tab(1), 'pointerdown', 50, { pointerType: 'touch' });
        vi.advanceTimersByTime(299);
        await settle();
        expect(dragging()).toEqual([]);
        vi.advanceTimersByTime(1);
        await settle();
        expect(dragging()).toEqual(['1']);

        pointer(document, 'pointermove', 260, { pointerType: 'touch' });
        await settle();
        expect(indicator()!.style.left).toBe('299px');
        await drop(260, { pointerType: 'touch' });
        expect(details(move)).toEqual([{ id: 1, index: 2 }]);
      });

      it('keeps the browser from panning or opening a menu once the tab is picked up', async () => {
        const touchMove = () => {
          const event = new Event('touchmove', { bubbles: true, composed: true, cancelable: true });
          tab(1).dispatchEvent(event);
          return event.defaultPrevented;
        };
        const contextMenu = () => {
          const event = new MouseEvent('contextmenu', { bubbles: true, composed: true, cancelable: true });
          tab(1).dispatchEvent(event);
          return event.defaultPrevented;
        };
        expect(touchMove()).toBe(false);
        expect(contextMenu()).toBe(false);

        pointer(tab(1), 'pointerdown', 50, { pointerType: 'pen' });
        expect(touchMove()).toBe(false);
        vi.advanceTimersByTime(300);
        await settle();
        expect(touchMove()).toBe(true);
        expect(contextMenu()).toBe(true);

        await drop(50, { pointerType: 'pen' });
        expect(dragging()).toEqual([]);
        expect(touchMove()).toBe(false);
      });
    });
  });

  describe('read-only', () => {
    beforeEach(async () => {
      el.readOnly = true;
      await settle();
    });

    it('reflects the attribute and hides the add and options buttons', () => {
      expect(el.hasAttribute('read-only')).toBe(true);
      expect(addButton()).toBeNull();
      expect(menuButton()).toBeNull();
      expect(tabs()).toHaveLength(3);
    });

    it('still switches sheets by click and keyboard', () => {
      const select = listen('sheet-select');
      tab(1).click();
      key(tab(2), 'ArrowRight');
      expect(details(select)).toEqual([{ id: 1 }, { id: 3 }]);
      expect(root().activeElement).toBe(tab(3));
    });

    it('ignores renaming, deleting and moving', async () => {
      const move = listen('sheet-move');
      expect(key(tab(2), 'F2').defaultPrevented).toBe(false);
      expect(key(tab(2), 'Delete').defaultPrevented).toBe(false);
      expect(key(tab(2), 'ArrowLeft', { ctrlKey: true, shiftKey: true }).defaultPrevented).toBe(false);
      tab(2).dispatchEvent(new MouseEvent('dblclick', { bubbles: true, composed: true }));
      el.startRename(2);
      await settle();
      expect(renameInput()).toBeNull();
      expect(confirmDialog()).toBeNull();
      expect(move).not.toHaveBeenCalled();
    });

    it('closes what was open when it is switched on', async () => {
      el.readOnly = false;
      await settle();
      await startRename(2, 'Costs');
      const rename = listen('sheet-rename');
      el.readOnly = true;
      await settle();
      expect(renameInput()).toBeNull();
      expect(tab(2)).not.toBeNull();
      expect(rename).not.toHaveBeenCalled();
    });
  });
});
