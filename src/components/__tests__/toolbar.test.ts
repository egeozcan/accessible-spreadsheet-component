// @vitest-environment jsdom
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { Y11nToolbar, TOOLBAR_GROUPS, TOOLBAR_ITEMS } from '../y11n-toolbar.js';
import type { FormatActionDetail, ToolbarActionDetail } from '../y11n-toolbar.js';

describe('Y11nToolbar', () => {
  let toolbar: Y11nToolbar;

  beforeEach(async () => {
    // Clear any previously registered definition
    const tag = 'y11n-toolbar';
    if (!customElements.get(tag)) {
      customElements.define(tag, Y11nToolbar);
    }
    toolbar = document.createElement(tag) as Y11nToolbar;
    document.body.appendChild(toolbar);
    await toolbar.updateComplete;
  });

  afterEach(() => {
    toolbar.remove();
  });

  describe('rendering', () => {
    it('renders a toolbar element with correct role', () => {
      const role = toolbar.shadowRoot!.querySelector('[role="toolbar"]');
      expect(role).not.toBeNull();
    });

    it('renders bold, italic, underline, strikethrough buttons', () => {
      const root = toolbar.shadowRoot!;
      expect(root.querySelector('[aria-label="Bold"]')).not.toBeNull();
      expect(root.querySelector('[aria-label="Italic"]')).not.toBeNull();
      expect(root.querySelector('[aria-label="Underline"]')).not.toBeNull();
      expect(root.querySelector('[aria-label="Strikethrough"]')).not.toBeNull();
    });

    it('renders color inputs', () => {
      const root = toolbar.shadowRoot!;
      expect(root.querySelector('[aria-label="Text color"]')).not.toBeNull();
      expect(root.querySelector('[aria-label="Background color"]')).not.toBeNull();
    });

    it('renders alignment buttons', () => {
      const root = toolbar.shadowRoot!;
      expect(root.querySelector('[aria-label="Align left"]')).not.toBeNull();
      expect(root.querySelector('[aria-label="Align center"]')).not.toBeNull();
      expect(root.querySelector('[aria-label="Align right"]')).not.toBeNull();
    });

    it('renders font size select', () => {
      const root = toolbar.shadowRoot!;
      expect(root.querySelector('[aria-label="Font size"]')).not.toBeNull();
    });

    it('renders clear formatting button', () => {
      const root = toolbar.shadowRoot!;
      expect(root.querySelector('[aria-label="Clear formatting"]')).not.toBeNull();
    });
  });

  describe('aria-pressed reflects properties', () => {
    it('bold button reflects bold property', async () => {
      toolbar.bold = true;
      await toolbar.updateComplete;
      const btn = toolbar.shadowRoot!.querySelector('[aria-label="Bold"]')!;
      expect(btn.getAttribute('aria-pressed')).toBe('true');
    });

    it('italic button reflects italic property', async () => {
      toolbar.italic = true;
      await toolbar.updateComplete;
      const btn = toolbar.shadowRoot!.querySelector('[aria-label="Italic"]')!;
      expect(btn.getAttribute('aria-pressed')).toBe('true');
    });

    it('underline button reflects underline property', async () => {
      toolbar.underline = true;
      await toolbar.updateComplete;
      const btn = toolbar.shadowRoot!.querySelector('[aria-label="Underline"]')!;
      expect(btn.getAttribute('aria-pressed')).toBe('true');
    });

    it('strikethrough button reflects strikethrough property', async () => {
      toolbar.strikethrough = true;
      await toolbar.updateComplete;
      const btn = toolbar.shadowRoot!.querySelector('[aria-label="Strikethrough"]')!;
      expect(btn.getAttribute('aria-pressed')).toBe('true');
    });

    it('alignment buttons reflect textAlign property', async () => {
      toolbar.textAlign = 'center';
      await toolbar.updateComplete;
      const root = toolbar.shadowRoot!;
      expect(root.querySelector('[aria-label="Align left"]')!.getAttribute('aria-pressed')).toBe('false');
      expect(root.querySelector('[aria-label="Align center"]')!.getAttribute('aria-pressed')).toBe('true');
      expect(root.querySelector('[aria-label="Align right"]')!.getAttribute('aria-pressed')).toBe('false');
    });
  });

  describe('format-action events', () => {
    it('dispatches bold toggle on click', async () => {
      const handler = vi.fn();
      toolbar.addEventListener('format-action', handler);

      const btn = toolbar.shadowRoot!.querySelector('[aria-label="Bold"]') as HTMLButtonElement;
      btn.click();

      expect(handler).toHaveBeenCalledTimes(1);
      const detail = (handler.mock.calls[0][0] as CustomEvent<FormatActionDetail>).detail;
      expect(detail.action).toBe('bold');
      expect(detail.value).toBe(true); // !this.bold where bold defaults false
    });

    it('dispatches italic toggle on click', async () => {
      toolbar.italic = true;
      await toolbar.updateComplete;

      const handler = vi.fn();
      toolbar.addEventListener('format-action', handler);

      const btn = toolbar.shadowRoot!.querySelector('[aria-label="Italic"]') as HTMLButtonElement;
      btn.click();

      const detail = (handler.mock.calls[0][0] as CustomEvent<FormatActionDetail>).detail;
      expect(detail.action).toBe('italic');
      expect(detail.value).toBe(false); // toggle off
    });

    it('dispatches clearFormat on clear button click', async () => {
      const handler = vi.fn();
      toolbar.addEventListener('format-action', handler);

      const btn = toolbar.shadowRoot!.querySelector('[aria-label="Clear formatting"]') as HTMLButtonElement;
      btn.click();

      const detail = (handler.mock.calls[0][0] as CustomEvent<FormatActionDetail>).detail;
      expect(detail.action).toBe('clearFormat');
    });

    it('dispatches textAlign action', async () => {
      const handler = vi.fn();
      toolbar.addEventListener('format-action', handler);

      const btn = toolbar.shadowRoot!.querySelector('[aria-label="Align center"]') as HTMLButtonElement;
      btn.click();

      const detail = (handler.mock.calls[0][0] as CustomEvent<FormatActionDetail>).detail;
      expect(detail.action).toBe('textAlign');
      expect(detail.value).toBe('center');
    });
  });

  describe('read-only mode', () => {
    it('disables the buttons that change cells when readOnly is true', async () => {
      toolbar.readOnly = true;
      await toolbar.updateComplete;

      const root = toolbar.shadowRoot!;
      // What a read-only grid still does stays available
      const viewOnly = ['Copy', 'Fit column width', 'Filter', 'Freeze panes', 'Find', 'Name manager'];
      const buttons = root.querySelectorAll<HTMLButtonElement>('button');
      expect(buttons.length).toBeGreaterThan(30);
      for (const btn of buttons) {
        const label = btn.getAttribute('aria-label')!;
        // No filter to reapply either way
        expect(btn.disabled, label).toBe(!viewOnly.includes(label));
      }
    });

    it('disables both selects when readOnly is true', async () => {
      toolbar.readOnly = true;
      await toolbar.updateComplete;

      const selects = toolbar.shadowRoot!.querySelectorAll('select');
      expect(selects.length).toBe(2);
      for (const select of selects) expect(select.disabled).toBe(true);
    });

    it('disables color inputs when readOnly is true', async () => {
      toolbar.readOnly = true;
      await toolbar.updateComplete;

      const inputs = toolbar.shadowRoot!.querySelectorAll('input[type="color"]');
      for (const input of inputs) {
        expect((input as HTMLInputElement).disabled).toBe(true);
      }
    });
  });

  describe('conditional formatting menu', () => {
    const root = () => toolbar.shadowRoot!;
    const button = () => root().querySelector('[aria-label="Conditional formatting"]') as HTMLButtonElement;
    const menu = () => root().querySelector('[role="menu"]') as HTMLElement | null;
    const items = () => [...root().querySelectorAll<HTMLElement>('[role="menuitem"]')];
    const key = (el: Element, name: string) => {
      const event = new KeyboardEvent('keydown', { key: name, bubbles: true, composed: true, cancelable: true });
      el.dispatchEvent(event);
      return event;
    };

    async function open() {
      button().click();
      await toolbar.updateComplete;
      await Promise.resolve();
    }

    it('is a menu button that is closed at first', () => {
      expect(button().getAttribute('aria-haspopup')).toBe('menu');
      expect(button().getAttribute('aria-expanded')).toBe('false');
      expect(menu()).toBeNull();
    });

    it('opens on click and focuses the first item', async () => {
      await open();
      expect(button().getAttribute('aria-expanded')).toBe('true');
      expect(items().map((i) => i.textContent)).toEqual([
        'Highlight duplicate values',
        'Highlight unique values',
        'Highlight above average',
        'Highlight below average',
        'Highlight top 10',
        'Highlight bottom 10',
        'Color scale',
        'Data bars',
        'Icon set',
        'Clear rules from selection',
        'Manage rules\u2026',
      ]);
      expect(root().activeElement).toBe(items()[0]);
    });

    it('moves through the items with the arrow keys, Home and End', async () => {
      await open();
      const all = items();
      expect(key(all[0], 'ArrowDown').defaultPrevented).toBe(true);
      expect(root().activeElement).toBe(all[1]);
      key(all[1], 'ArrowUp');
      key(all[0], 'ArrowUp');
      expect(root().activeElement).toBe(all[all.length - 1]);
      key(all[all.length - 1], 'Home');
      expect(root().activeElement).toBe(all[0]);
      key(all[0], 'End');
      expect(root().activeElement).toBe(all[all.length - 1]);
      // Left/Right are not passed on to the toolbar while the menu is open
      key(all[all.length - 1], 'ArrowRight');
      expect(root().activeElement).toBe(all[all.length - 1]);
    });

    it('opens from the keyboard with ArrowDown and ArrowUp', async () => {
      key(button(), 'ArrowUp');
      await toolbar.updateComplete;
      await Promise.resolve();
      const all = items();
      expect(root().activeElement).toBe(all[all.length - 1]);
    });

    it('closes on Escape and returns focus to the button', async () => {
      await open();
      key(items()[2], 'Escape');
      await toolbar.updateComplete;
      expect(menu()).toBeNull();
      expect(root().activeElement).toBe(button());
    });

    it('closes when something outside is pressed', async () => {
      await open();
      document.body.dispatchEvent(new Event('pointerdown', { bubbles: true, composed: true }));
      await toolbar.updateComplete;
      expect(menu()).toBeNull();
    });

    it('dispatches conditional-format-action and closes when an item is chosen', async () => {
      const handler = vi.fn();
      const formatHandler = vi.fn();
      toolbar.addEventListener('conditional-format-action', handler);
      toolbar.addEventListener('format-action', formatHandler);
      await open();
      (root().querySelector('[data-action="aboveAverage"]') as HTMLElement).click();
      await toolbar.updateComplete;

      expect(handler).toHaveBeenCalledOnce();
      expect(handler.mock.calls[0][0].detail).toEqual({ action: 'aboveAverage' });
      // Hosts that pass every format-action on to setRangeFormat are not affected
      expect(formatHandler).not.toHaveBeenCalled();
      expect(menu()).toBeNull();
      expect(root().activeElement).toBe(button());

      await open();
      (root().querySelector('[data-action="manage"]') as HTMLElement).click();
      expect(handler.mock.calls[1][0].detail).toEqual({ action: 'manage' });
    });

    it('keeps the menu items out of the toolbar arrow-key order', async () => {
      await open();
      key(button(), 'ArrowRight');
      expect((root().activeElement as HTMLElement).getAttribute('aria-label')).toBe('Clear formatting');
    });
  });
  describe('items', () => {
    const labels = () =>
      [...toolbar.shadowRoot!.querySelectorAll('[data-item]')].map((el) => el.getAttribute('aria-label'));
    const separators = () => toolbar.shadowRoot!.querySelectorAll('.toolbar > [role="separator"]').length;

    it('shows every item in group order by default', () => {
      expect(TOOLBAR_ITEMS.length).toBe(33);
      expect(new Set(TOOLBAR_ITEMS).size).toBe(33);
      expect(labels().slice(0, 3)).toEqual(['Undo', 'Redo', 'Cut']);
      expect(labels().pop()).toBe('Name manager');
      expect(separators()).toBe(Object.keys(TOOLBAR_GROUPS).length - 1);
    });

    it('takes items, groups and separators as an array', async () => {
      toolbar.items = ['bold', 'italic', '|', 'history', 'filter'];
      await toolbar.updateComplete;
      expect(labels()).toEqual(['Bold', 'Italic', 'Undo', 'Redo', 'Filter']);
      // bold italic | history (a group of its own) | filter
      expect(separators()).toBe(2);
    });

    it('takes them as text and as an attribute', async () => {
      toolbar.items = 'find, bold';
      await toolbar.updateComplete;
      expect(labels()).toEqual(['Find', 'Bold']);

      toolbar.setAttribute('items', 'textAlign | wrap');
      await toolbar.updateComplete;
      expect(labels()).toEqual(['Align left', 'Align center', 'Align right', 'Wrap text']);
    });

    it('skips repeated and unknown names, with a warning for the unknown', async () => {
      const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
      toolbar.items = 'bold font nope | | bold';
      await toolbar.updateComplete;
      expect(labels()).toEqual(['Bold', 'Italic', 'Underline', 'Strikethrough', 'Font size', 'Text color', 'Background color']);
      expect(warn).toHaveBeenCalledOnce();
      expect(warn.mock.calls[0][0]).toContain('"nope"');
      warn.mockRestore();
    });

    it('keeps exactly one control in the Tab order', async () => {
      const stops = () => [...toolbar.shadowRoot!.querySelectorAll<HTMLElement>('[data-item]')].filter((el) => el.tabIndex === 0);
      expect(stops().map((el) => el.getAttribute('aria-label'))).toEqual(['Undo']);

      toolbar.canUndo = false;
      toolbar.canRedo = false;
      await toolbar.updateComplete;
      expect(stops().map((el) => el.getAttribute('aria-label'))).toEqual(['Cut']);

      const bold = toolbar.shadowRoot!.querySelector<HTMLElement>('[aria-label="Bold"]')!;
      bold.focus();
      await toolbar.updateComplete;
      expect(stops()).toEqual([bold]);
    });

    it('has a label that can be changed', async () => {
      const bar = toolbar.shadowRoot!.querySelector('[role="toolbar"]')!;
      expect(bar.getAttribute('aria-label')).toBe('Spreadsheet toolbar');
      toolbar.label = 'View';
      await toolbar.updateComplete;
      expect(bar.getAttribute('aria-label')).toBe('View');
    });
  });

  describe('state', () => {
    const button = (label: string) => toolbar.shadowRoot!.querySelector(`[aria-label="${label}"]`) as HTMLButtonElement;

    it('undo, redo and reapply filter follow their properties', async () => {
      expect(button('Undo').disabled).toBe(false);
      expect(button('Reapply filter').disabled).toBe(true);
      toolbar.canUndo = false;
      toolbar.canRedo = false;
      toolbar.filtered = true;
      await toolbar.updateComplete;
      expect(button('Undo').disabled).toBe(true);
      expect(button('Redo').disabled).toBe(true);
      expect(button('Reapply filter').disabled).toBe(false);
      expect(button('Filter').getAttribute('aria-pressed')).toBe('true');
    });

    it('wrap, vertical alignment and number format follow their properties', async () => {
      toolbar.wrap = true;
      toolbar.verticalAlign = 'bottom';
      toolbar.numberFormat = { type: 'currency', decimals: 0 };
      await toolbar.updateComplete;
      expect(button('Wrap text').getAttribute('aria-pressed')).toBe('true');
      expect(button('Align bottom').getAttribute('aria-pressed')).toBe('true');
      expect(button('Align middle').getAttribute('aria-pressed')).toBe('false');
      const select = toolbar.shadowRoot!.querySelector('[aria-label="Number format"]') as HTMLSelectElement;
      expect(select.value).toBe('currency');
    });
  });

  describe('toolbar-action and the new format actions', () => {
    const button = (label: string) => toolbar.shadowRoot!.querySelector(`[aria-label="${label}"]`) as HTMLButtonElement;

    it('fires toolbar-action for the buttons that are not formats', () => {
      const handler = vi.fn();
      const formatHandler = vi.fn();
      toolbar.addEventListener('toolbar-action', handler);
      toolbar.addEventListener('format-action', formatHandler);
      for (const label of ['Undo', 'Copy', 'Sort descending', 'Filter', 'AutoSum', 'Find', 'Increase decimals']) {
        button(label).click();
      }
      expect(handler.mock.calls.map(([e]) => (e as CustomEvent<ToolbarActionDetail>).detail.action)).toEqual([
        'undo',
        'copy',
        'sortDescending',
        'filter',
        'autoSum',
        'find',
        'increaseDecimals',
      ]);
      expect(handler.mock.calls[0][0].cancelable).toBe(true);
      expect(formatHandler).not.toHaveBeenCalled();
    });

    it('fires toolbar-action from the Insert, Delete and Freeze menus', async () => {
      const handler = vi.fn();
      toolbar.addEventListener('toolbar-action', handler);
      toolbar.frozen = true;
      for (const [menu, action] of [
        ['Insert rows or columns', 'insertColumnLeft'],
        ['Delete rows or columns', 'deleteRow'],
        ['Freeze panes', 'unfreeze'],
      ]) {
        button(menu).click();
        await toolbar.updateComplete;
        await Promise.resolve();
        (toolbar.shadowRoot!.querySelector(`[data-action="${action}"]`) as HTMLElement).click();
        await toolbar.updateComplete;
      }
      expect(handler.mock.calls.map(([e]) => e.detail.action)).toEqual(['insertColumnLeft', 'deleteRow', 'unfreeze']);
    });

    it('fires format-action with values that are cell formats', () => {
      const handler = vi.fn();
      toolbar.addEventListener('format-action', handler);
      button('Wrap text').click();
      button('Align top').click();
      const select = toolbar.shadowRoot!.querySelector('[aria-label="Number format"]') as HTMLSelectElement;
      select.value = 'percent';
      select.dispatchEvent(new Event('change'));
      select.value = 'general';
      select.dispatchEvent(new Event('change'));
      expect(handler.mock.calls.map(([e]) => (e as CustomEvent<FormatActionDetail>).detail)).toEqual([
        { action: 'wrap', value: true },
        { action: 'verticalAlign', value: 'top' },
        { action: 'numberFormat', value: { type: 'percent', decimals: 0 } },
        { action: 'numberFormat', value: undefined },
      ]);
    });
  });

  describe('target', () => {
    /** The part of <y11n-spreadsheet> the toolbar uses */
    function fakeGrid() {
      const grid = document.createElement('div') as HTMLDivElement & Record<string, any>;
      Object.assign(grid, {
        readOnly: false,
        canUndo: false,
        canRedo: true,
        frozenRows: 0,
        frozenCols: 0,
        format: { bold: true, fontSize: 18, verticalAlign: 'top' } as Record<string, unknown>,
        getSelectionFormat: () => grid.format,
        getFilter: () => null,
        toggleFormat: vi.fn(),
        formatSelection: vi.fn(),
        clearSelectionFormat: vi.fn(),
        undo: vi.fn(),
        sortSelection: vi.fn(),
        insertRowsAtSelection: vi.fn(),
      });
      return grid;
    }
    const button = (label: string) => toolbar.shadowRoot!.querySelector(`[aria-label="${label}"]`) as HTMLButtonElement;
    const settle = async () => {
      await Promise.resolve();
      await toolbar.updateComplete;
    };

    it('shows the state of its target and follows its events', async () => {
      const grid = fakeGrid();
      toolbar.target = grid as never;
      await toolbar.updateComplete;
      await settle();
      expect(toolbar.bold).toBe(true);
      expect(toolbar.fontSize).toBe(18);
      expect(toolbar.verticalAlign).toBe('top');
      expect(toolbar.textColor).toBe('');
      expect(button('Undo').disabled).toBe(true);
      expect(button('Redo').disabled).toBe(false);

      grid.format = {};
      grid.canUndo = true;
      grid.readOnly = true;
      grid.dispatchEvent(new Event('selection-change'));
      await settle();
      expect(toolbar.bold).toBe(false);
      expect(toolbar.readOnly).toBe(true);
      expect(toolbar.canUndo).toBe(true);
    });

    it('finds its target by id with the for attribute', async () => {
      const grid = fakeGrid();
      grid.id = 'the-grid';
      document.body.append(grid);
      toolbar.setAttribute('for', 'the-grid');
      await toolbar.updateComplete;
      await settle();
      expect(toolbar.bold).toBe(true);
      grid.remove();
    });

    it('acts on its target, unless the event is cancelled', async () => {
      const grid = fakeGrid();
      grid.canUndo = true;
      toolbar.target = grid as never;
      await toolbar.updateComplete;
      await settle();

      button('Bold').click();
      button('Align center').click();
      button('Clear formatting').click();
      button('Undo').click();
      button('Sort ascending').click();
      expect(grid.toggleFormat).toHaveBeenCalledWith('bold');
      expect(grid.formatSelection).toHaveBeenCalledWith({ textAlign: 'center' });
      expect(grid.clearSelectionFormat).toHaveBeenCalledOnce();
      expect(grid.undo).toHaveBeenCalledOnce();
      expect(grid.sortSelection).toHaveBeenCalledWith('ascending');

      toolbar.addEventListener('toolbar-action', (e) => e.preventDefault());
      toolbar.addEventListener('format-action', (e) => e.preventDefault());
      await settle();
      button('Undo').click();
      button('Italic').click();
      expect(grid.undo).toHaveBeenCalledOnce();
      expect(grid.toggleFormat).toHaveBeenCalledOnce();
    });

    it('stops following a target it is taken off', async () => {
      const grid = fakeGrid();
      toolbar.target = grid as never;
      await toolbar.updateComplete;
      await settle();
      toolbar.target = null;
      await toolbar.updateComplete;
      grid.format = { italic: true };
      grid.dispatchEvent(new Event('selection-change'));
      await settle();
      expect(toolbar.italic).toBe(false);
    });
  });
});

