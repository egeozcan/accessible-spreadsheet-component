// @vitest-environment jsdom
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { Y11nFormatToolbar } from '../y11n-format-toolbar.js';
import type { FormatActionDetail } from '../y11n-format-toolbar.js';

describe('Y11nFormatToolbar', () => {
  let toolbar: Y11nFormatToolbar;

  beforeEach(async () => {
    // Clear any previously registered definition
    const tag = 'y11n-format-toolbar';
    if (!customElements.get(tag)) {
      customElements.define(tag, Y11nFormatToolbar);
    }
    toolbar = document.createElement(tag) as Y11nFormatToolbar;
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
    it('disables all buttons when readOnly is true', async () => {
      toolbar.readOnly = true;
      await toolbar.updateComplete;

      const root = toolbar.shadowRoot!;
      const buttons = root.querySelectorAll('button');
      for (const btn of buttons) {
        expect(btn.disabled).toBe(true);
      }
    });

    it('disables font size select when readOnly is true', async () => {
      toolbar.readOnly = true;
      await toolbar.updateComplete;

      const select = toolbar.shadowRoot!.querySelector('select') as HTMLSelectElement;
      expect(select.disabled).toBe(true);
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
});

