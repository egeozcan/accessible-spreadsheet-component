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
});

