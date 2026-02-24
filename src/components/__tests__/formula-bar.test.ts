// @vitest-environment jsdom
import { describe, it, expect, beforeEach, vi } from 'vitest';
import type { Y11nFormulaBar } from '../y11n-formula-bar.js';

// Import the component to register the custom element
import '../y11n-formula-bar.js';

/**
 * Helper: create a formula bar element, append to DOM, and wait for it to render.
 */
async function createFormulaBar(
  attrs: Partial<{
    cellRef: string;
    rawValue: string;
    displayValue: string;
    mode: 'raw' | 'formatted';
    readOnly: boolean;
  }> = {}
): Promise<Y11nFormulaBar> {
  const el = document.createElement('y11n-formula-bar') as Y11nFormulaBar;
  if (attrs.cellRef !== undefined) el.cellRef = attrs.cellRef;
  if (attrs.rawValue !== undefined) el.rawValue = attrs.rawValue;
  if (attrs.displayValue !== undefined) el.displayValue = attrs.displayValue;
  if (attrs.mode !== undefined) el.mode = attrs.mode;
  if (attrs.readOnly !== undefined) el.readOnly = attrs.readOnly;
  document.body.appendChild(el);
  await el.updateComplete;
  return el;
}

function getInput(el: Y11nFormulaBar): HTMLInputElement {
  return el.shadowRoot!.querySelector('input')!;
}

function getCellRef(el: Y11nFormulaBar): HTMLElement {
  return el.shadowRoot!.querySelector('.cell-ref')!;
}

function getModeButtons(el: Y11nFormulaBar): HTMLButtonElement[] {
  return Array.from(el.shadowRoot!.querySelectorAll('.mode-toggle button')) as HTMLButtonElement[];
}

describe('Y11nFormulaBar', () => {
  beforeEach(() => {
    // Clean up any previously created elements
    document.body.innerHTML = '';
  });

  describe('default properties', () => {
    it('renders with default properties', async () => {
      const el = await createFormulaBar();
      expect(el.cellRef).toBe('A1');
      expect(el.rawValue).toBe('');
      expect(el.mode).toBe('raw');
      expect(el.readOnly).toBe(false);
    });
  });

  describe('cell reference display', () => {
    it('displays the cell reference', async () => {
      const el = await createFormulaBar({ cellRef: 'B5' });
      const ref = getCellRef(el);
      expect(ref.textContent).toBe('B5');
    });

    it('updates when cellRef changes', async () => {
      const el = await createFormulaBar({ cellRef: 'A1' });
      el.cellRef = 'C3';
      await el.updateComplete;
      const ref = getCellRef(el);
      expect(ref.textContent).toBe('C3');
    });
  });

  describe('raw mode', () => {
    it('shows raw value in raw mode', async () => {
      const el = await createFormulaBar({
        rawValue: '=A1+B1',
        displayValue: '42',
        mode: 'raw',
      });
      const input = getInput(el);
      expect(input.value).toBe('=A1+B1');
    });
  });

  describe('formatted mode', () => {
    it('shows display value in formatted mode', async () => {
      const el = await createFormulaBar({
        rawValue: '=A1+B1',
        displayValue: '42',
        mode: 'formatted',
      });
      const input = getInput(el);
      expect(input.value).toBe('42');
    });
  });

  describe('readOnly mode', () => {
    it('disables input when readOnly is true', async () => {
      const el = await createFormulaBar({ readOnly: true });
      const input = getInput(el);
      expect(input.disabled).toBe(true);
    });

    it('enables input when readOnly is false', async () => {
      const el = await createFormulaBar({ readOnly: false });
      const input = getInput(el);
      expect(input.disabled).toBe(false);
    });
  });

  describe('mode change event', () => {
    it('dispatches formula-bar-mode-change on mode toggle button click', async () => {
      const el = await createFormulaBar({ mode: 'raw' });
      const handler = vi.fn();
      el.addEventListener('formula-bar-mode-change', handler);

      const buttons = getModeButtons(el);
      // The second button is "Formatted"
      const formattedBtn = buttons[1];
      formattedBtn.click();

      expect(handler).toHaveBeenCalledTimes(1);
      const event = handler.mock.calls[0][0] as CustomEvent;
      expect(event.detail.mode).toBe('formatted');
    });

    it('does not dispatch mode-change when clicking already-active mode', async () => {
      const el = await createFormulaBar({ mode: 'raw' });
      const handler = vi.fn();
      el.addEventListener('formula-bar-mode-change', handler);

      const buttons = getModeButtons(el);
      // The first button is "Raw" which is already active
      buttons[0].click();

      expect(handler).not.toHaveBeenCalled();
    });
  });

  describe('commit on Enter', () => {
    it('dispatches formula-bar-commit on Enter keydown', async () => {
      const el = await createFormulaBar({ rawValue: 'test value' });
      const handler = vi.fn();
      el.addEventListener('formula-bar-commit', handler);

      const input = getInput(el);
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

      expect(handler).toHaveBeenCalledTimes(1);
      const event = handler.mock.calls[0][0] as CustomEvent;
      expect(event.detail.value).toBe('test value');
    });
  });

  describe('revert on Escape', () => {
    it('reverts draft to source value on Escape', async () => {
      const el = await createFormulaBar({ rawValue: 'original', mode: 'raw' });
      const input = getInput(el);

      // Simulate typing a new value
      input.value = 'modified';
      input.dispatchEvent(new Event('input', { bubbles: true }));
      await el.updateComplete;

      // Press Escape to revert
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      await el.updateComplete;

      // The draft should revert to the raw value
      expect(input.value).toBe('original');
    });
  });
});
