// @vitest-environment jsdom
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import '../y11n-name-manager.js';
import type { Y11nNameManager } from '../y11n-name-manager.js';
import type { DefinedName } from '../../types.js';

const NAMES: DefinedName[] = [
  { name: 'taxRate', refersTo: '=0.19', sheet: 'Sheet1' },
  { name: 'Sales', refersTo: '=Budget!$A$1:$A$10', sheet: 'Budget' },
  { name: 'double', refersTo: '=LAMBDA(x, x * 2)', sheet: 'Sheet1' },
];

const INVALID_NAME =
  'A name starts with a letter or underscore, has only letters, digits, underscores and periods, and is not a cell reference';
const INVALID_REFERS_TO = 'Refers to is not a valid formula';

describe('Y11nNameManager', () => {
  let el: Y11nNameManager;

  const root = () => el.shadowRoot!;
  const q = <T extends HTMLElement = HTMLElement>(selector: string) => root().querySelector<T>(selector);
  const dialog = () => q('[role="dialog"]')!;
  const rows = () => [...root().querySelectorAll<HTMLTableRowElement>('tbody tr')];
  const rowButton = (name: string, action: 'go' | 'edit' | 'delete') =>
    q<HTMLButtonElement>(`button[data-name="${name}"][data-action="${action}"]`);
  const newButton = () => q<HTMLButtonElement>('#name-new');
  const closeButton = () => q<HTMLButtonElement>('#name-close');
  const nameField = () => q<HTMLInputElement>('#name-name');
  const refersField = () => q<HTMLInputElement>('#name-refers');
  const sheetField = () => q<HTMLSelectElement>('#name-sheet');
  const confirmDialog = () => q('[role="alertdialog"]');
  const alerts = () => [...root().querySelectorAll<HTMLElement>('[role="alert"]')].map((a) => a.textContent!.trim());
  const focused = () => root().activeElement;

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

  async function type(field: HTMLInputElement, text: string) {
    field.value = text;
    field.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
    await settle();
  }

  async function choose(field: HTMLSelectElement, value: string) {
    field.value = value;
    field.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
    await settle();
  }

  async function click(button: HTMLElement | null) {
    expect(button).not.toBeNull();
    button!.click();
    await settle();
  }

  async function openNew(name?: string, refersTo?: string) {
    await click(newButton());
    if (name !== undefined) await type(nameField()!, name);
    if (refersTo !== undefined) await type(refersField()!, refersTo);
  }

  async function save() {
    await click(q('#name-save'));
  }

  beforeEach(async () => {
    el = document.createElement('y11n-name-manager');
    el.names = NAMES.map((n) => ({ ...n }));
    el.sheets = ['Sheet1', 'Budget', 'Notes'];
    el.activeSheet = 'Budget';
    el.selectionRef = '=Budget!$B$2:$C$5';
    document.body.appendChild(el);
    await settle();
  });

  afterEach(() => {
    el.remove();
  });

  describe('list', () => {
    it('is a modal dialog called "Name manager"', () => {
      expect(dialog().getAttribute('aria-label')).toBe('Name manager');
      expect(dialog().getAttribute('aria-modal')).toBe('true');
    });

    it('lists the names in a labelled table with column headers', () => {
      const table = q('table')!;
      expect(table.getAttribute('aria-label')).toBe('Defined names');
      const headers = [...table.querySelectorAll('thead th')];
      expect(headers.map((h) => h.textContent)).toEqual(['Name', 'Refers to', 'Sheet', 'Actions']);
      expect(headers.every((h) => h.getAttribute('scope') === 'col')).toBe(true);
    });

    it('sorts the names ignoring case', () => {
      expect(rows().map((r) => r.querySelector('th')!.textContent)).toEqual(['double', 'Sales', 'taxRate']);
      expect(rows().map((r) => r.dataset.name)).toEqual(['double', 'Sales', 'taxRate']);
    });

    it('shows what each name refers to and its sheet, with the full text as a title', () => {
      const cells = [...rows()[1].children].map((c) => c.textContent!.trim());
      expect(cells.slice(0, 3)).toEqual(['Sales', '=Budget!$A$1:$A$10', 'Budget']);
      expect(rows()[1].querySelector('th')!.getAttribute('scope')).toBe('row');
      expect(rows()[1].querySelector('td .clip')!.getAttribute('title')).toBe('=Budget!$A$1:$A$10');
    });

    it('follows the names it is given', async () => {
      el.names = [...el.names, { name: 'alpha', refersTo: '=1', sheet: 'Sheet1' }];
      await settle();
      expect(rows().map((r) => r.dataset.name)).toEqual(['alpha', 'double', 'Sales', 'taxRate']);
    });

    it('says so when there are no names', async () => {
      el.names = [];
      await settle();
      expect(q('table')).toBeNull();
      expect(q('.empty')!.textContent).toContain('No names yet');
      expect(newButton()).not.toBeNull();
      expect(closeButton()).not.toBeNull();
    });

    it('names each row button after its name', () => {
      const labels = [...rows()[1].querySelectorAll('button')].map((b) => [b.textContent, b.getAttribute('aria-label')]);
      expect(labels).toEqual([
        ['Go to', 'Go to Sales'],
        ['Edit', 'Edit Sales'],
        ['Delete', 'Delete Sales'],
      ]);
    });

    it('has New name and Close buttons', () => {
      expect(newButton()!.textContent).toBe('New name');
      expect(closeButton()!.textContent).toBe('Close');
    });

    it('fires name-go for a name', async () => {
      const go = listen('name-go');
      await click(rowButton('Sales', 'go'));
      expect(details(go)).toEqual([{ name: 'Sales' }]);
      // Nothing else happens until the spreadsheet reacts
      expect(q('table')).not.toBeNull();
    });

    it('fires name-close from the Close button', async () => {
      const close = listen('name-close');
      await click(closeButton());
      expect(close).toHaveBeenCalledOnce();
    });

    it('fires events that cross shadow roots', async () => {
      const seen: string[] = [];
      const record = (e: Event) => seen.push(`${e.type}:${e.composed}`);
      for (const type of ['name-go', 'name-close']) document.body.addEventListener(type, record);
      await click(rowButton('Sales', 'go'));
      await click(closeButton());
      for (const type of ['name-go', 'name-close']) document.body.removeEventListener(type, record);
      expect(seen).toEqual(['name-go:true', 'name-close:true']);
    });
  });

  describe('status', () => {
    it('shows the status in a live region', async () => {
      const status = q('[role="status"]')!;
      expect(status.textContent).toBe('');
      el.status = 'Added the name total';
      await settle();
      // The same element: a region that is replaced is not announced
      expect(q('[role="status"]')).toBe(status);
      expect(status.textContent).toBe('Added the name total');
    });

    it('keeps the live region while the form is open', async () => {
      el.status = 'Deleted the name total';
      await settle();
      const status = q('[role="status"]')!;
      await openNew();
      expect(q('[role="status"]')).toBe(status);
      expect(status.textContent).toBe('Deleted the name total');
    });
  });

  describe('focusInitial', () => {
    it('focuses the New name button', () => {
      el.focusInitial();
      expect(focused()).toBe(newButton());
    });

    it('focuses the first name in read-only mode', async () => {
      el.readOnly = true;
      await settle();
      el.focusInitial();
      expect(focused()).toBe(rowButton('double', 'go'));
    });

    it('focuses the Close button in read-only mode without names', async () => {
      el.readOnly = true;
      el.names = [];
      await settle();
      el.focusInitial();
      expect(focused()).toBe(closeButton());
    });
  });

  describe('new name', () => {
    it('opens a form with labelled fields and focus on Name', async () => {
      await openNew();
      expect(q('table')).toBeNull();
      expect(q('#name-title')!.textContent).toBe('New name');
      expect(focused()).toBe(nameField());
      expect(q('label[for="name-name"]')!.textContent).toBe('Name');
      expect(q('label[for="name-refers"]')!.textContent).toBe('Refers to');
      expect(q('label[for="name-sheet"]')!.textContent).toBe('Sheet');
      expect(nameField()!.getAttribute('aria-invalid')).toBe('false');
      expect(alerts()).toEqual([]);
    });

    it('starts from the selection and the sheet being shown', async () => {
      await openNew();
      expect(nameField()!.value).toBe('');
      expect(refersField()!.value).toBe('=Budget!$B$2:$C$5');
      expect([...sheetField()!.options].map((o) => o.value)).toEqual(['Sheet1', 'Budget', 'Notes']);
      expect(sheetField()!.value).toBe('Budget');
    });

    it('explains what the sheet is for', async () => {
      await openNew();
      const hint = q(`#${sheetField()!.getAttribute('aria-describedby')}`)!;
      expect(hint.textContent).toBe('Sheet that references without a sheet name point to');
    });

    it('fires name-save with the trimmed name, then shows the list with focus on New name', async () => {
      const saved = listen('name-save');
      await openNew('  total ');
      await save();
      expect(details(saved)).toEqual([{ name: 'total', refersTo: '=Budget!$B$2:$C$5', sheet: 'Budget' }]);
      expect('previous' in details(saved)[0]).toBe(false);
      expect(nameField()).toBeNull();
      expect(q('table')).not.toBeNull();
      expect(focused()).toBe(newButton());
    });

    it('saves the sheet that was chosen', async () => {
      const saved = listen('name-save');
      await openNew('total', '=A1:A5');
      await choose(sheetField()!, 'Notes');
      await save();
      expect(details(saved)).toEqual([{ name: 'total', refersTo: '=A1:A5', sheet: 'Notes' }]);
    });

    it('adds the = that "Refers to" is missing', async () => {
      const saved = listen('name-save');
      await openNew('total', ' SUM(Sheet1!A1:A5) ');
      await save();
      expect(details(saved)[0].refersTo).toBe('=SUM(Sheet1!A1:A5)');
    });

    it('accepts ranges, constants and functions', async () => {
      const saved = listen('name-save');
      for (const [name, refersTo] of [
        ['range', "='My Sheet'!$A$1:$B$3"],
        ['rate', '=0.19'],
        ['twice', '=LAMBDA(x, x * 2)'],
      ]) {
        await openNew(name, refersTo);
        await save();
      }
      expect(details(saved).map((d) => d.refersTo)).toEqual(["='My Sheet'!$A$1:$B$3", '=0.19', '=LAMBDA(x, x * 2)']);
    });

    it('saves on Enter in a text field', async () => {
      const saved = listen('name-save');
      await openNew('total');
      expect(key(nameField()!, 'Enter').defaultPrevented).toBe(true);
      await settle();
      expect(details(saved)).toEqual([{ name: 'total', refersTo: '=Budget!$B$2:$C$5', sheet: 'Budget' }]);

      await openNew('other', '=1');
      key(refersField()!, 'Enter');
      await settle();
      expect(details(saved)[1]).toEqual({ name: 'other', refersTo: '=1', sheet: 'Budget' });
      expect(focused()).toBe(newButton());
    });

    it('does not save on the Enter that ends an IME composition', async () => {
      const saved = listen('name-save');
      await openNew('total');
      key(nameField()!, 'Enter', { isComposing: true });
      await settle();
      expect(saved).not.toHaveBeenCalled();
      expect(nameField()).not.toBeNull();
    });

    it('cancels without saving and returns focus to New name', async () => {
      const saved = listen('name-save');
      await openNew('total');
      await click(q('#name-cancel'));
      expect(saved).not.toHaveBeenCalled();
      expect(nameField()).toBeNull();
      expect(focused()).toBe(newButton());
    });

    it('starts each new name afresh', async () => {
      await openNew('total', '=1');
      await click(q('#name-cancel'));
      el.selectionRef = '=Notes!$D$4';
      el.activeSheet = 'Notes';
      await openNew();
      expect(nameField()!.value).toBe('');
      expect(refersField()!.value).toBe('=Notes!$D$4');
      expect(sheetField()!.value).toBe('Notes');
    });
  });

  describe('edit', () => {
    it('opens the form with the name, what it refers to and its sheet', async () => {
      await click(rowButton('taxRate', 'edit'));
      expect(q('#name-title')!.textContent).toBe('Edit name');
      expect(focused()).toBe(nameField());
      expect(nameField()!.value).toBe('taxRate');
      expect(refersField()!.value).toBe('=0.19');
      // Its own sheet, not the one being shown
      expect(sheetField()!.value).toBe('Sheet1');
    });

    it('fires name-save with the name it had before', async () => {
      const saved = listen('name-save');
      await click(rowButton('taxRate', 'edit'));
      await type(nameField()!, 'vat');
      await type(refersField()!, '0.21');
      await choose(sheetField()!, 'Notes');
      await save();
      expect(details(saved)).toEqual([{ name: 'vat', refersTo: '=0.21', sheet: 'Notes', previous: 'taxRate' }]);
      expect(focused()).toBe(newButton());
    });

    it('accepts its own name, also in another case', async () => {
      const saved = listen('name-save');
      await click(rowButton('taxRate', 'edit'));
      await type(refersField()!, '=0.2');
      await save();
      await click(rowButton('taxRate', 'edit'));
      await type(nameField()!, 'TAXRATE');
      await save();
      expect(details(saved)).toEqual([
        { name: 'taxRate', refersTo: '=0.2', sheet: 'Sheet1', previous: 'taxRate' },
        { name: 'TAXRATE', refersTo: '=0.19', sheet: 'Sheet1', previous: 'taxRate' },
      ]);
    });

    it('returns focus to the Edit button on Cancel', async () => {
      await click(rowButton('Sales', 'edit'));
      await click(q('#name-cancel'));
      expect(focused()).toBe(rowButton('Sales', 'edit'));
    });
  });

  describe('validation', () => {
    async function expectNameError(name: string, message: string) {
      const saved = listen('name-save');
      await openNew(name, '=A1');
      await save();
      expect(saved).not.toHaveBeenCalled();
      expect(alerts()).toEqual([message]);
      const field = nameField()!;
      expect(field.getAttribute('aria-invalid')).toBe('true');
      expect(q(`#${field.getAttribute('aria-describedby')}`)!.textContent).toBe(message);
      expect(refersField()!.getAttribute('aria-invalid')).toBe('false');
      expect(focused()).toBe(field);
    }

    async function expectRefersError(refersTo: string) {
      const saved = listen('name-save');
      await openNew('total', refersTo);
      await save();
      expect(saved).not.toHaveBeenCalled();
      expect(alerts()).toEqual([INVALID_REFERS_TO]);
      const field = refersField()!;
      expect(field.getAttribute('aria-invalid')).toBe('true');
      const described = field.getAttribute('aria-describedby')!.split(' ');
      expect(q(`#${described[0]}`)!.textContent).toBe(INVALID_REFERS_TO);
      expect(nameField()!.getAttribute('aria-invalid')).toBe('false');
      expect(focused()).toBe(field);
    }

    it('rejects a name that is a cell reference', () => expectNameError('A1', INVALID_NAME));

    it('rejects a name that starts with a digit', () => expectNameError('1abc', INVALID_NAME));

    it('rejects a name with a space', () => expectNameError('net total', INVALID_NAME));

    it('rejects an empty name', () => expectNameError('   ', INVALID_NAME));

    it('rejects a name that exists, ignoring case', () => expectNameError('SALES', 'There is already a name Sales'));

    it('rejects renaming to another existing name', async () => {
      const saved = listen('name-save');
      await click(rowButton('taxRate', 'edit'));
      await type(nameField()!, 'sales');
      await save();
      expect(saved).not.toHaveBeenCalled();
      expect(alerts()).toEqual(['There is already a name Sales']);
      expect(nameField()!.getAttribute('aria-invalid')).toBe('true');
    });

    it('rejects an empty "Refers to"', () => expectRefersError(''));

    it('rejects a "Refers to" that is only =', () => expectRefersError(' = '));

    it('rejects a "Refers to" that does not parse', () => expectRefersError('=SUM(A1:'));

    it('rejects a "Refers to" that does not parse after adding the =', () => expectRefersError('1 +'));

    it('shows both messages and focuses the first invalid field', async () => {
      const saved = listen('name-save');
      await openNew('B2', '=(');
      await save();
      expect(saved).not.toHaveBeenCalled();
      expect(alerts()).toEqual([INVALID_NAME, INVALID_REFERS_TO]);
      expect(focused()).toBe(nameField());
    });

    it('keeps what was typed and saves once it is fixed', async () => {
      const saved = listen('name-save');
      await openNew('2nd', '=A1:A3');
      await save();
      expect(nameField()!.value).toBe('2nd');
      expect(refersField()!.value).toBe('=A1:A3');
      await type(nameField()!, 'second');
      await save();
      expect(details(saved)).toEqual([{ name: 'second', refersTo: '=A1:A3', sheet: 'Budget' }]);
    });

    it('announces the same mistake again on the next try', async () => {
      await openNew('A1', '=1');
      await save();
      const first = q('[role="alert"]');
      await save();
      // A new element: a message that just stays is not read out again
      expect(q('[role="alert"]')).not.toBe(first);
      expect(alerts()).toEqual([INVALID_NAME]);
    });

    it('forgets the messages when the form is opened again', async () => {
      await openNew('A1', '');
      await save();
      await click(q('#name-cancel'));
      await openNew();
      expect(alerts()).toEqual([]);
      expect(nameField()!.getAttribute('aria-invalid')).toBe('false');
      expect(nameField()!.hasAttribute('aria-describedby')).toBe(false);
    });
  });

  describe('delete', () => {
    it('asks first, with focus on Cancel', async () => {
      const removed = listen('name-remove');
      await click(rowButton('Sales', 'delete'));
      expect(removed).not.toHaveBeenCalled();
      const confirm = confirmDialog()!;
      expect(confirm.getAttribute('aria-label')).toBe('Delete name');
      expect(q(`#${confirm.getAttribute('aria-describedby')}`)!.textContent).toContain('Delete the name "Sales"?');
      expect(focused()).toBe(q('#name-delete-cancel'));
      expect(q('tr[data-name="Sales"]')!.classList.contains('confirming')).toBe(true);
    });

    it('does nothing on Cancel and returns focus to the Delete button', async () => {
      const removed = listen('name-remove');
      await click(rowButton('Sales', 'delete'));
      await click(q('#name-delete-cancel'));
      expect(removed).not.toHaveBeenCalled();
      expect(confirmDialog()).toBeNull();
      expect(rows()).toHaveLength(3);
      expect(focused()).toBe(rowButton('Sales', 'delete'));
    });

    it('cancels on Escape without closing the dialog', async () => {
      const removed = listen('name-remove');
      const close = listen('name-close');
      await click(rowButton('Sales', 'delete'));
      expect(key(q('#name-delete-cancel')!, 'Escape').defaultPrevented).toBe(true);
      await settle();
      expect(confirmDialog()).toBeNull();
      expect(removed).not.toHaveBeenCalled();
      expect(close).not.toHaveBeenCalled();
      expect(focused()).toBe(rowButton('Sales', 'delete'));
    });

    it('keeps Tab on the two buttons of the confirmation', async () => {
      await click(rowButton('Sales', 'delete'));
      const cancel = q('#name-delete-cancel')!;
      const confirm = q('#name-delete-confirm')!;
      expect(key(cancel, 'Tab').defaultPrevented).toBe(true);
      expect(focused()).toBe(confirm);
      key(confirm, 'Tab');
      expect(focused()).toBe(cancel);
      key(cancel, 'Tab', { shiftKey: true });
      expect(focused()).toBe(confirm);
    });

    it('fires name-remove once confirmed, and focuses New name when the name is gone', async () => {
      const removed = listen('name-remove');
      // The spreadsheet removes the name and passes the rest back in
      el.addEventListener('name-remove', (e) => {
        const { name } = (e as CustomEvent<{ name: string }>).detail;
        el.names = el.names.filter((n) => n.name !== name);
        el.status = `Deleted the name ${name}`;
      });
      await click(rowButton('Sales', 'delete'));
      await click(q('#name-delete-confirm'));
      expect(details(removed)).toEqual([{ name: 'Sales' }]);
      expect(confirmDialog()).toBeNull();
      expect(rows().map((r) => r.dataset.name)).toEqual(['double', 'taxRate']);
      expect(focused()).toBe(newButton());
      expect(q('[role="status"]')!.textContent).toBe('Deleted the name Sales');
    });

    it('closes the confirmation when its name goes away', async () => {
      await click(rowButton('Sales', 'delete'));
      el.names = el.names.filter((n) => n.name !== 'Sales');
      await settle();
      expect(confirmDialog()).toBeNull();
      expect(focused()).toBe(newButton());
    });

    it('closes the confirmation when another name is edited instead', async () => {
      const removed = listen('name-remove');
      await click(rowButton('Sales', 'delete'));
      await click(rowButton('taxRate', 'edit'));
      expect(confirmDialog()).toBeNull();
      expect(nameField()!.value).toBe('taxRate');
      await click(q('#name-cancel'));
      expect(confirmDialog()).toBeNull();
      expect(removed).not.toHaveBeenCalled();
    });
  });

  describe('keyboard', () => {
    it('fires name-close on Escape in the list', () => {
      const close = listen('name-close');
      el.focusInitial();
      expect(key(newButton()!, 'Escape').defaultPrevented).toBe(true);
      expect(close).toHaveBeenCalledOnce();
    });

    it('closes the form on the first Escape and the dialog on the second', async () => {
      const close = listen('name-close');
      const saved = listen('name-save');
      await openNew('total');
      key(nameField()!, 'Escape');
      await settle();
      expect(close).not.toHaveBeenCalled();
      expect(saved).not.toHaveBeenCalled();
      expect(nameField()).toBeNull();
      expect(focused()).toBe(newButton());

      key(newButton()!, 'Escape');
      expect(close).toHaveBeenCalledOnce();
    });

    it('returns focus to the Edit button when Escape closes an edit', async () => {
      await click(rowButton('double', 'edit'));
      key(refersField()!, 'Escape');
      await settle();
      expect(focused()).toBe(rowButton('double', 'edit'));
    });

    it('keeps its key presses away from the grid', async () => {
      const outside = vi.fn();
      const parent = document.createElement('div');
      document.body.appendChild(parent);
      parent.appendChild(el);
      await settle();
      parent.addEventListener('keydown', outside);
      el.addEventListener('keydown', outside);

      for (const name of ['a', 'ArrowDown', 'Delete', 'Enter', 'Tab', 'Escape']) key(newButton()!, name);
      key(newButton()!, 'z', { ctrlKey: true });
      await openNew();
      for (const name of ['a', 'Backspace', 'Home', 'F2']) key(nameField()!, name);
      key(nameField()!, 'Escape');

      expect(outside).not.toHaveBeenCalled();
      document.body.appendChild(el);
      parent.remove();
    });

    it('keeps Tab inside the dialog, wrapping at its ends', async () => {
      // Between controls the browser moves focus; from the last one it wraps to the first
      newButton()!.focus();
      expect(key(newButton()!, 'Tab').defaultPrevented).toBe(false);
      closeButton()!.focus();
      expect(key(closeButton()!, 'Tab').defaultPrevented).toBe(true);
      const first = el.shadowRoot!.querySelector<HTMLElement>('button:not(:disabled), input, select')!;
      expect(el.shadowRoot!.activeElement).toBe(first);
      expect(key(first, 'Tab', { shiftKey: true }).defaultPrevented).toBe(true);
      expect(el.shadowRoot!.activeElement).toBe(closeButton());

      await openNew();
      q<HTMLElement>('#name-cancel')!.focus();
      expect(key(q('#name-cancel')!, 'Tab').defaultPrevented).toBe(true);
      expect(el.shadowRoot!.activeElement).toBe(nameField());
    });
  });

  describe('read-only', () => {
    beforeEach(async () => {
      el.readOnly = true;
      await settle();
    });

    it('reflects the read-only attribute', () => {
      expect(el.hasAttribute('read-only')).toBe(true);
    });

    it('has no controls to add, edit or delete', () => {
      expect(newButton()).toBeNull();
      expect(root().querySelectorAll('[data-action="edit"], [data-action="delete"]')).toHaveLength(0);
      expect([...root().querySelectorAll('button')].map((b) => b.getAttribute('aria-label') ?? b.textContent)).toEqual([
        'Go to double',
        'Go to Sales',
        'Go to taxRate',
        'Close',
      ]);
    });

    it('still goes to a name and closes', async () => {
      const go = listen('name-go');
      const close = listen('name-close');
      await click(rowButton('taxRate', 'go'));
      key(rowButton('taxRate', 'go')!, 'Escape');
      expect(details(go)).toEqual([{ name: 'taxRate' }]);
      expect(close).toHaveBeenCalledOnce();
    });

    it('closes a form that was open when it became read-only', async () => {
      el.readOnly = false;
      await settle();
      const saved = listen('name-save');
      await openNew('total');
      el.readOnly = true;
      await settle();
      expect(nameField()).toBeNull();
      expect(q('table')).not.toBeNull();
      expect(saved).not.toHaveBeenCalled();
      // Focus was in the form: it stays in the dialog
      expect(focused()).toBe(rowButton('double', 'go'));
    });
  });
});
