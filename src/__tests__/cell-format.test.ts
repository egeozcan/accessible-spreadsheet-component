// @vitest-environment jsdom
import { describe, it, expect, beforeEach } from 'vitest';
import { ClipboardManager } from '../controllers/clipboard-manager.js';
import { computeSelectionFormat } from '../components/y11n-format-toolbar.js';
import { formatsEqual } from '../types.js';
import type { CellFormat, SelectionRange } from '../types.js';

// Type-safe wrappers for testing private methods
let formatToInlineStyle: (fmt: CellFormat) => string;
let parseStyleToFormat: (style: string) => CellFormat | undefined;

describe('Cell Format Utilities', () => {
  let manager: ClipboardManager;

  beforeEach(() => {
    manager = new ClipboardManager();
    formatToInlineStyle = (fmt) => (manager as any)._formatToInlineStyle(fmt);
    parseStyleToFormat = (style) => (manager as any)._parseStyleToFormat(style);
  });

  describe('_formatToInlineStyle', () => {
    it('returns empty string for empty format', () => {
      expect(formatToInlineStyle({})).toBe('');
    });

    it('produces font-weight: bold', () => {
      expect(formatToInlineStyle({ bold: true })).toBe('font-weight: bold');
    });

    it('produces font-style: italic', () => {
      expect(formatToInlineStyle({ italic: true })).toBe('font-style: italic');
    });

    it('produces text-decoration: underline', () => {
      expect(formatToInlineStyle({ underline: true })).toBe(
        'text-decoration: underline'
      );
    });

    it('produces text-decoration: line-through', () => {
      expect(formatToInlineStyle({ strikethrough: true })).toBe(
        'text-decoration: line-through'
      );
    });

    it('merges underline and strikethrough into one text-decoration', () => {
      expect(
        formatToInlineStyle({ underline: true, strikethrough: true })
      ).toBe('text-decoration: underline line-through');
    });

    it('produces color for textColor', () => {
      expect(formatToInlineStyle({ textColor: '#ff0000' })).toBe(
        'color: #ff0000'
      );
    });

    it('produces background-color for backgroundColor', () => {
      expect(formatToInlineStyle({ backgroundColor: '#00ff00' })).toBe(
        'background-color: #00ff00'
      );
    });

    it('produces text-align', () => {
      expect(formatToInlineStyle({ textAlign: 'center' })).toBe(
        'text-align: center'
      );
    });

    it('produces font-size in px', () => {
      expect(formatToInlineStyle({ fontSize: 16 })).toBe('font-size: 16px');
    });

    it('produces combined style for all properties', () => {
      const fmt: CellFormat = {
        bold: true,
        italic: true,
        underline: true,
        strikethrough: true,
        textColor: '#333',
        backgroundColor: '#fff',
        textAlign: 'right',
        fontSize: 20,
      };
      const result = formatToInlineStyle(fmt);
      expect(result).toContain('font-weight: bold');
      expect(result).toContain('font-style: italic');
      expect(result).toContain('text-decoration: underline line-through');
      expect(result).toContain('color: #333');
      expect(result).toContain('background-color: #fff');
      expect(result).toContain('text-align: right');
      expect(result).toContain('font-size: 20px');
    });
  });

  describe('_parseStyleToFormat', () => {
    it('returns undefined for empty string', () => {
      expect(parseStyleToFormat('')).toBeUndefined();
    });

    it('parses font-weight: bold', () => {
      expect(parseStyleToFormat('font-weight: bold')).toEqual({
        bold: true,
      });
    });

    it('parses font-weight: 700 as bold', () => {
      expect(parseStyleToFormat('font-weight: 700')).toEqual({
        bold: true,
      });
    });

    it('parses font-style: italic', () => {
      expect(parseStyleToFormat('font-style: italic')).toEqual({
        italic: true,
      });
    });

    it('parses text-decoration: underline', () => {
      expect(parseStyleToFormat('text-decoration: underline')).toEqual({
        underline: true,
      });
    });

    it('parses text-decoration: line-through', () => {
      expect(parseStyleToFormat('text-decoration: line-through')).toEqual({
        strikethrough: true,
      });
    });

    it('parses combined text-decoration: underline line-through', () => {
      expect(
        parseStyleToFormat('text-decoration: underline line-through')
      ).toEqual({ underline: true, strikethrough: true });
    });

    it('parses color', () => {
      expect(parseStyleToFormat('color: #ff0000')).toEqual({
        textColor: '#ff0000',
      });
    });

    it('parses background-color', () => {
      expect(parseStyleToFormat('background-color: #00ff00')).toEqual({
        backgroundColor: '#00ff00',
      });
    });

    it('parses text-align', () => {
      expect(parseStyleToFormat('text-align: center')).toEqual({
        textAlign: 'center',
      });
    });

    it('parses font-size', () => {
      expect(parseStyleToFormat('font-size: 18px')).toEqual({
        fontSize: 18,
      });
    });

    it('parses multiple properties', () => {
      const style = 'font-weight: bold; font-style: italic; color: red; font-size: 14px';
      const result = parseStyleToFormat(style);
      expect(result?.bold).toBe(true);
      expect(result?.italic).toBe(true);
      expect(result?.textColor).toBe('red');
      expect(result?.fontSize).toBe(14);
    });
  });

  describe('round-trip: serialize → parse → match', () => {
    it('round-trips bold', () => {
      const fmt: CellFormat = { bold: true };
      const style = formatToInlineStyle(fmt);
      const parsed = parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips italic', () => {
      const fmt: CellFormat = { italic: true };
      const style = formatToInlineStyle(fmt);
      const parsed = parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips underline + strikethrough', () => {
      const fmt: CellFormat = { underline: true, strikethrough: true };
      const style = formatToInlineStyle(fmt);
      const parsed = parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips textColor', () => {
      const fmt: CellFormat = { textColor: '#c62828' };
      const style = formatToInlineStyle(fmt);
      const parsed = parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips backgroundColor', () => {
      const fmt: CellFormat = { backgroundColor: '#e3f2fd' };
      const style = formatToInlineStyle(fmt);
      const parsed = parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips textAlign', () => {
      const fmt: CellFormat = { textAlign: 'right' };
      const style = formatToInlineStyle(fmt);
      const parsed = parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips fontSize', () => {
      const fmt: CellFormat = { fontSize: 24 };
      const style = formatToInlineStyle(fmt);
      const parsed = parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips a complex format', () => {
      const fmt: CellFormat = {
        bold: true,
        italic: true,
        textColor: '#1565c0',
        backgroundColor: '#fff9c4',
        textAlign: 'center',
        fontSize: 16,
      };
      const style = formatToInlineStyle(fmt);
      const parsed = parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });
  });

  describe('_extractStyleProp regex fix (color vs background-color)', () => {
    it('does not match color inside background-color', () => {
      const result = parseStyleToFormat('background-color: #00ff00');
      expect(result?.textColor).toBeUndefined();
      expect(result?.backgroundColor).toBe('#00ff00');
    });

    it('correctly parses both color and background-color together', () => {
      const result = parseStyleToFormat('color: red; background-color: blue');
      expect(result?.textColor).toBe('red');
      expect(result?.backgroundColor).toBe('blue');
    });

    it('parses color alone without background-color interference', () => {
      const result = parseStyleToFormat('color: #333');
      expect(result?.textColor).toBe('#333');
      expect(result?.backgroundColor).toBeUndefined();
    });

    it('parses background-color before color in same string', () => {
      const result = parseStyleToFormat('background-color: yellow; color: green');
      expect(result?.textColor).toBe('green');
      expect(result?.backgroundColor).toBe('yellow');
    });

    it('handles font-weight: 600 as bold (external source compat)', () => {
      const result = parseStyleToFormat('font-weight: 600');
      expect(result?.bold).toBe(true);
    });

    it('handles font-weight: 800 as bold', () => {
      const result = parseStyleToFormat('font-weight: 800');
      expect(result?.bold).toBe(true);
    });

    it('does not treat font-weight: 400 as bold', () => {
      const result = parseStyleToFormat('font-weight: 400');
      expect(result).toBeUndefined();
    });
  });

  describe('parseHTMLTableWithFormat', () => {
    it('returns null for non-table HTML', () => {
      expect(manager.parseHTMLTableWithFormat('<p>hello</p>')).toBeNull();
    });

    it('parses table with data-format (preferred over inline style)', () => {
      const fmt: CellFormat = { bold: true, textColor: '#ff0000' };
      const html = `<table><tr><td data-format='${JSON.stringify(fmt)}' style="font-style: italic">hello</td></tr></table>`;
      const result = manager.parseHTMLTableWithFormat(html);
      expect(result).not.toBeNull();
      expect(result![0][0].value).toBe('hello');
      // data-format should win over inline style
      expect(result![0][0].format?.bold).toBe(true);
      expect(result![0][0].format?.textColor).toBe('#ff0000');
      expect(result![0][0].format?.italic).toBeUndefined();
    });

    it('falls back to inline style when data-format is malformed JSON', () => {
      const html = '<table><tr><td data-format="not-json" style="font-weight: bold">test</td></tr></table>';
      const result = manager.parseHTMLTableWithFormat(html);
      expect(result).not.toBeNull();
      expect(result![0][0].value).toBe('test');
      expect(result![0][0].format?.bold).toBe(true);
    });

    it('parses table without any formatting attributes', () => {
      const html = '<table><tr><td>plain</td><td>text</td></tr></table>';
      const result = manager.parseHTMLTableWithFormat(html);
      expect(result).not.toBeNull();
      expect(result![0][0].value).toBe('plain');
      expect(result![0][0].format).toBeUndefined();
      expect(result![0][1].value).toBe('text');
      expect(result![0][1].format).toBeUndefined();
    });

    it('parses multi-row multi-col table with mixed formatting', () => {
      const html = `<table>
        <tr><td style="font-weight: bold">A</td><td>B</td></tr>
        <tr><td>C</td><td style="color: red">D</td></tr>
      </table>`;
      const result = manager.parseHTMLTableWithFormat(html);
      expect(result).not.toBeNull();
      expect(result).toHaveLength(2);
      expect(result![0][0].format?.bold).toBe(true);
      expect(result![0][1].format).toBeUndefined();
      expect(result![1][0].format).toBeUndefined();
      expect(result![1][1].format?.textColor).toBe('red');
    });

    it('uses data-raw for cell value when present', () => {
      const html = '<table><tr><td data-raw="=A1+1">42</td></tr></table>';
      const result = manager.parseHTMLTableWithFormat(html);
      expect(result![0][0].value).toBe('=A1+1');
    });

    it('sanitizes data-format rejecting invalid types', () => {
      const html = '<table><tr><td data-format=\'{"bold":"yes","fontSize":-5,"textAlign":"justify"}\'>x</td></tr></table>';
      const result = manager.parseHTMLTableWithFormat(html);
      expect(result).not.toBeNull();
      // "yes" is not boolean, -5 is out of range, "justify" is invalid align
      expect(result![0][0].format).toBeUndefined();
    });
  });
});

describe('computeSelectionFormat', () => {
  function makeData(cells: Record<string, { format?: CellFormat }>) {
    return (cellId: string) => cells[cellId];
  }

  const singleCell: SelectionRange = {
    start: { row: 0, col: 0 },
    end: { row: 0, col: 0 },
  };

  it('returns format for a single formatted cell', () => {
    const data = makeData({
      '0:0': { format: { bold: true, textColor: '#f00' } },
    });
    const result = computeSelectionFormat(data, singleCell);
    expect(result.bold).toBe(true);
    expect(result.textColor).toBe('#f00');
  });

  it('returns empty format for a single unformatted cell', () => {
    const data = makeData({ '0:0': {} });
    const result = computeSelectionFormat(data, singleCell);
    expect(result).toEqual({});
  });

  it('returns empty format for a single missing cell', () => {
    const data = makeData({});
    const result = computeSelectionFormat(data, singleCell);
    expect(result).toEqual({});
  });

  it('returns consensus for uniform multi-cell bold', () => {
    const data = makeData({
      '0:0': { format: { bold: true } },
      '0:1': { format: { bold: true } },
      '1:0': { format: { bold: true } },
      '1:1': { format: { bold: true } },
    });
    const range: SelectionRange = {
      start: { row: 0, col: 0 },
      end: { row: 1, col: 1 },
    };
    const result = computeSelectionFormat(data, range);
    expect(result.bold).toBe(true);
  });

  it('omits property when cells disagree (mixed)', () => {
    const data = makeData({
      '0:0': { format: { bold: true } },
      '0:1': { format: { bold: false } },
    });
    const range: SelectionRange = {
      start: { row: 0, col: 0 },
      end: { row: 0, col: 1 },
    };
    const result = computeSelectionFormat(data, range);
    // Mixed: bold is undefined, so not in result
    expect(result.bold).toBeUndefined();
  });

  it('returns uniform textAlign across all cells', () => {
    const data = makeData({
      '0:0': { format: { textAlign: 'center' } },
      '0:1': { format: { textAlign: 'center' } },
    });
    const range: SelectionRange = {
      start: { row: 0, col: 0 },
      end: { row: 0, col: 1 },
    };
    const result = computeSelectionFormat(data, range);
    expect(result.textAlign).toBe('center');
  });

  it('omits textAlign when cells have different alignment', () => {
    const data = makeData({
      '0:0': { format: { textAlign: 'left' } },
      '0:1': { format: { textAlign: 'right' } },
    });
    const range: SelectionRange = {
      start: { row: 0, col: 0 },
      end: { row: 0, col: 1 },
    };
    const result = computeSelectionFormat(data, range);
    expect(result.textAlign).toBeUndefined();
  });

  it('handles mix of formatted and empty cells', () => {
    const data = makeData({
      '0:0': { format: { bold: true, fontSize: 14 } },
      '0:1': {},
    });
    const range: SelectionRange = {
      start: { row: 0, col: 0 },
      end: { row: 0, col: 1 },
    };
    const result = computeSelectionFormat(data, range);
    // bold: true vs false → mixed → omitted
    expect(result.bold).toBeUndefined();
    // fontSize: 14 vs undefined → mixed → omitted
    expect(result.fontSize).toBeUndefined();
  });

  it('does not include bold when all cells are explicitly false', () => {
    const data = makeData({
      '0:0': { format: { bold: false } },
      '0:1': { format: { bold: false } },
    });
    const range: SelectionRange = {
      start: { row: 0, col: 0 },
      end: { row: 0, col: 1 },
    };
    const result = computeSelectionFormat(data, range);
    expect(result.bold).toBeUndefined();
  });
});

describe('_sanitizeFormat', () => {
  let manager: ClipboardManager;
  beforeEach(() => {
    manager = new ClipboardManager();
  });

  it('returns undefined for null', () => {
    expect((manager as any)._sanitizeFormat(null)).toBeUndefined();
  });

  it('returns undefined for array', () => {
    expect((manager as any)._sanitizeFormat([1, 2, 3])).toBeUndefined();
  });

  it('returns undefined for non-object primitive', () => {
    expect((manager as any)._sanitizeFormat('hello')).toBeUndefined();
    expect((manager as any)._sanitizeFormat(42)).toBeUndefined();
  });

  it('keeps valid fields and strips invalid ones (partial validity)', () => {
    const result = (manager as any)._sanitizeFormat({
      bold: true,
      fontSize: -5,
      textAlign: 'justify',
    });
    expect(result).toEqual({ bold: true });
  });

  it('rejects non-boolean bold', () => {
    expect((manager as any)._sanitizeFormat({ bold: 'yes' })).toBeUndefined();
  });

  it('rejects color with CSS injection characters', () => {
    expect(
      (manager as any)._sanitizeFormat({ textColor: 'red; background-image: url(evil)' })
    ).toBeUndefined();
  });

  it('rejects color with quotes', () => {
    expect(
      (manager as any)._sanitizeFormat({ textColor: '"onmouseover=alert(1)' })
    ).toBeUndefined();
  });

  it('accepts valid CSS color values', () => {
    expect((manager as any)._sanitizeFormat({ textColor: '#ff0000' })).toEqual({
      textColor: '#ff0000',
    });
    expect((manager as any)._sanitizeFormat({ textColor: 'rgb(255, 0, 0)' })).toEqual({
      textColor: 'rgb(255, 0, 0)',
    });
  });

  it('accepts modern CSS color syntax with /', () => {
    expect((manager as any)._sanitizeFormat({ textColor: 'rgb(255 0 0 / 0.5)' })).toEqual({
      textColor: 'rgb(255 0 0 / 0.5)',
    });
  });

  it('rejects color strings longer than 100 chars', () => {
    const longColor = 'a'.repeat(101);
    expect((manager as any)._sanitizeFormat({ textColor: longColor })).toBeUndefined();
  });

  it('fontSize boundary: 0 is rejected', () => {
    expect((manager as any)._sanitizeFormat({ fontSize: 0 })).toBeUndefined();
  });

  it('fontSize boundary: 200 is accepted', () => {
    expect((manager as any)._sanitizeFormat({ fontSize: 200 })).toEqual({ fontSize: 200 });
  });

  it('fontSize boundary: 201 is rejected', () => {
    expect((manager as any)._sanitizeFormat({ fontSize: 201 })).toBeUndefined();
  });

  it('accepts a fully valid format object', () => {
    const fmt = {
      bold: true,
      italic: false,
      textColor: '#333',
      backgroundColor: 'rgb(0, 128, 255)',
      textAlign: 'center',
      fontSize: 14,
    };
    const result = (manager as any)._sanitizeFormat(fmt);
    expect(result).toEqual({
      bold: true,
      italic: false,
      textColor: '#333',
      backgroundColor: 'rgb(0, 128, 255)',
      textAlign: 'center',
      fontSize: 14,
    });
  });
});

describe('formatsEqual', () => {
  it('returns true for both undefined', () => {
    expect(formatsEqual(undefined, undefined)).toBe(true);
  });

  it('returns false for undefined vs empty object', () => {
    expect(formatsEqual(undefined, {})).toBe(false);
  });

  it('returns false for empty object vs undefined', () => {
    expect(formatsEqual({}, undefined)).toBe(false);
  });

  it('returns true for two empty objects', () => {
    expect(formatsEqual({}, {})).toBe(true);
  });

  it('returns true for same keys and values', () => {
    expect(formatsEqual({ bold: true, fontSize: 14 }, { bold: true, fontSize: 14 })).toBe(true);
  });

  it('returns true regardless of key order', () => {
    expect(formatsEqual({ bold: true, italic: true }, { italic: true, bold: true })).toBe(true);
  });

  it('returns false for same keys different values', () => {
    expect(formatsEqual({ bold: true }, { bold: false })).toBe(false);
  });

  it('returns false when one has extra keys', () => {
    expect(formatsEqual({ bold: true }, { bold: true, italic: true })).toBe(false);
  });

  it('returns false when the other has extra keys', () => {
    expect(formatsEqual({ bold: true, italic: true }, { bold: true })).toBe(false);
  });
});

describe('_escapeHTML', () => {
  let manager: ClipboardManager;
  beforeEach(() => {
    manager = new ClipboardManager();
  });

  it('escapes ampersands', () => {
    expect((manager as any)._escapeHTML('a & b')).toBe('a &amp; b');
  });

  it('escapes angle brackets', () => {
    expect((manager as any)._escapeHTML('<script>alert(1)</script>')).toBe(
      '&lt;script&gt;alert(1)&lt;/script&gt;'
    );
  });

  it('escapes double quotes', () => {
    expect((manager as any)._escapeHTML('a "b" c')).toBe('a &quot;b&quot; c');
  });

  it('handles strings with multiple special characters', () => {
    expect((manager as any)._escapeHTML('<div class="x">&</div>')).toBe(
      '&lt;div class=&quot;x&quot;&gt;&amp;&lt;/div&gt;'
    );
  });

  it('returns plain strings unchanged', () => {
    expect((manager as any)._escapeHTML('hello world')).toBe('hello world');
  });
});
