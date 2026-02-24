// @vitest-environment jsdom
import { describe, it, expect, beforeEach } from 'vitest';
import { ClipboardManager } from '../controllers/clipboard-manager.js';
import { computeSelectionFormat } from '../components/y11n-format-toolbar.js';
import type { CellFormat, SelectionRange } from '../types.js';

describe('Cell Format Utilities', () => {
  let manager: ClipboardManager;

  beforeEach(() => {
    manager = new ClipboardManager();
  });

  describe('_formatToInlineStyle', () => {
    it('returns empty string for empty format', () => {
      expect((manager as any)._formatToInlineStyle({})).toBe('');
    });

    it('produces font-weight: bold', () => {
      expect((manager as any)._formatToInlineStyle({ bold: true })).toBe('font-weight: bold');
    });

    it('produces font-style: italic', () => {
      expect((manager as any)._formatToInlineStyle({ italic: true })).toBe('font-style: italic');
    });

    it('produces text-decoration: underline', () => {
      expect((manager as any)._formatToInlineStyle({ underline: true })).toBe(
        'text-decoration: underline'
      );
    });

    it('produces text-decoration: line-through', () => {
      expect((manager as any)._formatToInlineStyle({ strikethrough: true })).toBe(
        'text-decoration: line-through'
      );
    });

    it('merges underline and strikethrough into one text-decoration', () => {
      expect(
        (manager as any)._formatToInlineStyle({ underline: true, strikethrough: true })
      ).toBe('text-decoration: underline line-through');
    });

    it('produces color for textColor', () => {
      expect((manager as any)._formatToInlineStyle({ textColor: '#ff0000' })).toBe(
        'color: #ff0000'
      );
    });

    it('produces background-color for backgroundColor', () => {
      expect((manager as any)._formatToInlineStyle({ backgroundColor: '#00ff00' })).toBe(
        'background-color: #00ff00'
      );
    });

    it('produces text-align', () => {
      expect((manager as any)._formatToInlineStyle({ textAlign: 'center' })).toBe(
        'text-align: center'
      );
    });

    it('produces font-size in px', () => {
      expect((manager as any)._formatToInlineStyle({ fontSize: 16 })).toBe('font-size: 16px');
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
      const result = (manager as any)._formatToInlineStyle(fmt);
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
      expect((manager as any)._parseStyleToFormat('')).toBeUndefined();
    });

    it('parses font-weight: bold', () => {
      expect((manager as any)._parseStyleToFormat('font-weight: bold')).toEqual({
        bold: true,
      });
    });

    it('parses font-weight: 700 as bold', () => {
      expect((manager as any)._parseStyleToFormat('font-weight: 700')).toEqual({
        bold: true,
      });
    });

    it('parses font-style: italic', () => {
      expect((manager as any)._parseStyleToFormat('font-style: italic')).toEqual({
        italic: true,
      });
    });

    it('parses text-decoration: underline', () => {
      expect((manager as any)._parseStyleToFormat('text-decoration: underline')).toEqual({
        underline: true,
      });
    });

    it('parses text-decoration: line-through', () => {
      expect((manager as any)._parseStyleToFormat('text-decoration: line-through')).toEqual({
        strikethrough: true,
      });
    });

    it('parses combined text-decoration: underline line-through', () => {
      expect(
        (manager as any)._parseStyleToFormat('text-decoration: underline line-through')
      ).toEqual({ underline: true, strikethrough: true });
    });

    it('parses color', () => {
      expect((manager as any)._parseStyleToFormat('color: #ff0000')).toEqual({
        textColor: '#ff0000',
      });
    });

    it('parses background-color', () => {
      expect((manager as any)._parseStyleToFormat('background-color: #00ff00')).toEqual({
        backgroundColor: '#00ff00',
      });
    });

    it('parses text-align', () => {
      expect((manager as any)._parseStyleToFormat('text-align: center')).toEqual({
        textAlign: 'center',
      });
    });

    it('parses font-size', () => {
      expect((manager as any)._parseStyleToFormat('font-size: 18px')).toEqual({
        fontSize: 18,
      });
    });

    it('parses multiple properties', () => {
      const style = 'font-weight: bold; font-style: italic; color: red; font-size: 14px';
      const result = (manager as any)._parseStyleToFormat(style);
      expect(result?.bold).toBe(true);
      expect(result?.italic).toBe(true);
      expect(result?.textColor).toBe('red');
      expect(result?.fontSize).toBe(14);
    });
  });

  describe('round-trip: serialize → parse → match', () => {
    it('round-trips bold', () => {
      const fmt: CellFormat = { bold: true };
      const style = (manager as any)._formatToInlineStyle(fmt);
      const parsed = (manager as any)._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips italic', () => {
      const fmt: CellFormat = { italic: true };
      const style = (manager as any)._formatToInlineStyle(fmt);
      const parsed = (manager as any)._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips underline + strikethrough', () => {
      const fmt: CellFormat = { underline: true, strikethrough: true };
      const style = (manager as any)._formatToInlineStyle(fmt);
      const parsed = (manager as any)._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips textColor', () => {
      const fmt: CellFormat = { textColor: '#c62828' };
      const style = (manager as any)._formatToInlineStyle(fmt);
      const parsed = (manager as any)._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips backgroundColor', () => {
      const fmt: CellFormat = { backgroundColor: '#e3f2fd' };
      const style = (manager as any)._formatToInlineStyle(fmt);
      const parsed = (manager as any)._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips textAlign', () => {
      const fmt: CellFormat = { textAlign: 'right' };
      const style = (manager as any)._formatToInlineStyle(fmt);
      const parsed = (manager as any)._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips fontSize', () => {
      const fmt: CellFormat = { fontSize: 24 };
      const style = (manager as any)._formatToInlineStyle(fmt);
      const parsed = (manager as any)._parseStyleToFormat(style);
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
      const style = (manager as any)._formatToInlineStyle(fmt);
      const parsed = (manager as any)._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });
  });

  describe('_extractStyleProp regex fix (color vs background-color)', () => {
    it('does not match color inside background-color', () => {
      const result = (manager as any)._parseStyleToFormat('background-color: #00ff00');
      expect(result?.textColor).toBeUndefined();
      expect(result?.backgroundColor).toBe('#00ff00');
    });

    it('correctly parses both color and background-color together', () => {
      const result = (manager as any)._parseStyleToFormat('color: red; background-color: blue');
      expect(result?.textColor).toBe('red');
      expect(result?.backgroundColor).toBe('blue');
    });

    it('parses color alone without background-color interference', () => {
      const result = (manager as any)._parseStyleToFormat('color: #333');
      expect(result?.textColor).toBe('#333');
      expect(result?.backgroundColor).toBeUndefined();
    });

    it('parses background-color before color in same string', () => {
      const result = (manager as any)._parseStyleToFormat('background-color: yellow; color: green');
      expect(result?.textColor).toBe('green');
      expect(result?.backgroundColor).toBe('yellow');
    });

    it('handles font-weight: 600 as bold (external source compat)', () => {
      const result = (manager as any)._parseStyleToFormat('font-weight: 600');
      expect(result?.bold).toBe(true);
    });

    it('handles font-weight: 800 as bold', () => {
      const result = (manager as any)._parseStyleToFormat('font-weight: 800');
      expect(result?.bold).toBe(true);
    });

    it('does not treat font-weight: 400 as bold', () => {
      const result = (manager as any)._parseStyleToFormat('font-weight: 400');
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
});
