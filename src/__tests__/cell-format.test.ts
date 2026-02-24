// @vitest-environment jsdom
import { describe, it, expect, beforeEach } from 'vitest';
import { ClipboardManager } from '../controllers/clipboard-manager.js';
import type { CellFormat } from '../types.js';

describe('Cell Format Utilities', () => {
  let manager: ClipboardManager;

  beforeEach(() => {
    manager = new ClipboardManager();
  });

  describe('_formatToInlineStyle', () => {
    it('returns empty string for empty format', () => {
      expect(manager._formatToInlineStyle({})).toBe('');
    });

    it('produces font-weight: bold', () => {
      expect(manager._formatToInlineStyle({ bold: true })).toBe('font-weight: bold');
    });

    it('produces font-style: italic', () => {
      expect(manager._formatToInlineStyle({ italic: true })).toBe('font-style: italic');
    });

    it('produces text-decoration: underline', () => {
      expect(manager._formatToInlineStyle({ underline: true })).toBe(
        'text-decoration: underline'
      );
    });

    it('produces text-decoration: line-through', () => {
      expect(manager._formatToInlineStyle({ strikethrough: true })).toBe(
        'text-decoration: line-through'
      );
    });

    it('merges underline and strikethrough into one text-decoration', () => {
      expect(
        manager._formatToInlineStyle({ underline: true, strikethrough: true })
      ).toBe('text-decoration: underline line-through');
    });

    it('produces color for textColor', () => {
      expect(manager._formatToInlineStyle({ textColor: '#ff0000' })).toBe(
        'color: #ff0000'
      );
    });

    it('produces background-color for backgroundColor', () => {
      expect(manager._formatToInlineStyle({ backgroundColor: '#00ff00' })).toBe(
        'background-color: #00ff00'
      );
    });

    it('produces text-align', () => {
      expect(manager._formatToInlineStyle({ textAlign: 'center' })).toBe(
        'text-align: center'
      );
    });

    it('produces font-size in px', () => {
      expect(manager._formatToInlineStyle({ fontSize: 16 })).toBe('font-size: 16px');
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
      const result = manager._formatToInlineStyle(fmt);
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
      expect(manager._parseStyleToFormat('')).toBeUndefined();
    });

    it('parses font-weight: bold', () => {
      expect(manager._parseStyleToFormat('font-weight: bold')).toEqual({
        bold: true,
      });
    });

    it('parses font-weight: 700 as bold', () => {
      expect(manager._parseStyleToFormat('font-weight: 700')).toEqual({
        bold: true,
      });
    });

    it('parses font-style: italic', () => {
      expect(manager._parseStyleToFormat('font-style: italic')).toEqual({
        italic: true,
      });
    });

    it('parses text-decoration: underline', () => {
      expect(manager._parseStyleToFormat('text-decoration: underline')).toEqual({
        underline: true,
      });
    });

    it('parses text-decoration: line-through', () => {
      expect(manager._parseStyleToFormat('text-decoration: line-through')).toEqual({
        strikethrough: true,
      });
    });

    it('parses combined text-decoration: underline line-through', () => {
      expect(
        manager._parseStyleToFormat('text-decoration: underline line-through')
      ).toEqual({ underline: true, strikethrough: true });
    });

    it('parses color', () => {
      expect(manager._parseStyleToFormat('color: #ff0000')).toEqual({
        textColor: '#ff0000',
      });
    });

    it('parses background-color', () => {
      expect(manager._parseStyleToFormat('background-color: #00ff00')).toEqual({
        backgroundColor: '#00ff00',
      });
    });

    it('parses text-align', () => {
      expect(manager._parseStyleToFormat('text-align: center')).toEqual({
        textAlign: 'center',
      });
    });

    it('parses font-size', () => {
      expect(manager._parseStyleToFormat('font-size: 18px')).toEqual({
        fontSize: 18,
      });
    });

    it('parses multiple properties', () => {
      const style = 'font-weight: bold; font-style: italic; color: red; font-size: 14px';
      const result = manager._parseStyleToFormat(style);
      expect(result?.bold).toBe(true);
      expect(result?.italic).toBe(true);
      expect(result?.textColor).toBe('red');
      expect(result?.fontSize).toBe(14);
    });
  });

  describe('round-trip: serialize → parse → match', () => {
    it('round-trips bold', () => {
      const fmt: CellFormat = { bold: true };
      const style = manager._formatToInlineStyle(fmt);
      const parsed = manager._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips italic', () => {
      const fmt: CellFormat = { italic: true };
      const style = manager._formatToInlineStyle(fmt);
      const parsed = manager._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips underline + strikethrough', () => {
      const fmt: CellFormat = { underline: true, strikethrough: true };
      const style = manager._formatToInlineStyle(fmt);
      const parsed = manager._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips textColor', () => {
      const fmt: CellFormat = { textColor: '#c62828' };
      const style = manager._formatToInlineStyle(fmt);
      const parsed = manager._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips backgroundColor', () => {
      const fmt: CellFormat = { backgroundColor: '#e3f2fd' };
      const style = manager._formatToInlineStyle(fmt);
      const parsed = manager._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips textAlign', () => {
      const fmt: CellFormat = { textAlign: 'right' };
      const style = manager._formatToInlineStyle(fmt);
      const parsed = manager._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });

    it('round-trips fontSize', () => {
      const fmt: CellFormat = { fontSize: 24 };
      const style = manager._formatToInlineStyle(fmt);
      const parsed = manager._parseStyleToFormat(style);
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
      const style = manager._formatToInlineStyle(fmt);
      const parsed = manager._parseStyleToFormat(style);
      expect(parsed).toEqual(fmt);
    });
  });
});
