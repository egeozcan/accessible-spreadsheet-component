import { describe, it, expect } from 'vitest';
import { parseDelimited, quoteDelimitedField, serializeDelimited } from '../delimited.js';

describe('parseDelimited', () => {
  it('parses simple CSV', () => {
    expect(parseDelimited('a,b\n1,2')).toEqual([['a', 'b'], ['1', '2']]);
  });

  it('handles quoted fields with delimiters, quotes and newlines', () => {
    expect(parseDelimited('"a,b","say ""hi""","line1\nline2"')).toEqual([
      ['a,b', 'say "hi"', 'line1\nline2'],
    ]);
  });

  it('treats a mid-field quote as literal', () => {
    expect(parseDelimited('5" screen,x')).toEqual([['5" screen', 'x']]);
  });

  it('accepts CRLF, LF and bare CR row endings', () => {
    expect(parseDelimited('a\r\nb\nc\rd')).toEqual([['a'], ['b'], ['c'], ['d']]);
  });

  it('ignores a single trailing newline but keeps empty fields', () => {
    expect(parseDelimited('a,,c\n')).toEqual([['a', '', 'c']]);
  });

  it('keeps blank rows in the middle', () => {
    expect(parseDelimited('a\n\nb')).toEqual([['a'], [''], ['b']]);
  });

  it('supports other delimiters', () => {
    expect(parseDelimited('a\tb;c', '\t')).toEqual([['a', 'b;c']]);
    expect(parseDelimited('a;b', ';')).toEqual([['a', 'b']]);
  });

  it('keeps an explicitly quoted empty final record', () => {
    expect(parseDelimited('""')).toEqual([['']]);
    expect(parseDelimited('x\n""')).toEqual([['x'], ['']]);
    expect(parseDelimited('x\n""\n')).toEqual([['x'], ['']]);
  });

  it('ignores a leading UTF-8 BOM', () => {
    expect(parseDelimited('\uFEFF"a,b",c')).toEqual([['a,b', 'c']]);
  });

  it('returns no rows for empty input', () => {
    expect(parseDelimited('')).toEqual([]);
  });
});

describe('serializeDelimited', () => {
  it('quotes only when needed', () => {
    expect(quoteDelimitedField('plain')).toBe('plain');
    expect(quoteDelimitedField('a,b')).toBe('"a,b"');
    expect(quoteDelimitedField('a\tb', '\t')).toBe('"a\tb"');
    expect(quoteDelimitedField('a\tb', ',')).toBe('a\tb');
    expect(quoteDelimitedField('say "hi"')).toBe('"say ""hi"""');
  });

  it('writes a single-empty-field row as "" so trailing empty rows survive', () => {
    expect(serializeDelimited([['x'], ['']])).toBe('x\n""');
    expect(parseDelimited(serializeDelimited([['x'], ['']]))).toEqual([['x'], ['']]);
    expect(parseDelimited(serializeDelimited([['']]))).toEqual([['']]);
  });

  it('round-trips through parseDelimited', () => {
    const rows = [['a,b', 'x"y', 'multi\nline'], ['', '1', '=SUM(A1:A2)']];
    expect(parseDelimited(serializeDelimited(rows))).toEqual(rows);
  });
});
