import { describe, expect, it } from 'vitest';
import { type ListItem, listItemsOf, listToFormula, listToText, looksLikeList, parseListText } from '../list.js';

describe('listToFormula', () => {
  it('writes the items as a column array kept in one cell', () => {
    expect(listToFormula(['Ann', 'Bob', 3])).toBe('={{"Ann";"Bob";3}}');
    expect(listToFormula(['x'])).toBe('={{"x"}}');
  });

  it('writes numbers and booleans as typed constants', () => {
    expect(listToFormula([true, false, 1.5, -2, 0])).toBe('={{TRUE;FALSE;1.5;-2;0}}');
  });

  it('keeps text that looks like a number or boolean as text', () => {
    expect(listToFormula(['007', 'TRUE', '1e3'])).toBe('={{"007";"TRUE";"1e3"}}');
  });

  it('escapes quotes and keeps commas, semicolons and spaces', () => {
    expect(listToFormula(['say "hi"', 'a,b', 'c;d', ' pad ', ''])).toBe('={{"say ""hi""";"a,b";"c;d";" pad ";""}}');
  });

  it('is an empty string for no items', () => {
    expect(listToFormula([])).toBe('');
  });

  it('writes numbers that are not finite as 0', () => {
    expect(listToFormula([NaN, Infinity])).toBe('={{0;0}}');
  });
});

describe('listItemsOf', () => {
  it('reads the items of a list formula', () => {
    expect(listItemsOf('={{"Ann";"Bob";3}}')).toEqual(['Ann', 'Bob', 3]);
    expect(listItemsOf('={{1;2}}')).toEqual([1, 2]);
    expect(listItemsOf('={{1}}')).toEqual([1]);
    expect(listItemsOf('={{-1;"a";TRUE;FALSE}}')).toEqual([-1, 'a', true, false]);
  });

  it('accepts a row of items and spaces between them', () => {
    expect(listItemsOf('={{1,2,3}}')).toEqual([1, 2, 3]);
    expect(listItemsOf('={{ 1 ; "a b" ; 2 }}')).toEqual([1, 'a b', 2]);
  });

  it('unescapes quotes in text items', () => {
    expect(listItemsOf('={{"a""b";"";"x;y"}}')).toEqual(['a"b', '', 'x;y']);
  });

  it('rejects formulas that are not a list', () => {
    expect(listItemsOf('={1,2}')).toBeNull();
    expect(listItemsOf('=SUM(1)')).toBeNull();
    expect(listItemsOf('={{1;2}}+1')).toBeNull();
    expect(listItemsOf('={{1;2};{3;4}}')).toBeNull();
    expect(listItemsOf('={{{1;2}}}')).toBeNull();
    // A table is not a list
    expect(listItemsOf('={{1,2;3,4}}')).toBeNull();
  });

  it('rejects lists with anything but constants in them', () => {
    expect(listItemsOf('={{A1;2}}')).toBeNull();
    expect(listItemsOf('={{1+1}}')).toBeNull();
    expect(listItemsOf('={{SUM(1);2}}')).toBeNull();
    expect(listItemsOf('={{-"a"}}')).toBeNull();
    expect(listItemsOf('={{--1}}')).toBeNull();
    expect(listItemsOf('={{1%}}')).toBeNull();
    expect(listItemsOf('={{#N/A}}')).toBeNull();
    expect(listItemsOf('={{Name}}')).toBeNull();
  });

  it('rejects non-formulas and formulas that do not parse', () => {
    expect(listItemsOf('')).toBeNull();
    expect(listItemsOf('hello')).toBeNull();
    expect(listItemsOf('Ann, Bob')).toBeNull();
    expect(listItemsOf('{{1;2}}')).toBeNull();
    expect(listItemsOf('={{1;2}')).toBeNull();
    expect(listItemsOf('={{1;;2}}')).toBeNull();
  });
});

describe('listToText', () => {
  it('joins the items with a comma and a space', () => {
    expect(listToText(['Ann', 'Bob', 3])).toBe('Ann, Bob, 3');
    expect(listToText([true, false, 1.5, -2])).toBe('TRUE, FALSE, 1.5, -2');
    expect(listToText([])).toBe('');
  });

  it('quotes text that contains the separator or a quote', () => {
    expect(listToText(['Smith, J', 'x'])).toBe('"Smith, J", x');
    expect(listToText(['say "hi"'])).toBe('"say ""hi"""');
  });

  it('quotes text that would read back as a number or boolean', () => {
    expect(listToText(['007', 7])).toBe('"007", 7');
    expect(listToText(['TRUE', true, 'false'])).toBe('"TRUE", TRUE, "false"');
    expect(listToText(['1e3', '-5', '.5'])).toBe('"1e3", "-5", ".5"');
  });

  it('quotes empty text and text with outer spaces', () => {
    expect(listToText(['', ' pad ', 'in side'])).toBe('"", " pad ", in side');
  });

  it('uses another separator when given', () => {
    expect(listToText(['a;b', 'c', 'x,y'], ';')).toBe('"a;b"; c; x,y');
  });
});

describe('parseListText', () => {
  it('splits at commas and trims the items', () => {
    expect(parseListText('a, b, c')).toEqual(['a', 'b', 'c']);
    expect(parseListText('  a ,b  ,  c d  ')).toEqual(['a', 'b', 'c d']);
    expect(parseListText('one')).toEqual(['one']);
  });

  it('types numbers and booleans', () => {
    expect(parseListText('1, 2,3')).toEqual([1, 2, 3]);
    expect(parseListText('1e3, -2.5, .5')).toEqual([1000, -2.5, 0.5]);
    expect(parseListText('TRUE, false, x')).toEqual([true, false, 'x']);
    // Not numbers a user means
    expect(parseListText('Infinity, NaN')).toEqual(['Infinity', 'NaN']);
  });

  it('keeps quoted items as text exactly as written', () => {
    expect(parseListText('"007", 7, 007')).toEqual(['007', 7, 7]);
    expect(parseListText('"TRUE", TRUE')).toEqual(['TRUE', true]);
    expect(parseListText('"a,b", c')).toEqual(['a,b', 'c']);
    expect(parseListText('" pad ",x')).toEqual([' pad ', 'x']);
    expect(parseListText('"say ""hi""", x')).toEqual(['say "hi"', 'x']);
  });

  it('drops empty items, but keeps a quoted empty one', () => {
    expect(parseListText('a,,b, ')).toEqual(['a', 'b']);
    expect(parseListText('a,"",b')).toEqual(['a', '', 'b']);
    expect(parseListText('')).toEqual([]);
    expect(parseListText('  ')).toEqual([]);
    expect(parseListText(',')).toEqual([]);
  });

  it('is lenient with malformed quoting', () => {
    // Text after a closing quote is ignored up to the separator
    expect(parseListText('"abc" junk, d')).toEqual(['abc', 'd']);
    // A quote that is never closed runs to the end
    expect(parseListText('"unterminated, x')).toEqual(['unterminated, x']);
  });

  it('splits at another separator when given', () => {
    expect(parseListText('a; b;c', ';')).toEqual(['a', 'b', 'c']);
    expect(parseListText('a, b; c', ';')).toEqual(['a, b', 'c']);
    expect(parseListText('a; b; c')).toEqual(['a; b; c']);
  });
});

describe('looksLikeList', () => {
  it('is true for text with at least two items', () => {
    expect(looksLikeList('a, b')).toBe(true);
    expect(looksLikeList('a,b')).toBe(true);
    expect(looksLikeList('1, 2, 3')).toBe(true);
    expect(looksLikeList('"a,b", c')).toBe(true);
  });

  it('is false for a single item', () => {
    expect(looksLikeList('one')).toBe(false);
    expect(looksLikeList('a,')).toBe(false);
    expect(looksLikeList(',')).toBe(false);
    expect(looksLikeList('')).toBe(false);
    // The comma is inside quotes
    expect(looksLikeList('"a,b"')).toBe(false);
  });

  it('is false for formulas', () => {
    expect(looksLikeList('=SUM(A1,B1)')).toBe(false);
    expect(looksLikeList('={{1;2}}')).toBe(false);
  });

  it('follows the separator', () => {
    expect(looksLikeList('a;b', ';')).toBe(true);
    expect(looksLikeList('a;b')).toBe(false);
    expect(looksLikeList('a,b', ';')).toBe(false);
  });
});

describe('list round trips', () => {
  const lists: ListItem[][] = [
    ['Ann', 'Bob', 3],
    ['a,b', 'say "hi"', '007', ' pad ', ''],
    [true, false, 1.5, -2, 0],
    ['TRUE', 'true', '1e3', '-1'],
    ['x'],
    ['semi;colon', 'new\nline', "apostrophe's", '{brace}', '=formula'],
    [0.1, 1e21, -0.000001, 123456789.123],
  ];

  it('items → formula → items', () => {
    for (const items of lists) {
      expect(listItemsOf(listToFormula(items)), JSON.stringify(items)).toEqual(items);
    }
  });

  it('items → text → items', () => {
    for (const items of lists) {
      expect(parseListText(listToText(items)), JSON.stringify(items)).toEqual(items);
    }
  });

  it('items → text → items with another separator', () => {
    for (const items of lists) {
      expect(parseListText(listToText(items, ';'), ';'), JSON.stringify(items)).toEqual(items);
    }
  });

  it('typed text → formula → text', () => {
    const typed = 'Ann, "Smith, J", 42, TRUE, "007"';
    const items = parseListText(typed);
    expect(items).toEqual(['Ann', 'Smith, J', 42, true, '007']);
    const formula = listToFormula(items);
    expect(formula).toBe('={{"Ann";"Smith, J";42;TRUE;"007"}}');
    expect(listToText(listItemsOf(formula) ?? [])).toBe(typed);
  });

  it('nothing typed gives no list', () => {
    expect(listToFormula(parseListText(' , '))).toBe('');
    expect(listItemsOf(listToFormula([]))).toBeNull();
  });
});
