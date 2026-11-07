/**
 * Parsing and serialization for delimiter-separated text (CSV, TSV).
 *
 * Follows RFC 4180 quoting: a field that contains the delimiter, a double
 * quote, or a line break is wrapped in double quotes, and embedded quotes are
 * doubled (`""`). This is also the format Excel and Google Sheets put on the
 * clipboard as TSV.
 */

/**
 * Parse delimited text into a 2D array of strings.
 *
 * - Quoted fields may contain delimiters, line breaks, and escaped quotes (`""`).
 * - A quote only opens a quoted field at the start of a field; a quote in the
 *   middle of a field (e.g. `5" screen`, which Excel leaves unquoted) is literal.
 * - `\r\n`, `\n`, and bare `\r` all end a row.
 * - A single trailing line break does not produce an extra empty row.
 * - A leading UTF-8 byte-order mark (common in Excel exports) is ignored.
 *
 * @param text - The text to parse
 * @param delimiter - Field separator (default `,`)
 * @returns Rows of fields
 */
export function parseDelimited(text: string, delimiter = ','): string[][] {
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentField = '';
  let inQuotes = false;
  // Whether the current field was quoted, so `""` counts as a real field
  let fieldQuoted = false;
  let i = text.charCodeAt(0) === 0xfeff ? 1 : 0;

  const endRow = () => {
    currentRow.push(currentField);
    currentField = '';
    fieldQuoted = false;
    rows.push(currentRow);
    currentRow = [];
  };

  while (i < text.length) {
    const ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          currentField += '"';
          i += 2;
        } else {
          inQuotes = false;
          i++;
        }
      } else {
        currentField += ch;
        i++;
      }
      continue;
    }

    if (ch === '"' && currentField === '' && !fieldQuoted) {
      inQuotes = true;
      fieldQuoted = true;
      i++;
    } else if (ch === delimiter) {
      currentRow.push(currentField);
      currentField = '';
      fieldQuoted = false;
      i++;
    } else if (ch === '\r') {
      endRow();
      i += text[i + 1] === '\n' ? 2 : 1;
    } else if (ch === '\n') {
      endRow();
      i++;
    } else {
      currentField += ch;
      i++;
    }
  }

  // Push the last field and row, unless it is the empty remainder after a
  // trailing line break.
  currentRow.push(currentField);
  if (currentRow.length > 1 || currentRow[0] !== '' || fieldQuoted) {
    rows.push(currentRow);
  }

  return rows;
}

/**
 * Quote a single field if it contains the delimiter, a quote, or a line break.
 *
 * @param value - The field value
 * @param delimiter - Field separator (default `,`)
 * @returns The field, quoted and escaped when necessary
 */
export function quoteDelimitedField(value: string, delimiter = ','): string {
  if (
    value.includes(delimiter) ||
    value.includes('"') ||
    value.includes('\n') ||
    value.includes('\r')
  ) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

/**
 * Serialize a 2D array of strings as delimited text. Rows are joined with `\n`.
 *
 * @param rows - Rows of fields
 * @param delimiter - Field separator (default `,`)
 * @returns The delimited text
 */
export function serializeDelimited(rows: string[][], delimiter = ','): string {
  return rows
    .map((row) =>
      // A row that is a single empty field is written as "" so it survives a
      // round trip (a bare empty line at the end would be dropped as trailing).
      row.length === 1 && row[0] === ''
        ? '""'
        : row.map((field) => quoteDelimitedField(field, delimiter)).join(delimiter)
    )
    .join('\n');
}
