import type { CellData, GridData } from '../src/types.js';

/** Convenience to build a CellData object */
function cell(rawValue: string, type: CellData['type'] = 'text', displayValue?: string): CellData {
  return {
    rawValue,
    displayValue: displayValue ?? rawValue,
    type,
  };
}

/** Build a GridData map from a 2D array (row-major). Empty strings are skipped. */
export function gridFromRows(rows: string[][]): GridData {
  const data: GridData = new Map();
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const val = rows[r][c];
      if (val === '') continue;
      const isNum = val !== '' && !isNaN(Number(val));
      const isFormula = val.startsWith('=');
      data.set(`${r}:${c}`, cell(val, isFormula ? 'text' : isNum ? 'number' : 'text'));
    }
  }
  return data;
}

/** Wrap the spreadsheet element in a container with fixed height for stories */
export function wrapInContainer(el: Element, height = '500px'): HTMLDivElement {
  const container = document.createElement('div');
  container.style.height = height;
  container.style.border = '1px solid #ccc';
  container.style.borderRadius = '4px';
  container.style.overflow = 'hidden';
  container.appendChild(el);
  return container;
}

// ─── Data Generators for Stories ─────────────────────

/** Generate a product catalog for VLOOKUP/INDEX/MATCH demos */
export function generateProductCatalog(): string[][] {
  return [
    ['Product ID', 'Product Name',    'Category',      'Price',  'Stock'],
    ['P001',       'Wireless Mouse',  'Electronics',   '29.99',  '150'],
    ['P002',       'Mechanical KB',   'Electronics',   '89.99',  '75'],
    ['P003',       'USB-C Hub',       'Electronics',   '45.00',  '200'],
    ['P004',       'Desk Lamp',       'Office',        '34.50',  '120'],
    ['P005',       'Monitor Stand',   'Office',        '59.99',  '90'],
    ['P006',       'Webcam HD',       'Electronics',   '69.99',  '60'],
    ['P007',       'Notebook A5',     'Stationery',    '4.99',   '500'],
    ['P008',       'Pen Set',         'Stationery',    '12.50',  '300'],
  ];
}

/** Generate categorized expense data for SUMIF/COUNTIF demos */
export function generateExpenseData(): string[][] {
  return [
    ['Date',       'Description',         'Category',      'Amount'],
    ['2024-01-05', 'Office supplies',     'Office',        '125.00'],
    ['2024-01-08', 'Cloud hosting',       'Technology',    '450.00'],
    ['2024-01-12', 'Team lunch',          'Food',          '89.50'],
    ['2024-01-15', 'Software license',    'Technology',    '299.00'],
    ['2024-01-18', 'Printer paper',       'Office',        '35.00'],
    ['2024-01-20', 'Coffee beans',        'Food',          '42.00'],
    ['2024-01-22', 'Domain renewal',      'Technology',    '15.00'],
    ['2024-01-25', 'Desk organizer',      'Office',        '28.50'],
    ['2024-01-28', 'Client dinner',       'Food',          '156.00'],
    ['2024-01-30', 'Antivirus renewal',   'Technology',    '79.99'],
  ];
}

/** Generate data for logic function demos (AND/OR/NOT/IFERROR) */
export function generateLogicData(): string[][] {
  return [
    ['Student',  'Math', 'Science', 'English', 'Attendance %'],
    ['Alice',    '92',   '88',      '95',      '98'],
    ['Bob',      '65',   '72',      '58',      '85'],
    ['Charlie',  '78',   '81',      '74',      '92'],
    ['Diana',    '45',   '52',      '48',      '70'],
    ['Eve',      '88',   '91',      '85',      '96'],
    ['Frank',    '55',   '60',      '62',      '78'],
  ];
}

/** Generate data for string function demos (LEFT/RIGHT/MID/SUBSTITUTE/FIND) */
export function generateStringData(): string[][] {
  return [
    ['Full Name',        'Email',                          'Phone'],
    ['Alice Johnson',    'alice.johnson@example.com',      '(555) 123-4567'],
    ['Bob Smith',        'bob.smith@example.com',          '(555) 987-6543'],
    ['Charlie Brown',    'charlie.brown@example.com',      '(555) 456-7890'],
    ['Diana Ross',       'diana.ross@example.com',         '(555) 321-0987'],
    ['Eve Williams',     'eve.williams@example.com',       '(555) 654-3210'],
  ];
}

/** Generate data for absolute reference demos ($A$1 vs A1) */
export function generateAbsoluteRefData(): string[][] {
  return [
    ['',       'Jan',   'Feb',   'Mar',   'Apr'],
    ['Sales',  '10000', '12000', '11500', '13000'],
    ['Costs',  '7000',  '7500',  '7200',  '7800'],
  ];
}
