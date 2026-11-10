/** Rules for sheet names, shared by the spreadsheet and its sheet tabs. */

/** Longest sheet name allowed (as in Excel). */
export const MAX_SHEET_NAME_LENGTH = 31;

/**
 * Check a name for a new or renamed sheet: 1 to 31 characters, none of
 * `: \ / ? * [ ]`, no apostrophe at either end, and not the name of another
 * sheet (ignoring case).
 *
 * @param name - The proposed name
 * @param others - Names of the other sheets
 * @returns What is wrong with the name, or `null` when it is fine
 */
export function validateSheetName(name: string, others: readonly string[]): string | null {
  if (name.trim() === '') return 'A sheet name cannot be empty';
  if (name !== name.trim()) return 'A sheet name cannot start or end with a space';
  if (name.length > MAX_SHEET_NAME_LENGTH) return `A sheet name cannot be longer than ${MAX_SHEET_NAME_LENGTH} characters`;
  if (/[:\\/?*[\]]/.test(name)) return 'A sheet name cannot contain : \\ / ? * [ or ]';
  if (name.startsWith("'") || name.endsWith("'")) return 'A sheet name cannot start or end with an apostrophe';
  const lower = name.toLowerCase();
  if (others.some((other) => other.toLowerCase() === lower)) return `There is already a sheet named "${name}"`;
  return null;
}

/** The first of `Sheet1`, `Sheet2`, ... that no sheet is named yet. */
export function nextSheetName(existing: readonly string[]): string {
  const taken = new Set(existing.map((n) => n.toLowerCase()));
  for (let i = existing.length + 1; ; i++) {
    const name = `Sheet${i}`;
    if (!taken.has(name.toLowerCase())) return name;
  }
}
