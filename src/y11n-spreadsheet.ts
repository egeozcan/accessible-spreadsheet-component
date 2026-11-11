import { LitElement, html, css, nothing, svg } from 'lit';
import { customElement, property, state, query } from 'lit/decorators.js';
import { repeat } from 'lit/directives/repeat.js';
import { styleMap } from 'lit/directives/style-map.js';
import {
  type CellCoord,
  type CellData,
  type CellFormat,
  type GridData,
  type SelectionRange,
  type FormulaFunction,
  type CellChangeDetail,
  type SelectionChangeDetail,
  type SelectionInfo,
  type HistoryChangeDetail,
  type DataChangeDetail,
  type FormatChangeDetail,
  type DataChangeOperation,
  type CellEditStartDetail,
  type ProposedChange,
  type BeforeChangeDetail,
  type CellValidator,
  type ValidationErrorDetail,
  type StructureChangeDetail,
  type SortDirection,
  type SortKey,
  type SortOptions,
  type SortDetail,
  type ColumnResizeDetail,
  type FreezeChangeDetail,
  type FilterColumn,
  type FilterInfo,
  type FilterChangeDetail,
  type NumberFormatOptions,
  type ConditionalFormatChangeDetail,
  type ConditionalFormatPreset,
  type ConditionalFormatResult,
  type ConditionalFormatRule,
  type ConditionalFormatRuleInput,
  type ConditionalFormatStyle,
  type CellArray,
  type CellArrayItem,
  type ListConversionDetail,
  type AddSheetOptions,
  type SheetChangeDetail,
  type SheetsChangeDetail,
  type DefinedName,
  type NamesChangeDetail,
  type FormulaFunctionOptions,
  cellKey,
  colToLetter,
  coordToRef,
  dateToSerial,
  formatNumber,
  formatsEqual,
  parseKey,
  refToCoord,
} from './types.js';
import { SelectionManager, findDataEdge } from './controllers/selection-manager.js';
import { parseDelimited, serializeDelimited } from './utils/delimited.js';
import {
  adjustFormulaForStructureChange,
  renameIsCaptured,
  renameNameInFormula,
  type StructureChange,
} from './utils/references.js';
import { compareSortValues, sortValueOf, type SortValue } from './utils/sort.js';
import { parseInput } from './utils/input-parse.js';
import { arrayToFormula, listItemsOf, listToFormula, listToText, looksLikeList, parseListText } from './utils/list.js';
import { formatSeriesNumber, seriesFor } from './utils/series.js';
import { FormulaEngine, isValidName, renameSheetInFormula } from './engine/formula-engine.js';
import { sheetPrefix } from './engine/parser.js';
import { nextSheetName, validateSheetName } from './utils/sheet-names.js';
import { ClipboardManager, computedLiteral } from './controllers/clipboard-manager.js';
import {
  ConditionalFormatEvaluator,
  adjustRulesForStructureChange,
  clearRulesInRange,
  cloneConditionalRule,
  conditionalFormatPreset,
  conditionalRulesEqual,
  describeCondition,
  normalizeConditionalRule,
  rangeToA1,
} from './utils/conditional-format.js';
import './components/y11n-formula-bar.js';
import type { FormulaBarMode, NameBoxCommitDetail, Y11nFormulaBar } from './components/y11n-formula-bar.js';
import './components/y11n-conditional-format-dialog.js';
import type {
  ConditionalFormatMoveDetail,
  ConditionalFormatSaveDetail,
  Y11nConditionalFormatDialog,
} from './components/y11n-conditional-format-dialog.js';
import { conditionalIcon, conditionalIconStyles } from './components/conditional-icons.js';
import './components/y11n-sheet-tabs.js';
import './components/y11n-name-manager.js';
import type { NameGoDetail, NameRemoveDetail, NameSaveDetail, Y11nNameManager } from './components/y11n-name-manager.js';

type ChangeSource = NonNullable<DataChangeDetail['source']>;

function cloneCell(cell: CellData): CellData {
  const copy: CellData = { ...cell };
  if (cell.format) {
    copy.format = { ...cell.format };
    if (cell.format.numberFormat) {
      copy.format.numberFormat = { ...cell.format.numberFormat };
    }
  }
  return copy;
}

/** Whether a cell shows something: a value of its own, or one spilled into it by another cell's formula. */
function hasContent(cell: CellData | undefined): cell is CellData {
  return cell !== undefined && (cell.rawValue !== '' || cell.spilledFrom !== undefined);
}

/**
 * The part of a cell that is stored rather than computed. A cell that only
 * shows a spilled value is nothing (or just its format): sorting, inserting
 * rows and replacing the data must not carry spilled values around, since the
 * formula that spills them recomputes them.
 */
function storedCell(cell: CellData | undefined): CellData | undefined {
  if (!cell || cell.spilledFrom === undefined || cell.rawValue !== '') return cell;
  if (!cell.format && !cell.readOnly) return undefined;
  const stored: CellData = { rawValue: '', displayValue: '', type: 'text' };
  if (cell.format) stored.format = cell.format;
  if (cell.readOnly) stored.readOnly = true;
  return stored;
}

/**
 * Lay out a sorted list of rendered indices across `total` slots, emitting
 * a spacer for every run of unrendered slots so the CSS grid keeps its geometry.
 */
function layoutWithGaps<T>(
  indices: number[],
  total: number,
  item: (index: number) => T,
  gap: (start: number, size: number) => T
): T[] {
  const out: T[] = [];
  let next = 0;
  for (const i of indices) {
    if (i > next) out.push(gap(next, i - next));
    out.push(item(i));
    next = i + 1;
  }
  if (total > next) out.push(gap(next, total - next));
  return out;
}

/** Index of the first element of a sorted array that is >= `value` (its length if none). */
function lowerBound(sorted: ArrayLike<number>, value: number): number {
  let lo = 0;
  let hi = sorted.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (sorted[mid] < value) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** Contiguous `[start, end)` indices plus any `extras` outside that window, sorted. */
function windowWithExtras(start: number, end: number, extras: number[]): number[] {
  const out: number[] = [];
  for (let i = start; i < end; i++) out.push(i);
  for (const e of extras) {
    if ((e < start || e >= end) && !out.includes(e)) out.push(e);
  }
  return out.sort((a, b) => a - b);
}

/** The text of one value of a cell that holds several, as shown in the values popup and the filter menu. */
function arrayItemText(item: CellArrayItem): string {
  if (item === null) return '';
  if (typeof item === 'object') return `{${arrayText(item)}}`;
  return typeof item === 'boolean' ? (item ? 'TRUE' : 'FALSE') : String(item);
}

function arrayText(array: CellArray): string {
  const rows: string[] = [];
  for (let r = 0; r < array.rows; r++) {
    rows.push(array.values.slice(r * array.cols, (r + 1) * array.cols).map(arrayItemText).join(', '));
  }
  return array.cols === 1 ? rows.join(', ') : rows.join('; ');
}

/** Every single value in an array, going into nested arrays. */
function arrayScalars(
  array: CellArray,
  out: Array<string | number | boolean | null> = []
): Array<string | number | boolean | null> {
  for (const item of array.values) {
    if (item !== null && typeof item === 'object') arrayScalars(item, out);
    else out.push(item);
  }
  return out;
}

/** "list of 3 items" / "array of 2 rows and 3 columns", for announcements. */
function describeArray(array: CellArray): string {
  if (array.cols === 1 || array.rows === 1) {
    const n = array.rows * array.cols;
    return `list of ${n} item${n === 1 ? '' : 's'}`;
  }
  return `array of ${array.rows} rows and ${array.cols} columns`;
}

/** Icons of a cell that holds several values: a list, and a grid for a 2D array. */
const LIST_ICON = svg`<svg viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><path d="M1 2.5h1.5M4.5 2.5H11M1 6h1.5M4.5 6H11M1 9.5h1.5M4.5 9.5H11" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>`;
const ARRAY_ICON = svg`<svg viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><path d="M1.5 1.5h9v9h-9zM1.5 6h9M6 1.5v9" fill="none" stroke="currentColor" stroke-width="1.2"/></svg>`;

/** Filter button icons: a dropdown chevron, and a funnel for a filtered column. */
const CHEVRON_ICON = svg`<svg viewBox="0 0 10 10" width="10" height="10" aria-hidden="true"><path d="M1.5 3.5 5 7l3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.5"/></svg>`;
const FUNNEL_ICON = svg`<svg viewBox="0 0 10 10" width="10" height="10" aria-hidden="true"><path d="M1 1.5h8L6 5.2V9L4 8V5.2z" fill="currentColor"/></svg>`;

const ARROW_DELTAS: Record<string, [number, number]> = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
};

/**
 * A cell write for `_buildCellBatch`. `format: undefined` leaves the cell's
 * format alone, `null` clears it, and an object replaces it.
 */
/** What a paste brings over: everything, only values (no formulas or formats), or only formats. */
export type PasteMode = 'all' | 'values' | 'formats';

/** Options for `paste()`. */
export interface PasteSpecialOptions {
  /** Default: 'all' */
  mode?: PasteMode;
  /** Swap rows and columns. Default: false */
  transpose?: boolean;
}

/** What the status bar shows for the selection (`_selectionSummary()`). */
interface SelectionSummary {
  count: number;
  /** Formatted sum of the number cells; undefined when there are none */
  sum?: string;
  average?: string;
}

/** Order cell keys in reading order (row, then column). */
function compareKeys(a: string, b: string): number {
  const pa = parseKey(a);
  const pb = parseKey(b);
  return pa.row - pb.row || pa.col - pb.col;
}

interface CellUpdate {
  id: string;
  value: string;
  format?: CellFormat | null;
}

/** Options for `toCSV()`. */
export interface CSVExportOptions {
  /** Range to export. Default: from A1 to the last row/column that has content. */
  range?: SelectionRange;
  /** Field separator. Default: `,` */
  delimiter?: string;
  /** Export computed display values (default) or raw values (formulas). */
  values?: 'display' | 'raw';
}

/** Options for `findCells()`. */
export interface FindOptions {
  /** Match upper/lower case exactly. Default: false */
  matchCase?: boolean;
  /** Search raw values (formula text) instead of displayed values. Default: false */
  formulas?: boolean;
}

/** Options for `fromCSV()`. */
export interface CSVImportOptions {
  /** Field separator. Default: `,` */
  delimiter?: string;
  /** Top-left cell to write to. Default: A1 */
  start?: CellCoord;
}

interface CellDelta {
  id: string;
  before: string;
  after: string;
  formatBefore?: CellFormat;
  formatAfter?: CellFormat;
}

interface SelectionSnapshot {
  anchor: CellCoord;
  head: CellCoord;
  /** Active-cell override (whole-row/column selections); null when it follows the head */
  active: CellCoord | null;
}

/** Whole-cell before/after state for a row/column insertion or deletion. */
interface StructureCellDelta {
  id: string;
  before?: CellData;
  after?: CellData;
}

interface StructureBatch {
  change: StructureChange;
  cells: StructureCellDelta[];
  /** Formulas on other sheets that refer to the changed sheet, rewritten with it */
  others?: Array<{ sheetId: number; id: string; before: string; after: string }>;
  /** Conditional-format rules of other sheets whose formulas refer to the changed sheet */
  otherRules?: Array<{
    sheetId: number;
    before: readonly ConditionalFormatRule[];
    after: readonly ConditionalFormatRule[];
  }>;
  /**
   * Defined names whose formula the change rewrites, with the formula before
   * and after and the sheet the name was based on. Undo and redo only touch a
   * name that still has that formula and that sheet.
   */
  names?: Array<{ name: string; sheet: string; before: string; after: string }>;
  rowsBefore: number;
  colsBefore: number;
  rowsAfter: number;
  colsAfter: number;
  /** Custom column widths before/after, when columns with custom widths move */
  colWidths?: { before: Map<number, number>; after: Map<number, number> };
}

/** A column width change; `null` is the default width (`--ls-cell-width`). */
interface WidthDelta {
  col: number;
  before: number | null;
  after: number | null;
}

/** Whole-cell before/after state for a sort, plus what was sorted. */
interface SortBatch {
  detail: Omit<SortDetail, 'source'>;
  cells: StructureCellDelta[];
}

/**
 * A filter: its range (the first row is the header), the displayed values
 * each filtered column keeps, and the rows it hides. Treated as immutable:
 * every change builds a new one, so undo can hold on to old states.
 */
interface FilterState {
  range: SelectionRange;
  /** Column -> displayed values to keep (`''` keeps blank cells) */
  criteria: Map<number, Set<string>>;
  /**
   * Rows hidden as of when the filter was last applied. Editing a cell
   * doesn't re-filter (as in Excel): rows only hide or reappear when the
   * filter is changed or reapplied.
   */
  hidden: Set<number>;
}

/** State of the open filter menu. */
interface FilterMenuState {
  col: number;
  /** The column's distinct displayed values, in ascending sort order; `''` is (Blanks) */
  values: string[];
  /** Values that are checked */
  checked: Set<string>;
  search: string;
  /** Index of the list item with the roving tab stop (0 = Select all) */
  focus: number;
}

/**
 * One sheet of the workbook. For the sheet being shown only `id` and `name`
 * are current: its data and view live in the component's own fields
 * (`_internalData`, `rows`, `_filter`, ...) and are stored here when another
 * sheet is shown.
 */
interface SheetState {
  id: number;
  name: string;
  data: GridData;
  rows: number;
  cols: number;
  frozenRows: number;
  frozenCols: number;
  colWidths: Map<number, number>;
  filter: FilterState | null;
  cfRules: readonly ConditionalFormatRule[];
  selection: SelectionSnapshot;
  scrollTop: number;
  scrollLeft: number;
}

function newSheetState(id: number, name: string, data: GridData, rows: number, cols: number): SheetState {
  const origin = { row: 0, col: 0 };
  return {
    id,
    name,
    data,
    rows,
    cols,
    frozenRows: 0,
    frozenCols: 0,
    colWidths: new Map(),
    filter: null,
    cfRules: [],
    selection: { anchor: origin, head: origin, active: null },
    scrollTop: 0,
    scrollLeft: 0,
  };
}

/** Apply `fn` to every formula in conditional-format rules; returns the same array when nothing changes. */
function mapRuleFormulas(
  rules: readonly ConditionalFormatRule[],
  fn: (formula: string) => string
): readonly ConditionalFormatRule[] {
  let changed = false;
  const next = rules.map((rule) => {
    const json = JSON.stringify(rule);
    if (!json.includes('"=')) return rule;
    const mapped = JSON.parse(json, (_key, value) =>
      typeof value === 'string' && value.startsWith('=') ? fn(value) : value
    ) as ConditionalFormatRule;
    if (JSON.stringify(mapped) === json) return rule;
    changed = true;
    return mapped;
  });
  return changed ? next : rules;
}

interface CommandBatch {
  /** The sheet the batch changes (set when it is recorded) */
  sheetId?: number;
  deltas: CellDelta[];
  selectionBefore: SelectionSnapshot;
  selectionAfter: SelectionSnapshot;
  operation: DataChangeOperation;
  /** Set for row/column insertions and deletions, which replace whole cells */
  structure?: StructureBatch;
  /** Set for sorts, which move whole cells */
  sort?: SortBatch;
  /** Set for column width changes, which touch no cells */
  widths?: WidthDelta[];
  /**
   * Filter before/after. On its own (no deltas, structure or sort), the
   * batch is a filter change; structure batches carry it when they move the filter.
   */
  filter?: { before: FilterState | null; after: FilterState | null };
  /**
   * Conditional-format rules before/after. On its own, the batch is a rule
   * change; structure batches carry it when they move or drop rules.
   */
  conditional?: { before: readonly ConditionalFormatRule[]; after: readonly ConditionalFormatRule[] };
}

const STRUCTURE_OPERATIONS: Record<string, DataChangeOperation> = {
  'insert:row': 'insert-rows',
  'delete:row': 'delete-rows',
  'insert:col': 'insert-columns',
  'delete:col': 'delete-columns',
};

/** Today's date as `YYYY-MM-DD`, which typed input reads back as a date (Ctrl+;). */
function todayText(): string {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/** Today's date as a date serial number, for conditional-format date rules. */
function todaySerial(): number {
  const d = new Date();
  return dateToSerial(d.getFullYear(), d.getMonth(), d.getDate());
}

/** The current time as `HH:mm`, which typed input reads back as a time (Ctrl+Shift+;). */
function currentTimeText(): string {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/**
 * `<y11n-spreadsheet>` -- An accessible spreadsheet web component.
 *
 * Implements the WAI-ARIA Grid pattern with:
 * - Roving tabindex and full keyboard navigation (arrows, Tab, Home/End, Ctrl+arrows)
 * - Recursive-descent formula engine (`=SUM(A1:B5)`, `=IF(...)`, etc.)
 * - Copy/cut/paste via the Clipboard API (HTML table + TSV formats)
 * - Undo/redo command history (up to 100 entries)
 * - Virtual rendering for large grids (only visible cells + buffer are in the DOM)
 * - Cell formatting (bold, italic, colors, number formats, etc.)
 *
 * @fires cell-change - When a single cell value is edited
 * @fires selection-change - When the selected range changes
 * @fires history-change - When undo or redo becomes available or stops being available
 * @fires data-change - On bulk updates (paste, cut, clear, undo/redo)
 * @fires format-change - When cell formatting is modified
 * @fires cell-edit-start - Before the inline editor opens (cancelable)
 * @fires before-change - Before a user edit, paste, cut, clear or fill is applied (cancelable)
 * @fires validation-error - When `validateCell` rejects one or more values
 * @fires before-structure-change - Before a user inserts or deletes rows/columns (cancelable)
 * @fires structure-change - After rows/columns are inserted or deleted
 * @fires before-sort - Before a user sorts a range (cancelable)
 * @fires sort-change - After a range is sorted (also on undo/redo)
 * @fires before-column-resize - Before a user resizes columns (cancelable)
 * @fires column-resize - After column widths change (also on undo/redo)
 * @fires freeze-change - After the frozen rows/columns change (not when the properties are set directly)
 * @fires before-conditional-format-change - Before a user changes the conditional-format rules (cancelable)
 * @fires conditional-format-change - After the conditional-format rules change (also on undo/redo)
 *
 * @csspart editor - The inline cell editor input element
 * @csspart validation-message - The message shown under the editor when a value is rejected
 */
@customElement('y11n-spreadsheet')
export class Y11nSpreadsheet extends LitElement {
  // ─── Public Properties ──────────────────────────────

  /** Number of rows in the grid. Default: 100 */
  @property({ type: Number }) rows = 100;
  /** Number of columns in the grid. Default: 26 */
  @property({ type: Number }) cols = 26;
  /** The sparse grid data map. Set externally or via `setData()`. */
  @property({ attribute: false }) data: GridData = new Map();
  /** When true, all editing is disabled. Default: false */
  @property({ type: Boolean, attribute: 'read-only', reflect: true }) readOnly = false;
  /** Accessible label for the grid. Default: 'Spreadsheet' */
  @property({ type: String, attribute: 'grid-label' }) gridLabel = 'Spreadsheet';
  /** Record of custom formula functions to register, keyed by name. */
  @property({ attribute: false }) functions: Record<string, FormulaFunction> = {};
  /**
   * Optional validator for values the user enters (typing, formula bar,
   * paste, fill). Return an error message to reject a value, or null/undefined
   * to accept it. Clearing a cell is never validated.
   */
  @property({ attribute: false }) validateCell?: CellValidator;
  /**
   * Number of rows at the top that stay in view while scrolling. At least
   * one row always scrolls. Default: 0
   */
  @property({ type: Number, attribute: 'frozen-rows', reflect: true }) frozenRows = 0;
  /**
   * Number of columns at the left that stay in view while scrolling. At
   * least one column always scrolls. Default: 0
   */
  @property({ type: Number, attribute: 'frozen-cols', reflect: true }) frozenCols = 0;
  /**
   * When true, typed values are stored exactly as entered. By default, as in
   * Excel, input like `15%`, `$1,200`, `1,234.5`, `2026-01-05` or `14:30` is
   * stored as a number and the cell gets the matching number format.
   */
  @property({ type: Boolean, attribute: 'raw-input' }) rawInput = false;
  /**
   * When true, the sheet tabs are shown even with a single sheet, so users
   * can add sheets. With more than one sheet the tabs always show.
   */
  @property({ type: Boolean, attribute: 'sheet-tabs' }) sheetTabs = false;
  /** When true, the status bar under the grid (Sum, Average, Count of the selection) is hidden. */
  @property({ type: Boolean, attribute: 'hide-status-bar' }) hideStatusBar = false;

  // ─── Internal State ─────────────────────────────────

  @state() private _isEditing = false;
  @state() private _editValue = '';
  @state() private _formulaBarMode: FormulaBarMode = 'raw';
  /** When true, the next render should reposition the editor cursor (set on edit start / ref mode) */
  private _editorNeedsCursorReset = false;

  /** Whether arrow keys are currently selecting a cell reference in a formula */
  @state() private _refMode = false;

  /** Screen reader announcement text for mode changes */
  @state() private _modeAnnouncement = '';

  /** Message from `validateCell` for the value in the editor, shown below it */
  @state() private _validationMessage = '';

  /** Find bar state */
  @state() private _findOpen = false;
  @state() private _findQuery = '';
  @state() private _findStatus = '';

  /** Cells a fill-handle drag would fill, for the drag preview */
  @state() private _fillPreview: SelectionRange | null = null;

  /** The open filter menu, if any */
  @state() private _filterMenu: FilterMenuState | null = null;
  /** The open Paste Special dialog's choices, or null when it is closed */
  @state() private _pasteSpecial: Required<PasteSpecialOptions> | null = null;
  /** Whether the name manager is open */
  @state() private _nameManagerOpen = false;
  /** Outcome of the last change made in the name manager, announced inside it */
  @state() private _nameManagerStatus = '';
  private _nameManagerCleanup: (() => void) | null = null;
  /** Whether the conditional-format rule manager is open */
  @state() private _cfManagerOpen = false;
  /** Outcome of the last change made in the rule manager, announced inside it */
  @state() private _cfManagerStatus = '';

  /** The open values popup of a cell that holds several values: its cell, and the path into nested arrays */
  @state() private _arrayCard: { key: string; path: number[] } | null = null;
  private _arrayCardCleanup: (() => void) | null = null;
  /** Whether the text in the editor is a list's items (so committing it writes a list) */
  private _editListMode = false;
  /**
   * The sheet of the cell being edited. While a formula is edited, another
   * sheet can be shown to pick a reference on it; committing or cancelling
   * shows this sheet again.
   */
  private _editSheetId = 1;
  /** Set while a reference is being picked with the pointer, so the editor losing focus doesn't commit */
  private _refPickPointer = false;
  /** The ranges of the defined names that refer to cells, for the Name Box; null until needed */
  private _nameRanges: Array<{ name: string; sheet: string; start: CellCoord; end: CellCoord }> | null = null;
  /** The spill range the active cell is part of, outlined while it is active */
  private _activeSpill: SelectionRange | null = null;

  /** Track scroll position for virtual rendering */
  @state() private _scrollTop = 0;
  @state() private _scrollLeft = 0;

  // ─── DOM Refs ───────────────────────────────────────

  @query('#editor') private _editor!: HTMLInputElement;
  @query('#find-input') private _findInput?: HTMLInputElement;
  @query('.ls-grid') private _grid!: HTMLDivElement;

  // ─── Controllers / Managers ─────────────────────────

  private _selection = new SelectionManager(this, this.rows, this.cols, (row) => this._isRowHidden(row));
  private _formulaEngine = new FormulaEngine();
  private _clipboardManager = new ClipboardManager();

  /** Internal data store - we keep a separate mutable copy */
  private _internalData: GridData = new Map();
  private _lastSyncedDataRef: GridData | null = null;

  /** The sheets, in tab order (see `SheetState`) */
  private _sheets: SheetState[] = [newSheetState(1, 'Sheet1', new Map(), 100, 26)];
  private _activeSheetId = 1;
  private _nextSheetId = 2;

  /** Track which cell started editing for commit logic */
  private _editingCellKey: string | null = null;

  private _undoStack: CommandBatch[] = [];
  private _redoStack: CommandBatch[] = [];
  /** What the last `history-change` reported: undo and redo available */
  private _historyReported = '00:0';
  /** What had the focus, outside the grid, when a dialog was opened (see `_returnFocus`) */
  private _dialogOpener: HTMLElement | null = null;
  private readonly _maxHistory = 100;

  /** Reference mode cursor position */
  private _refCursorRow = 0;
  private _refCursorCol = 0;

  /** Start/end positions of the reference text within the formula string */
  private _refInsertStart = 0;
  private _refInsertEnd = 0;

  /** Whether an async paste operation is currently in progress */
  private _pasteInProgress = false;

  /**
   * Bumped whenever rows/columns are inserted or deleted (including undo/redo),
   * so async clipboard work can tell that its cell coordinates went stale.
   */
  private _structureRevision = 0;

  /**
   * Bumped when the data is replaced (`setData()` / `data`), which clears the
   * undo history. Operations that started before then must not push history
   * entries or keep writing (e.g. a progressive fill between frames).
   */
  private _dataGeneration = 0;

  /**
   * Bumped by every data change. Cancelable events compare it before and after
   * dispatch to detect a listener that changed the grid reentrantly.
   */
  private _mutationCount = 0;
  /** `_selectionSummary()` result for the selection and data it was computed for */
  private _summaryCache: { key: string; value: SelectionSummary | null } | null = null;

  /** Number of cells to process per animation frame during large pastes */
  private static readonly PASTE_CHUNK_SIZE = 500;
  private static readonly INITIAL_VIEW_ROWS = 15;
  private static readonly INITIAL_VIEW_COLS = 8;
  private static readonly ROW_HEADER_WIDTH = 50;

  /**
   * Rendered cell size in px, measured from `--ls-cell-width` /
   * `--ls-cell-height` (see `_measureCellSize()`). Defaults match the CSS
   * fallbacks until the first measurement.
   */
  private _cellSize = { width: 100, height: 28 };

  /** Column widths set by the user or API, in px; other columns use `_cellSize.width` */
  private _colWidths = new Map<number, number>();
  /** Live widths while a header border is dragged, before the resize is committed */
  private _resizePreview: { cols: number[]; width: number } | null = null;
  /** Bumped whenever a column width (or preview) changes, to invalidate `_colLayout` */
  private _colWidthsRevision = 0;
  /** Cached left offsets of every column (`cols + 1` entries), see `_colOffsets()` */
  private _colLayout: { key: string; offsets: Float64Array } | null = null;
  private _measureCanvas: CanvasRenderingContext2D | null | undefined;

  /** The filter, if any (see `FilterState`) */
  private _filter: FilterState | null = null;
  /** Bumped whenever the filter changes, to invalidate `_rowLayout` and stale async work */
  private _filterRevision = 0;
  /** Cached ascending list of the rows not hidden by the filter, see `_visibleRowList()` */
  private _rowLayout: { key: string; visible: Int32Array | null } | null = null;
  /** Removes the filter menu's outside-click listener */
  private _filterMenuCleanup: (() => void) | null = null;
  private _pasteSpecialCleanup: (() => void) | null = null;

  /**
   * Conditional-format rules in priority order. Treated as immutable like
   * `_filter`: every change builds a new array, so undo can hold on to old ones.
   */
  private _cfRules: readonly ConditionalFormatRule[] = [];
  /** Bumped whenever the rules change, to detect a reentrant change during `before-conditional-format-change` */
  private _cfRevision = 0;
  private _cfNextId = 1;
  private _cfManagerCleanup: (() => void) | null = null;
  /** Works out what the rules give each cell, reading values the way sorting does */
  private _cf = new ConditionalFormatEvaluator({
    valueAt: (row, col) => this._cfValue(cellKey(row, col)),
    forEachValue: (range, visit) => {
      const r1 = Math.min(range.end.row, this.rows - 1);
      const c1 = Math.min(range.end.col, this.cols - 1);
      const area = Math.max(0, r1 - range.start.row + 1) * Math.max(0, c1 - range.start.col + 1);
      if (area <= this._internalData.size) {
        for (let r = range.start.row; r <= r1; r++) {
          for (let c = range.start.col; c <= c1; c++) {
            const key = cellKey(r, c);
            if (this._internalData.has(key)) visit(this._cfValue(key));
          }
        }
        return;
      }
      // A range much larger than the data (whole columns): walk the data instead
      for (const key of this._internalData.keys()) {
        const { row, col } = parseKey(key);
        if (row >= range.start.row && row <= r1 && col >= range.start.col && col <= c1) visit(this._cfValue(key));
      }
    },
    evaluate: (formula, at) => {
      // Against what the grid shows: precedents are read as computed, not computed again
      const result = this._formulaEngine.evaluateCommitted(formula, undefined, at);
      switch (result.type) {
        case 'number':
          return { kind: 'number', value: Number(result.displayValue) };
        case 'boolean':
          return { kind: 'boolean', value: result.displayValue === 'TRUE' };
        case 'error':
          return { kind: 'error' };
        default:
          return result.displayValue === '' ? { kind: 'blank' } : { kind: 'text', value: result.displayValue };
      }
    },
    today: todaySerial,
  });
  /** Most rows / columns the values popup shows of one array */
  private static readonly ARRAY_CARD_MAX_ROWS = 200;
  private static readonly ARRAY_CARD_MAX_COLS = 20;
  /** Most values the filter menu lists at once; the search box narrows the rest */
  private static readonly FILTER_MENU_MAX_VALUES = 500;

  private static readonly MIN_COL_WIDTH = 20;
  private static readonly MAX_COL_WIDTH = 2000;
  /** Alt+Shift+Left/Right step, in px */
  private static readonly COL_RESIZE_STEP = 8;
  /** The `_cellSize` object the last render used */
  private _renderedCellSize = this._cellSize;
  private _cellSizeObserver: ResizeObserver | null = null;

  // ─── Lifecycle ──────────────────────────────────────

  connectedCallback(): void {
    super.connectedCallback();
    // Re-attach after a move in the DOM; the first attach happens in firstUpdated
    if (this.hasUpdated) this._observeCellSize();
  }

  disconnectedCallback(): void {
    super.disconnectedCallback();
    this._cellSizeObserver?.disconnect();
    this._closeFilterMenu(false);
    this._closePasteSpecial(false);
    this.closeConditionalFormatManager(false);
    this.closeCellValues(false);
    this.closeNameManager(false);
  }

  protected firstUpdated(): void {
    // Measure synchronously so API calls right after the first render
    // (scrolling, editor placement) already use the real size. The re-render
    // with it is left to the ResizeObserver's initial callback: requesting an
    // update from firstUpdated trips Lit's change-in-update warning.
    this._measureCellSize();
    if (!this._observeCellSize() && this._cellSize !== this._renderedCellSize) this.requestUpdate();
  }

  willUpdate(changedProps: Map<string, unknown>): void {
    let resized = false;
    if (changedProps.has('rows') || changedProps.has('cols')) {
      this.rows = Math.max(1, Math.round(this.rows));
      this.cols = Math.max(1, Math.round(this.cols));
      this._selection.setBounds(this.rows, this.cols);
      resized = this._syncEngineSize();
    }
    if (changedProps.has('data') && this.data !== this._lastSyncedDataRef) {
      this._syncData();
    }
    if (resized) {
      // The size bounds spills and whole-column references, also for formulas on other sheets
      this._recalcAll();
    }
    if (changedProps.has('readOnly') && this.readOnly && this._isEditing) {
      // The grid became read-only mid-edit: discard the edit
      this._exitRefMode();
      this._cancelEdit();
    }
    if (changedProps.has('functions')) {
      // On first render the old value is undefined and this.functions is the
      // default empty object.  Skip the setFunctions call in that case so we
      // don't wipe out functions added imperatively via registerFunction().
      const prev = changedProps.get('functions') as Record<string, FormulaFunction> | undefined;
      if (prev !== undefined || Object.keys(this.functions).length > 0) {
        this._formulaEngine.setFunctions(this.functions);
        this._recalcAll();
      }
    }
  }

  // ─── Public API Methods ─────────────────────────────

  /**
   * Returns a shallow copy of the current grid data.
   *
   * @returns A new Map containing all cell data
   */
  getData(): GridData {
    const copy: GridData = new Map();
    for (const [key, cell] of this._internalData) {
      copy.set(key, cloneCell(cell));
    }
    return copy;
  }

  /**
   * Replaces the entire grid data and resets undo/redo history.
   *
   * @param data - The new sparse grid data map
   */
  setData(data: GridData): void {
    this.data = data;
    this._syncData();
    this.requestUpdate();
  }

  /**
   * Register a custom formula function at runtime.
   * Triggers a full recalculation of all formulas.
   *
   * @param name - Function name (will be uppercased for lookup)
   * @param fn - The function implementation
   */
  registerFunction(name: string, fn: FormulaFunction, options?: FormulaFunctionOptions): void {
    this._formulaEngine.registerFunction(name, fn, options);
    this._recalcAll();
  }

  /**
   * Serialize grid contents as CSV (or TSV etc. via `delimiter`).
   * By default exports display values from A1 to the last row and column
   * that contain a value.
   *
   * @param options - Range, delimiter, and whether to export display or raw values
   * @returns The delimited text (rows joined with `\n`); empty string for an empty grid
   */
  toCSV(options: CSVExportOptions = {}): string {
    const { delimiter = ',', values = 'display' } = options;
    let range = options.range;
    if (!range) {
      let maxRow = -1;
      let maxCol = -1;
      for (const [key, cell] of this._internalData) {
        if (!hasContent(cell)) continue;
        const { row, col } = parseKey(key);
        if (row > maxRow) maxRow = row;
        if (col > maxCol) maxCol = col;
      }
      if (maxRow < 0) return '';
      range = { start: { row: 0, col: 0 }, end: { row: maxRow, col: maxCol } };
    }

    const rows: string[][] = [];
    for (let r = range.start.row; r <= range.end.row; r++) {
      const fields: string[] = [];
      for (let c = range.start.col; c <= range.end.col; c++) {
        const cell = this._internalData.get(cellKey(r, c));
        fields.push((values === 'raw' ? cell?.rawValue : cell?.displayValue) ?? '');
      }
      rows.push(fields);
    }
    return serializeDelimited(rows, delimiter);
  }

  /**
   * Write CSV (or TSV etc. via `delimiter`) text into the grid starting at
   * `start` (default A1). Values beginning with `=` become formulas. Values
   * outside the grid's `rows`/`cols` are dropped, so grow the grid first if
   * needed. The import is a single undoable step and fires `data-change`
   * with `operation: 'import'`.
   *
   * @param text - Delimited text
   * @param options - Delimiter and top-left target cell
   */
  fromCSV(text: string, options: CSVImportOptions = {}): void {
    if (this._pasteInProgress) return;
    this._syncPendingData();
    const { delimiter = ',', start = { row: 0, col: 0 } } = options;
    if (!Number.isInteger(start.row) || !Number.isInteger(start.col) || start.row < 0 || start.col < 0) {
      throw new RangeError('fromCSV: start must be a cell with non-negative integer row and col');
    }
    const updates: CellUpdate[] = [];
    parseDelimited(text, delimiter).forEach((fields, r) => {
      fields.forEach((value, c) => {
        const row = start.row + r;
        const col = start.col + c;
        if (row < this.rows && col < this.cols) {
          updates.push({ id: cellKey(row, col), value });
        }
      });
    });

    const selection = this._snapshotSelection();
    const batch = this._buildCellBatch(updates, 'import', selection);
    if (!batch) return;
    this._commitBatch(batch, 'programmatic');
  }

  // ─── Rows / Columns API ──────────────────────────────

  /**
   * Insert `count` empty rows before row `index` (0-based; `index === rows`
   * appends). Cells below move down, the grid grows by `count` rows, and
   * formula references are rewritten to follow the cells they point to.
   * Undoable; fires `data-change` and `structure-change`.
   *
   * @param index - Row to insert before
   * @param count - Number of rows to insert (default 1)
   */
  insertRows(index: number, count = 1): void {
    this._changeStructure({ axis: 'row', kind: 'insert', index, count }, 'programmatic');
  }

  /**
   * Delete `count` rows starting at row `index` (0-based). Cells below move
   * up and the grid shrinks; references to deleted cells become `#REF!`.
   * At least one row must remain. Undoable; fires `data-change` and
   * `structure-change`.
   *
   * @param index - First row to delete
   * @param count - Number of rows to delete (default 1)
   */
  deleteRows(index: number, count = 1): void {
    this._changeStructure({ axis: 'row', kind: 'delete', index, count }, 'programmatic');
  }

  /**
   * Insert `count` empty columns before column `index` (0-based;
   * `index === cols` appends). See `insertRows()`.
   *
   * @param index - Column to insert before
   * @param count - Number of columns to insert (default 1)
   */
  insertColumns(index: number, count = 1): void {
    this._changeStructure({ axis: 'col', kind: 'insert', index, count }, 'programmatic');
  }

  /**
   * Delete `count` columns starting at column `index` (0-based). See
   * `deleteRows()`.
   *
   * @param index - First column to delete
   * @param count - Number of columns to delete (default 1)
   */
  deleteColumns(index: number, count = 1): void {
    this._changeStructure({ axis: 'col', kind: 'delete', index, count }, 'programmatic');
  }

  // ─── Sort API ────────────────────────────────────────

  /**
   * Sort the rows of `range` by one or more key columns. Only cells inside
   * the range move; with `hasHeader`, its first row stays in place.
   *
   * Ascending order is numbers (smallest first), text (A to Z, ignoring
   * case), FALSE, TRUE, then errors; descending reverses it. Blank cells go
   * last either way. The sort is stable, and later keys break ties of
   * earlier ones. Cells move with their formats and read-only flags.
   * Formulas in moved cells shift their relative references by the rows they
   * moved, as if copied; references from outside the range are not changed.
   *
   * One undoable step; fires `data-change` and `sort-change` (not
   * `before-sort`). Throws a RangeError for a range outside the grid or a key
   * column outside the range.
   *
   * @param range - The cells to sort
   * @param keys - Sort key(s) in priority order
   * @param options - Whether the first row is a header
   * @returns Whether any cell moved
   */
  sort(range: SelectionRange, keys: SortKey | SortKey[], options: SortOptions = {}): boolean {
    const keyList = Array.isArray(keys) ? keys : [keys];
    const normalized = this._normalizeRange(range, 'sort');
    if (keyList.length === 0) throw new RangeError('sort: at least one key is required');
    const resolved = keyList.map((key) => {
      if (!key || !Number.isInteger(key.col) || key.col < normalized.start.col || key.col > normalized.end.col) {
        throw new RangeError(`sort: key column ${key?.col} is outside the range`);
      }
      const direction = key.direction ?? 'ascending';
      if (direction !== 'ascending' && direction !== 'descending') {
        throw new RangeError(`sort: direction must be 'ascending' or 'descending'`);
      }
      return { col: key.col, direction };
    });
    return this._sortRange(
      { range: normalized, keys: resolved, hasHeader: options.hasHeader === true },
      'programmatic'
    );
  }

  /**
   * Sort as the user would with Alt+Shift+Up / Alt+Shift+Down: by the
   * active cell's column, over the selection, or over the block of data
   * around the active cell when a single cell or a single column of a wider
   * table is selected. A first row of
   * text headers over other data is detected and kept in place. Respects
   * `readOnly`, refuses to move read-only cells, and fires the cancelable
   * `before-sort`. The outcome is announced to screen readers. Use this to
   * back a "Sort A to Z" button.
   *
   * @param direction - Default: `'ascending'`
   * @returns Whether any cell moved
   */
  sortSelection(direction: SortDirection = 'ascending'): boolean {
    if (this.readOnly || this._pasteInProgress) return false;
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();

    const active = this._selection.activeCell;
    const selected = this._selection.range;
    const single = selected.start.row === selected.end.row && selected.start.col === selected.end.col;
    const filterRange = this._filter?.range;
    if (
      single &&
      filterRange &&
      active.row >= filterRange.start.row &&
      active.row <= filterRange.end.row &&
      active.col >= filterRange.start.col &&
      active.col <= filterRange.end.col
    ) {
      // Inside a filter: sort its table, below its header row
      return this._sortRange({ range: filterRange, keys: [{ col: active.col, direction }], hasHeader: true }, 'user');
    }
    let range = single ? this._currentRegion(active.row, active.col) : this._clipToData(selected, active);
    // Sorting one column of a wider table would split its records apart, so
    // sort the whole table instead (Excel's default when it asks)
    if (range && range.start.col === range.end.col) {
      // Grow from the column's first filled cell: the active cell may be below the data
      let row = range.start.row;
      while (row < range.end.row && !this._isCellFilled(row, range.start.col)) row++;
      const region = this._currentRegion(row, range.start.col);
      if (region && region.start.col < region.end.col) range = region;
    }
    if (!range) {
      this._modeAnnouncement = 'Nothing to sort';
      return false;
    }
    return this._sortRange(
      { range, keys: [{ col: active.col, direction }], hasHeader: this._looksLikeHeader(range) },
      'user'
    );
  }

  // ─── Column Widths API ───────────────────────────────

  /**
   * The width of a column in px: its custom width, or the default from
   * `--ls-cell-width`.
   *
   * @param col - 0-based column index
   */
  getColumnWidth(col: number): number {
    return this._colWidth(col);
  }

  /**
   * The custom column widths, keyed by 0-based column index. Columns at the
   * default width are left out. Pass the result to `setColumnWidths()` to
   * restore a saved layout.
   */
  getColumnWidths(): Record<number, number> {
    const out: Record<number, number> = {};
    for (const [col, width] of [...this._colWidths].sort((a, b) => a[0] - b[0])) {
      if (col < this.cols) out[col] = width;
    }
    return out;
  }

  /**
   * Set the width of a column in px, or `null` for the default width.
   * Widths are rounded and clamped to 20–2000px. Undoable; fires
   * `column-resize`.
   *
   * @param col - 0-based column index
   * @param width - Width in px, or null to reset
   */
  setColumnWidth(col: number, width: number | null): void {
    this.setColumnWidths({ [col]: width });
  }

  /**
   * Set several column widths at once (one undo step). Keys are 0-based
   * column indices; `null` resets a column to the default width. Throws a
   * RangeError for a column outside the grid or a width that isn't a
   * positive number.
   *
   * @param widths - Widths in px keyed by column index
   */
  setColumnWidths(widths: Record<number, number | null>): void {
    const changes = Object.entries(widths).map(([key, width]) => {
      const col = Number(key);
      if (!Number.isInteger(col) || col < 0 || col >= this.cols) {
        throw new RangeError(`setColumnWidths: column ${key} is outside the grid (${this.cols} columns)`);
      }
      if (width !== null && !(typeof width === 'number' && Number.isFinite(width) && width > 0)) {
        throw new RangeError(`setColumnWidths: the width of column ${key} must be a positive number or null`);
      }
      return { col, width };
    });
    this._resizeColumns(changes, 'programmatic');
  }

  /**
   * Fit columns to their widest value, as a double-click on a header border
   * does. Empty columns go back to the default width. Undoable; fires
   * `column-resize`.
   *
   * @param cols - 0-based column indices
   */
  autoFitColumns(cols: number[]): void {
    for (const col of cols) {
      if (!Number.isInteger(col) || col < 0 || col >= this.cols) {
        throw new RangeError(`autoFitColumns: column ${col} is outside the grid (${this.cols} columns)`);
      }
    }
    this._resizeColumns(cols.map((col) => ({ col, width: this._fitWidth(col) })), 'programmatic');
  }

  // ─── Frozen Panes API ────────────────────────────────

  /**
   * Freeze the top `rows` rows and left `cols` columns (default 0) so they
   * stay in view while the rest scrolls. Without arguments, freezes the rows
   * above and the columns left of the active cell, as Alt+Shift+F does (at
   * A1 that unfreezes). Values are
   * clamped so at least one row and column scroll. Not undoable (a view
   * setting); fires `freeze-change`.
   *
   * @param rows - Number of rows to freeze
   * @param cols - Number of columns to freeze
   */
  freezePanes(rows?: number, cols?: number): void {
    const active = this._selection.activeCell;
    const r = rows ?? active.row;
    // freezePanes(1) freezes one row, not also the columns left of the active cell
    const c = cols ?? (rows === undefined ? active.col : 0);
    if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0) {
      throw new RangeError('freezePanes: rows and cols must be non-negative integers');
    }
    this._setFrozen(r, c, 'programmatic');
  }

  /** Unfreeze all rows and columns. Fires `freeze-change`. */
  unfreezePanes(): void {
    this._setFrozen(0, 0, 'programmatic');
  }

  /**
   * Freeze panes as a user action, to back a button: like `freezePanes()`,
   * and without arguments what Alt+Shift+F freezes (the rows above and the
   * columns left of the active cell), but refused when the frozen area would
   * not leave room to scroll, announced to screen readers, and reported by
   * `freeze-change` with `source: 'user'`. `freezePanesAsUser(0, 0)`
   * unfreezes.
   *
   * @returns Whether the frozen panes changed
   */
  freezePanesAsUser(rows?: number, cols?: number): boolean {
    if (this._endEditElsewhere()) return false;
    const active = this._selection.activeCell;
    const r = rows ?? active.row;
    const c = cols ?? (rows === undefined ? active.col : 0);
    if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0) {
      throw new RangeError('freezePanesAsUser: rows and cols must be non-negative integers');
    }
    return this._freezeAsUser(r, c);
  }

  // ─── Filter API ──────────────────────────────────────

  /**
   * Filter `range`: its first row is the header row (never hidden), and each
   * row below it stays visible only if, in every column of `columns`, it
   * displays one of that column's `values` (`''` matches blank cells).
   * Values are compared as displayed, number formats included. Replaces any
   * existing filter; with no `columns`, adds filter buttons without hiding
   * rows. Rows don't re-filter as cells are edited: call `reapplyFilter()`.
   *
   * One undoable step; fires `filter-change` (not `before-filter`). Throws a
   * RangeError for a range outside the grid or a column outside the range.
   *
   * @param range - The table to filter, header row first
   * @param columns - Criteria per column
   */
  setFilter(range: SelectionRange, columns: FilterColumn[] = []): void {
    const normalized = this._normalizeRange(range, 'setFilter');
    const criteria = new Map<number, Set<string>>();
    for (const column of columns) {
      if (!column || !Number.isInteger(column.col) || column.col < normalized.start.col || column.col > normalized.end.col) {
        throw new RangeError(`setFilter: column ${column?.col} is outside the range`);
      }
      criteria.set(column.col, new Set(Array.from(column.values ?? [], String)));
    }
    this._syncPendingData();
    this._commitFilter(this._evaluateFilter(normalized, criteria), 'programmatic');
  }

  /**
   * Set (or, with `null`, clear) one column's criteria in the current filter
   * and re-filter the rows. See `setFilter()`. Throws if there is no filter
   * or `col` is outside its range.
   *
   * @param col - 0-based column inside the filter range
   * @param values - Displayed values to keep (`''` keeps blanks), or null for all
   */
  setColumnFilter(col: number, values: string[] | null): void {
    const filter = this._filter;
    if (!filter) throw new RangeError('setColumnFilter: there is no filter; call setFilter() first');
    if (!Number.isInteger(col) || col < filter.range.start.col || col > filter.range.end.col) {
      throw new RangeError(`setColumnFilter: column ${col} is outside the filter range`);
    }
    const criteria = new Map(filter.criteria);
    if (values === null) criteria.delete(col);
    else criteria.set(col, new Set(Array.from(values, String)));
    this._syncPendingData();
    this._commitFilter(this._evaluateFilter(filter.range, criteria), 'programmatic');
  }

  /** Re-filter the rows with the current criteria (after cells were edited). Undoable. */
  reapplyFilter(): void {
    const filter = this._filter;
    if (!filter) return;
    this._syncPendingData();
    this._commitFilter(this._evaluateFilter(filter.range, filter.criteria), 'programmatic');
  }

  /** Remove the filter, showing every row. Undoable; fires `filter-change`. */
  removeFilter(): void {
    this._commitFilter(null, 'programmatic');
  }

  /** The current filter, or null. */
  getFilter(): FilterInfo | null {
    return this._filter ? this._filterInfo(this._filter) : null;
  }

  /**
   * Whether the filter hides a row.
   *
   * @param row - 0-based row index
   */
  isRowHidden(row: number): boolean {
    return this._isRowHidden(row);
  }

  /**
   * Reapply the filter as a user action, as Ctrl+Alt+L does, to back a
   * button: unlike `reapplyFilter()` it takes in rows added under the table,
   * fires the cancelable `before-filter`, and announces how many rows are
   * shown. Allowed in read-only mode.
   */
  reapplyFilterAsUser(): void {
    if (this._endEditElsewhere()) return;
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();
    this._reapplyUserFilter();
  }

  /**
   * Add or remove a filter as the user would with Ctrl+Shift+L: removes the
   * filter if there is one, otherwise filters the selection, or the block of
   * data around the active cell when a single cell is selected, using its
   * first row as the header. Allowed in read-only mode, fires the cancelable
   * `before-filter`, and announces the outcome. Use this to back a "Filter" button.
   *
   * @returns Whether the filter was added or removed
   */
  toggleFilter(): boolean {
    if (this._pasteInProgress) return false;
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();
    if (this._filter) {
      return this._announceFilterCommit(this._commitFilter(null, 'user'), 'Filter removed');
    }
    const active = this._selection.activeCell;
    const selected = this._selection.range;
    const single = selected.start.row === selected.end.row && selected.start.col === selected.end.col;
    const range = single ? this._currentRegion(active.row, active.col) : this._clipToData(selected, active);
    if (!range) {
      this._modeAnnouncement = 'Nothing to filter: select a table whose first row has its headers';
      return false;
    }
    const ref = `${coordToRef(range.start)}:${coordToRef(range.end)}`;
    return this._announceFilterCommit(
      this._commitFilter(this._evaluateFilter(range, new Map()), 'user'),
      `Filter added to ${ref}. Press Alt+Down on a header cell for filter options`
    );
  }

  /**
   * Open the filter menu of a column (default: the active cell's), as
   * Alt+Down on a filter header cell does, moving the active cell to that
   * header and focusing the menu's search box.
   *
   * @param col - 0-based column inside the filter range
   * @returns Whether the menu opened (false without a filter or outside its columns)
   */
  async openFilterMenu(col?: number): Promise<boolean> {
    const filter = this._filter;
    if (!filter) return false;
    const c = col ?? this._selection.activeCell.col;
    if (c < filter.range.start.col || c > filter.range.end.col) return false;
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();
    const headerRow = filter.range.start.row;
    const active = this._selection.activeCell;
    if (active.row !== headerRow || active.col !== c || this._selection.range.end.row !== headerRow) {
      this._selection.moveTo(headerRow, c);
      this._dispatchSelectionChange();
    }
    this._ensureCellVisible(headerRow, c);
    this.closeConditionalFormatManager(false);
    const values = this._filterMenuValues(c);
    this._filterMenu = {
      col: c,
      values,
      checked: this._checkedForCriteria(values, filter.criteria.get(c)),
      search: '',
      focus: 0,
    };
    if (!this._filterMenuCleanup) {
      // Clicking anywhere outside the menu closes it, like a native dropdown.
      // Bubble phase, so the menu's own filter button can toggle it closed.
      const onPointerDown = (e: PointerEvent) => {
        const menuEl = this.renderRoot.querySelector('.ls-filter-menu');
        if (menuEl && e.composedPath().includes(menuEl)) return;
        this._closeFilterMenu(false);
      };
      document.addEventListener('pointerdown', onPointerDown);
      this._filterMenuCleanup = () => document.removeEventListener('pointerdown', onPointerDown);
    }
    this._modeAnnouncement = '';
    await this.updateComplete;
    this.renderRoot.querySelector<HTMLInputElement>('#filter-search')?.focus();
    return true;
  }

  // ─── Sheets API ──────────────────────────────────────

  /** Names of the sheets, in tab order. */
  get sheets(): string[] {
    return this._sheets.map((s) => s.name);
  }

  /**
   * Name of the sheet being shown. `data`, `rows`, `cols`, the selection,
   * the filter and every method that takes cell keys work on this sheet.
   * Assign a sheet's name to show it.
   */
  get activeSheet(): string {
    return this._activeState.name;
  }

  set activeSheet(name: string) {
    this.activateSheet(name);
  }

  /**
   * Show another sheet. Each sheet keeps its own data, size, selection,
   * scroll position, column widths, frozen panes, filter and conditional formats.
   *
   * @returns False when it couldn't be shown right now (a paste is in progress)
   * @throws {RangeError} If there is no sheet of that name
   */
  activateSheet(name: string): boolean {
    return this._activateSheet(this._requireSheet(name, 'activateSheet'), 'programmatic');
  }

  /**
   * Add a sheet. Formulas refer to its cells as `Name!A1` (`'My sheet'!A1`
   * when the name has spaces). Not undoable. With more than one sheet the
   * sheet tabs appear under the grid.
   *
   * @param name - Default: the next free `Sheet2`, `Sheet3`, ...
   * @param options - Data, size, position, and whether to show the sheet
   * @returns The sheet's name
   * @throws {RangeError} If the name is not a valid, unused sheet name
   */
  addSheet(name?: string, options: AddSheetOptions = {}): string {
    this._syncPendingData();
    const sheetName = name ?? nextSheetName(this.sheets);
    const problem = validateSheetName(sheetName, this.sheets);
    if (problem) throw new RangeError(`addSheet: ${problem}`);
    return this._addSheet(sheetName, options, 'programmatic');
  }

  /**
   * Remove a sheet and its data. Not undoable; changes made on that sheet
   * leave the undo history. Formulas that refer to the sheet keep their text
   * and show `#REF!` (until a sheet of that name exists again).
   *
   * @returns False when it couldn't be removed right now (a paste is in progress)
   * @throws {RangeError} If there is no sheet of that name, or it is the only sheet
   */
  removeSheet(name: string): boolean {
    const state = this._requireSheet(name, 'removeSheet');
    if (this._sheets.length === 1) throw new RangeError('removeSheet: a workbook needs at least one sheet');
    return this._removeSheet(state, 'programmatic');
  }

  /**
   * Rename a sheet. Formulas, defined names and conditional formats that
   * refer to it are rewritten, on every sheet. Not undoable.
   *
   * @throws {RangeError} If there is no sheet of that name, or the new name is invalid or taken
   */
  renameSheet(name: string, newName: string): void {
    const state = this._requireSheet(name, 'renameSheet');
    const problem = validateSheetName(newName, this.sheets.filter((n) => n !== state.name));
    if (problem) throw new RangeError(`renameSheet: ${problem}`);
    this._renameSheet(state, newName, 'programmatic');
  }

  /**
   * Move a sheet to another position among the tabs (0 = first).
   *
   * @throws {RangeError} If there is no sheet of that name
   */
  moveSheet(name: string, index: number): void {
    this._moveSheet(this._requireSheet(name, 'moveSheet'), index, 'programmatic');
  }

  /**
   * A copy of a sheet's data, like `getData()` for the sheet being shown.
   *
   * @throws {RangeError} If there is no sheet of that name
   */
  getSheetData(name: string): GridData {
    const state = this._requireSheet(name, 'getSheetData');
    if (state.id === this._activeSheetId) return this.getData();
    const copy: GridData = new Map();
    for (const [key, cell] of state.data) copy.set(key, cloneCell(cell));
    return copy;
  }

  /**
   * Replace a sheet's data, like `setData()` for the sheet being shown.
   * Changes made on that sheet leave the undo history.
   *
   * @throws {RangeError} If there is no sheet of that name
   */
  setSheetData(name: string, data: GridData): void {
    const state = this._requireSheet(name, 'setSheetData');
    if (state.id === this._activeSheetId) {
      this.setData(data);
      return;
    }
    this._syncPendingData();
    state.data = this._prepareSheetData(data);
    this._dropHistoryOf(state.id);
    this._formulaEngine.setData(state.data, state.name);
    this._recalcAll();
    this.requestUpdate();
  }

  // ─── Names API ───────────────────────────────────────

  /**
   * Define a name that formulas can use: for a range (`=A1:A10`,
   * `=Sheet2!B2`), a constant (`=0.19`) or a function (`=LAMBDA(x, x * 2)`,
   * called as `Name(3)`). Names are workbook-wide and not case-sensitive;
   * defining an existing name replaces it. The references a name holds move
   * when rows or columns are inserted or deleted. Not undoable.
   *
   * @param name - Letters, digits, `_` and `.`, not starting with a digit, and not a cell reference like `A1`
   * @param refersTo - A formula, with or without the leading `=`
   * @param sheet - Sheet that references without a sheet name point to. Default: the sheet being shown
   * @throws {RangeError} If the name is invalid or there is no such sheet
   */
  defineName(name: string, refersTo: string, sheet?: string): void {
    this._syncPendingData();
    if (!isValidName(name)) throw new RangeError(`defineName: "${name}" is not a valid name`);
    const scope = sheet === undefined ? this._activeState : this._requireSheet(sheet, 'defineName');
    this._formulaEngine.defineName(name, refersTo, scope.name);
    this._namesChanged('programmatic');
  }

  /**
   * Remove a defined name. Formulas that use it show `#NAME?`.
   *
   * @returns Whether the name existed
   */
  removeName(name: string): boolean {
    this._syncPendingData();
    if (!this._formulaEngine.removeName(name)) return false;
    this._namesChanged('programmatic');
    return true;
  }

  /** The defined names. */
  getNames(): DefinedName[] {
    return this._formulaEngine.getNames().map((n) => ({ name: n.name, refersTo: n.formula, sheet: n.sheet }));
  }

  /**
   * Replace all defined names (e.g. when loading a saved workbook):
   * `setNames(getNames())` changes nothing. A name without `sheet` is based
   * on the sheet being shown. A name whose `sheet` is not a sheet of the
   * workbook (`getNames()` gives `''` for a name whose sheet was removed)
   * keeps having none: it still works as a constant or with sheet-qualified
   * references, and its references without a sheet name are `#REF!`.
   *
   * @throws {RangeError} If a name is invalid
   */
  setNames(names: Array<{ name: string; refersTo: string; sheet?: string }>): void {
    this._syncPendingData();
    for (const n of names) {
      if (!isValidName(n.name)) throw new RangeError(`setNames: "${n.name}" is not a valid name`);
    }
    // A name based on a sheet that doesn't exist (as getNames() reports one whose
    // sheet was removed) stays that way; only a name without a sheet gets this one
    this._formulaEngine.setNames(
      names.map((n) => ({ name: n.name, formula: n.refersTo, sheet: n.sheet ?? this._activeState.name }))
    );
    this._namesChanged('programmatic');
  }

  /**
   * Select the cells a reference or a defined name stands for, showing
   * their sheet: `goTo('B5')`, `goTo('A1:C3')`, `goTo('Orders!A2')`,
   * `goTo('sales')`. What typing into the Name Box and pressing Enter does.
   *
   * @returns False when the text is not a reference or a name for cells, or the cells are outside the grid
   */
  goTo(reference: string): boolean {
    this._syncPendingData();
    const target = this._formulaEngine.referenceOf(reference, this._activeState.name);
    return !!target && this._goToRange(target, 'programmatic');
  }

  /**
   * Open the name manager (what Ctrl+F3 does): a dialog that lists the
   * defined names and lets users add, edit, delete and go to them. Users'
   * changes fire the cancelable `before-names-change`; in read-only mode the
   * names can only be looked at.
   */
  async openNameManager(): Promise<boolean> {
    this._syncPendingData();
    if (this._isEditing) {
      this._exitRefMode();
      this._commitEdit();
    }
    this._closeFilterMenu(false);
    this._closePasteSpecial(false);
    this.closeConditionalFormatManager(false);
    this.closeCellValues(false);
    this._dialogOpener = this._focusedOutside();
    this._nameManagerStatus = '';
    this._nameManagerOpen = true;
    if (!this._nameManagerCleanup) {
      const onPointerDown = (e: PointerEvent) => {
        const el = this.renderRoot.querySelector('y11n-name-manager');
        if (el && e.composedPath().includes(el)) return;
        this.closeNameManager(false);
      };
      document.addEventListener('pointerdown', onPointerDown);
      this._nameManagerCleanup = () => document.removeEventListener('pointerdown', onPointerDown);
    }
    this._modeAnnouncement = '';
    await this.updateComplete;
    const dialog = this.renderRoot.querySelector<Y11nNameManager>('y11n-name-manager');
    await dialog?.updateComplete;
    dialog?.focusInitial();
    return true;
  }

  /**
   * Close the name manager.
   *
   * @param focusCell - Return focus to the active cell. Default: true
   */
  closeNameManager(focusCell = true): void {
    this._nameManagerCleanup?.();
    this._nameManagerCleanup = null;
    if (!this._nameManagerOpen) return;
    this._nameManagerOpen = false;
    if (focusCell) this._returnFocus();
  }

  /** Move the keyboard focus to the Name Box (what Ctrl+G does). */
  focusNameBox(): void {
    this.shadowRoot?.querySelector<Y11nFormulaBar>('y11n-formula-bar')?.focusNameBox();
  }

  // ─── Name Internals ─────────────────────────────────

  private _goToRange(
    target: { sheet: string; start: CellCoord; end: CellCoord },
    source: 'user' | 'programmatic'
  ): boolean {
    const state = this._findSheet(target.sheet);
    if (!state) return false;
    if (this._isEditing) {
      this._exitRefMode();
      this._commitEdit();
    }
    const rows = state.id === this._activeSheetId ? this.rows : state.rows;
    const cols = state.id === this._activeSheetId ? this.cols : state.cols;
    if (target.start.row >= rows || target.start.col >= cols) return false;
    if (!this._activateSheet(state, source)) return false;
    const clamp = (c: CellCoord): CellCoord => ({ row: Math.min(c.row, rows - 1), col: Math.min(c.col, cols - 1) });
    const start = clamp(target.start);
    // The first cell of the range becomes the active one
    this._restoreSelection({ anchor: clamp(target.end), head: start, active: null });
    this._ensureCellVisible(this._selection.activeCell.row, start.col);
    this.requestUpdate();
    return true;
  }

  /** The current selection as an absolute, sheet-qualified reference: `=Sheet1!$A$1:$B$3`. */
  private _selectionReference(): string {
    const { start, end } = this._selection.range;
    const abs = (c: CellCoord) => `$${colToLetter(c.col)}$${c.row + 1}`;
    const cells = start.row === end.row && start.col === end.col ? abs(start) : `${abs(start)}:${abs(end)}`;
    return `=${sheetPrefix(this._activeState.name)}${cells}`;
  }

  /** The defined name whose cells are exactly the selection, for the Name Box. */
  private _selectionName(): string {
    if (!this._nameRanges) {
      this._nameRanges = [];
      for (const n of this._formulaEngine.getNames()) {
        const range = this._formulaEngine.referenceOf(n.name);
        if (range) this._nameRanges.push({ name: n.name, ...range });
      }
    }
    if (this._nameRanges.length === 0) return '';
    const { start, end } = this._selection.range;
    const sheet = this._activeState.name;
    const match = this._nameRanges.find(
      (r) =>
        r.sheet === sheet &&
        r.start.row === start.row &&
        r.start.col === start.col &&
        r.end.row === end.row &&
        r.end.col === end.col
    );
    return match?.name ?? '';
  }

  private _nameListCache: { key: string; list: string[] } | null = null;

  /** The defined names in alphabetical order, for the Name Box's suggestions. */
  private _nameList(): string[] {
    const names = this._formulaEngine.getNames().map((n) => n.name);
    const key = names.join('\n');
    if (this._nameListCache?.key !== key) {
      const list = names.sort((x, y) => x.localeCompare(y, undefined, { sensitivity: 'base' }));
      this._nameListCache = { key, list };
    }
    return this._nameListCache.list;
  }

  /**
   * Change the defined names for a user: refused in read-only mode, and
   * `before-names-change` (describing the result) can veto it.
   */
  private _changeNamesByUser(next: DefinedName[], apply: () => void): boolean {
    if (this.readOnly) {
      this._modeAnnouncement = 'The spreadsheet is read-only: names cannot be changed';
      return false;
    }
    const event = new CustomEvent<NamesChangeDetail>('before-names-change', {
      detail: { names: next, source: 'user' },
      bubbles: true,
      composed: true,
      cancelable: true,
    });
    if (!this.dispatchEvent(event)) return false;
    apply();
    this._namesChanged('user');
    return true;
  }

  /** Show an outcome in the name manager. The same text twice in a row is announced again. */
  private _setNameStatus(text: string): void {
    if (this._nameManagerStatus !== text) {
      this._nameManagerStatus = text;
      return;
    }
    this._nameManagerStatus = '';
    this.updateComplete.then(() => {
      this._nameManagerStatus = text;
    });
  }

  private _handleNameSave(e: CustomEvent<NameSaveDetail>): void {
    const { name, refersTo, previous } = e.detail;
    const scope = this._findSheet(e.detail.sheet) ?? this._activeState;
    const same = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();
    const renamed = previous !== undefined && !same(previous, name);
    if (renamed) {
      const conflict = this._renameConflict(previous!, name);
      if (conflict) {
        this._setNameStatus(`${previous} was not renamed to ${name}: ${conflict} would no longer mean the same`);
        return;
      }
    }
    const next = this.getNames().filter((n) => !same(n.name, previous ?? name) && !same(n.name, name));
    next.push({ name, refersTo, sheet: scope.name });
    const applied = this._changeNamesByUser(next, () => {
      if (previous !== undefined) this._formulaEngine.removeName(previous);
      // Formulas that use the name keep working under its new one, its own definition included (a recursive LAMBDA)
      const definition = renamed ? this._renameNameEverywhere(previous!, name, refersTo) : refersTo;
      this._formulaEngine.defineName(name, definition, scope.name);
    });
    this._setNameStatus(
      !applied ? 'The change was not applied' : previous === undefined ? `Name added: ${name}` : `Name changed: ${name}`
    );
  }

  private _handleNameRemove(e: CustomEvent<NameRemoveDetail>): void {
    const { name } = e.detail;
    const next = this.getNames().filter((n) => n.name.toLowerCase() !== name.toLowerCase());
    const applied = this._changeNamesByUser(next, () => this._formulaEngine.removeName(name));
    this._setNameStatus(applied ? `Name deleted: ${name}` : 'The change was not applied');
  }

  private _handleNameGo(e: CustomEvent<NameGoDetail>): void {
    const { name } = e.detail;
    const target = this._formulaEngine.referenceOf(name, this._activeState.name);
    if (!target) {
      this._setNameStatus(`${name} does not refer to cells`);
      return;
    }
    this.closeNameManager(false);
    if (!this._goToRange(target, 'user')) this._modeAnnouncement = `${name} is outside the sheet`;
  }

  /**
   * What stands in the way of renaming a name: a formula that uses it and
   * has a LET or LAMBDA variable of the old or the new name, or that calls it
   * when the new name is a function, so rewriting it would change what it means. Looks at every formula the rename would
   * rewrite: cells of all sheets, other names, conditional formats, and what
   * undo would write back. Undefined when the rename is safe.
   */
  private _renameConflict(oldName: string, newName: string): string | undefined {
    const functions = new Set(this._formulaEngine.functionNames());
    const captured = (formula: string) => renameIsCaptured(formula, oldName, newName, (fn) => functions.has(fn));
    for (const sheet of this._sheets) {
      const data = sheet.id === this._activeSheetId ? this._internalData : sheet.data;
      for (const [key, cell] of data) {
        if (captured(cell.rawValue)) return `the formula in ${sheet.name}!${coordToRef(parseKey(key))}`;
      }
    }
    const other = this._formulaEngine.getNames().find((n) => captured(n.formula));
    if (other) return `the name ${other.name}`;
    let stored: string | undefined;
    // Visits without changing anything
    this._rewriteStoredFormulas((formula) => {
      if (stored === undefined && captured(formula)) stored = formula;
      return formula;
    });
    return stored === undefined ? undefined : `${stored} (a conditional format, or a formula undo would bring back)`;
  }

  /**
   * Rewrite a renamed name in the formulas of every sheet, in the other
   * names, and in what undo would write back. Returns `definition` (the
   * name's own formula) rewritten the same way.
   */
  private _renameNameEverywhere(oldName: string, newName: string, definition: string): string {
    const functions = new Set(this._formulaEngine.functionNames());
    const rewrite = (formula: string) => renameNameInFormula(formula, oldName, newName, (fn) => functions.has(fn));
    for (const sheet of this._sheets) {
      const data = sheet.id === this._activeSheetId ? this._internalData : sheet.data;
      for (const cell of data.values()) {
        if (cell.rawValue.startsWith('=')) cell.rawValue = rewrite(cell.rawValue);
      }
    }
    this._formulaEngine.mapNameFormulas(rewrite);
    this._rewriteStoredFormulas(rewrite);
    // Undo of a row or column change finds the names it moved by name
    const lower = oldName.toLowerCase();
    for (const batch of [...this._undoStack, ...this._redoStack]) {
      for (const moved of batch.structure?.names ?? []) {
        if (moved.name.toLowerCase() === lower) moved.name = newName;
      }
    }
    return rewrite(definition.startsWith('=') ? definition : '=' + definition);
  }

  /** Enter in the Name Box: go to the cells typed, or name the selection. */
  private _handleNameBoxCommit(e: CustomEvent<NameBoxCommitDetail>): void {
    this._syncPendingData();
    const bar = e.currentTarget as Y11nFormulaBar;
    const value = e.detail.value;
    const target = this._formulaEngine.referenceOf(value, this._activeState.name);
    if (target) {
      if (!this._goToRange(target, 'user')) this._modeAnnouncement = `${value} is outside the sheet`;
      bar.resetNameBox();
      return;
    }
    const names = this.getNames();
    const existing = names.find((n) => n.name.toLowerCase() === value.toLowerCase());
    if (existing) {
      this._modeAnnouncement = `${existing.name} does not refer to cells: it is ${existing.refersTo}`;
      bar.resetNameBox();
      return;
    }
    if (!isValidName(value)) {
      this._modeAnnouncement = `${value} is not a cell reference or a valid name`;
      bar.resetNameBox();
      return;
    }
    const sheet = this._activeState.name;
    const refersTo = this._selectionReference();
    const defined = this._changeNamesByUser([...names, { name: value, refersTo, sheet }], () =>
      this._formulaEngine.defineName(value, refersTo, sheet)
    );
    bar.resetNameBox();
    if (!defined) return;
    this._modeAnnouncement = `Named ${refersTo.slice(1)} ${value}`;
    this.updateComplete.then(() => this._focusActiveCell());
  }

  // ─── Sheet Internals ────────────────────────────────

  /** The sheet being shown. */
  private get _activeState(): SheetState {
    return this._sheets.find((s) => s.id === this._activeSheetId)!;
  }

  private _findSheet(name: string): SheetState | undefined {
    const lower = name.toLowerCase();
    return this._sheets.find((s) => s.name.toLowerCase() === lower);
  }

  private _requireSheet(name: string, method: string): SheetState {
    const state = this._findSheet(name);
    if (!state) throw new RangeError(`${method}: there is no sheet named "${name}"`);
    return state;
  }

  /** Cells for a sheet that is not being shown: copies, with literal values displayed and formulas left to the engine. */
  private _prepareSheetData(data: GridData | undefined): GridData {
    const out: GridData = new Map();
    if (!data) return out;
    for (const [key, cell] of data) {
      const stored = storedCell(cell);
      if (!stored) continue;
      const copy = cloneCell(stored);
      delete copy.spill;
      delete copy.array;
      if (copy.rawValue.startsWith('=')) {
        // Computed (and number-formatted) by the next recalculation
        copy.displayValue = '';
        copy.type = 'text';
      } else {
        const evaluated = this._formulaEngine.evaluate(copy.rawValue);
        copy.type = evaluated.type;
        copy.displayValue =
          evaluated.type === 'number' && copy.format?.numberFormat
            ? formatNumber(Number(copy.rawValue), copy.format.numberFormat)
            : evaluated.displayValue;
      }
      out.set(key, copy);
    }
    return out;
  }

  /** Keep the view of the sheet being shown, before another one replaces it. */
  private _stashSheet(state: SheetState): void {
    state.data = this._internalData;
    state.rows = this.rows;
    state.cols = this.cols;
    state.frozenRows = this.frozenRows;
    state.frozenCols = this.frozenCols;
    state.colWidths = this._colWidths;
    state.filter = this._filter;
    state.cfRules = this._cfRules;
    state.selection = this._snapshotSelection();
    state.scrollTop = this._scrollTop;
    state.scrollLeft = this._scrollLeft;
  }

  private _loadSheet(state: SheetState): void {
    this._internalData = state.data;
    this.rows = state.rows;
    this.cols = state.cols;
    this.frozenRows = state.frozenRows;
    this.frozenCols = state.frozenCols;
    this._colWidths = state.colWidths;
    this._colWidthsRevision++;
    this._resizePreview = null;
    this._filter = state.filter;
    this._filterRevision++;
    this._formulaEngine.setActiveSheet(state.name);
    this._syncEngineSize();
    this._setConditionalRules(state.cfRules);
    this._cf.invalidate();
    this._selection.setBounds(this.rows, this.cols);
    this._selection.setSelection(state.selection.anchor, state.selection.head, state.selection.active);
    this._fillPreview = null;
    this._summaryCache = null;
    // Work that is waiting on the clipboard must not land on another sheet
    this._structureRevision++;
    this._mutationCount++;
    this._scrollTop = state.scrollTop;
    this._scrollLeft = state.scrollLeft;
  }

  private _activateSheet(state: SheetState, source: SheetChangeDetail['source']): boolean {
    if (state.id === this._activeSheetId) return true;
    if (this._pasteInProgress) return false;
    this._syncPendingData();
    if (this._isEditing) {
      if (source === 'user' && this._editValue.startsWith('=')) {
        // A formula stays open while another sheet is shown, so a reference can be picked on it.
        // A reference already being picked stays replaceable: the next pick takes its place.
        this._switchSheetView(state, source);
        const edited = this._sheets.find((s) => s.id === this._editSheetId);
        const cell = this._editingCellKey ? coordToRef(parseKey(this._editingCellKey)) : '';
        this._modeAnnouncement =
          state.id === this._editSheetId || !edited
            ? `${state.name}, still editing ${cell}`
            : `${state.name}. Pick a cell for the formula in ${edited.name}!${cell}; Enter finishes`;
        return true;
      }
      this._exitRefMode();
      this._commitEdit();
      // Committing shows the edited cell's sheet again, which may be this one
      if (state.id === this._activeSheetId) return true;
    }
    this._closeFilterMenu(false);
    this._closePasteSpecial(false);
    this.closeConditionalFormatManager(false);
    this.closeCellValues(false);
    this._switchSheetView(state, source);
    if (source !== 'programmatic') {
      this._modeAnnouncement = `${state.name}, sheet ${this._sheets.indexOf(state) + 1} of ${this._sheets.length}`;
    }
    return true;
  }

  /** Swap the view to another sheet: store this sheet's view, load the other's, tell listeners. */
  private _switchSheetView(state: SheetState, source: SheetChangeDetail['source']): void {
    const hadFocus = this._gridHasFocus();
    const previous = this._activeState;
    // A size set just before must reach the engine while this sheet is still the active one
    if (this._syncEngineSize()) this._recalcAll();
    this._stashSheet(previous);
    this._activeSheetId = state.id;
    this._loadSheet(state);
    // Text typed into the formula bar or the Name Box was for the sheet that was shown
    const bar = this.shadowRoot?.querySelector<Y11nFormulaBar>('y11n-formula-bar');
    bar?.resetDraft();
    bar?.resetNameBox();
    this.dispatchEvent(
      new CustomEvent<SheetChangeDetail>('sheet-change', {
        detail: { sheet: state.name, previous: previous.name, source },
        bubbles: true,
        composed: true,
      })
    );
    this._dispatchSelectionChange();
    this.requestUpdate();
    this.updateComplete.then(() => {
      if (this._activeSheetId !== state.id) return;
      const grid = this._grid;
      if (grid) {
        grid.scrollTop = state.scrollTop;
        grid.scrollLeft = state.scrollLeft;
      }
      if (hadFocus && !this._isEditing) this._focusActiveCell();
    });
  }

  /** Show the sheet a recorded change belongs to. False when that sheet no longer exists. */
  private _showSheetOf(batch: CommandBatch, source: 'undo' | 'redo'): boolean {
    if (batch.sheetId === undefined || batch.sheetId === this._activeSheetId) return true;
    const state = this._sheets.find((s) => s.id === batch.sheetId);
    return !!state && this._activateSheet(state, source);
  }

  private _stepSheet(delta: number): void {
    const count = this._sheets.length;
    const index = this._sheets.indexOf(this._activeState);
    const next = this._sheets[(index + delta + count) % count];
    if (this._activateSheet(next, 'user')) {
      this.updateComplete.then(() => {
        if (!this._isEditing) this._focusActiveCell();
      });
    }
  }

  /** Adding, removing or renaming sheets ends an edit first (and shows its sheet again). */
  private _endEditForSheetChange(): void {
    if (!this._isEditing) return;
    this._exitRefMode();
    this._commitEdit();
  }

  private _dropHistoryOf(sheetId: number): void {
    this._undoStack = this._undoStack.filter((b) => b.sheetId !== sheetId);
    this._redoStack = this._redoStack.filter((b) => b.sheetId !== sheetId);
    this.requestUpdate();
  }

  private _addSheet(name: string, options: AddSheetOptions, source: 'user' | 'programmatic'): string {
    const rows = Math.max(1, Math.round(options.rows ?? this.rows));
    const cols = Math.max(1, Math.round(options.cols ?? this.cols));
    const state = newSheetState(this._nextSheetId++, name, this._prepareSheetData(options.data), rows, cols);
    const index = Math.max(0, Math.min(Math.trunc(options.index ?? this._sheets.length), this._sheets.length));
    const sheets = [...this._sheets];
    sheets.splice(index, 0, state);
    this._sheets = sheets;
    this._formulaEngine.addSheet(name, state.data, rows, cols);
    this._formulaEngine.moveSheet(name, index);
    // Formulas that already named this sheet now resolve
    this._recalcAll();
    this._dispatchSheetsChange('add', name, source);
    if (options.activate) this._activateSheet(state, source);
    this.requestUpdate();
    return name;
  }

  private _removeSheet(state: SheetState, source: 'user' | 'programmatic'): boolean {
    if (this._pasteInProgress || this._sheets.length === 1) return false;
    this._syncPendingData();
    this._endEditForSheetChange();
    const index = this._sheets.indexOf(state);
    if (state.id === this._activeSheetId) {
      const next = this._sheets[index + 1] ?? this._sheets[index - 1];
      if (!this._activateSheet(next, source)) return false;
    }
    const namesBefore = JSON.stringify(this.getNames());
    this._sheets = this._sheets.filter((s) => s !== state);
    this._formulaEngine.removeSheet(state.name);
    this._dropHistoryOf(state.id);
    this._recalcAll();
    this._dispatchSheetsChange('remove', state.name, source);
    // Names that were based on the sheet no longer have one
    if (JSON.stringify(this.getNames()) !== namesBefore) this._dispatchNamesChange(source);
    this.requestUpdate();
    return true;
  }

  private _renameSheet(state: SheetState, newName: string, source: 'user' | 'programmatic'): void {
    const oldName = state.name;
    if (oldName === newName) return;
    this._syncPendingData();
    if (this._isEditing) {
      this._exitRefMode();
      this._commitEdit();
    }
    const namesBefore = JSON.stringify(this.getNames());
    // Rewrites the formulas of every sheet, and the defined names
    this._formulaEngine.renameSheet(oldName, newName);
    state.name = newName;

    this._rewriteStoredFormulas((formula) => renameSheetInFormula(formula, oldName, newName));
    for (const batch of [...this._undoStack, ...this._redoStack]) {
      for (const moved of batch.structure?.names ?? []) {
        if (moved.sheet === oldName) moved.sheet = newName;
      }
    }
    this._recalcAll();
    this._dispatchSheetsChange('rename', newName, source, oldName);
    // Names based on the sheet, or referring to it, changed with it
    if (JSON.stringify(this.getNames()) !== namesBefore) this._dispatchNamesChange(source);
    this.requestUpdate();
  }

  /**
   * Apply a rewrite to the formulas kept outside the cells: what undo and
   * redo would write back, and conditional-format rules of every sheet.
   */
  private _rewriteStoredFormulas(rewrite: (formula: string) => string): void {
    const rewriteCell = (cell: CellData | undefined) => {
      if (cell) cell.rawValue = rewrite(cell.rawValue);
    };
    // Undo knows a sheet's rules by identity, and one array can be in several
    // batches: every array is rewritten once, and everything pointing at it follows
    const rewritten = new Map<readonly ConditionalFormatRule[], readonly ConditionalFormatRule[]>();
    const rewriteRules = (rules: readonly ConditionalFormatRule[]) => {
      let next = rewritten.get(rules);
      if (!next) rewritten.set(rules, (next = mapRuleFormulas(rules, rewrite)));
      return next;
    };
    for (const batch of [...this._undoStack, ...this._redoStack]) {
      for (const delta of batch.deltas) {
        delta.before = rewrite(delta.before);
        delta.after = rewrite(delta.after);
      }
      for (const cell of [...(batch.structure?.cells ?? []), ...(batch.sort?.cells ?? [])]) {
        rewriteCell(cell.before);
        rewriteCell(cell.after);
      }
      for (const other of batch.structure?.others ?? []) {
        other.before = rewrite(other.before);
        other.after = rewrite(other.after);
      }
      for (const moved of batch.structure?.names ?? []) {
        moved.before = rewrite(moved.before);
        moved.after = rewrite(moved.after);
      }
      for (const rules of batch.structure?.otherRules ?? []) {
        rules.before = rewriteRules(rules.before);
        rules.after = rewriteRules(rules.after);
      }
      if (batch.conditional) {
        batch.conditional = {
          before: rewriteRules(batch.conditional.before),
          after: rewriteRules(batch.conditional.after),
        };
      }
    }
    for (const sheet of this._sheets) {
      if (sheet.id === this._activeSheetId) {
        const rules = rewriteRules(this._cfRules);
        if (rules !== this._cfRules) this._setConditionalRules(rules);
      } else {
        sheet.cfRules = rewriteRules(sheet.cfRules);
      }
    }
  }

  private _moveSheet(state: SheetState, index: number, source: 'user' | 'programmatic'): void {
    const from = this._sheets.indexOf(state);
    const to = Math.max(0, Math.min(Math.trunc(index) || 0, this._sheets.length - 1));
    if (from === to) return;
    const sheets = this._sheets.filter((s) => s !== state);
    sheets.splice(to, 0, state);
    this._sheets = sheets;
    this._formulaEngine.moveSheet(state.name, to);
    this._dispatchSheetsChange('move', state.name, source);
    this.requestUpdate();
  }

  private _dispatchSheetsChange(
    kind: SheetsChangeDetail['kind'],
    sheet: string,
    source: 'user' | 'programmatic',
    previous?: string
  ): void {
    const detail: SheetsChangeDetail = { sheets: this.sheets, kind, sheet, source };
    if (previous !== undefined) detail.previous = previous;
    this.dispatchEvent(new CustomEvent<SheetsChangeDetail>('sheets-change', { detail, bubbles: true, composed: true }));
  }

  /**
   * Ask `before-sheets-change` listeners whether a user may change the
   * sheets as described. Returns false when cancelled.
   */
  private _confirmSheetsChange(detail: SheetsChangeDetail): boolean {
    const event = new CustomEvent<SheetsChangeDetail>('before-sheets-change', {
      detail,
      bubbles: true,
      composed: true,
      cancelable: true,
    });
    return this.dispatchEvent(event);
  }

  private _namesChanged(source: NamesChangeDetail['source']): void {
    this._recalcAll();
    this._dispatchNamesChange(source);
    this.requestUpdate();
  }

  private _dispatchNamesChange(source: NamesChangeDetail['source']): void {
    this.dispatchEvent(
      new CustomEvent<NamesChangeDetail>('names-change', {
        detail: { names: this.getNames(), source },
        bubbles: true,
        composed: true,
      })
    );
  }

  /**
   * Add to a structure batch what a row/column change on this sheet does
   * elsewhere: formulas on other sheets that refer to it, and defined names.
   */
  private _addCrossSheetStructure(structure: StructureBatch): void {
    const { change } = structure;
    const changedSheet = this._activeState.name;
    const others: NonNullable<StructureBatch['others']> = [];
    for (const sheet of this._sheets) {
      if (sheet.id === this._activeSheetId) continue;
      for (const [id, cell] of sheet.data) {
        // Only a sheet-qualified reference can point here
        if (!cell.rawValue.startsWith('=') || !cell.rawValue.includes('!')) continue;
        const after = adjustFormulaForStructureChange(cell.rawValue, change, { formulaSheet: sheet.name, changedSheet });
        if (after !== cell.rawValue) others.push({ sheetId: sheet.id, id, before: cell.rawValue, after });
      }
    }
    if (others.length > 0) structure.others = others;

    const otherRules: NonNullable<StructureBatch['otherRules']> = [];
    for (const sheet of this._sheets) {
      if (sheet.id === this._activeSheetId) continue;
      const after = mapRuleFormulas(sheet.cfRules, (formula) =>
        adjustFormulaForStructureChange(formula, change, { formulaSheet: sheet.name, changedSheet })
      );
      if (after !== sheet.cfRules) otherRules.push({ sheetId: sheet.id, before: sheet.cfRules, after });
    }
    if (otherRules.length > 0) structure.otherRules = otherRules;

    const names: NonNullable<StructureBatch['names']> = [];
    for (const n of this._formulaEngine.getNames()) {
      const after = adjustFormulaForStructureChange(n.formula, change, { formulaSheet: n.sheet, changedSheet });
      if (after !== n.formula) names.push({ name: n.name, sheet: n.sheet, before: n.formula, after });
    }
    if (names.length > 0) structure.names = names;
  }

  // ─── Sheet tab events ───────────────────────────────

  private _handleSheetSelect(e: CustomEvent<{ id: number }>): void {
    const state = this._sheets.find((s) => s.id === e.detail.id);
    if (state) this._activateSheet(state, 'user');
  }

  private _handleSheetAdd(): void {
    if (this.readOnly) return;
    this._syncPendingData();
    this._endEditForSheetChange();
    const name = nextSheetName(this.sheets);
    if (!this._confirmSheetsChange({ sheets: [...this.sheets, name], kind: 'add', sheet: name, source: 'user' })) return;
    this._addSheet(name, { activate: true }, 'user');
    this._modeAnnouncement = `Added ${name}`;
  }

  private _handleSheetRename(e: CustomEvent<{ id: number; name: string }>): void {
    if (this.readOnly) return;
    const state = this._sheets.find((s) => s.id === e.detail.id);
    const newName = e.detail.name;
    if (!state || state.name === newName) return;
    const problem = validateSheetName(newName, this.sheets.filter((n) => n !== state.name));
    if (problem) {
      this._modeAnnouncement = problem;
      return;
    }
    const oldName = state.name;
    const sheets = this.sheets.map((n) => (n === oldName ? newName : n));
    if (!this._confirmSheetsChange({ sheets, kind: 'rename', sheet: newName, previous: oldName, source: 'user' })) return;
    this._renameSheet(state, newName, 'user');
    this._modeAnnouncement = `Renamed ${oldName} to ${newName}`;
  }

  private _handleSheetRemove(e: CustomEvent<{ id: number }>): void {
    if (this.readOnly || this._sheets.length === 1) return;
    const state = this._sheets.find((s) => s.id === e.detail.id);
    if (!state) return;
    const sheets = this.sheets.filter((n) => n !== state.name);
    if (!this._confirmSheetsChange({ sheets, kind: 'remove', sheet: state.name, source: 'user' })) return;
    if (this._removeSheet(state, 'user')) this._modeAnnouncement = `Deleted ${state.name}`;
  }

  private _handleSheetMove(e: CustomEvent<{ id: number; index: number }>): void {
    if (this.readOnly) return;
    const state = this._sheets.find((s) => s.id === e.detail.id);
    if (!state) return;
    const sheets = this.sheets.filter((n) => n !== state.name);
    const index = Math.max(0, Math.min(e.detail.index, sheets.length));
    sheets.splice(index, 0, state.name);
    if (!this._confirmSheetsChange({ sheets, kind: 'move', sheet: state.name, source: 'user' })) return;
    this._moveSheet(state, index, 'user');
  }

  // ─── Lists API ───────────────────────────────────────

  /**
   * Turn the text of the cells in a range into lists: several values in one
   * cell. `Ann, Bob` becomes a list of two items, which can be filtered by
   * item and used in formulas (`=HAS(A2, "Bob")`, `=COUNTA(A2)`). Formula
   * cells are left alone. A list is stored as the formula `={{"Ann";"Bob"}}`.
   * What Ctrl+J does for the selection. Undoable.
   *
   * @param range - Default: the selection
   * @returns The number of cells converted
   */
  convertToList(range?: SelectionRange): number {
    return this._convertLists(this._normalizeRange(range ?? this._selection.range, 'convertToList'), true, 'programmatic');
  }

  /**
   * Turn the lists in a range back into comma-separated text. Undoable.
   *
   * @param range - Default: the selection
   * @returns The number of cells converted
   */
  convertListToText(range?: SelectionRange): number {
    return this._convertLists(
      this._normalizeRange(range ?? this._selection.range, 'convertListToText'),
      false,
      'programmatic'
    );
  }

  /**
   * The values of a cell that holds several (a list, or an array kept in the
   * cell with `={...}`), or `null` for any other cell.
   *
   * @param cellId - Cell key in `"row:col"` format
   */
  getCellValues(cellId: string): CellArray | null {
    this._syncPendingData();
    const array = this._internalData.get(cellId)?.array;
    return array ? (JSON.parse(JSON.stringify(array)) as CellArray) : null;
  }

  /**
   * Open the popup that shows the values of a cell holding several values
   * (what Alt+Down and the cell's icon do). Nested arrays can be opened from
   * it. Escape closes it.
   *
   * @param cellId - Default: the active cell
   * @returns Whether it opened (false when the cell holds a single value)
   */
  async showCellValues(cellId?: string): Promise<boolean> {
    this._syncPendingData();
    const key = cellId ?? this._getActiveCellKey();
    if (!this._internalData.get(key)?.array) return false;
    if (this._isEditing) this._commitEdit();
    this._closeFilterMenu(false);
    const coord = parseKey(key);
    if (key !== this._getActiveCellKey()) {
      this._selection.setSelection(coord, coord, null);
      this._dispatchSelectionChange();
    }
    this._ensureCellVisible(coord.row, coord.col);
    this._arrayCard = { key, path: [] };
    if (!this._arrayCardCleanup) {
      const onPointerDown = (e: PointerEvent) => {
        const cardEl = this.renderRoot.querySelector('.ls-array-card');
        if (cardEl && e.composedPath().includes(cardEl)) return;
        this.closeCellValues(false);
      };
      document.addEventListener('pointerdown', onPointerDown);
      this._arrayCardCleanup = () => document.removeEventListener('pointerdown', onPointerDown);
    }
    this._modeAnnouncement = '';
    await this.updateComplete;
    this.renderRoot.querySelector<HTMLElement>('.ls-array-card-body')?.focus();
    return true;
  }

  /** Close the values popup. */
  closeCellValues(focusCell = true): void {
    this._arrayCardCleanup?.();
    this._arrayCardCleanup = null;
    if (!this._arrayCard) return;
    this._arrayCard = null;
    if (focusCell) this.updateComplete.then(() => this._focusActiveCell());
  }

  // ─── Paste API ───────────────────────────────────────

  /**
   * Paste the clipboard at the selection, as Ctrl+V does, with Paste Special
   * options: `mode: 'values'` pastes values without formulas or formats (what
   * Ctrl+Shift+V does), `mode: 'formats'` only the formats, and `transpose`
   * swaps rows and columns. Reads the system clipboard, so browsers may ask
   * for permission or require a user gesture. Undoable.
   */
  async paste(options: PasteSpecialOptions = {}): Promise<void> {
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();
    await this._handlePaste({ mode: options.mode ?? 'all', transpose: options.transpose ?? false });
  }

  /**
   * Open the Paste Special dialog (what Ctrl+Alt+V does): choose All, Values
   * or Formats, and whether to transpose. Returns false in read-only mode.
   */
  async openPasteSpecial(): Promise<boolean> {
    if (this.readOnly) return false;
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();
    this._closeFilterMenu(false);
    this.closeConditionalFormatManager(false);
    this._dialogOpener = this._focusedOutside();
    this._pasteSpecial = { mode: 'all', transpose: false };
    if (!this._pasteSpecialCleanup) {
      const onPointerDown = (e: PointerEvent) => {
        const el = this.renderRoot.querySelector('.ls-paste-special');
        if (el && e.composedPath().includes(el)) return;
        this._closePasteSpecial(false);
      };
      document.addEventListener('pointerdown', onPointerDown);
      this._pasteSpecialCleanup = () => document.removeEventListener('pointerdown', onPointerDown);
    }
    await this.updateComplete;
    this.renderRoot.querySelector<HTMLInputElement>('.ls-paste-special input:checked')?.focus();
    return true;
  }

  // ─── Find API ────────────────────────────────────────

  /**
   * Find the cells whose displayed value (or formula text, with
   * `formulas: true`) contains `query`. Case-insensitive by default.
   *
   * @param query - Text to search for
   * @param options - Case sensitivity and whether to search formulas
   * @returns Matching cell keys in reading order (row by row)
   */
  findCells(query: string, options: FindOptions = {}): string[] {
    if (!query) return [];
    const { matchCase = false, formulas = false } = options;
    const needle = matchCase ? query : query.toLowerCase();
    const hits: CellCoord[] = [];
    for (const [key, cell] of this._internalData) {
      const text = formulas ? cell.rawValue : cell.displayValue;
      if (!text) continue;
      const haystack = matchCase ? text : text.toLowerCase();
      if (!haystack.includes(needle)) continue;
      const coord = parseKey(key);
      if (coord.row < this.rows && coord.col < this.cols) hits.push(coord);
    }
    hits.sort((a, b) => a.row - b.row || a.col - b.col);
    return hits.map((c) => cellKey(c.row, c.col));
  }

  /** Open the find bar (as Ctrl/Cmd+F does) and focus its input. */
  async openFind(): Promise<void> {
    if (this._isEditing) this._commitEdit();
    this._findOpen = true;
    await this.updateComplete;
    this._findInput?.focus();
    this._findInput?.select();
  }

  /** Close the find bar and return focus to the active cell. */
  closeFind(): void {
    if (!this._findOpen) return;
    this._findOpen = false;
    this._findStatus = '';
    this.updateComplete.then(() => this._focusActiveCell());
  }

  // ─── Format API ──────────────────────────────────────

  /**
   * Returns the format for a cell, or undefined if unformatted.
   *
   * @param cellId - Cell key in `"row:col"` format
   * @returns The cell's format, or undefined
   */
  getCellFormat(cellId: string): CellFormat | undefined {
    return this._internalData.get(cellId)?.format;
  }

  /**
   * Merges format properties into a cell. Creates the cell if it does not exist.
   *
   * @param cellId - Cell key in `"row:col"` format
   * @param format - Format properties to merge
   */
  setCellFormat(cellId: string, format: CellFormat): void {
    this._applyFormat([cellId], format, 'programmatic');
  }

  /**
   * Applies format to all cells in a rectangular range.
   *
   * @param range - The selection range to format
   * @param format - Format properties to apply
   */
  setRangeFormat(range: SelectionRange, format: CellFormat): void {
    const cellIds: string[] = [];
    for (let r = range.start.row; r <= range.end.row; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        cellIds.push(cellKey(r, c));
      }
    }
    this._applyFormat(cellIds, format, 'programmatic');
  }

  /**
   * Removes all formatting from a single cell.
   *
   * @param cellId - Cell key in `"row:col"` format
   */
  clearCellFormat(cellId: string): void {
    this._clearFormat([cellId], 'programmatic');
  }

  /**
   * Removes all formatting from every cell in a range.
   *
   * @param range - The selection range to clear formatting from
   */
  clearRangeFormat(range: SelectionRange): void {
    const cellIds: string[] = [];
    for (let r = range.start.row; r <= range.end.row; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        cellIds.push(cellKey(r, c));
      }
    }
    this._clearFormat(cellIds, 'programmatic');
  }

  /**
   * Toggles a boolean format property on the current selection.
   * If all selected cells have the property set, it is removed; otherwise it is applied to all.
   *
   * @param prop - The boolean format property to toggle
   */
  toggleFormat(prop: 'bold' | 'italic' | 'underline' | 'strikethrough'): void {
    if (this.readOnly || this._pasteInProgress) return;
    const range = this._selection.range;
    const cellIds: string[] = [];
    let allHaveProp = true;
    let locked = 0;

    // Rows hidden by the filter are left alone, as by every user edit
    for (const r of this._visibleRowsBetween(range.start.row, range.end.row)) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        const id = cellKey(r, c);
        if (this._isCellLocked(id)) {
          locked++;
          continue;
        }
        cellIds.push(id);
        const cell = this._internalData.get(id);
        if (!cell?.format?.[prop]) {
          allHaveProp = false;
        }
      }
    }

    if (locked > 0) this._announceSkipped(locked, []);
    if (cellIds.length === 0) return;
    const format: CellFormat = { [prop]: !allHaveProp };
    this._applyFormat(cellIds, format, 'user');
  }

  /**
   * The format that the selected cells share: a property is set when every
   * selected cell has the same value for it, and left out when they differ
   * (or none has it). Rows hidden by the filter don't count, as they are
   * left alone by formatting. Use it to show the state of formatting
   * controls; `<y11n-toolbar>` does.
   */
  getSelectionFormat(): CellFormat {
    const { start, end } = this._selection.range;
    const rows = this._visibleRowsBetween(start.row, end.row);
    const area = rows.length * (end.col - start.col + 1);
    // More cells than the sheet stores: some have no format, so nothing is shared
    if (area > this._internalData.size) return {};
    const formats: CellFormat[] = [];
    for (const r of rows) {
      for (let c = start.col; c <= end.col; c++) {
        const format = this._internalData.get(cellKey(r, c))?.format;
        if (!format) return {};
        formats.push(format);
      }
    }
    const shared: CellFormat = {};
    const same = (a: unknown, b: unknown) => a === b || JSON.stringify(a) === JSON.stringify(b);
    for (const prop of Object.keys(formats[0] ?? {}) as (keyof CellFormat)[]) {
      const value = formats[0][prop];
      if (value === undefined || value === false) continue;
      if (formats.every((f) => same(f[prop], value))) {
        (shared as Record<string, unknown>)[prop] = typeof value === 'object' ? { ...value } : value;
      }
    }
    return shared;
  }

  /**
   * Merge format properties into the selected cells, as a user action: like
   * `toggleFormat()` it is refused in read-only mode, leaves read-only cells
   * and rows hidden by the filter alone, goes through `before-change` and is
   * one undo step. A property set to `undefined` is removed.
   *
   * @param format - Format properties to merge, e.g. `{ textAlign: 'center' }`
   */
  formatSelection(format: CellFormat): void {
    if (this.readOnly || this._pasteInProgress || this._endEditElsewhere()) return;
    const { cellIds } = this._selectionCellsToFormat();
    if (cellIds.length > 0) this._applyFormat(cellIds, format, 'user');
  }

  /** Remove all formatting from the selected cells, as a user action (see `formatSelection()`). */
  clearSelectionFormat(): void {
    if (this.readOnly || this._pasteInProgress || this._endEditElsewhere()) return;
    const { cellIds } = this._selectionCellsToFormat();
    if (cellIds.length > 0) this._clearFormat(cellIds, 'user');
  }

  /**
   * Show more or fewer decimal places in the selected cells, as a user
   * action. Every cell keeps its own kind of number format (currency stays
   * currency); a number without a format starts from the decimals it shows.
   * Dates, text and empty cells are left alone. One undo step.
   *
   * @param delta - Decimal places to add (negative: to remove)
   */
  changeDecimals(delta: number): void {
    if (this.readOnly || this._pasteInProgress || this._endEditElsewhere()) return;
    const step = Math.trunc(delta);
    const { cellIds, locked } = this._selectionCellsToFormat();
    const selection = this._snapshotSelection();
    const deltas: CellDelta[] = [];
    const changed: string[] = [];
    let shown: number | null = null;
    const activeKey = this._getActiveCellKey();
    for (const id of cellIds) {
      const cell = this._internalData.get(id);
      const current = cell?.format?.numberFormat;
      if (current?.type === 'date') continue;
      let decimals: number;
      if (current) decimals = current.decimals ?? 2;
      else if (cell?.type === 'number') decimals = /\.(\d+)/.exec(cell.displayValue)?.[1].length ?? 0;
      else continue;
      const next = Math.max(0, Math.min(10, decimals + step));
      if (id === activeKey || shown === null) shown = next;
      if (current && next === decimals) continue;
      deltas.push({
        id,
        before: cell!.rawValue,
        after: cell!.rawValue,
        formatBefore: cell!.format ? { ...cell!.format } : undefined,
        formatAfter: { ...cell!.format, numberFormat: { ...(current ?? { type: 'number' as const }), decimals: next } },
      });
      changed.push(id);
    }
    if (shown === null) {
      if (locked === 0) this._modeAnnouncement = 'No numbers in the selection';
      return;
    }
    if (deltas.length > 0) {
      this._executeUserBatch({ deltas, selectionBefore: selection, selectionAfter: selection, operation: 'format' });
      const first = this._internalData.get(changed[0])?.format?.numberFormat;
      this._dispatchFormatChange({ cellIds: changed, format: { numberFormat: first }, source: 'user' });
    }
    if (locked === 0) this._modeAnnouncement = shown === 1 ? '1 decimal place' : `${shown} decimal places`;
  }

  /**
   * The cells of the selection that a user format change applies to: not the
   * read-only ones (announced) and not those on rows hidden by the filter.
   */
  private _selectionCellsToFormat(): { cellIds: string[]; locked: number } {
    const range = this._selection.range;
    const cellIds: string[] = [];
    let locked = 0;
    for (const r of this._visibleRowsBetween(range.start.row, range.end.row)) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        const id = cellKey(r, c);
        if (this._isCellLocked(id)) locked++;
        else cellIds.push(id);
      }
    }
    if (locked > 0) this._announceSkipped(locked, []);
    return { cellIds, locked };
  }

  // ─── Selection and Command API ───────────────────────

  /**
   * The selected range and the active cell. `selection-change` reports the
   * range as it changes.
   */
  getSelection(): SelectionInfo {
    const { start, end } = this._selection.range;
    return {
      range: { start: { ...start }, end: { ...end } },
      activeCell: { ...this._selection.activeCell },
    };
  }

  /**
   * Whether there is a change to undo. `history-change` reports when this or
   * `canRedo` changes, and when new data (`setData()`) starts a fresh history.
   */
  get canUndo(): boolean {
    return this._undoStack.length > 0;
  }

  /** Whether there is an undone change to redo. */
  get canRedo(): boolean {
    return this._redoStack.length > 0;
  }

  /**
   * Undo the last change, as Ctrl+Z does. Refused in read-only mode.
   *
   * @returns Whether a change was undone
   */
  undo(): boolean {
    if (this._endEditElsewhere()) return false;
    return this._undo();
  }

  /**
   * Redo the change that was undone last, as Ctrl+Y does.
   *
   * @returns Whether a change was redone
   */
  redo(): boolean {
    if (this._endEditElsewhere()) return false;
    return this._redo();
  }

  /** Copy the selection to the clipboard, as Ctrl+C does. */
  async copy(): Promise<void> {
    if (this._endEditElsewhere()) return;
    await this._handleCopy();
  }

  /** Cut the selection to the clipboard, as Ctrl+X does. Refused in read-only mode. */
  async cut(): Promise<void> {
    if (this._endEditElsewhere()) return;
    await this._handleCut();
  }

  /**
   * Insert as many rows as the selection spans, above or below it, as a user
   * action: refused in read-only mode, fires the cancelable
   * `before-structure-change`, undoable and announced. Unlike the keyboard
   * shortcut it does not need whole rows selected.
   *
   * @param where - Default: `'above'`
   * @returns Whether rows were inserted
   */
  insertRowsAtSelection(where: 'above' | 'below' = 'above'): boolean {
    return this._changeStructureAtSelection('row', 'insert', where === 'below');
  }

  /**
   * Insert as many columns as the selection spans, left or right of it, as a
   * user action (see `insertRowsAtSelection()`).
   *
   * @param where - Default: `'left'`
   * @returns Whether columns were inserted
   */
  insertColumnsAtSelection(where: 'left' | 'right' = 'left'): boolean {
    return this._changeStructureAtSelection('col', 'insert', where === 'right');
  }

  /**
   * Delete the rows the selection spans, as a user action (see
   * `insertRowsAtSelection()`).
   *
   * @returns Whether rows were deleted
   */
  deleteSelectedRows(): boolean {
    return this._changeStructureAtSelection('row', 'delete');
  }

  /**
   * Delete the columns the selection spans, as a user action (see
   * `insertRowsAtSelection()`).
   *
   * @returns Whether columns were deleted
   */
  deleteSelectedColumns(): boolean {
    return this._changeStructureAtSelection('col', 'delete');
  }

  /** Fit the selected columns to their content, as Alt+Shift+0 does. Works in read-only mode. */
  autoFitSelectedColumns(): void {
    if (this._endEditElsewhere()) return;
    const { start, end } = this._selection.range;
    this._autoFitUserColumns(Array.from({ length: end.col - start.col + 1 }, (_, i) => start.col + i));
  }

  /**
   * AutoSum, as Alt+= does: on a single cell, start a SUM of the numbers
   * above it (or to its left); on a range, write a SUM under each column
   * (or right of a single row). Refused in read-only mode.
   */
  autoSum(): void {
    if (this._endEditElsewhere()) return;
    this._autoSum();
  }

  /**
   * Turn the comma-separated text of the selected cells into lists, or lists
   * back into text, as Ctrl+J does. On an empty cell it starts a new list.
   * Refused in read-only mode.
   */
  toggleLists(): void {
    if (this._endEditElsewhere()) return;
    this._toggleListsInSelection();
  }

  /**
   * Have screen readers announce a message, through the grid's own live
   * region. For controls outside the grid that change it without a message
   * of their own.
   */
  announce(message: string): void {
    this._modeAnnouncement = message;
  }

  private _changeStructureAtSelection(axis: 'row' | 'col', kind: 'insert' | 'delete', after = false): boolean {
    if (this._pasteInProgress) return false;
    if (this.readOnly) {
      this._modeAnnouncement = 'The spreadsheet is read-only';
      return false;
    }
    // A formula open on another sheet: what is selected here is not what the edit is about
    if (this._endEditElsewhere()) return false;
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();
    const { start, end } = this._selection.range;
    const first = axis === 'row' ? start.row : start.col;
    const last = axis === 'row' ? end.row : end.col;
    const count = last - first + 1;
    if (kind === 'delete' && count >= (axis === 'row' ? this.rows : this.cols)) {
      this._modeAnnouncement = `Cannot delete every ${axis === 'row' ? 'row' : 'column'}`;
      return false;
    }
    return this._changeStructure({ axis, kind, index: kind === 'insert' && after ? last + 1 : first, count }, 'user');
  }

  private _reportHistory(): void {
    const detail: HistoryChangeDetail = { canUndo: this.canUndo, canRedo: this.canRedo };
    // The generation counts the times the data was replaced: a fresh history
    const state = `${Number(detail.canUndo)}${Number(detail.canRedo)}:${this._dataGeneration}`;
    if (state === this._historyReported) return;
    this._historyReported = state;
    this.dispatchEvent(new CustomEvent<HistoryChangeDetail>('history-change', { detail, bubbles: true, composed: true }));
  }

  // ─── Conditional Format API ─────────────────────────

  /**
   * The conditional-format rules in priority order: where two rules set the
   * same thing for a cell, the earlier one wins. Returns copies.
   */
  getConditionalFormats(): ConditionalFormatRule[] {
    return this._cfRules.map(cloneConditionalRule);
  }

  /**
   * Add a conditional-format rule: a range, a condition and a format. Rules
   * live beside the data and are merged over the cells' own formats when they
   * are drawn, so `getData()`, copy, sort and CSV export are unaffected.
   * Undoable; fires `conditional-format-change`.
   *
   * @param rule - The rule; `id` is generated when left out
   * @param index - Priority position (0 = first). Default: 0, as in Excel
   * @returns The rule's id
   * @throws {TypeError} If the rule is invalid or its id is taken
   * @throws {RangeError} If the range is outside the grid
   */
  addConditionalFormat(rule: ConditionalFormatRuleInput, index = 0): string {
    const normalized = this._normalizeCfRule(rule, 'addConditionalFormat');
    if (this._cfRules.some((r) => r.id === normalized.id)) {
      throw new TypeError(`addConditionalFormat: a rule with id "${normalized.id}" already exists`);
    }
    const next = [...this._cfRules];
    next.splice(Math.max(0, Math.min(Math.trunc(index) || 0, next.length)), 0, normalized);
    this._commitConditional(next, 'programmatic');
    return normalized.id;
  }

  /**
   * Change a rule in place, keeping its priority. `changes` replaces the
   * given properties (`range`, `condition`, `format`, `stopIfTrue`, `label`).
   *
   * @returns False when there is no rule with that id
   */
  updateConditionalFormat(id: string, changes: Partial<Omit<ConditionalFormatRule, 'id'>>): boolean {
    const i = this._cfRules.findIndex((r) => r.id === id);
    if (i === -1) return false;
    const next = [...this._cfRules];
    next[i] = this._normalizeCfRule({ ...this._cfRules[i], ...changes, id }, 'updateConditionalFormat');
    this._commitConditional(next, 'programmatic');
    return true;
  }

  /**
   * Remove a rule.
   *
   * @returns False when there is no rule with that id
   */
  removeConditionalFormat(id: string): boolean {
    if (!this._cfRules.some((r) => r.id === id)) return false;
    this._commitConditional(this._cfRules.filter((r) => r.id !== id), 'programmatic');
    return true;
  }

  /**
   * Move a rule to another priority position (0 = first).
   *
   * @returns False when there is no rule with that id
   */
  moveConditionalFormat(id: string, index: number): boolean {
    const next = this._movedRules(id, index);
    if (!next) return false;
    this._commitConditional(next, 'programmatic');
    return true;
  }

  /**
   * Replace all rules, in priority order (for example, to restore saved ones).
   *
   * @throws {TypeError} If a rule is invalid or two share an id
   * @throws {RangeError} If a range is outside the grid
   */
  setConditionalFormats(rules: ConditionalFormatRuleInput[]): void {
    const next = rules.map((rule) => this._normalizeCfRule(rule, 'setConditionalFormats'));
    if (new Set(next.map((r) => r.id)).size !== next.length) {
      throw new TypeError('setConditionalFormats: rule ids must be unique');
    }
    this._commitConditional(next, 'programmatic');
  }

  /**
   * Remove conditional formatting: all of it, or from a range. Rules inside
   * the range are removed and rules it cuts across a whole edge are trimmed;
   * a rule that would no longer be one rectangle is kept as it is.
   */
  clearConditionalFormats(range?: SelectionRange): void {
    const next = range
      ? clearRulesInRange(this._cfRules, this._normalizeRange(range, 'clearConditionalFormats')).rules
      : [];
    this._commitConditional(next, 'programmatic');
  }

  /**
   * What the conditional-format rules give a cell: the merged style, the
   * labels exposed to assistive technology, and any data bar, icon or color
   * scale position. Null when no rule applies to it.
   *
   * @param cellId - Cell key in `"row:col"` format
   */
  getConditionalFormatResult(cellId: string): ConditionalFormatResult | null {
    this._syncPendingData();
    const { row, col } = parseKey(cellId);
    const result = this._cf.resultFor(row, col);
    return result ? (JSON.parse(JSON.stringify(result)) as ConditionalFormatResult) : null;
  }

  /**
   * Add a one-step rule to the selection, as the toolbar's
   * Conditional formatting menu does: `'duplicates'`, `'unique'`,
   * `'aboveAverage'`, `'belowAverage'`, `'top10'`, `'bottom10'`,
   * `'colorScale'`, `'dataBar'` or `'iconSet'`. A user action: refused in
   * read-only mode, fires the cancelable `before-conditional-format-change`,
   * and announces the outcome.
   *
   * @returns The new rule's id, or null when nothing was added
   */
  applyConditionalFormatPreset(preset: ConditionalFormatPreset): string | null {
    if (this.readOnly) return null;
    const { condition, format } = conditionalFormatPreset(preset);
    const { start, end } = this._selection.range;
    const rule = this._normalizeCfRule(
      { range: { start: { ...start }, end: { ...end } }, condition, format },
      'applyConditionalFormatPreset'
    );
    const result = this._commitConditional([rule, ...this._cfRules], 'user');
    this._modeAnnouncement =
      result === 'applied'
        ? `Conditional format added to ${rangeToA1(rule.range)}: ${describeCondition(rule.condition)}`
        : 'Conditional format change cancelled';
    return result === 'applied' ? rule.id : null;
  }

  /**
   * Clear conditional formatting from the selection (see
   * `clearConditionalFormats()`), as a user action: refused in read-only
   * mode, cancelable, announced.
   *
   * @returns Whether any rule was removed or trimmed
   */
  clearSelectionConditionalFormats(): boolean {
    if (this.readOnly) return false;
    const outcome = clearRulesInRange(this._cfRules, this._selection.range);
    const count = (n: number) => `${n} rule${n === 1 ? '' : 's'}`;
    const kept =
      outcome.kept > 0
        ? `. ${count(outcome.kept)} only partly in the selection ${outcome.kept === 1 ? 'was' : 'were'} kept`
        : '';
    if (outcome.removed + outcome.trimmed === 0) {
      this._modeAnnouncement = outcome.kept > 0 ? `Nothing cleared${kept}` : 'No conditional formats in the selection';
      return false;
    }
    if (this._commitConditional(outcome.rules, 'user') !== 'applied') {
      this._modeAnnouncement = 'Conditional format change cancelled';
      return false;
    }
    const parts: string[] = [];
    if (outcome.removed > 0) parts.push(`${count(outcome.removed)} removed`);
    if (outcome.trimmed > 0) parts.push(`${count(outcome.trimmed)} trimmed`);
    this._modeAnnouncement = `Conditional formats cleared from the selection: ${parts.join(', ')}${kept}`;
    return true;
  }

  /**
   * Open the conditional-format rule manager (what Alt+Shift+C does): a
   * dialog that lists the rules in priority order and lets the user add,
   * edit, reorder and delete them. In read-only mode it only shows them.
   */
  async openConditionalFormatManager(): Promise<boolean> {
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();
    this._closeFilterMenu(false);
    this._closePasteSpecial(false);
    this._dialogOpener = this._focusedOutside();
    this._cfManagerStatus = '';
    this._cfManagerOpen = true;
    if (!this._cfManagerCleanup) {
      const onPointerDown = (e: PointerEvent) => {
        const el = this.renderRoot.querySelector('y11n-conditional-format-dialog');
        if (el && e.composedPath().includes(el)) return;
        this.closeConditionalFormatManager(false);
      };
      document.addEventListener('pointerdown', onPointerDown);
      this._cfManagerCleanup = () => document.removeEventListener('pointerdown', onPointerDown);
    }
    this._modeAnnouncement = '';
    await this.updateComplete;
    const dialog = this.renderRoot.querySelector<Y11nConditionalFormatDialog>('y11n-conditional-format-dialog');
    await dialog?.updateComplete;
    dialog?.focusInitial();
    return true;
  }

  /**
   * Close the rule manager.
   *
   * @param focusCell - Return focus to the active cell. Default: true
   */
  closeConditionalFormatManager(focusCell = true): void {
    this._cfManagerCleanup?.();
    this._cfManagerCleanup = null;
    if (!this._cfManagerOpen) return;
    this._cfManagerOpen = false;
    if (focusCell) this._returnFocus();
  }

  // ─── Format Internals ────────────────────────────────

  private _applyFormat(cellIds: string[], format: CellFormat, source: FormatChangeDetail['source']): void {
    const selection = this._snapshotSelection();
    const batch = this._buildFormatBatch(cellIds, format, selection);
    if (!batch) return;

    if (source === 'user') {
      this._executeUserBatch(batch);
    } else {
      this._commitBatch(batch, source);
    }

    this._dispatchFormatChange({ cellIds, format, source });
  }

  private _clearFormat(cellIds: string[], source: FormatChangeDetail['source']): void {
    const selection = this._snapshotSelection();
    const deltas: CellDelta[] = [];

    for (const id of cellIds) {
      const cell = this._internalData.get(id);
      if (cell?.format) {
        deltas.push({
          id,
          before: cell.rawValue,
          after: cell.rawValue,
          formatBefore: { ...cell.format },
          formatAfter: undefined,
        });
      }
    }

    if (deltas.length === 0) return;

    const batch: CommandBatch = {
      deltas,
      selectionBefore: selection,
      selectionAfter: selection,
      operation: 'format',
    };

    if (source === 'user') {
      this._executeUserBatch(batch);
    } else {
      this._commitBatch(batch, source);
    }

    this._dispatchFormatChange({ cellIds, format: {}, source });
  }

  private _buildFormatBatch(
    cellIds: string[],
    format: CellFormat,
    selection: SelectionSnapshot
  ): CommandBatch | null {
    const deltas: CellDelta[] = [];

    for (const id of cellIds) {
      const cell = this._internalData.get(id);
      const existingFormat = cell?.format;
      const merged = { ...existingFormat, ...format };
      // Clean out undefined values
      for (const k of Object.keys(merged) as (keyof CellFormat)[]) {
        if (merged[k] === undefined) delete merged[k];
      }

      const formatChanged = !formatsEqual(existingFormat, merged);
      if (formatChanged) {
        deltas.push({
          id,
          before: cell?.rawValue ?? '',
          after: cell?.rawValue ?? '',
          formatBefore: existingFormat ? { ...existingFormat } : undefined,
          formatAfter: Object.keys(merged).length > 0 ? merged : undefined,
        });
      }
    }

    if (deltas.length === 0) return null;

    return {
      deltas,
      selectionBefore: selection,
      selectionAfter: selection,
      operation: 'format',
    };
  }

  /**
   * Convert CellFormat to inline styles for styleMap. `conditional` is the
   * style of the conditional-format rules that apply to the cell, merged over
   * its own format for display only.
   */
  private _getCellStyles(row: number, col: number, conditional?: ConditionalFormatStyle): Record<string, string> {
    const cell = this._getCell(row, col);
    if (!cell?.format && !conditional) return {};

    const styles: Record<string, string> = {};
    const fmt: CellFormat = conditional ? { ...cell?.format, ...conditional } : cell!.format!;

    if (fmt.bold) styles['font-weight'] = 'bold';
    if (fmt.italic) styles['font-style'] = 'italic';

    const decorations: string[] = [];
    if (fmt.underline) decorations.push('underline');
    if (fmt.strikethrough) decorations.push('line-through');
    if (decorations.length > 0) styles['text-decoration'] = decorations.join(' ');

    if (fmt.textColor) styles.color = fmt.textColor;
    if (fmt.backgroundColor) styles['background-color'] = fmt.backgroundColor;
    if (fmt.fontSize) styles['font-size'] = `${fmt.fontSize}px`;

    if (fmt.textAlign) {
      styles['text-align'] = fmt.textAlign;
      const justifyMap = { left: 'flex-start', center: 'center', right: 'flex-end' };
      styles['justify-content'] = justifyMap[fmt.textAlign];
    }

    if (fmt.verticalAlign) {
      const alignMap = { top: 'flex-start', middle: 'center', bottom: 'flex-end' };
      styles['align-items'] = alignMap[fmt.verticalAlign];
    }

    if (fmt.wrap) {
      styles['white-space'] = 'normal';
      styles['overflow-wrap'] = 'anywhere';
    }

    return styles;
  }

  // ─── Conditional Format Internals ────────────────────

  /** The value a cell is tested by: its number even when the display is formatted. */
  private _cfValue(key: string): SortValue {
    return sortValueOf(this._internalData.get(key), (cell) => this._sortNumber(key, cell));
  }

  /** Validate a rule from the API or the rule manager, giving it an id if it has none. */
  private _normalizeCfRule(input: ConditionalFormatRuleInput, method: string): ConditionalFormatRule {
    let id = input?.id;
    if (id !== undefined && (typeof id !== 'string' || id === '')) {
      throw new TypeError(`${method}: id must be a non-empty string`);
    }
    if (id === undefined) {
      do {
        id = `cf-${this._cfNextId++}`;
      } while (this._cfRules.some((r) => r.id === id));
    }
    const rule = normalizeConditionalRule(input, id);
    this._normalizeRange(rule.range, method);
    return rule;
  }

  /** The rules with one moved to `index`, or null when there is no rule with that id. */
  private _movedRules(id: string, index: number): ConditionalFormatRule[] | null {
    const from = this._cfRules.findIndex((r) => r.id === id);
    if (from === -1) return null;
    const next = [...this._cfRules];
    const [rule] = next.splice(from, 1);
    next.splice(Math.max(0, Math.min(Math.trunc(index) || 0, next.length)), 0, rule);
    return next;
  }

  private _setConditionalRules(rules: readonly ConditionalFormatRule[]): void {
    this._cfRules = rules;
    this._cfRevision++;
    this._cf.setRules(rules);
  }

  /**
   * Change the rules as one undoable step. User changes can be vetoed via
   * `before-conditional-format-change`, whose detail describes the result.
   */
  private _commitConditional(
    next: readonly ConditionalFormatRule[],
    source: 'user' | 'programmatic'
  ): 'applied' | 'unchanged' | 'cancelled' {
    const before = this._cfRules;
    if (conditionalRulesEqual(before, next)) return 'unchanged';

    if (source === 'user') {
      const event = new CustomEvent<ConditionalFormatChangeDetail>('before-conditional-format-change', {
        detail: { rules: next.map(cloneConditionalRule), source },
        bubbles: true,
        composed: true,
        cancelable: true,
      });
      const revision = this._cfRevision;
      if (!this.dispatchEvent(event)) return 'cancelled';
      // A listener changed the rules: `next` was built from stale ones
      if (revision !== this._cfRevision) return 'cancelled';
    }

    const selection = this._snapshotSelection();
    this._commitBatch(
      // `operation` is never reported for these batches: they fire conditional-format-change, not data-change
      {
        deltas: [],
        selectionBefore: selection,
        selectionAfter: selection,
        operation: 'format',
        conditional: { before, after: next },
      },
      source
    );
    return 'applied';
  }

  private _applyConditionalBatch(batch: CommandBatch, side: 'before' | 'after', source: ChangeSource): void {
    this._setConditionalRules(side === 'after' ? batch.conditional!.after : batch.conditional!.before);
    if (source === 'undo' || source === 'redo') {
      this._restoreSelection(side === 'after' ? batch.selectionAfter : batch.selectionBefore);
    }
    this.requestUpdate();
    this._dispatchConditionalChange(source);
  }

  private _dispatchConditionalChange(source: ConditionalFormatChangeDetail['source']): void {
    this.dispatchEvent(
      new CustomEvent<ConditionalFormatChangeDetail>('conditional-format-change', {
        detail: { rules: this._cfRules.map(cloneConditionalRule), source },
        bubbles: true,
        composed: true,
      })
    );
  }

  /** Report a change made in the rule manager inside it, where its user is. */
  private _reportCfChange(result: ReturnType<Y11nSpreadsheet['_commitConditional']>, applied: string): void {
    this._cfManagerStatus =
      result === 'applied' ? applied : result === 'cancelled' ? 'Change cancelled' : 'Nothing changed';
  }

  private _handleCfSave(e: CustomEvent<ConditionalFormatSaveDetail>): void {
    if (this.readOnly) return;
    const { id, rule } = e.detail;
    let next: ConditionalFormatRule[];
    let saved: ConditionalFormatRule;
    try {
      if (id === null) {
        saved = this._normalizeCfRule(rule, 'conditional format');
        next = [saved, ...this._cfRules];
      } else {
        const i = this._cfRules.findIndex((r) => r.id === id);
        if (i === -1) {
          this._cfManagerStatus = 'That rule no longer exists';
          return;
        }
        saved = this._normalizeCfRule({ ...rule, id }, 'conditional format');
        next = [...this._cfRules];
        next[i] = saved;
      }
    } catch (err) {
      this._cfManagerStatus = (err as Error).message;
      return;
    }
    const what = `${describeCondition(saved.condition)}, ${rangeToA1(saved.range)}`;
    this._reportCfChange(
      this._commitConditional(next, 'user'),
      id === null ? `Rule added: ${what}` : `Rule saved: ${what}`
    );
  }

  private _handleCfRemove(e: CustomEvent<{ id: string }>): void {
    if (this.readOnly) return;
    const rule = this._cfRules.find((r) => r.id === e.detail.id);
    if (!rule) return;
    this._reportCfChange(
      this._commitConditional(this._cfRules.filter((r) => r !== rule), 'user'),
      `Rule deleted: ${describeCondition(rule.condition)}, ${rangeToA1(rule.range)}`
    );
  }

  private _handleCfMove(e: CustomEvent<ConditionalFormatMoveDetail>): void {
    if (this.readOnly) return;
    const next = this._movedRules(e.detail.id, e.detail.index);
    if (!next) return;
    const position = next.findIndex((r) => r.id === e.detail.id) + 1;
    this._reportCfChange(
      this._commitConditional(next, 'user'),
      `Rule moved to position ${position} of ${next.length}`
    );
  }

  /** Classes for a cell that conditional formats apply to. */
  private _cfClasses(cf: ConditionalFormatResult): string {
    const graphics = (cf.dataBar ? 1 : 0) + (cf.icon ? 1 : 0) + (cf.scale !== undefined ? 1 : 0);
    // A highlight rule matched: forced-colors mode marks the cell, since it drops the fill
    return cf.ruleIds.length > graphics ? 'cf-match cf-highlight' : 'cf-match';
  }

  /** Whether a data bar or icon replaces the value on screen (it stays in the accessible name). */
  private _cfHidesValue(cf: ConditionalFormatResult): boolean {
    return cf.dataBar?.showValue === false || cf.icon?.showValue === false;
  }

  /** The data bar, icon and forced-colors scale meter of a cell. All decorative: the label carries the meaning. */
  private _renderCfGraphics(cf: ConditionalFormatResult) {
    const bar = cf.dataBar;
    const percent = (n: number) => `${Math.round(n * 1000) / 10}%`;
    return html`
      ${bar
        ? html`<span
            class="ls-cf-bar ${bar.negative ? 'negative' : ''}"
            aria-hidden="true"
            style=${styleMap({
              left: percent(bar.start),
              width: percent(bar.size),
              ...(bar.color ? { 'background-color': bar.color } : {}),
            })}
          ></span>`
        : nothing}
      ${cf.scale !== undefined
        ? html`<span class="ls-cf-meter" aria-hidden="true" style="width: ${percent(cf.scale)}"></span>`
        : nothing}
      ${cf.icon
        ? html`<span class="ls-cf-icon" aria-hidden="true">${conditionalIcon(cf.icon.set, cf.icon.index)}</span>`
        : nothing}
    `;
  }

  // ─── Data Sync ──────────────────────────────────────

  private _syncData(): void {
    const oldData = this._internalData;
    // Clone each cell so internal mutations (recalc, number formatting) never
    // leak into the caller's objects, and caller mutations never leak in.
    const newData: GridData = new Map();
    for (const [key, cell] of this.data) {
      // Spilled values are recomputed, never imported
      const stored = storedCell(cell);
      if (stored) newData.set(key, cloneCell(stored));
    }

    // Diff old vs new to collect changed keys for future targeted recalc
    const changedKeys: string[] = [];

    // Keys in new data that are different or absent in old data
    for (const [key, cell] of newData) {
      const oldCell = oldData.get(key);
      if (
        !oldCell ||
        oldCell.rawValue !== cell.rawValue ||
        !formatsEqual(oldCell.format, cell.format)
      ) {
        changedKeys.push(key);
        // What the caller's copy says it spills or holds is recomputed
        delete cell.spill;
        delete cell.array;
        delete cell.spilledFrom;
      } else {
        // Unchanged cells keep their computed values; the caller's copy may
        // carry a placeholder displayValue (e.g. '' for formulas).
        cell.displayValue = oldCell.displayValue;
        cell.type = oldCell.type;
        if (oldCell.spill) cell.spill = oldCell.spill;
        else delete cell.spill;
        if (oldCell.array) cell.array = oldCell.array;
        else delete cell.array;
        // A formatted cell that shows a spilled value still does
        if (oldCell.spilledFrom !== undefined && cell.rawValue === '') cell.spilledFrom = oldCell.spilledFrom;
      }
    }

    // Keys in old data that are absent in new data (deleted)
    for (const key of oldData.keys()) {
      if (!newData.has(key)) {
        changedKeys.push(key);
      }
    }

    this._internalData = newData;
    this._lastSyncedDataRef = this.data;
    // Changes made on this sheet can't be undone into the new data; other sheets keep theirs
    this._dropHistoryOf(this._activeSheetId);
    this._dataGeneration++;
    this._mutationCount++;

    if (changedKeys.length > 0 && oldData.size > 0) {
      // Incremental update: preserve dep graph for unchanged cells
      this._formulaEngine.updateData(this._internalData, changedKeys);
      // Literal cells aren't recalculated by the engine: derive their display
      // from the raw value so an added or removed number format takes effect.
      for (const key of changedKeys) {
        const cell = this._internalData.get(key);
        if (cell && !cell.rawValue.startsWith('=')) {
          const evaluated = this._formulaEngine.evaluate(cell.rawValue);
          cell.displayValue = evaluated.displayValue;
          cell.type = evaluated.type;
        }
      }
      this._recalcAffected(changedKeys);
      this._applyNumberFormatsToKeys(new Set(changedKeys));
    } else {
      // Full reset: first load or complete replacement
      this._formulaEngine.setData(this._internalData);
      this._recalcAll();
    }
  }

  /**
   * Whether the grid changed since `mutations` was read from `_mutationCount`,
   * counting a `data` assignment that hasn't been imported yet.
   */
  private _changedSince(mutations: number): boolean {
    return mutations !== this._mutationCount || this.data !== this._lastSyncedDataRef;
  }

  /**
   * Import a `data` assignment that Lit hasn't processed yet, so an API call
   * made right after it works on the new data instead of overwriting it.
   */
  private _syncPendingData(): void {
    if (this.data !== this._lastSyncedDataRef) this._syncData();
  }

  /**
   * Tell the engine the grid's size, which bounds spills and whole-column
   * references. Returns whether it changed.
   */
  private _syncEngineSize(): boolean {
    return this._formulaEngine.setSheetSize(Math.max(1, Math.round(this.rows)), Math.max(1, Math.round(this.cols)));
  }

  private _recalcAll(): void {
    this._syncEngineSize();
    this._formulaEngine.recalculate();
    this._reapplyNumberFormats();
    this._formatOtherSheets();
    this._cf.invalidate();
    this._nameRanges = null;
  }

  /** Apply number formats to the cells of other sheets that the last recalculation changed. */
  private _formatOtherSheets(): void {
    if (this._sheets.length === 1) return;
    const active = this._activeState.name;
    for (const [name, keys] of this._formulaEngine.changes) {
      if (name === active) continue;
      const data = this._findSheet(name)?.data;
      if (!data) continue;
      for (const key of keys) {
        const cell = data.get(key);
        if (cell && cell.type === 'number' && cell.format?.numberFormat) {
          const plain = this._formulaEngine.resultFor(key, cell.rawValue, name).displayValue;
          cell.displayValue = formatNumber(Number(plain), cell.format.numberFormat);
        }
      }
    }
  }

  private _recalcAffected(changedKeys: string[]): void {
    if (this._syncEngineSize()) {
      // The size bounds every spill and whole-column reference
      this._formulaEngine.updateData(this._internalData, changedKeys);
      this._recalcAll();
      return;
    }
    const changed = this._formulaEngine.recalculateAffected(changedKeys);
    this._applyNumberFormatsToKeys(changed);
    this._formatOtherSheets();
    // Conditional formats follow the values: the written cells and every formula that changed with them
    this._cf.invalidate(changedKeys);
    this._cf.invalidate(changed);
  }

  /** Re-apply number formatting to all cells that have a numberFormat */
  private _reapplyNumberFormats(): void {
    for (const [key, cell] of this._internalData) {
      this._applyNumberFormatToCell(key, cell);
    }
  }

  /** Re-apply number formatting only to the given set of cell keys */
  private _applyNumberFormatsToKeys(keys: Set<string>): void {
    for (const key of keys) {
      const cell = this._internalData.get(key);
      if (cell) {
        this._applyNumberFormatToCell(key, cell);
      }
    }
  }

  /** Apply number formatting to a single cell using rawValue for robustness */
  private _applyNumberFormatToCell(key: string, cell: CellData): void {
    if (cell.type === 'number' && cell.format?.numberFormat) {
      cell.displayValue = formatNumber(this._plainNumber(key, cell), cell.format.numberFormat);
    }
  }

  /** The number a number cell holds, whatever its display format: typed, computed, or spilled into it. */
  private _plainNumber(key: string, cell: CellData): number {
    return cell.rawValue.startsWith('=') || cell.spilledFrom !== undefined
      ? Number(this._formulaEngine.resultFor(key, cell.rawValue).displayValue)
      : Number(cell.rawValue);
  }

  // ─── Cell Access ────────────────────────────────────

  private _getCell(row: number, col: number): CellData | undefined {
    return this._internalData.get(cellKey(row, col));
  }

  private _getCellDisplay(row: number, col: number): string {
    const cell = this._getCell(row, col);
    return cell?.displayValue ?? '';
  }

  private _getActiveCellKey(): string {
    const { row, col } = this._selection.activeCell;
    return cellKey(row, col);
  }

  private _getActiveCellRef(): string {
    return coordToRef(this._selection.activeCell);
  }

  private _setCellRaw(key: string, rawValue: string): void {
    const existing = this._internalData.get(key);
    const evaluated = this._formulaEngine.evaluate(rawValue, key);

    // Apply number format to display value if the cell has one and the value is numeric
    let displayValue = evaluated.displayValue;
    const numFmt = existing?.format?.numberFormat;
    if (numFmt && evaluated.type === 'number') {
      displayValue = formatNumber(Number(evaluated.displayValue), numFmt);
    }

    const cell: CellData = {
      rawValue,
      displayValue,
      type: evaluated.type,
      format: existing?.format,
    };
    if (existing?.readOnly) cell.readOnly = true;
    this._internalData.set(key, cell);
  }

  private _applyRawValueByString(key: string, rawValue: string): void {
    if (rawValue === '') {
      const existing = this._internalData.get(key);
      if (existing?.format || existing?.readOnly) {
        // Keep the entry so the format / read-only flag survives
        const cell: CellData = { rawValue: '', displayValue: '', type: 'text' };
        if (existing.format) cell.format = existing.format;
        if (existing.readOnly) cell.readOnly = true;
        this._internalData.set(key, cell);
      } else {
        this._internalData.delete(key);
      }
      return;
    }
    this._setCellRaw(key, rawValue);
  }

  private _snapshotSelection(): SelectionSnapshot {
    return {
      anchor: { ...this._selection.anchor },
      head: { ...this._selection.head },
      active: this._selection.activeOverride,
    };
  }

  private _restoreSelection(snapshot: SelectionSnapshot): void {
    this._selection.setSelection(snapshot.anchor, snapshot.head, snapshot.active);
    const active = this._selection.activeCell;
    if (this._isRowHidden(active.row)) {
      // The active cell must be on a visible row (it holds the focus)
      const cell = { row: this._nearestVisibleRow(active.row), col: active.col };
      this._selection.setSelection(cell, cell, null);
    }
    this._dispatchSelectionChange();
    // A change made from outside (a toolbar, the page's own buttons) leaves the focus there
    if (!this._focusedOutside()) this.updateComplete.then(() => this._focusActiveCell());
  }

  /**
   * The element that has the keyboard focus, if it is outside this component:
   * a toolbar button, a control of the page. Null when the focus is in the
   * grid (or one of its parts) or nowhere.
   */
  private _focusedOutside(): HTMLElement | null {
    let active = document.activeElement as HTMLElement | null;
    while (active?.shadowRoot?.activeElement) active = active.shadowRoot.activeElement as HTMLElement;
    if (!active || active === document.body || active === document.documentElement) return null;
    for (let node: Node | null = active; node; node = node.parentNode ?? (node as ShadowRoot).host ?? null) {
      if (node === this) return null;
    }
    return active;
  }

  /**
   * After a dialog closes: focus what opened it, if that was outside the grid
   * (a toolbar button), otherwise the active cell.
   */
  private _returnFocus(): void {
    const opener = this._dialogOpener;
    this._dialogOpener = null;
    this.updateComplete.then(() => {
      if (opener?.isConnected) opener.focus();
      else this._focusActiveCell();
    });
  }

  private _buildCommandBatch(
    updates: Array<{ id: string; value: string }>,
    operation: DataChangeOperation,
    selectionBefore: SelectionSnapshot,
    selectionAfter: SelectionSnapshot
  ): CommandBatch | null {
    const deltas: CellDelta[] = [];

    for (const update of updates) {
      const before = this._internalData.get(update.id)?.rawValue ?? '';
      const after = update.value;
      if (before !== after) {
        deltas.push({ id: update.id, before, after });
      }
    }

    if (deltas.length === 0) {
      return null;
    }

    return {
      deltas,
      selectionBefore,
      selectionAfter,
      operation,
    };
  }

  /**
   * Build a batch that writes values and, optionally, formats. Cells whose
   * value and format would both be unchanged are skipped.
   */
  private _buildCellBatch(
    updates: CellUpdate[],
    operation: DataChangeOperation,
    selection: SelectionSnapshot
  ): CommandBatch | null {
    const deltas: CellDelta[] = [];
    for (const update of updates) {
      const existing = this._internalData.get(update.id);
      const before = existing?.rawValue ?? '';
      const delta: CellDelta = { id: update.id, before, after: update.value };
      if (update.format !== undefined) {
        const formatAfter = update.format && Object.keys(update.format).length > 0
          ? { ...update.format }
          : undefined;
        if (!formatsEqual(existing?.format, formatAfter)) {
          delta.formatBefore = existing?.format ? { ...existing.format } : undefined;
          delta.formatAfter = formatAfter;
        }
      }
      if (before !== update.value || 'formatAfter' in delta) {
        deltas.push(delta);
      }
    }
    if (deltas.length === 0) return null;
    return { deltas, selectionBefore: selection, selectionAfter: selection, operation };
  }

  /**
   * Execute and record a user batch. Large batches are applied in chunks
   * across animation frames so the UI stays responsive.
   */
  private async _runUserBatch(batch: CommandBatch): Promise<void> {
    if (batch.deltas.length <= Y11nSpreadsheet.PASTE_CHUNK_SIZE) {
      this._executeUserBatch(batch);
      return;
    }
    this._pasteInProgress = true;
    const generation = this._dataGeneration;
    try {
      if (await this._applyBatchProgressive(batch)) {
        this._pushHistory(batch, generation);
      }
    } finally {
      this._pasteInProgress = false;
    }
  }

  /**
   * Record a batch for undo, unless the data was replaced (which clears the
   * history) since `generation` was captured, before the batch was applied.
   */
  private _pushHistory(batch: CommandBatch, generation: number): void {
    if (generation !== this._dataGeneration) return;
    batch.sheetId ??= this._activeSheetId;
    this._undoStack.push(batch);
    this._redoStack = [];
    if (this._undoStack.length > this._maxHistory) {
      this._undoStack.shift();
    }
    // `history-change` is fired from updated()
    this.requestUpdate();
  }

  private _finalizeBatch(
    batch: CommandBatch,
    source: ChangeSource,
    valueSide: 'before' | 'after'
  ): void {
    this._recalcAffected(batch.deltas.map((d) => d.id));

    const shouldEmitDataChange = source !== 'user' || batch.operation !== 'edit';
    if (shouldEmitDataChange) {
      this._dispatchDataChange({
        updates: batch.deltas.map((delta) => ({ id: delta.id, value: delta[valueSide] })),
        source,
        operation: batch.operation,
      });
    }

    this.requestUpdate();
  }

  /** Whether a batch only changes the view (column widths, the filter, conditional formats), not the data. */
  private _isViewBatch(batch: CommandBatch): boolean {
    return !!batch.widths || ((!!batch.filter || !!batch.conditional) && !batch.structure);
  }

  private _applyBatch(batch: CommandBatch, source: ChangeSource): void {
    // Column widths and filters are views, not data: they must not make pending writes look stale
    if (!this._isViewBatch(batch)) this._mutationCount++;
    if (batch.structure) {
      this._applyStructureBatch(batch, 'after', source);
      return;
    }
    if (batch.sort) {
      this._applySortBatch(batch, 'after', source);
      return;
    }
    if (batch.widths) {
      this._applyWidthBatch(batch, 'after', source);
      return;
    }
    if (batch.filter) {
      this._applyFilterBatch(batch, 'after', source);
      return;
    }
    if (batch.conditional) {
      this._applyConditionalBatch(batch, 'after', source);
      return;
    }
    for (const delta of batch.deltas) {
      this._applyRawValueByString(delta.id, delta.after);
      if ('formatAfter' in delta) {
        this._applyCellFormat(delta.id, delta.formatAfter);
      }
    }
    this._lastSyncedDataRef = this.data;
    this._restoreSelection(batch.selectionAfter);
    this._finalizeBatch(batch, source, 'after');
  }

  private _revertBatch(batch: CommandBatch): void {
    if (!this._isViewBatch(batch)) this._mutationCount++;
    if (batch.structure) {
      this._applyStructureBatch(batch, 'before', 'undo');
      return;
    }
    if (batch.sort) {
      this._applySortBatch(batch, 'before', 'undo');
      return;
    }
    if (batch.widths) {
      this._applyWidthBatch(batch, 'before', 'undo');
      return;
    }
    if (batch.filter) {
      this._applyFilterBatch(batch, 'before', 'undo');
      return;
    }
    if (batch.conditional) {
      this._applyConditionalBatch(batch, 'before', 'undo');
      return;
    }
    for (const delta of batch.deltas) {
      this._applyRawValueByString(delta.id, delta.before);
      if ('formatBefore' in delta) {
        this._applyCellFormat(delta.id, delta.formatBefore);
      }
    }
    this._lastSyncedDataRef = this.data;
    this._restoreSelection(batch.selectionBefore);
    this._finalizeBatch(batch, 'undo', 'before');
  }

  /** Low-level: set or remove format on a cell */
  private _applyCellFormat(cellId: string, format: CellFormat | undefined): void {
    let cell = this._internalData.get(cellId);
    if (!cell) {
      // Create an empty cell to hold the format
      cell = { rawValue: '', displayValue: '', type: 'text' };
      this._internalData.set(cellId, cell);
    }
    if (format && Object.keys(format).length > 0) {
      cell.format = format;
    } else {
      delete cell.format;
      // Remove empty cells
      if (cell.rawValue === '' && !cell.readOnly && cell.spilledFrom === undefined) {
        this._internalData.delete(cellId);
        return;
      }
    }
    // Re-apply number format to display value
    if (cell.type === 'number' && cell.format?.numberFormat) {
      // For formula cells, get the unformatted result from the engine
      // (displayValue may already be formatted, e.g., "$1,234.50")
      cell.displayValue = formatNumber(this._plainNumber(cellId, cell), cell.format.numberFormat);
    } else if (cell.type === 'number' && !cell.format?.numberFormat) {
      // Remove number formatting — show the plain result
      const evaluated = this._formulaEngine.resultFor(cellId, cell.rawValue);
      cell.displayValue = evaluated.displayValue;
    }
  }

  private _executeUserBatch(batch: CommandBatch): void {
    this._commitBatch(batch, 'user');
  }

  /**
   * Record a batch for undo, then apply it. Recording first keeps the history
   * in order when a listener of the events fired while applying makes a
   * change of its own (say, re-sorting on every `data-change`): that change
   * is recorded after this one, so undo reverts them in the right order. If a
   * listener replaces the data instead, the history is cleared with it.
   */
  private _commitBatch(batch: CommandBatch, source: ChangeSource): void {
    this._pushHistory(batch, this._dataGeneration);
    this._applyBatch(batch, source);
  }

  private _commitRawValue(
    key: string,
    newValue: string,
    options: { validate?: boolean; format?: CellFormat | null } = {}
  ): boolean {
    const oldCell = this._internalData.get(key);
    const oldValue = oldCell?.rawValue ?? '';
    const selection = this._snapshotSelection();
    const batch =
      options.format !== undefined
        ? this._buildCellBatch([{ id: key, value: newValue, format: options.format }], 'edit', selection)
        : newValue === oldValue
          ? null
          : this._buildCommandBatch([{ id: key, value: newValue }], 'edit', selection, selection);

    const vetted = batch && this._vetUserBatch(batch, options);
    if (!vetted) return false;
    this._executeUserBatch(vetted);
    this._dispatchCellChange({ cellId: key, value: vetted.deltas[0].after, oldValue });
    return true;
  }

  // ─── Read-only cells, validation, before-change ─────

  /** Whether a cell is individually marked read-only (`CellData.readOnly`). */
  private _isCellLocked(key: string): boolean {
    return this._internalData.get(key)?.readOnly === true;
  }

  /** Run `validateCell`; returns the rejection message, or null if valid. */
  private _validate(cellId: string, value: string): string | null {
    if (!this.validateCell || value === '') return null;
    try {
      const result = this.validateCell(cellId, value);
      return typeof result === 'string' && result !== '' ? result : null;
    } catch (err) {
      console.error('y11n-spreadsheet: validateCell threw', err);
      return 'Invalid value';
    }
  }

  private _announceSkipped(locked: number, errors: ValidationErrorDetail['errors']): void {
    const parts: string[] = [];
    if (locked > 0) {
      parts.push(locked === 1 ? '1 read-only cell was not changed' : `${locked} read-only cells were not changed`);
    }
    if (errors.length === 1) {
      parts.push(`Invalid value: ${errors[0].message}`);
    } else if (errors.length > 1) {
      parts.push(`${errors.length} values were rejected: ${errors[0].message}`);
    }
    if (parts.length > 0) this._modeAnnouncement = parts.join('. ');
  }

  /**
   * Prepare a user batch for applying: drop read-only cells and values that
   * `validateCell` rejects, then fire the cancelable `before-change` event
   * and apply the host's rewrites. Returns null when nothing is left to apply.
   */
  private _vetUserBatch(batch: CommandBatch, options: { validate?: boolean } = {}): CommandBatch | null {
    const { validate = true } = options;
    // Re-checked here because async operations (paste, cut) await the
    // clipboard between their initial checks and applying.
    if (this.readOnly || this._pasteInProgress) return null;
    // Validators and event listeners may change the grid; if they do, the
    // batch's before-values are stale and it must not be applied.
    const mutations = this._mutationCount;
    const stale = () => {
      if (!this._changedSince(mutations)) return false;
      this._modeAnnouncement = 'Change not applied because the data changed';
      return true;
    };
    let locked = 0;
    const errors: ValidationErrorDetail['errors'] = [];
    let deltas = batch.deltas.filter((d) => {
      if (this._isCellLocked(d.id)) {
        locked++;
        return false;
      }
      if (validate && d.after !== d.before) {
        const message = this._validate(d.id, d.after);
        if (message) {
          errors.push({ cellId: d.id, value: d.after, message });
          return false;
        }
      }
      return true;
    });

    if (errors.length > 0) this._dispatchValidationError({ errors, operation: batch.operation });
    this._announceSkipped(locked, errors);
    if (stale() || deltas.length === 0) return null;

    if (batch.operation !== 'format') {
      const changes: ProposedChange[] = deltas.map((d) => ({
        cellId: d.id,
        value: d.after,
        oldValue: d.before,
      }));
      const detail: BeforeChangeDetail = { changes, operation: batch.operation };
      const event = new CustomEvent<BeforeChangeDetail>('before-change', {
        detail,
        bubbles: true,
        composed: true,
        cancelable: true,
      });
      if (!this.dispatchEvent(event)) {
        this._modeAnnouncement = 'Change cancelled';
        return null;
      }
      if (stale()) return null;
      // Apply the host's edits to the list: rewritten values and removed cells
      const finalValues = new Map<string, string>();
      for (const change of Array.isArray(detail.changes) ? detail.changes : []) {
        if (change && typeof change.cellId === 'string') {
          finalValues.set(change.cellId, String(change.value ?? ''));
        }
      }
      deltas = deltas.flatMap((d) => {
        const after = finalValues.get(d.id);
        if (after === undefined) return [];
        if (after === d.before && !('formatAfter' in d)) return [];
        return [{ ...d, after }];
      });
      if (deltas.length === 0) return null;
    }

    return { ...batch, deltas };
  }

  // ─── Row / Column Insertion and Deletion ────────────

  /**
   * Insert or delete rows/columns as one undoable step. Programmatic calls
   * throw a RangeError for out-of-range arguments; user-initiated ones are
   * pre-checked by the caller and can be vetoed via `before-structure-change`.
   */
  private _changeStructure(change: StructureChange, source: 'user' | 'programmatic'): boolean {
    const { axis, kind, index, count } = change;
    const size = axis === 'row' ? this.rows : this.cols;
    const noun = axis === 'row' ? 'row' : 'column';
    if (!Number.isInteger(index) || !Number.isInteger(count) || index < 0 || count < 1) {
      throw new RangeError(`${kind} ${noun}s: index must be a non-negative integer and count a positive integer`);
    }
    if (kind === 'insert' ? index > size : index + count > size) {
      throw new RangeError(`${kind} ${noun}s: ${index}+${count} is outside the grid (${size} ${noun}s)`);
    }
    if (kind === 'delete' && count >= size) {
      throw new RangeError(`delete ${noun}s: at least one ${noun} must remain`);
    }
    if (this._pasteInProgress) return false;
    // Import a pending data assignment first, so the edit commits into it
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();

    if (source === 'user') {
      if (this.readOnly) return false;
      if (kind === 'delete' && this._hasLockedCellIn(change)) {
        this._modeAnnouncement = `Cannot delete: the ${noun}s contain read-only cells`;
        return false;
      }
      if (kind === 'delete' && axis === 'row' && this._visibleRowsBetween(index, index + count - 1).length < count) {
        // Deleting would also delete rows the user can't see
        this._modeAnnouncement = 'Cannot delete: some of the rows are hidden by the filter';
        return false;
      }
      const detail: StructureChangeDetail = { ...change, source };
      const event = new CustomEvent<StructureChangeDetail>('before-structure-change', {
        detail,
        bubbles: true,
        composed: true,
        cancelable: true,
      });
      const mutations = this._mutationCount;
      if (!this.dispatchEvent(event)) return false;
      // A listener changed (or replaced) the grid: the checks above no longer hold
      if (this._changedSince(mutations)) return false;
    }

    const rowsBefore = this.rows;
    const colsBefore = this.cols;
    const delta = kind === 'insert' ? count : -count;
    const rowsAfter = axis === 'row' ? rowsBefore + delta : rowsBefore;
    const colsAfter = axis === 'col' ? colsBefore + delta : colsBefore;
    const selectionBefore = this._snapshotSelection();

    const batch: CommandBatch = {
      deltas: [],
      selectionBefore,
      selectionAfter: this._remapSelection(selectionBefore, rowsBefore, colsBefore, rowsAfter, colsAfter),
      operation: STRUCTURE_OPERATIONS[`${kind}:${axis}`],
      structure: {
        change,
        cells: this._computeStructureCells(change),
        rowsBefore,
        colsBefore,
        rowsAfter,
        colsAfter,
        colWidths: this._computeStructureWidths(change),
      },
    };
    if (this._filter) batch.filter = { before: this._filter, after: this._filterAfterStructureChange(change) };
    const rulesAfter = adjustRulesForStructureChange(this._cfRules, change, this._activeState.name);
    if (rulesAfter !== this._cfRules) batch.conditional = { before: this._cfRules, after: rulesAfter };
    this._addCrossSheetStructure(batch.structure!);
    this._commitBatch(batch, source);

    if (source === 'user') {
      const label = (i: number) => (axis === 'row' ? String(i + 1) : colToLetter(i));
      const what =
        count === 1
          ? `${noun} ${label(index)}`
          : `${noun}s ${label(index)} to ${label(index + count - 1)}`;
      this._modeAnnouncement = `${kind === 'insert' ? 'Inserted' : 'Deleted'} ${what}`;
    }
    return true;
  }

  /** Custom column widths after a column insertion or deletion: they move with their columns. */
  private _computeStructureWidths(change: StructureChange): StructureBatch['colWidths'] {
    if (change.axis !== 'col' || this._colWidths.size === 0) return undefined;
    const { kind, index, count } = change;
    const after = new Map<number, number>();
    for (const [col, width] of this._colWidths) {
      if (kind === 'insert') after.set(col >= index ? col + count : col, width);
      else if (col < index) after.set(col, width);
      else if (col >= index + count) after.set(col - count, width);
    }
    return { before: new Map(this._colWidths), after };
  }

  private _hasLockedCellIn(change: StructureChange): boolean {
    for (const [key, cell] of this._internalData) {
      if (!cell.readOnly) continue;
      const { row, col } = parseKey(key);
      const p = change.axis === 'row' ? row : col;
      if (p >= change.index && p < change.index + change.count) return true;
    }
    return false;
  }

  /** Before/after state of every cell a row/column insertion or deletion changes. */
  private _computeStructureCells(change: StructureChange): StructureCellDelta[] {
    const { axis, kind, index, count } = change;
    const next: GridData = new Map();
    for (const [key, shown] of this._internalData) {
      // Spilled values don't move: their formula does, and spills again
      const cell = storedCell(shown);
      if (!cell) continue;
      const { row, col } = parseKey(key);
      const p = axis === 'row' ? row : col;
      let moved = p;
      if (kind === 'insert') {
        if (p >= index) moved = p + count;
      } else if (p >= index + count) {
        moved = p - count;
      } else if (p >= index) {
        continue; // deleted
      }
      const copy = cloneCell(cell);
      copy.rawValue = adjustFormulaForStructureChange(cell.rawValue, change, { formulaSheet: this._activeState.name });
      next.set(axis === 'row' ? cellKey(moved, col) : cellKey(row, moved), copy);
    }

    const cells: StructureCellDelta[] = [];
    for (const id of new Set([...this._internalData.keys(), ...next.keys()])) {
      const before = storedCell(this._internalData.get(id));
      const after = next.get(id);
      if (!before && !after) continue;
      if (
        before &&
        after &&
        before.rawValue === after.rawValue &&
        formatsEqual(before.format, after.format) &&
        !!before.readOnly === !!after.readOnly
      ) {
        continue;
      }
      cells.push({ id, before: before && cloneCell(before), after });
    }
    return cells;
  }

  /**
   * Selection after a structure change: the same cells by position, clamped
   * to the new size. A selection of whole rows (or columns) stays whole.
   */
  private _remapSelection(
    snapshot: SelectionSnapshot,
    rowsBefore: number,
    colsBefore: number,
    rowsAfter: number,
    colsAfter: number
  ): SelectionSnapshot {
    const minCol = Math.min(snapshot.anchor.col, snapshot.head.col);
    const maxCol = Math.max(snapshot.anchor.col, snapshot.head.col);
    const minRow = Math.min(snapshot.anchor.row, snapshot.head.row);
    const maxRow = Math.max(snapshot.anchor.row, snapshot.head.row);
    const fullWidth = minCol === 0 && maxCol === colsBefore - 1;
    const fullHeight = minRow === 0 && maxRow === rowsBefore - 1;
    const remap = (c: CellCoord): CellCoord => ({
      row: fullHeight && c.row === rowsBefore - 1 ? rowsAfter - 1 : Math.min(c.row, rowsAfter - 1),
      col: fullWidth && c.col === colsBefore - 1 ? colsAfter - 1 : Math.min(c.col, colsAfter - 1),
    });
    return {
      anchor: remap(snapshot.anchor),
      head: remap(snapshot.head),
      active: snapshot.active
        ? { row: Math.min(snapshot.active.row, rowsAfter - 1), col: Math.min(snapshot.active.col, colsAfter - 1) }
        : null,
    };
  }

  private _applyStructureBatch(batch: CommandBatch, side: 'before' | 'after', source: ChangeSource): void {
    const structure = batch.structure!;
    for (const cell of structure.cells) {
      const value = side === 'after' ? cell.after : cell.before;
      if (value) {
        this._internalData.set(cell.id, cloneCell(value));
      } else {
        this._internalData.delete(cell.id);
      }
    }
    this._structureRevision++;
    let movedWidths: ColumnResizeDetail['columns'] = [];
    if (structure.colWidths) {
      const from = side === 'after' ? structure.colWidths.before : structure.colWidths.after;
      const to = side === 'after' ? structure.colWidths.after : structure.colWidths.before;
      this._colWidths = new Map(to);
      this._colWidthsRevision++;
      // Columns whose width changed as widths moved with their columns
      const cols = [...new Set([...from.keys(), ...to.keys()])].sort((a, b) => a - b);
      movedWidths = cols
        .filter((c) => from.get(c) !== to.get(c))
        .map((c) => ({ col: c, width: to.get(c) ?? this._cellSize.width }));
    }
    // The change as applied now: undo applies the inverse
    const { change } = structure;
    const applied: StructureChange =
      side === 'after' ? change : { ...change, kind: change.kind === 'insert' ? 'delete' : 'insert' };
    const frozenBefore = { rows: this._frozenRowCount, cols: this._frozenColCount };
    const frozenNext = this._frozenAfterStructureChange(applied);
    this.rows = side === 'after' ? structure.rowsAfter : structure.rowsBefore;
    this.cols = side === 'after' ? structure.colsAfter : structure.colsBefore;
    // Store the effective counts, so a clamped value can't come back later
    if (applied.axis === 'row') this.frozenRows = Math.min(frozenNext, this.rows - 1);
    else this.frozenCols = Math.min(frozenNext, this.cols - 1);
    if (applied.axis !== 'row') this.frozenRows = this._frozenRowCount;
    else this.frozenCols = this._frozenColCount;
    const frozen =
      this._frozenRowCount !== frozenBefore.rows || this._frozenColCount !== frozenBefore.cols
        ? { rows: this._frozenRowCount, cols: this._frozenColCount }
        : null;
    let filterMoved = false;
    if (batch.filter) {
      this._filter = side === 'after' ? batch.filter.after : batch.filter.before;
      this._filterRevision++;
      filterMoved = !this._filtersEqual(batch.filter.before, batch.filter.after);
      if (this._filterMenu) this._closeFilterMenu(false);
    }
    if (batch.conditional) {
      this._setConditionalRules(side === 'after' ? batch.conditional.after : batch.conditional.before);
    }
    this._selection.setBounds(this.rows, this.cols);
    this._lastSyncedDataRef = this.data;

    // Formulas on other sheets and defined names that refer to this sheet move with it
    // Only where the formula is still the one this change wrote (or found)
    const from = side === 'after' ? 'before' : 'after';
    for (const other of structure.others ?? []) {
      const cell = this._sheets.find((s) => s.id === other.sheetId)?.data.get(other.id);
      if (cell && cell.rawValue === other[from]) cell.rawValue = other[side];
    }
    for (const rules of structure.otherRules ?? []) {
      const sheet = this._sheets.find((s) => s.id === rules.sheetId);
      if (sheet && sheet.cfRules === rules[from]) sheet.cfRules = rules[side];
    }
    if (structure.names) {
      const current = new Map(this._formulaEngine.getNames().map((n) => [n.name.toLowerCase(), n]));
      for (const moved of structure.names) {
        const name = current.get(moved.name.toLowerCase());
        if (name && name.formula === moved[from] && name.sheet === moved.sheet) {
          this._formulaEngine.setNameFormula(name.name, moved[side]);
        }
      }
    }

    // Every reference may have moved, so rebuild the dependency graph.
    this._formulaEngine.setData(this._internalData);
    this._recalcAll();
    this._restoreSelection(side === 'after' ? batch.selectionAfter : batch.selectionBefore);

    this._dispatchDataChange({
      updates: structure.cells.map((cell) => ({
        id: cell.id,
        value: (side === 'after' ? cell.after : cell.before)?.rawValue ?? '',
      })),
      source,
      operation: batch.operation,
    });
    this._dispatchStructureChange({ ...change, kind: applied.kind, source });
    if (frozen) this._dispatchFreezeChange({ ...frozen, source });
    if (filterMoved) this._dispatchFilterChange(this._filterDetail(this._filter, source));
    if (batch.conditional) this._dispatchConditionalChange(source);
    if (structure.names) this._dispatchNamesChange(source);
    if (movedWidths.length > 0) {
      // So hosts that save widths on column-resize see them move
      this.dispatchEvent(
        new CustomEvent<ColumnResizeDetail>('column-resize', {
          detail: { columns: movedWidths, source },
          bubbles: true,
          composed: true,
        })
      );
    }
    this.requestUpdate();
  }

  // ─── Sorting ────────────────────────────────────────

  private _sortRange(spec: Omit<SortDetail, 'source'>, source: 'user' | 'programmatic'): boolean {
    if (this._pasteInProgress) return false;
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();
    if (source === 'user' && this.readOnly) return false;

    const { range, hasHeader } = spec;
    const first = range.start.row + (hasHeader ? 1 : 0);
    // Rows hidden by the filter stay where they are: only the visible rows are sorted
    const targets = this._visibleRowsBetween(first, range.end.row);
    const skipped = range.end.row - first + 1 - targets.length;
    const order = this._computeSortOrder(targets, spec.keys);
    const moved = order.filter((src, i) => src !== targets[i]);
    if (moved.length === 0) {
      if (source === 'user') this._modeAnnouncement = 'Already sorted';
      return false;
    }

    if (source === 'user') {
      const locked = moved.some((row) => {
        for (let c = range.start.col; c <= range.end.col; c++) {
          if (this._isCellLocked(cellKey(row, c))) return true;
        }
        return false;
      });
      if (locked) {
        this._modeAnnouncement = 'Cannot sort: the rows to move contain read-only cells';
        return false;
      }
      // A copy, so a listener that mutates the detail can't change what is sorted
      const detail: SortDetail = {
        range: { start: { ...range.start }, end: { ...range.end } },
        keys: spec.keys.map((k) => ({ ...k })),
        hasHeader,
        source,
      };
      const event = new CustomEvent<SortDetail>('before-sort', {
        detail,
        bubbles: true,
        composed: true,
        cancelable: true,
      });
      const mutations = this._mutationCount;
      if (!this.dispatchEvent(event)) {
        this._modeAnnouncement = 'Sort cancelled';
        return false;
      }
      if (this._changedSince(mutations)) {
        this._modeAnnouncement = 'Sort not applied because the data changed';
        return false;
      }
    }

    const cells = this._computeSortCells(order, targets, range.start.col, range.end.col);
    const selection = this._snapshotSelection();
    const batch: CommandBatch = {
      deltas: [],
      selectionBefore: selection,
      selectionAfter: selection,
      operation: 'sort',
      sort: {
        detail: {
          range: { start: { ...range.start }, end: { ...range.end } },
          keys: spec.keys.map((k) => ({ ...k })),
          hasHeader,
        },
        cells,
      },
    };
    this._commitBatch(batch, source);

    if (source === 'user') {
      const sortedRef = `${coordToRef({ row: first, col: range.start.col })}:${coordToRef(range.end)}`;
      const by = spec.keys.map((k) => `column ${colToLetter(k.col)}, ${k.direction}`).join(', then ');
      this._modeAnnouncement =
        `Sorted ${sortedRef} by ${by}` +
        (hasHeader ? `. Row ${range.start.row + 1} kept as the header` : '') +
        (skipped > 0 ? `. ${skipped} hidden row${skipped === 1 ? ' was' : 's were'} not moved` : '');
    }
    return true;
  }

  /** Source row for each of the `targets` rows, in sorted order (stable). */
  private _computeSortOrder(targets: number[], keys: SortDetail['keys']): number[] {
    const rows: Array<{ row: number; values: SortValue[] }> = [];
    for (const row of targets) {
      rows.push({
        row,
        values: keys.map((k) => {
          const key = cellKey(row, k.col);
          return sortValueOf(this._internalData.get(key), (cell) => this._sortNumber(key, cell));
        }),
      });
    }
    // Array.prototype.sort is stable (ES2019)
    rows.sort((a, b) => {
      for (let i = 0; i < keys.length; i++) {
        const result = compareSortValues(a.values[i], b.values[i], keys[i].direction);
        if (result !== 0) return result;
      }
      return 0;
    });
    return rows.map((r) => r.row);
  }

  /** The number behind a number cell, whose display may be formatted. */
  private _sortNumber(key: string, cell: CellData): number {
    if (!cell.format?.numberFormat) return Number(cell.displayValue);
    return Number(this._formulaEngine.resultFor(key, cell.rawValue).displayValue);
  }

  /** Before/after state of every cell in columns `c0..c1` that the sort changes. */
  private _computeSortCells(order: number[], targets: number[], c0: number, c1: number): StructureCellDelta[] {
    const cells: StructureCellDelta[] = [];
    order.forEach((src, i) => {
      const target = targets[i];
      if (src === target) return;
      for (let c = c0; c <= c1; c++) {
        const id = cellKey(target, c);
        const before = storedCell(this._internalData.get(id));
        const source = storedCell(this._internalData.get(cellKey(src, c)));
        let after: CellData | undefined;
        if (source) {
          after = cloneCell(source);
          after.rawValue = this._shiftFormula(source.rawValue, target - src, 0);
        }
        if (
          before &&
          after &&
          before.rawValue === after.rawValue &&
          formatsEqual(before.format, after.format) &&
          !!before.readOnly === !!after.readOnly
        ) {
          continue;
        }
        if (!before && !after) continue;
        cells.push({ id, before: before && cloneCell(before), after });
      }
    });
    return cells;
  }

  private _applySortBatch(batch: CommandBatch, side: 'before' | 'after', source: ChangeSource): void {
    const sort = batch.sort!;
    for (const cell of sort.cells) {
      const value = side === 'after' ? cell.after : cell.before;
      if (value) {
        this._internalData.set(cell.id, cloneCell(value));
      } else {
        this._internalData.delete(cell.id);
      }
    }
    this._lastSyncedDataRef = this.data;

    const ids = sort.cells.map((cell) => cell.id);
    this._formulaEngine.updateData(this._internalData, ids);
    this._recalcAffected(ids);
    this._applyNumberFormatsToKeys(new Set(ids));
    this._restoreSelection(side === 'after' ? batch.selectionAfter : batch.selectionBefore);

    this._dispatchDataChange({
      updates: sort.cells.map((cell) => ({
        id: cell.id,
        value: (side === 'after' ? cell.after : cell.before)?.rawValue ?? '',
      })),
      source,
      operation: 'sort',
    });
    const { range, keys, hasHeader } = sort.detail;
    this._dispatchSortChange({
      range: { start: { ...range.start }, end: { ...range.end } },
      keys: keys.map((k) => ({ ...k })),
      hasHeader,
      source,
    });
    this.requestUpdate();
  }

  /**
   * The block of data around a cell, like Excel's current region: grown
   * while any cell bordering it (diagonals included) is filled. Null when the
   * block is a single empty cell.
   */
  private _currentRegion(row: number, col: number): SelectionRange | null {
    let r0 = row;
    let r1 = row;
    let c0 = col;
    let c1 = col;
    const rowHasData = (r: number) => {
      for (let c = Math.max(0, c0 - 1); c <= Math.min(this.cols - 1, c1 + 1); c++) {
        if (this._isCellFilled(r, c)) return true;
      }
      return false;
    };
    const colHasData = (c: number) => {
      for (let r = Math.max(0, r0 - 1); r <= Math.min(this.rows - 1, r1 + 1); r++) {
        if (this._isCellFilled(r, c)) return true;
      }
      return false;
    };
    let grew = true;
    while (grew) {
      const before = [r0, r1, c0, c1].join();
      if (r0 > 0 && rowHasData(r0 - 1)) r0--;
      if (r1 < this.rows - 1 && rowHasData(r1 + 1)) r1++;
      if (c0 > 0 && colHasData(c0 - 1)) c0--;
      if (c1 < this.cols - 1 && colHasData(c1 + 1)) c1++;
      grew = [r0, r1, c0, c1].join() !== before;
    }
    if (r0 === r1 && c0 === c1 && !this._isCellFilled(row, col)) return null;
    return { start: { row: r0, col: c0 }, end: { row: r1, col: c1 } };
  }

  /**
   * A selection trimmed to its last filled row and column, so sorting whole
   * columns doesn't move blank rows around. Keeps the active cell's column
   * (the sort key) inside. Null when the selection has no data.
   */
  private _clipToData(range: SelectionRange, active: CellCoord): SelectionRange | null {
    let lastRow = -1;
    let lastCol = -1;
    for (const [key, cell] of this._internalData) {
      if (!hasContent(cell)) continue;
      const { row, col } = parseKey(key);
      if (row < range.start.row || row > range.end.row || col < range.start.col || col > range.end.col) continue;
      if (row > lastRow) lastRow = row;
      if (col > lastCol) lastCol = col;
    }
    if (lastRow < 0) return null;
    return {
      start: { ...range.start },
      end: { row: lastRow, col: Math.max(lastCol, active.col) },
    };
  }

  /**
   * Whether the range's first row looks like column headers: every cell is
   * non-empty text, and either some column has non-text data below it or
   * the first row is bold while the second isn't.
   */
  private _looksLikeHeader(range: SelectionRange): boolean {
    const { start, end } = range;
    if (end.row <= start.row) return false;
    for (let c = start.col; c <= end.col; c++) {
      const cell = this._internalData.get(cellKey(start.row, c));
      if (!hasContent(cell) || cell.type !== 'text') return false;
    }
    for (const [key, cell] of this._internalData) {
      if (!hasContent(cell) || cell.type === 'text') continue;
      const { row, col } = parseKey(key);
      if (row > start.row && row <= end.row && col >= start.col && col <= end.col) return true;
    }
    const bold = (row: number) => {
      for (let c = start.col; c <= end.col; c++) {
        if (!this._internalData.get(cellKey(row, c))?.format?.bold) return false;
      }
      return true;
    };
    return bold(start.row) && !bold(start.row + 1);
  }

  // ─── Filtering ──────────────────────────────────────

  /** Normalize a range given to an API method; throws if it isn't inside the grid. */
  private _normalizeRange(range: SelectionRange, method: string): SelectionRange {
    const start = range?.start;
    const end = range?.end;
    const isIndex = (n: unknown): n is number => Number.isInteger(n) && (n as number) >= 0;
    if (!start || !end || ![start.row, start.col, end.row, end.col].every(isIndex)) {
      throw new RangeError(`${method}: range must have non-negative integer start and end coordinates`);
    }
    const normalized: SelectionRange = {
      start: { row: Math.min(start.row, end.row), col: Math.min(start.col, end.col) },
      end: { row: Math.max(start.row, end.row), col: Math.max(start.col, end.col) },
    };
    if (normalized.end.row >= this.rows || normalized.end.col >= this.cols) {
      throw new RangeError(`${method}: the range is outside the grid (${this.rows} rows, ${this.cols} columns)`);
    }
    return normalized;
  }

  /** The displayed value a filter compares for a cell. */
  private _filterValue(row: number, col: number): string {
    return this._internalData.get(cellKey(row, col))?.displayValue ?? '';
  }

  /**
   * Whether a row meets every column's criteria, optionally leaving one
   * column out. A cell that holds several values matches when any one of
   * them is kept, so a list can be filtered by its items.
   */
  private _rowMatches(row: number, criteria: Map<number, Set<string>>, ignoreCol = -1): boolean {
    for (const [col, values] of criteria) {
      if (col === ignoreCol) continue;
      const array = this._internalData.get(cellKey(row, col))?.array;
      if (array) {
        if (!arrayScalars(array).some((item) => values.has(arrayItemText(item)))) return false;
      } else if (!values.has(this._filterValue(row, col))) {
        return false;
      }
    }
    return true;
  }

  /** A filter with the rows it hides worked out from the current data. */
  private _evaluateFilter(range: SelectionRange, criteria: Map<number, Set<string>>): FilterState {
    const hidden = new Set<number>();
    if (criteria.size > 0) {
      for (let r = range.start.row + 1; r <= Math.min(range.end.row, this.rows - 1); r++) {
        if (!this._rowMatches(r, criteria)) hidden.add(r);
      }
    }
    return { range: { start: { ...range.start }, end: { ...range.end } }, criteria, hidden };
  }

  /**
   * A filter range grown down over rows of data added right below it, as
   * Excel does when the user changes or reapplies a filter.
   */
  private _extendFilterRange(range: SelectionRange): SelectionRange {
    const filled = (r: number) => {
      for (let c = range.start.col; c <= range.end.col; c++) if (this._isCellFilled(r, c)) return true;
      return false;
    };
    let end = range.end.row;
    while (end + 1 < this.rows && filled(end + 1)) end++;
    return { start: { ...range.start }, end: { row: end, col: range.end.col } };
  }

  private _filterInfo(filter: FilterState): FilterInfo {
    return {
      range: { start: { ...filter.range.start }, end: { ...filter.range.end } },
      columns: [...filter.criteria.keys()]
        .sort((a, b) => a - b)
        .map((col) => ({ col, values: [...filter.criteria.get(col)!] })),
      hiddenRows: [...filter.hidden].sort((a, b) => a - b),
    };
  }

  private _filterDetail(filter: FilterState | null, source: FilterChangeDetail['source']): FilterChangeDetail {
    if (!filter) return { range: null, columns: [], hiddenRows: [], source };
    return { ...this._filterInfo(filter), source };
  }

  private _filtersEqual(a: FilterState | null, b: FilterState | null): boolean {
    if (a === b) return true;
    if (!a || !b) return false;
    const sameSet = (x: Set<unknown>, y: Set<unknown>) => x.size === y.size && [...x].every((v) => y.has(v));
    const ra = a.range;
    const rb = b.range;
    return (
      ra.start.row === rb.start.row &&
      ra.start.col === rb.start.col &&
      ra.end.row === rb.end.row &&
      ra.end.col === rb.end.col &&
      a.criteria.size === b.criteria.size &&
      [...a.criteria].every(([col, values]) => b.criteria.has(col) && sameSet(values, b.criteria.get(col)!)) &&
      sameSet(a.hidden, b.hidden)
    );
  }

  /**
   * Change the filter as one undoable step. User changes can be vetoed via
   * `before-filter`. Allowed in read-only mode: filtering changes the view,
   * not the data.
   */
  private _commitFilter(next: FilterState | null, source: 'user' | 'programmatic'): 'applied' | 'unchanged' | 'cancelled' {
    if (this._pasteInProgress) return 'unchanged';
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();
    const before = this._filter;
    if (this._filtersEqual(before, next)) return 'unchanged';

    if (source === 'user') {
      const event = new CustomEvent<FilterChangeDetail>('before-filter', {
        detail: this._filterDetail(next, source),
        bubbles: true,
        composed: true,
        cancelable: true,
      });
      const mutations = this._mutationCount;
      const revision = this._filterRevision;
      if (!this.dispatchEvent(event)) return 'cancelled';
      // A listener changed the grid or the filter: `next` was worked out from stale state
      if (this._changedSince(mutations) || revision !== this._filterRevision) return 'cancelled';
    }

    const selectionBefore = this._snapshotSelection();
    let selectionAfter = selectionBefore;
    const active = this._selection.activeCell;
    if (next?.hidden.has(active.row)) {
      // The active cell's row is about to hide: move to the nearest visible row
      const cell = { row: this._nearestVisibleRow(active.row, next.hidden), col: active.col };
      selectionAfter = { anchor: { ...cell }, head: { ...cell }, active: null };
    }
    this._commitBatch(
      // `operation` is never reported for filter batches: they fire filter-change, not data-change
      { deltas: [], selectionBefore, selectionAfter, operation: 'format', filter: { before, after: next } },
      source
    );
    return 'applied';
  }

  /** Announce a user filter change; returns whether it was applied. */
  private _announceFilterCommit(result: ReturnType<Y11nSpreadsheet['_commitFilter']>, applied?: string): boolean {
    if (result === 'cancelled') {
      this._modeAnnouncement = 'Filter change cancelled';
    } else if (result === 'unchanged') {
      this._modeAnnouncement = 'Filter unchanged';
    } else {
      this._modeAnnouncement = applied ?? this._filterSummary();
    }
    return result === 'applied';
  }

  /** "Showing 12 of 40 rows", for the current filter's data rows. */
  private _filterSummary(): string {
    const filter = this._filter;
    if (!filter) return 'Filter removed';
    const total = Math.max(0, Math.min(filter.range.end.row, this.rows - 1) - filter.range.start.row);
    const shown = total - filter.hidden.size;
    return filter.criteria.size === 0 ? `Showing all ${total} rows` : `Showing ${shown} of ${total} rows`;
  }

  private _applyFilterBatch(batch: CommandBatch, side: 'before' | 'after', source: ChangeSource): void {
    const hadFocus = this._gridHasFocus();
    this._filter = side === 'after' ? batch.filter!.after : batch.filter!.before;
    this._filterRevision++;
    if (this._filterMenu) this._closeFilterMenu(false);
    const selection = side === 'after' ? batch.selectionAfter : batch.selectionBefore;
    if (source === 'undo' || source === 'redo') {
      this._restoreSelection(selection);
    } else if (this._isRowHidden(this._selection.activeCell.row)) {
      // Don't move the selection (or focus) unless the active cell was hidden
      this._selection.setSelection(selection.anchor, selection.head, selection.active);
      this._dispatchSelectionChange();
      if (hadFocus) this.updateComplete.then(() => this._focusActiveCell());
    }
    this.requestUpdate();
    this.updateComplete.then(() => {
      // Fewer rows can shrink the scroll range under the current position
      const gridEl = this._grid;
      if (gridEl && gridEl.scrollTop !== this._scrollTop) this._scrollTop = gridEl.scrollTop;
      if (source === 'user') {
        const { row, col } = this._selection.activeCell;
        this._ensureCellVisible(row, col);
      }
    });
    this._dispatchFilterChange(this._filterDetail(this._filter, source));
  }

  private _dispatchFilterChange(detail: FilterChangeDetail): void {
    this.dispatchEvent(new CustomEvent<FilterChangeDetail>('filter-change', { detail, bubbles: true, composed: true }));
  }

  /** Whether focus is on something inside the grid (a cell). */
  private _gridHasFocus(): boolean {
    const active = this.shadowRoot?.activeElement as HTMLElement | null | undefined;
    return !!active && !!this._grid?.contains(active);
  }

  /**
   * The filter after a row/column insertion or deletion: it moves with its
   * cells like a range reference. Rows inserted inside it are visible;
   * deleting its header row (or all its columns) removes it.
   */
  private _filterAfterStructureChange(change: StructureChange): FilterState | null {
    const filter = this._filter;
    if (!filter) return null;
    const { axis, kind, index, count } = change;
    // A position after the change, or null if it was deleted
    const move = (p: number): number | null => {
      if (kind === 'insert') return p >= index ? p + count : p;
      if (p >= index + count) return p - count;
      return p >= index ? null : p;
    };
    // The last surviving position at or before `p`, for a range end
    const moveEnd = (p: number): number => {
      if (kind === 'insert') return p >= index ? p + count : p;
      return p >= index + count ? p - count : Math.min(p, index - 1);
    };
    const { start, end } = filter.range;
    if (axis === 'row') {
      const header = move(start.row);
      if (header === null) return null;
      const hidden = new Set<number>();
      for (const r of filter.hidden) {
        const moved = move(r);
        if (moved !== null) hidden.add(moved);
      }
      return {
        range: { start: { row: header, col: start.col }, end: { row: moveEnd(end.row), col: end.col } },
        criteria: filter.criteria,
        hidden,
      };
    }
    // Columns: the first surviving column, the last surviving one
    let first = start.col;
    while (first <= end.col && move(first) === null) first++;
    if (first > end.col) return null;
    const criteria = new Map<number, Set<string>>();
    for (const [col, values] of filter.criteria) {
      const moved = move(col);
      if (moved !== null) criteria.set(moved, values);
    }
    return {
      range: { start: { row: start.row, col: move(first)! }, end: { row: end.row, col: moveEnd(end.col) } },
      criteria,
      hidden: filter.hidden,
    };
  }

  // ─── Filter Menu ────────────────────────────────────

  /**
   * The distinct displayed values the menu lists for a column: from the data
   * rows that pass the other columns' criteria, in ascending sort order with
   * blanks last.
   */
  private _filterMenuValues(col: number): string[] {
    const filter = this._filter!;
    const range = this._extendFilterRange(filter.range);
    // One cell per distinct value, to sort the values by
    const cells = new Map<string, SortValue>();
    for (let r = range.start.row + 1; r <= range.end.row; r++) {
      if (!this._rowMatches(r, filter.criteria, col)) continue;
      const key = cellKey(r, col);
      const cell = this._internalData.get(key);
      if (cell?.array) {
        // A cell that holds several values lists each of them
        for (const item of arrayScalars(cell.array)) {
          const text = arrayItemText(item);
          if (cells.has(text)) continue;
          cells.set(
            text,
            text === ''
              ? { kind: 'blank' }
              : typeof item === 'number'
                ? { kind: 'number', value: item }
                : typeof item === 'boolean'
                  ? { kind: 'boolean', value: item }
                  : { kind: 'text', value: text }
          );
        }
        continue;
      }
      const value = cell?.displayValue ?? '';
      if (!cells.has(value)) {
        cells.set(value, value === '' ? { kind: 'blank' } : this._cfValue(key));
      }
    }
    const entries = [...cells].map(([value, sort]) => ({ value, sort }));
    entries.sort((a, b) => compareSortValues(a.sort, b.sort, 'ascending') || (a.value < b.value ? -1 : a.value > b.value ? 1 : 0));
    return entries.map((e) => e.value);
  }

  /** The values checked when the menu opens: those the criteria keep, or all. */
  private _checkedForCriteria(values: string[], criteria: Set<string> | undefined): Set<string> {
    return new Set(criteria ? values.filter((v) => criteria.has(v)) : values);
  }

  // ─── Spilled values ─────────────────────────────────

  /**
   * The range a formula's array result spills over, when the cell at
   * (row, col) is part of one: the formula's own cell or a cell showing one
   * of its values.
   */
  private _spillRangeAt(row: number, col: number): SelectionRange | null {
    const cell = this._getCell(row, col);
    if (!cell) return null;
    let anchor: CellCoord = { row, col };
    let spill = cell.spill;
    if (cell.spilledFrom !== undefined) {
      anchor = parseKey(cell.spilledFrom);
      spill = this._internalData.get(cell.spilledFrom)?.spill;
    }
    if (!spill) return null;
    return { start: anchor, end: { row: anchor.row + spill.rows - 1, col: anchor.col + spill.cols - 1 } };
  }

  /** Classes that outline the spill range the active cell belongs to. */
  private _spillClasses(row: number, col: number): string {
    const spill = this._activeSpill;
    if (!spill || row < spill.start.row || row > spill.end.row || col < spill.start.col || col > spill.end.col) {
      return '';
    }
    return (
      'in-spill' +
      (row === spill.start.row ? ' spill-top' : '') +
      (row === spill.end.row ? ' spill-bottom' : '') +
      (col === spill.start.col ? ' spill-left' : '') +
      (col === spill.end.col ? ' spill-right' : '')
    );
  }

  // ─── Lists and arrays in cells ──────────────────────

  /** Ctrl+J: turn the selected cells into lists, or back into text when they all are lists. */
  private _toggleListsInSelection(): void {
    if (this.readOnly || this._pasteInProgress) return;
    const range = this._selection.range;
    let lists = 0;
    let others = 0;
    for (const r of this._visibleRowsBetween(range.start.row, range.end.row)) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        const raw = this._internalData.get(cellKey(r, c))?.rawValue ?? '';
        if (raw === '') continue;
        if (listItemsOf(raw)) lists++;
        else if (!raw.startsWith('=')) others++;
      }
    }
    if (lists === 0 && others === 0) {
      // Nothing to convert: type the items of a new list
      this._startEditing('');
      if (this._isEditing) {
        this._editListMode = true;
        this._modeAnnouncement = `Editing ${coordToRef(this._selection.activeCell)} as a list. Separate the items with commas`;
      }
      return;
    }
    this._convertLists(range, others > 0, 'user');
  }

  /**
   * Convert the cells of a range between text and lists. Returns how many
   * cells changed. Formulas are left alone.
   */
  private _convertLists(range: SelectionRange, toList: boolean, source: 'user' | 'programmatic'): number {
    this._syncPendingData();
    if (this._isEditing) this._commitEdit();
    if (source === 'user' && this.readOnly) return 0;
    const updates: Array<{ id: string; value: string }> = [];
    for (const r of this._visibleRowsBetween(range.start.row, Math.min(range.end.row, this.rows - 1))) {
      for (let c = range.start.col; c <= Math.min(range.end.col, this.cols - 1); c++) {
        const id = cellKey(r, c);
        const raw = this._internalData.get(id)?.rawValue ?? '';
        if (raw === '') continue;
        const items = listItemsOf(raw);
        if (toList) {
          if (items || raw.startsWith('=')) continue;
          const parsed = parseListText(raw);
          if (parsed.length > 0) updates.push({ id, value: listToFormula(parsed) });
        } else if (items) {
          updates.push({ id, value: listToText(items) });
        }
      }
    }
    const selection = this._snapshotSelection();
    let batch = this._buildCommandBatch(updates, 'edit', selection, selection);
    if (batch && source === 'user') batch = this._vetUserBatch(batch);
    if (!batch) {
      if (source === 'user') this._modeAnnouncement = toList ? 'No text to turn into lists' : 'No lists to turn into text';
      return 0;
    }
    this._commitBatch(batch, source);
    const count = batch.deltas.length;
    if (source === 'user') {
      this._modeAnnouncement = toList
        ? `Turned ${count} cell${count === 1 ? '' : 's'} into ${count === 1 ? 'a list' : 'lists'}`
        : `Turned ${count} list${count === 1 ? '' : 's'} into text`;
    }
    this.dispatchEvent(
      new CustomEvent<ListConversionDetail>('list-change', {
        detail: { cellIds: batch.deltas.map((d) => d.id), toList, source },
        bubbles: true,
        composed: true,
      })
    );
    return count;
  }

  private _handleArrayButtonPointerDown(e: PointerEvent): void {
    if (e.button !== 0) return;
    if (this._endEditElsewhere()) {
      e.preventDefault();
      e.stopPropagation();
      return;
    }
    const cellEl = (e.currentTarget as HTMLElement).closest<HTMLElement>('.ls-cell');
    if (!cellEl) return;
    e.preventDefault();
    e.stopPropagation();
    const key = cellEl.dataset.key!;
    if (this._arrayCard?.key === key) {
      this.closeCellValues(true);
      return;
    }
    if (this._isEditing) this._commitEdit();
    const coord = parseKey(key);
    this._selection.setSelection(coord, coord, null);
    this._dispatchSelectionChange();
    void this.showCellValues(key);
  }

  private _handleArrayCardKeydown(e: KeyboardEvent): void {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      this.closeCellValues(true);
    } else if (e.key === 'Backspace' && this._arrayCard && this._arrayCard.path.length > 0) {
      e.preventDefault();
      this._arrayCardGo(this._arrayCard.path.slice(0, -1));
    } else if (e.key === 'Tab') {
      // Modal: keep focus inside
      const card = e.currentTarget as HTMLElement;
      const focusables = [...card.querySelectorAll<HTMLElement>('button, [tabindex="0"]')];
      const first = focusables[0];
      const last = focusables[focusables.length - 1];
      const current = this.shadowRoot?.activeElement;
      if (e.shiftKey && current === first) {
        e.preventDefault();
        last?.focus();
      } else if (!e.shiftKey && current === last) {
        e.preventDefault();
        first?.focus();
      }
    }
  }

  /** Show another level of the values popup (into a nested array, or back out). */
  private _arrayCardGo(path: number[]): void {
    if (!this._arrayCard) return;
    this._arrayCard = { key: this._arrayCard.key, path };
    this.updateComplete.then(() => this.renderRoot.querySelector<HTMLElement>('.ls-array-card-body')?.focus());
  }

  private _renderArrayCard(card: { key: string; path: number[] }) {
    const root = this._internalData.get(card.key)?.array;
    if (!root) return nothing;
    let array = root;
    const trail: string[] = [];
    for (const index of card.path) {
      const item = array.values[index];
      if (item === null || typeof item !== 'object') break;
      trail.push(`item ${index + 1}`);
      array = item;
    }
    const { row, col } = parseKey(card.key);
    const ref = coordToRef({ row, col });
    const title = [ref, ...trail].join(' › ');
    const isList = array.cols === 1 || array.rows === 1;
    const maxRows = Y11nSpreadsheet.ARRAY_CARD_MAX_ROWS;
    const maxCols = Y11nSpreadsheet.ARRAY_CARD_MAX_COLS;
    const total = array.rows * array.cols;
    const item = (index: number) => {
      const value = array.values[index];
      if (value !== null && typeof value === 'object') {
        return html`<button
          type="button"
          class="ls-array-nested"
          @click="${() => this._arrayCardGo([...card.path.slice(0, trail.length), index])}"
        >${describeArray(value)}: ${arrayText(value)}</button>`;
      }
      return value === null ? html`<span class="ls-array-blank">(blank)</span>` : arrayItemText(value);
    };

    let body;
    let more = 0;
    if (isList) {
      const shown = Math.min(total, maxRows);
      more = total - shown;
      body = html`<ol class="ls-array-list">
        ${Array.from({ length: shown }, (_, i) => html`<li>${item(i)}</li>`)}
      </ol>`;
    } else {
      const shownRows = Math.min(array.rows, maxRows);
      const shownCols = Math.min(array.cols, maxCols);
      more = total - shownRows * shownCols;
      body = html`<table class="ls-array-table">
        <thead>
          <tr>
            <td></td>
            ${Array.from({ length: shownCols }, (_, c) => html`<th scope="col">${c + 1}</th>`)}
          </tr>
        </thead>
        <tbody>
          ${Array.from(
            { length: shownRows },
            (_, r) => html`<tr>
              <th scope="row">${r + 1}</th>
              ${Array.from({ length: shownCols }, (_, c) => html`<td>${item(r * array.cols + c)}</td>`)}
            </tr>`
          )}
        </tbody>
      </table>`;
    }

    return html`
      <div
        class="ls-array-card"
        role="dialog"
        aria-modal="true"
        aria-label="Values in ${title}"
        style=${styleMap(this._getFilterMenuStyle(col, row))}
        @keydown="${this._handleArrayCardKeydown}"
      >
        <div class="ls-array-card-head">
          ${trail.length > 0
            ? html`<button
                type="button"
                class="ls-array-back"
                @click="${() => this._arrayCardGo(card.path.slice(0, trail.length - 1))}"
              >Back</button>`
            : nothing}
          <span class="ls-array-card-title" id="array-card-title">${title}: ${describeArray(array)}</span>
          <button type="button" class="ls-array-close" aria-label="Close" @click="${() => this.closeCellValues(true)}">×</button>
        </div>
        <div class="ls-array-card-body" tabindex="0" role="group" aria-labelledby="array-card-title">
          ${body}
          ${more > 0 ? html`<p class="ls-array-more">and ${more} more</p>` : nothing}
        </div>
      </div>
    `;
  }

  /** The menu's values that match its search text (case-insensitive). */
  private _filterMenuMatches(menu: FilterMenuState): string[] {
    const needle = menu.search.trim().toLowerCase();
    if (!needle) return menu.values;
    return menu.values.filter((v) => this._filterValueLabel(v).toLowerCase().includes(needle));
  }

  private _filterValueLabel(value: string): string {
    return value === '' ? '(Blanks)' : value;
  }

  private _closeFilterMenu(focusCell: boolean): void {
    this._filterMenuCleanup?.();
    this._filterMenuCleanup = null;
    if (!this._filterMenu) return;
    this._filterMenu = null;
    if (focusCell) this.updateComplete.then(() => this._focusActiveCell());
  }

  /** OK: filter the column to the checked values (with a search, only the checked matches). */
  private _applyFilterMenu(): void {
    const menu = this._filterMenu;
    const filter = this._filter;
    if (!menu || !filter) return;
    const matches = this._filterMenuMatches(menu);
    const keep = matches.filter((v) => menu.checked.has(v));
    if (keep.length === 0) return;
    const criteria = new Map(filter.criteria);
    if (!menu.search.trim() && keep.length === menu.values.length) criteria.delete(menu.col);
    else criteria.set(menu.col, new Set(keep));
    this._closeFilterMenu(true);
    const next = this._evaluateFilter(this._extendFilterRange(filter.range), criteria);
    this._announceFilterCommit(this._commitFilter(next, 'user'));
  }

  private _clearColumnFromFilterMenu(): void {
    const menu = this._filterMenu;
    const filter = this._filter;
    if (!menu || !filter) return;
    const criteria = new Map(filter.criteria);
    criteria.delete(menu.col);
    this._closeFilterMenu(true);
    const next = this._evaluateFilter(this._extendFilterRange(filter.range), criteria);
    this._announceFilterCommit(this._commitFilter(next, 'user'));
  }

  private _sortFromFilterMenu(direction: SortDirection): void {
    const menu = this._filterMenu;
    const filter = this._filter;
    if (!menu || !filter || this.readOnly) return;
    this._closeFilterMenu(true);
    this._sortRange({ range: filter.range, keys: [{ col: menu.col, direction }], hasHeader: true }, 'user');
  }

  /** Ctrl+Alt+L: re-filter the rows with the current criteria. */
  private _reapplyUserFilter(): void {
    const filter = this._filter;
    if (!filter) {
      this._modeAnnouncement = 'There is no filter to reapply';
      return;
    }
    const next = this._evaluateFilter(this._extendFilterRange(filter.range), filter.criteria);
    this._announceFilterCommit(this._commitFilter(next, 'user'));
  }

  private _updateFilterMenu(changes: Partial<FilterMenuState>): void {
    if (this._filterMenu) this._filterMenu = { ...this._filterMenu, ...changes };
  }

  private _handleFilterSearchInput(e: Event): void {
    const menu = this._filterMenu;
    if (!menu) return;
    const search = (e.target as HTMLInputElement).value;
    const next = { ...menu, search, focus: 0 };
    // As in Excel: searching checks every match; clearing the search goes back to the criteria
    next.checked = search.trim()
      ? new Set(this._filterMenuMatches(next))
      : this._checkedForCriteria(menu.values, this._filter?.criteria.get(menu.col));
    this._filterMenu = next;
  }

  private _toggleFilterValue(value: string): void {
    const menu = this._filterMenu;
    if (!menu) return;
    const checked = new Set(menu.checked);
    if (checked.has(value)) checked.delete(value);
    else checked.add(value);
    this._updateFilterMenu({ checked });
  }

  private _toggleAllFilterValues(): void {
    const menu = this._filterMenu;
    if (!menu) return;
    const matches = this._filterMenuMatches(menu);
    const all = matches.every((v) => menu.checked.has(v));
    const checked = new Set(menu.checked);
    for (const v of matches) {
      if (all) checked.delete(v);
      else checked.add(v);
    }
    this._updateFilterMenu({ checked });
  }

  private _handleFilterMenuKeydown(e: KeyboardEvent): void {
    const menu = this._filterMenu;
    if (!menu) return;
    // Keep keys from reaching the page while the menu has focus
    e.stopPropagation();
    const target = e.target as HTMLElement;
    const menuEl = e.currentTarget as HTMLElement;

    if (e.key === 'Escape') {
      e.preventDefault();
      this._closeFilterMenu(true);
      this._modeAnnouncement = 'Filter menu closed';
      return;
    }
    if (e.key === 'Tab') {
      // Keep focus inside the menu (it is modal)
      const focusables = [...menuEl.querySelectorAll<HTMLElement>('button, input')].filter(
        (el) => !(el as HTMLButtonElement).disabled && el.tabIndex >= 0
      );
      const first = focusables[0];
      const last = focusables[focusables.length - 1];
      const current = this.shadowRoot?.activeElement;
      if (e.shiftKey && current === first) {
        e.preventDefault();
        last?.focus();
      } else if (!e.shiftKey && current === last) {
        e.preventDefault();
        first?.focus();
      }
      return;
    }

    const shown = Math.min(this._filterMenuMatches(menu).length, Y11nSpreadsheet.FILTER_MENU_MAX_VALUES);
    const inList = target.dataset.filterItem !== undefined;
    const fromSearch = target.id === 'filter-search' && e.key === 'ArrowDown';
    if (fromSearch || (inList && ['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key))) {
      e.preventDefault();
      let focus = menu.focus;
      if (!fromSearch) {
        if (e.key === 'ArrowDown') focus = Math.min(shown, focus + 1);
        else if (e.key === 'ArrowUp') focus = Math.max(0, focus - 1);
        else if (e.key === 'Home') focus = 0;
        else focus = shown;
      }
      this._updateFilterMenu({ focus });
      this.updateComplete.then(() => {
        this.renderRoot.querySelector<HTMLElement>(`[data-filter-item="${focus}"]`)?.focus();
      });
      return;
    }
    // Enter anywhere but on a button means OK, as in Excel
    if (e.key === 'Enter' && target.tagName !== 'BUTTON') {
      e.preventDefault();
      this._applyFilterMenu();
    }
  }

  private _handleFilterButtonPointerDown(e: PointerEvent): void {
    if (e.button !== 0) return;
    if (this._endEditElsewhere()) {
      e.preventDefault();
      e.stopPropagation();
      return;
    }
    // Keep the cell underneath from starting a selection, and the
    // outside-click listener from closing the menu this button toggles
    e.preventDefault();
    e.stopPropagation();
    const col = parseInt((e.currentTarget as HTMLElement).dataset.filterCol!, 10);
    if (this._filterMenu?.col === col) {
      this._closeFilterMenu(true);
      return;
    }
    void this.openFilterMenu(col);
  }

  /** Whether cell (row, col) is a header cell of the filter, with a filter button. */
  private _isFilterHeader(row: number, col: number): boolean {
    const range = this._filter?.range;
    return !!range && row === range.start.row && col >= range.start.col && col <= range.end.col;
  }

  // ─── Row Geometry (hidden rows) ─────────────────────

  /** Whether the filter hides a row. */
  private _isRowHidden(row: number): boolean {
    return this._filter !== null && this._filter.hidden.has(row);
  }

  /** The rows not hidden by the filter, ascending, or null when no row is hidden. */
  private _visibleRowList(): Int32Array | null {
    const hidden = this._filter?.hidden;
    if (!hidden || hidden.size === 0) return null;
    const key = `${this.rows}|${this._filterRevision}`;
    if (this._rowLayout?.key !== key) {
      const visible: number[] = [];
      for (let r = 0; r < this.rows; r++) if (!hidden.has(r)) visible.push(r);
      this._rowLayout = { key, visible: Int32Array.from(visible) };
    }
    return this._rowLayout.visible;
  }

  /** Number of visible rows before `row`: its position in the laid-out grid. */
  private _visibleIndex(row: number): number {
    const visible = this._visibleRowList();
    const r = Math.max(0, Math.min(row, this.rows));
    return visible ? lowerBound(visible, r) : r;
  }

  /** Top offset in px of a row within the data area; hidden rows take no height. */
  private _rowTop(row: number): number {
    return this._visibleIndex(row) * this._cellSize.height;
  }

  /** The visible rows from `from` to `to` (inclusive), ascending. */
  private _visibleRowsBetween(from: number, to: number): number[] {
    const rows: number[] = [];
    for (let r = Math.max(0, from); r <= Math.min(to, this.rows - 1); r++) {
      if (!this._isRowHidden(r)) rows.push(r);
    }
    return rows;
  }

  /** The visible rows of a range, or undefined when none of them is hidden. */
  private _visibleRowsIn(range: SelectionRange): number[] | undefined {
    if (!this._filter || this._filter.hidden.size === 0) return undefined;
    const rows = this._visibleRowsBetween(range.start.row, range.end.row);
    return rows.length === range.end.row - range.start.row + 1 ? undefined : rows;
  }

  /** `row` if it is visible, else the nearest visible row below it, else above it. */
  private _nearestVisibleRow(row: number, hidden: Set<number> | undefined = this._filter?.hidden): number {
    if (!hidden?.has(row)) return row;
    for (let r = row + 1; r < this.rows; r++) if (!hidden.has(r)) return r;
    for (let r = row - 1; r >= 0; r--) if (!hidden.has(r)) return r;
    return row;
  }

  /** `row` if it is visible, else the nearest visible row in direction `dir` (falling back to the other way). */
  private _visibleRowToward(row: number, dir: 1 | -1): number {
    for (let r = row; r >= 0 && r < this.rows; r += dir) if (!this._isRowHidden(r)) return r;
    return this._nearestVisibleRow(row);
  }

  // ─── Frozen Panes ───────────────────────────────────

  /** Frozen rows in effect: the property clamped so at least one row scrolls. */
  private get _frozenRowCount(): number {
    const n = Math.floor(Number(this.frozenRows));
    return Number.isFinite(n) ? Math.max(0, Math.min(n, this.rows - 1)) : 0;
  }

  private get _frozenColCount(): number {
    const n = Math.floor(Number(this.frozenCols));
    return Number.isFinite(n) ? Math.max(0, Math.min(n, this.cols - 1)) : 0;
  }

  /** Height in px of the frozen rows (those not hidden by a filter), below the column header. */
  private get _frozenHeight(): number {
    return this._rowTop(this._frozenRowCount);
  }

  /** Width in px of the frozen columns, right of the row header. */
  private get _frozenWidth(): number {
    return this._colLeft(this._frozenColCount);
  }

  private _setFrozen(rows: number, cols: number, source: FreezeChangeDetail['source']): boolean {
    const r = Math.max(0, Math.min(rows, this.rows - 1));
    const c = Math.max(0, Math.min(cols, this.cols - 1));
    if (r === this._frozenRowCount && c === this._frozenColCount) {
      // Normalize an out-of-range property value without an event
      this.frozenRows = r;
      this.frozenCols = c;
      return false;
    }
    this.frozenRows = r;
    this.frozenCols = c;
    // The frozen rows/columns now cover the start of the viewport, which may
    // hide the active cell: scroll it clear (without moving focus), after the
    // render applies the new scroll padding its snap position depends on
    this.updateComplete.then(() => {
      const { row, col } = this._selection.activeCell;
      this._ensureCellVisible(row, col);
    });
    this._dispatchFreezeChange({ rows: r, cols: c, source });
    return true;
  }

  private _dispatchFreezeChange(detail: FreezeChangeDetail): void {
    this.dispatchEvent(new CustomEvent<FreezeChangeDetail>('freeze-change', { detail, bubbles: true, composed: true }));
  }

  /** Alt+Shift+F: unfreeze, or freeze the rows above and columns left of the active cell. */
  private _toggleFreeze(): void {
    if (this._frozenRowCount > 0 || this._frozenColCount > 0) {
      this._freezeAsUser(0, 0);
      return;
    }
    const { row, col } = this._selection.activeCell;
    this._freezeAsUser(row, col);
  }

  /**
   * Freeze `row` rows and `col` columns as the user (0, 0 unfreezes): refused
   * when the frozen area would not leave room to scroll, and announced.
   */
  private _freezeAsUser(rows: number, cols: number): boolean {
    const row = Math.max(0, Math.min(rows, this.rows - 1));
    const col = Math.max(0, Math.min(cols, this.cols - 1));
    if (row === 0 && col === 0) {
      if (this._frozenRowCount === 0 && this._frozenColCount === 0) {
        this._modeAnnouncement = 'Nothing to freeze: move below or right of the rows and columns to keep in view';
        return false;
      }
      this._setFrozen(0, 0, 'user');
      this._modeAnnouncement = 'Panes unfrozen';
      return true;
    }
    const gridEl = this._grid;
    const h = this._cellSize.height;
    // Leave at least two rows (one column) of scrolling area, when the grid has a size
    if (
      gridEl?.clientHeight &&
      (this._rowTop(row) > gridEl.clientHeight - 3 * h ||
        this._colLeft(col) > gridEl.clientWidth - Y11nSpreadsheet.ROW_HEADER_WIDTH - this._colWidth(col))
    ) {
      this._modeAnnouncement = 'Cannot freeze: the frozen area would not fit on screen';
      return false;
    }
    const changed = this._setFrozen(row, col, 'user');
    const parts: string[] = [];
    if (row > 0) parts.push(row === 1 ? 'row 1' : `rows 1 to ${row}`);
    if (col > 0) parts.push(col === 1 ? 'column A' : `columns A to ${colToLetter(col - 1)}`);
    this._modeAnnouncement = `Froze ${parts.join(' and ')}`;
    return changed;
  }

  /**
   * Frozen count on `change.axis` after a row/column insertion or deletion
   * (the freeze line follows the content, as in Excel), from the current
   * effective count. Undo/redo pass the inverse change, so a freeze made
   * after the original change is respected.
   */
  private _frozenAfterStructureChange(change: StructureChange): number {
    const frozen = change.axis === 'row' ? this._frozenRowCount : this._frozenColCount;
    if (change.kind === 'insert') return change.index < frozen ? frozen + change.count : frozen;
    const overlap = Math.max(0, Math.min(frozen, change.index + change.count) - change.index);
    return frozen - overlap;
  }

  /**
   * The grid's scroll padding: the sticky header row and frozen rows at the
   * top, the row header and frozen columns at the left. Scroll snapping then
   * aligns a row just below the frozen area (its snap position is exactly
   * where `_ensureCellVisible` scrolls a row to), and the browser's own
   * focus scrolling keeps cells out from under the sticky parts.
   */
  private _scrollPadding(): string {
    return (
      `scroll-padding-top: calc(var(--_cell-height) * ${1 + this._visibleIndex(this._frozenRowCount)}); ` +
      `scroll-padding-left: ${this._stickyLeft(this._frozenColCount)};`
    );
  }

  /** CSS for a frozen row's sticky top offset: below the header and the (visible) frozen rows above it. */
  private _stickyTop(row: number): string {
    return `calc(var(--_cell-height) * ${this._visibleIndex(row) + 1})`;
  }

  /** Classes marking a cell or header as frozen (pass -1 for a header's other axis). */
  private _frozenClasses(row: number, col: number): string {
    const fr = this._frozenRowCount;
    const fc = this._frozenColCount;
    const classes: string[] = [];
    // Headers pass -1 for the other axis, which must never match an edge
    if (row >= 0 && row < fr) classes.push('frozen-row');
    // The freeze line goes under the last frozen row a filter leaves visible
    if (row >= 0 && row < fr && row === this._visibleRowToward(fr - 1, -1)) classes.push('frozen-row-edge');
    if (col >= 0 && col < fc) classes.push('frozen-col');
    if (col >= 0 && col === fc - 1) classes.push('frozen-col-edge');
    return classes.join(' ');
  }

  /**
   * CSS for a frozen column's sticky left offset: the row header plus the
   * widths of the columns before it, with default-width columns kept as
   * `var(--_cell-width)` so the offset follows the CSS variable exactly.
   */
  private _stickyLeft(col: number): string {
    let px = 0;
    let defaults = 0;
    for (let c = 0; c < col; c++) {
      const custom = this._resizePreview?.cols.includes(c) ? this._resizePreview.width : this._colWidths.get(c);
      if (custom === undefined) defaults++;
      else px += custom;
    }
    return `calc(var(--_row-header-width) + ${px}px + ${defaults} * var(--_cell-width))`;
  }

  // ─── Column Widths ──────────────────────────────────

  private _colWidth(col: number): number {
    if (this._resizePreview?.cols.includes(col)) return this._resizePreview.width;
    return this._colWidths.get(col) ?? this._cellSize.width;
  }

  /** Left offset of every column within the data area, plus the total width at `[cols]`. */
  private _colOffsets(): Float64Array {
    const key = `${this.cols}|${this._cellSize.width}|${this._colWidthsRevision}`;
    if (this._colLayout?.key !== key) {
      const offsets = new Float64Array(this.cols + 1);
      for (let c = 0; c < this.cols; c++) offsets[c + 1] = offsets[c] + this._colWidth(c);
      this._colLayout = { key, offsets };
    }
    return this._colLayout.offsets;
  }

  private _colLeft(col: number): number {
    return this._colOffsets()[Math.max(0, Math.min(col, this.cols))];
  }

  /** The column at `x` px into the data area (clamped to the grid). */
  private _colAt(x: number): number {
    const offsets = this._colOffsets();
    let lo = 0;
    let hi = this.cols - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (offsets[mid] <= x) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }

  /** `grid-template-columns`, with runs of default-width columns kept as `repeat()`. */
  private _gridTemplateColumns(): string {
    const parts = ['var(--_row-header-width)'];
    let run = 0;
    const flush = () => {
      if (run > 0) parts.push(run === 1 ? 'var(--_cell-width)' : `repeat(${run}, var(--_cell-width))`);
      run = 0;
    };
    for (let c = 0; c < this.cols; c++) {
      const custom = this._resizePreview?.cols.includes(c) ? this._resizePreview.width : this._colWidths.get(c);
      if (custom === undefined) {
        run++;
      } else {
        flush();
        parts.push(`${custom}px`);
      }
    }
    flush();
    return parts.join(' ');
  }

  private _clampWidth(width: number): number {
    return Math.round(Math.min(Y11nSpreadsheet.MAX_COL_WIDTH, Math.max(Y11nSpreadsheet.MIN_COL_WIDTH, width)));
  }

  /**
   * Resize columns as one undoable step. User resizes (allowed in read-only
   * mode: widths are presentation, not data) can be vetoed via
   * `before-column-resize`. Returns whether anything changed.
   */
  private _resizeColumns(
    changes: Array<{ col: number; width: number | null }>,
    source: 'user' | 'programmatic'
  ): 'applied' | 'unchanged' | 'cancelled' {
    const deltas: WidthDelta[] = [];
    const seen = new Set<number>();
    for (const { col, width } of changes) {
      if (seen.has(col)) continue;
      seen.add(col);
      const before = this._colWidths.get(col) ?? null;
      const after = width === null ? null : this._clampWidth(width);
      if (after !== before) deltas.push({ col, before, after });
    }
    if (deltas.length === 0) return 'unchanged';

    if (source === 'user') {
      const detail: ColumnResizeDetail = { columns: this._resizedColumns(deltas, 'after'), source };
      const event = new CustomEvent<ColumnResizeDetail>('before-column-resize', {
        detail,
        bubbles: true,
        composed: true,
        cancelable: true,
      });
      if (!this.dispatchEvent(event)) return 'cancelled';
    }

    const selection = this._snapshotSelection();
    this._commitBatch(
      // `operation` is never reported for width batches: they fire column-resize, not data-change
      { deltas: [], selectionBefore: selection, selectionAfter: selection, operation: 'format', widths: deltas },
      source
    );
    return 'applied';
  }

  /** Effective widths of the columns in `deltas` on one side of the change. */
  private _resizedColumns(deltas: WidthDelta[], side: 'before' | 'after'): ColumnResizeDetail['columns'] {
    return deltas.map((d) => ({ col: d.col, width: d[side] ?? this._cellSize.width }));
  }

  private _applyWidthBatch(batch: CommandBatch, side: 'before' | 'after', source: ChangeSource): void {
    const deltas = batch.widths!;
    const { row, col } = this._selection.activeCell;
    const activeInView = this._isCellInView(row, col);
    for (const delta of deltas) {
      const width = delta[side];
      if (width === null) this._colWidths.delete(delta.col);
      else this._colWidths.set(delta.col, width);
    }
    this._colWidthsRevision++;
    // The selection is unchanged, so don't restore it (which would move focus,
    // e.g. when a page restores saved widths). If the active cell was in view,
    // keep it there: widening its column can push it off-screen. Scroll after
    // the render, once the grid has its new width (requested first, so
    // updateComplete waits for that render).
    this.requestUpdate();
    this.updateComplete.then(() => {
      if (activeInView && source !== 'programmatic') this._ensureCellVisible(row, col);
      // Narrower columns can shrink the scroll range under the current position
      const gridEl = this._grid;
      if (gridEl && gridEl.scrollLeft !== this._scrollLeft) this._scrollLeft = gridEl.scrollLeft;
    });
    this.dispatchEvent(
      new CustomEvent<ColumnResizeDetail>('column-resize', {
        detail: { columns: this._resizedColumns(deltas, side), source },
        bubbles: true,
        composed: true,
      })
    );
  }

  /** Whether any part of a cell is inside the scrolled data viewport. */
  private _isCellInView(row: number, col: number): boolean {
    const gridEl = this._grid;
    if (!gridEl) return false;
    const left = this._colLeft(col);
    const right = left + this._colWidth(col);
    const top = this._rowTop(row);
    const viewWidth = gridEl.clientWidth - Y11nSpreadsheet.ROW_HEADER_WIDTH;
    const viewHeight = gridEl.clientHeight - this._cellSize.height;
    // Frozen cells are always in view; the others scroll beside the frozen area
    const inX =
      col < this._frozenColCount ||
      (right > this._scrollLeft + this._frozenWidth && left < this._scrollLeft + viewWidth);
    const inY =
      row < this._frozenRowCount ||
      (top + this._cellSize.height > this._scrollTop + this._frozenHeight && top < this._scrollTop + viewHeight);
    return inX && inY;
  }

  /** The columns a resize of `col` applies to: all selected whole columns, if `col` is one of them. */
  private _columnsToResize(col: number): number[] {
    const { start, end } = this._selection.range;
    const wholeColumns = start.row === 0 && end.row === this.rows - 1;
    if (wholeColumns && col >= start.col && col <= end.col) {
      return Array.from({ length: end.col - start.col + 1 }, (_, i) => start.col + i);
    }
    return [col];
  }

  /** Alt+Shift+Left/Right: resize the selected columns from the active column's width. */
  private _resizeSelectedColumns(delta: number): void {
    const { start, end } = this._selection.range;
    const active = this._selection.activeCell.col;
    const width = this._clampWidth(this._colWidth(active) + delta);
    const cols = Array.from({ length: end.col - start.col + 1 }, (_, i) => start.col + i);
    this._announceWidths(cols, this._resizeColumns(cols.map((col) => ({ col, width })), 'user'));
  }

  /** Alt+Shift+0 and header-border double-click: fit columns to their content. */
  private _autoFitUserColumns(cols: number[]): void {
    this._announceWidths(cols, this._resizeColumns(cols.map((col) => ({ col, width: this._fitWidth(col) })), 'user'));
  }

  /** Announce the widths after a user resize (also when they didn't change, e.g. at the minimum). */
  private _announceWidths(cols: number[], result: ReturnType<Y11nSpreadsheet['_resizeColumns']>): void {
    if (result === 'cancelled') {
      this._modeAnnouncement = 'Resize cancelled';
      return;
    }
    const first = colToLetter(cols[0]);
    const last = colToLetter(cols[cols.length - 1]);
    const widths = new Set(cols.map((c) => Math.round(this._colWidth(c))));
    const what = cols.length === 1 ? `Column ${first}` : `Columns ${first} to ${last}`;
    this._modeAnnouncement =
      widths.size === 1 ? `${what} width ${[...widths][0]} pixels` : `${what} resized to fit their content`;
  }

  /**
   * The width that fits a column's widest displayed value, or null (the
   * default width) for an empty column. Text is measured on a canvas, so
   * rows outside the rendered window count too.
   */
  private _fitWidth(col: number): number | null {
    const host = getComputedStyle(this);
    const family = host.fontFamily || 'sans-serif';
    const baseSize = parseFloat(host.fontSize) || 13;
    if (this._measureCanvas === undefined) {
      try {
        this._measureCanvas = document.createElement('canvas').getContext('2d');
      } catch {
        this._measureCanvas = null;
      }
    }
    const ctx = this._measureCanvas;
    let widest = 0;
    for (const [key, cell] of this._internalData) {
      if (!cell.displayValue) continue;
      const pos = parseKey(key);
      if (pos.col !== col || pos.row >= this.rows) continue;
      const fmt = cell.format;
      const size = fmt?.fontSize ?? baseSize;
      let width: number;
      if (ctx) {
        ctx.font = `${fmt?.italic ? 'italic ' : ''}${fmt?.bold ? 'bold ' : ''}${size}px ${family}`;
        width = ctx.measureText(cell.displayValue).width;
      } else {
        width = cell.displayValue.length * size * 0.6; // no canvas (e.g. a test DOM): estimate
      }
      widest = Math.max(widest, width);
    }
    if (widest === 0) return null;
    // Cell padding (2 x 4px), right border, and a little room so text isn't cut off
    return this._clampWidth(Math.ceil(widest) + 15);
  }

  /** Pointer drag on a column header's right border: resize with a live preview. */
  private _handleResizerPointerDown(e: PointerEvent): void {
    if (e.button !== 0) return;
    if (this._endEditElsewhere()) {
      e.preventDefault();
      e.stopPropagation();
      return;
    }
    const handle = e.currentTarget as HTMLElement;
    const col = parseInt(handle.dataset.resizeCol!, 10);
    e.preventDefault();
    e.stopPropagation();
    if (this._isEditing) this._commitEdit();

    const cols = this._columnsToResize(col);
    const startX = e.clientX;
    const startWidth = this._colWidth(col);
    const controller = new AbortController();
    const { signal } = controller;
    const setPreview = (preview: typeof this._resizePreview) => {
      this._resizePreview = preview;
      this._colWidthsRevision++;
      this.requestUpdate();
    };

    const onMove = (ev: PointerEvent) => {
      setPreview({ cols, width: this._clampWidth(startWidth + ev.clientX - startX) });
    };
    const finish = (commit: boolean) => {
      controller.abort();
      const width = this._resizePreview?.width;
      setPreview(null);
      if (commit && width !== undefined) {
        this._announceWidths(cols, this._resizeColumns(cols.map((c) => ({ col: c, width })), 'user'));
      }
    };
    document.addEventListener('pointermove', onMove, { signal });
    document.addEventListener('pointerup', () => finish(true), { signal });
    document.addEventListener('pointercancel', () => finish(false), { signal });
    // Escape cancels the drag, restoring the width
    document.addEventListener(
      'keydown',
      (ev) => {
        if (ev.key !== 'Escape') return;
        ev.preventDefault();
        ev.stopPropagation();
        finish(false);
      },
      { signal, capture: true }
    );
    window.addEventListener('blur', () => finish(false), { signal });
  }

  private _handleResizerDblClick(e: MouseEvent): void {
    e.preventDefault();
    e.stopPropagation();
    if (this._endEditElsewhere()) return;
    const col = parseInt((e.currentTarget as HTMLElement).dataset.resizeCol!, 10);
    this._autoFitUserColumns(this._columnsToResize(col));
  }

  private _undo(): boolean {
    if (this.readOnly || this._pasteInProgress) return false;

    const batch = this._undoStack.pop();
    if (!batch) return false;
    this.requestUpdate();
    // A change made on another sheet is undone there, in view
    if (!this._showSheetOf(batch, 'undo')) return false;

    // Recorded before reverting, for the same reason as in _commitBatch: a
    // change a listener makes while it reverts clears the redo stack
    this._redoStack.push(batch);
    this._revertBatch(batch);
    this._dispatchFormatChangeForBatch(batch, 'undo');
    return true;
  }

  private _redo(): boolean {
    if (this.readOnly || this._pasteInProgress) return false;

    const batch = this._redoStack.pop();
    if (!batch) return false;
    this.requestUpdate();
    if (!this._showSheetOf(batch, 'redo')) return false;

    // Recorded before applying (see _commitBatch)
    this._undoStack.push(batch);
    if (this._undoStack.length > this._maxHistory) {
      this._undoStack.shift();
    }
    this._applyBatch(batch, 'redo');
    this._dispatchFormatChangeForBatch(batch, 'redo');
    return true;
  }

  /** Dispatch format-change for any batch that contains format deltas.
   *  NOTE: The event carries the format from the first delta only. For heterogeneous
   *  batches (e.g., paste with varied per-cell formatting), this is an approximation —
   *  the cellIds are accurate but the format may not represent every cell. Consumers
   *  that need per-cell accuracy should read cell data directly. */
  private _dispatchFormatChangeForBatch(batch: CommandBatch, source: 'undo' | 'redo'): void {
    const formatDeltas = batch.deltas.filter(
      (d) => d.formatBefore !== undefined || d.formatAfter !== undefined
    );
    if (formatDeltas.length === 0) return;

    const cellIds = formatDeltas.map((d) => d.id);
    // For undo, the applied format is formatBefore; for redo, formatAfter
    const format: CellFormat = source === 'undo'
      ? (formatDeltas[0].formatBefore ?? {})
      : (formatDeltas[0].formatAfter ?? {});
    this._dispatchFormatChange({ cellIds, format, source });
  }

  // ─── Editing ────────────────────────────────────────

  private _startEditing(initialValue?: string): void {
    if (this.readOnly || this._isEditing || this._pasteInProgress) return;

    const { row, col } = this._selection.activeCell;
    const key = cellKey(row, col);
    if (this._isCellLocked(key)) {
      this._modeAnnouncement = `${coordToRef({ row, col })} is read-only`;
      return;
    }

    const cell = this._internalData.get(key);
    const detail: CellEditStartDetail = { cellId: key, value: cell?.rawValue ?? '' };
    const event = new CustomEvent<CellEditStartDetail>('cell-edit-start', {
      detail,
      bubbles: true,
      composed: true,
      cancelable: true,
    });
    if (!this.dispatchEvent(event)) return;

    this._editingCellKey = key;
    this._editSheetId = this._activeSheetId;
    // Editing a list edits its items; typing over it replaces it
    this._editListMode = initialValue === undefined && listItemsOf(cell?.rawValue ?? '') !== null;
    this._editValue = initialValue ?? this._editTextFor(cell);
    this._validationMessage = '';
    this._isEditing = true;
    this._editorNeedsCursorReset = true;
    this._modeAnnouncement = `Editing ${coordToRef(this._selection.activeCell)}`;
  }

  /**
   * Commit the inline edit. Returns false if `validateCell` rejected the
   * value: with `onInvalid: 'keep'` (keyboard commits) the editor stays open
   * and shows the message; with `'cancel'` (focus moving away) the edit is
   * discarded and the message announced.
   */
  private _commitEdit(onInvalid: 'keep' | 'cancel' = 'cancel'): boolean {
    if (!this._isEditing || !this._editingCellKey) return true;
    this._returnToEditSheet();

    const key = this._editingCellKey;
    const { value: newValue, format } = this._interpretInput(key, this._editValue, this._editListMode);
    if (!this._checkEditValue(key, newValue, onInvalid)) return false;

    this._isEditing = false;
    this._editingCellKey = null;
    this._validationMessage = '';
    this._modeAnnouncement = 'Edit committed';

    this._commitRawValue(key, newValue, { validate: false, format });

    // Return focus to the grid
    this.updateComplete.then(() => this._focusActiveCell());
    return true;
  }

  /**
   * Validate the value being edited in `key`. On rejection, fires
   * `validation-error` and either keeps the editor open with the message or
   * cancels the edit. Unchanged values are not validated.
   */
  private _checkEditValue(key: string, value: string, onInvalid: 'keep' | 'cancel'): boolean {
    if (value === (this._internalData.get(key)?.rawValue ?? '')) return true;
    const message = this._validate(key, value);
    if (!message) return true;
    this._dispatchValidationError({ errors: [{ cellId: key, value, message }], operation: 'edit' });
    if (onInvalid === 'keep') {
      this._validationMessage = message;
      this._modeAnnouncement = `Invalid value: ${message}`;
    } else {
      this._cancelEdit(`Invalid value: ${message}. Edit cancelled`);
    }
    return false;
  }

  /**
   * How typed input is stored in `key`: `15%`, `$1,200`, `2026-01-05` and
   * the like become a number, and the cell gets the number format the input
   * implies, unless it already has a number format of that type (so a custom
   * date pattern survives typing a new date). A plain number typed into a
   * percent cell is a percentage, as with Excel's automatic percent entry:
   * `13` and `0.13` both store 0.13 (numbers of 1 or more are divided by 100).
   * With `raw-input`, or for anything else, the value is stored as typed and
   * `format` is undefined.
   *
   * Text becomes a list (several values in the cell) when `asList` is set
   * (a list's items were edited, or Ctrl+J), and when comma-separated text is
   * typed right below a list, as Excel continues a column of lists.
   */
  private _interpretInput(key: string, value: string, asList = false): { value: string; format?: CellFormat } {
    const listText = value.trim();
    if (listText !== '' && !listText.startsWith('=')) {
      let list = asList;
      if (!list && !this.rawInput && looksLikeList(listText)) {
        const { row, col } = parseKey(key);
        list = row > 0 && listItemsOf(this._internalData.get(cellKey(row - 1, col))?.rawValue ?? '') !== null;
      }
      if (list) {
        const items = parseListText(listText);
        if (items.length > 0) return { value: listToFormula(items) };
      }
    }
    if (this.rawInput) return { value };
    const existing = this._internalData.get(key)?.format;
    const text = value.trim();
    if (existing?.numberFormat?.type === 'percent' && /^[+-]?(\d+\.?\d*|\.\d+)$/.test(text)) {
      return { value: Math.abs(Number(text)) >= 1 ? String(parseFloat(`${text}e-2`)) : String(Number(text)) };
    }
    const parsed = parseInput(value);
    if (!parsed) return { value };
    if (existing?.numberFormat?.type === parsed.numberFormat.type) return { value: parsed.value };
    return { value: parsed.value, format: { ...existing, numberFormat: parsed.numberFormat } };
  }

  /**
   * The text the inline editor starts with for a cell. Percent and date
   * cells holding a plain number are edited as displayed (`15%`,
   * `2026-01-05`), as in Excel, when that text reads back as the same number;
   * everything else is edited as its raw value.
   */
  private _editTextFor(cell: CellData | undefined): string {
    const raw = cell?.rawValue ?? '';
    // A list is edited as its items: Ann, Bob
    const items = listItemsOf(raw);
    if (items) return listToText(items);
    const type = cell?.format?.numberFormat?.type;
    if (this.rawInput || cell?.type !== 'number' || raw.startsWith('=') || (type !== 'percent' && type !== 'date')) {
      return raw;
    }
    const parsed = parseInput(cell.displayValue);
    return parsed && Math.abs(Number(parsed.value) - Number(raw)) < 1e-9 ? cell.displayValue : raw;
  }

  private _cancelEdit(announcement = 'Edit cancelled'): void {
    this._returnToEditSheet();
    this._isEditing = false;
    this._editingCellKey = null;
    this._validationMessage = '';
    this._modeAnnouncement = announcement;
    this.updateComplete.then(() => this._focusActiveCell());
  }

  // ─── Reference Selection ────────────────────────────

  /**
   * Check if the cursor is at a position in the formula where a cell reference
   * can be inserted (after =, +, -, *, /, &, (, ,, or comparison operators).
   */
  private _isRefPositionInFormula(): boolean {
    if (!this._editValue.startsWith('=')) return false;

    const cursorPos = this._editor?.selectionStart ?? this._editValue.length;
    const before = this._editValue.substring(0, cursorPos).trimEnd();
    if (before === '=') return true;

    const lastChar = before[before.length - 1];
    return ['+', '-', '*', '/', '^', '&', '(', ',', ':', ';', '{', '<', '>', '=', '@'].includes(lastChar);
  }

  /**
   * Handle an arrow key press during formula editing to insert or move
   * a cell reference.
   */
  private _handleRefArrow(key: string): void {
    const dRow = key === 'ArrowDown' ? 1 : key === 'ArrowUp' ? -1 : 0;
    const dCol = key === 'ArrowRight' ? 1 : key === 'ArrowLeft' ? -1 : 0;

    if (!this._refMode) {
      // Enter reference mode - start from the editing cell and move
      const { row, col } = this._selection.activeCell;
      const clamped = this._selection.clamp({ row: this._selection.stepRows(row, dRow), col: col + dCol });
      this._refCursorRow = clamped.row;
      this._refCursorCol = clamped.col;
      this._refMode = true;
      this._modeAnnouncement = `Reference selection mode. Selecting ${coordToRef({ row: this._refCursorRow, col: this._refCursorCol })}`;

      const cursorPos = this._editor?.selectionStart ?? this._editValue.length;
      this._refInsertStart = cursorPos;
      const ref = this._refText({ row: this._refCursorRow, col: this._refCursorCol });
      this._editValue =
        this._editValue.substring(0, cursorPos) +
        ref +
        this._editValue.substring(cursorPos);
      this._refInsertEnd = cursorPos + ref.length;
    } else {
      // Already in reference mode - move the reference cursor
      const moved = this._selection.clamp({
        row: this._selection.stepRows(this._refCursorRow, dRow),
        col: this._refCursorCol + dCol,
      });
      this._refCursorRow = moved.row;
      this._refCursorCol = moved.col;

      const ref = this._refText({ row: this._refCursorRow, col: this._refCursorCol });
      this._modeAnnouncement = `Selecting ${ref}`;
      this._editValue =
        this._editValue.substring(0, this._refInsertStart) +
        ref +
        this._editValue.substring(this._refInsertEnd);
      this._refInsertEnd = this._refInsertStart + ref.length;
    }

    this._ensureCellVisible(this._refCursorRow, this._refCursorCol);
  }

  /** Whether the cell being edited is on another sheet than the one shown (a reference is being picked there). */
  private get _editingElsewhere(): boolean {
    return this._isEditing && this._editSheetId !== this._activeSheetId;
  }

  /** Show the sheet of the cell being edited again, before the edit ends. */
  private _returnToEditSheet(): void {
    if (!this._editingElsewhere) return;
    const state = this._sheets.find((s) => s.id === this._editSheetId);
    if (state) this._switchSheetView(state, 'user');
  }

  /**
   * The text of a reference picked while editing a formula: `B2`, `B2:C4`,
   * and with the sheet's name when it is picked on another sheet.
   */
  private _refText(a: CellCoord, b: CellCoord = a): string {
    const start = { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) };
    const end = { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) };
    const cells =
      start.row === end.row && start.col === end.col ? coordToRef(start) : `${coordToRef(start)}:${coordToRef(end)}`;
    return (this._editingElsewhere ? sheetPrefix(this._activeState.name) : '') + cells;
  }

  /** Put a picked reference into the formula at the caret, replacing the one being picked. */
  private _insertReference(ref: string): void {
    if (!this._refMode) {
      // Selected text is replaced, as typing would replace it
      const cursorPos = this._editor?.selectionStart ?? this._editValue.length;
      this._refInsertStart = cursorPos;
      this._refInsertEnd = Math.max(cursorPos, this._editor?.selectionEnd ?? cursorPos);
      this._refMode = true;
    }
    this._editValue =
      this._editValue.substring(0, this._refInsertStart) + ref + this._editValue.substring(this._refInsertEnd);
    this._refInsertEnd = this._refInsertStart + ref.length;
  }

  /**
   * A click (or drag) on cells while a formula is being typed where a
   * reference fits: the cell or range goes into the formula.
   */
  private _pickReferenceWithPointer(row: number, col: number, e: PointerEvent): void {
    e.preventDefault();
    this._refPickPointer = true;
    const anchor = { row, col };
    const pick = (head: CellCoord) => {
      this._refCursorRow = head.row;
      this._refCursorCol = head.col;
      const ref = this._refText(anchor, head);
      this._insertReference(ref);
      this._modeAnnouncement = `Selecting ${ref}`;
    };
    pick(anchor);

    const controller = new AbortController();
    const { signal } = controller;
    const onPointerMove = (ev: PointerEvent) => {
      const el = this.shadowRoot?.elementFromPoint(ev.clientX, ev.clientY) as HTMLElement | null;
      const nearest = el?.closest?.('[data-row]') as HTMLElement | null;
      if (!nearest) return;
      const head = { row: parseInt(nearest.dataset.row!, 10), col: parseInt(nearest.dataset.col!, 10) };
      if (head.row !== this._refCursorRow || head.col !== this._refCursorCol) pick(head);
    };
    const onPointerUp = () => {
      controller.abort();
      this._refPickPointer = false;
      // Typing goes on after the reference
      this.updateComplete.then(() => {
        if (!this._isEditing) return;
        this._editor?.focus();
        this._editor?.setSelectionRange(this._refInsertEnd, this._refInsertEnd);
      });
    };
    document.addEventListener('pointermove', onPointerMove, { signal });
    document.addEventListener('pointerup', onPointerUp, { signal });
    document.addEventListener('pointercancel', onPointerUp, { signal });
    window.addEventListener('blur', onPointerUp, { signal });
  }

  /**
   * For grid controls other than cells (headers, resize handles, filter
   * buttons): with a formula open whose cell is on another sheet, the click
   * ends the edit, which shows that sheet again, and does nothing else. What
   * was clicked belonged to the sheet that is no longer shown.
   */
  private _endEditElsewhere(): boolean {
    if (!this._editingElsewhere) return false;
    this._exitRefMode();
    this._commitEdit();
    return true;
  }

  private _exitRefMode(): void {
    if (this._refMode) {
      this._modeAnnouncement = 'Reference selection ended';
    }
    this._refMode = false;
  }

  /**
   * Cycle the cell reference at/before the cursor through reference modes:
   * A1 → $A$1 → A$1 → $A1 → A1
   */
  private _cycleReferenceMode(): void {
    const cursorPos = this._editor?.selectionStart ?? this._editValue.length;
    const text = this._editValue;

    // Find the reference at or just before the cursor
    const refPattern = /(?<![A-Z0-9_.$])(\$?)([A-Z]+)(\$?)(\d+)(?![A-Z0-9_.(])/gi;
    let match: RegExpExecArray | null;
    let bestMatch: RegExpExecArray | null = null;

    while ((match = refPattern.exec(text)) !== null) {
      const start = match.index;
      const end = start + match[0].length;
      // The ref contains or is immediately before the cursor
      if (start <= cursorPos && end >= cursorPos - 1) {
        bestMatch = match;
      }
    }

    if (!bestMatch) return;

    const [full, colDollar, colLetters, rowDollar, rowDigits] = bestMatch;
    const start = bestMatch.index;
    const hasColDollar = colDollar === '$';
    const hasRowDollar = rowDollar === '$';

    let newRef: string;
    if (!hasColDollar && !hasRowDollar) {
      // A1 → $A$1
      newRef = `$${colLetters}$${rowDigits}`;
    } else if (hasColDollar && hasRowDollar) {
      // $A$1 → A$1
      newRef = `${colLetters}$${rowDigits}`;
    } else if (!hasColDollar && hasRowDollar) {
      // A$1 → $A1
      newRef = `$${colLetters}${rowDigits}`;
    } else {
      // $A1 → A1
      newRef = `${colLetters}${rowDigits}`;
    }

    this._editValue = text.substring(0, start) + newRef + text.substring(start + full.length);
    // A reference that is being picked stays the one the next pick replaces
    if (this._refMode && start >= this._refInsertStart && start + full.length <= this._refInsertEnd) {
      this._refInsertEnd += newRef.length - full.length;
    }

    // Adjust cursor position to stay at the end of the reference
    const newCursorPos = start + newRef.length;
    this.updateComplete.then(() => {
      if (this._editor) {
        this._editor.setSelectionRange(newCursorPos, newCursorPos);
      }
    });
  }

  // ─── Focus Management ───────────────────────────────

  private _focusActiveCell(): void {
    const { row, col } = this._selection.activeCell;
    this._ensureCellVisible(row, col);
    const cellEl = this.shadowRoot?.querySelector(
      `[data-row="${row}"][data-col="${col}"]`
    ) as HTMLElement | null;

    if (cellEl) {
      cellEl.focus();
    }
  }

  // ─── Keyboard Handling ──────────────────────────────

  private _handleGridKeydown(e: KeyboardEvent): void {
    if (this._isEditing) {
      this._handleEditorKeydown(e);
      return;
    }

    const shift = e.shiftKey;
    const ctrl = e.ctrlKey || e.metaKey;

    // Ctrl+Space selects columns, Shift+Space selects rows, both select all.
    // Checked before the printable-character branch, which would otherwise
    // treat Shift+Space as typing a space.
    if (e.key === ' ' && (ctrl || shift) && !e.altKey) {
      e.preventDefault();
      e.stopPropagation();
      this._selectWholeRowsOrColumns(ctrl, shift);
      return;
    }

    // Only stop propagation for keys we actually handle, so unhandled keys
    // (e.g. Tab) propagate normally and focus can leave the grid.
    // Start editing on printable character input. This must run before the
    // shortcut switch so plain letters like "a" or "c" aren't swallowed by the
    // Ctrl+letter cases. Count code points so emoji also start editing.
    if (!ctrl && !e.altKey && [...e.key].length === 1) {
      if (!this.readOnly) {
        e.preventDefault();
        e.stopPropagation();
        this._startEditing(e.key);
      }
      return;
    }

    // Insert/delete rows or columns: Ctrl+Shift+= / Ctrl+Alt+= / Ctrl+NumpadAdd
    // and Ctrl+- / Ctrl+Alt+- / Ctrl+NumpadSubtract.
    if (ctrl && this._handleStructureShortcut(e)) return;

    // Ctrl+Alt+D / Ctrl+Alt+R: fill series down / right (the fill handle's
    // keyboard equivalent). Matched by e.code because Option changes e.key on macOS.
    if (ctrl && e.altKey && !shift && (e.code === 'KeyD' || e.code === 'KeyR')) {
      e.preventDefault();
      e.stopPropagation();
      this._fillSeriesInSelection(e.code === 'KeyD' ? 'down' : 'right');
      return;
    }

    // Ctrl+Shift+L: add or remove a filter. Ctrl+Alt+L: reapply it. Views,
    // so allowed in read-only mode. Matched by e.code (Option changes e.key on macOS).
    if (ctrl && e.code === 'KeyL' && (shift !== e.altKey)) {
      e.preventDefault();
      e.stopPropagation();
      if (shift) this.toggleFilter();
      else this._reapplyUserFilter();
      return;
    }

    // Ctrl+PageDown / Ctrl+PageUp (or with Alt, where the browser keeps
    // Ctrl+Page keys for its own tabs): show the next / previous sheet
    if ((ctrl || e.altKey) && !shift && (e.key === 'PageDown' || e.key === 'PageUp') && this._sheets.length > 1) {
      e.preventDefault();
      e.stopPropagation();
      this._stepSheet(e.key === 'PageDown' ? 1 : -1);
      return;
    }

    // Ctrl+J: turn the selected cells into lists (several values in one
    // cell), or lists back into text. Matched by e.code (Option changes e.key on macOS).
    if (ctrl && !e.altKey && !shift && e.code === 'KeyJ') {
      e.preventDefault();
      e.stopPropagation();
      this._toggleListsInSelection();
      return;
    }

    // Alt+Down on a filter header cell: open its filter menu. On a cell that
    // holds several values: show them.
    if (e.altKey && !shift && !ctrl && e.key === 'ArrowDown') {
      const { row, col } = this._selection.activeCell;
      if (this._isFilterHeader(row, col)) {
        e.preventDefault();
        e.stopPropagation();
        void this.openFilterMenu(col);
        return;
      }
      if (this._getCell(row, col)?.array) {
        e.preventDefault();
        e.stopPropagation();
        void this.showCellValues();
        return;
      }
    }

    // Alt+Shift+F: freeze the rows above and columns left of the active cell,
    // or unfreeze. A view setting, so allowed in read-only mode.
    if (e.altKey && shift && !ctrl && e.code === 'KeyF') {
      e.preventDefault();
      e.stopPropagation();
      this._toggleFreeze();
      return;
    }

    // Alt+Shift+Left / Alt+Shift+Right: narrow / widen the selected columns.
    // Alt+Shift+0: fit them to their content. Allowed in read-only mode.
    if (e.altKey && shift && !ctrl && (e.key === 'ArrowLeft' || e.key === 'ArrowRight' || e.code === 'Digit0')) {
      e.preventDefault();
      e.stopPropagation();
      if (e.code === 'Digit0') {
        const { start, end } = this._selection.range;
        this._autoFitUserColumns(Array.from({ length: end.col - start.col + 1 }, (_, i) => start.col + i));
      } else {
        const step = Y11nSpreadsheet.COL_RESIZE_STEP;
        this._resizeSelectedColumns(e.key === 'ArrowLeft' ? -step : step);
      }
      return;
    }

    // Alt+Shift+Up / Alt+Shift+Down: sort ascending / descending by the
    // active cell's column
    if (e.altKey && shift && !ctrl && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
      e.preventDefault();
      e.stopPropagation();
      this.sortSelection(e.key === 'ArrowUp' ? 'ascending' : 'descending');
      return;
    }

    if (this._handleQuickShortcut(e, ctrl, shift)) {
      e.preventDefault();
      e.stopPropagation();
      return;
    }

    let handled = true;

    // Normalize letter keys so shortcuts work with Caps Lock / Shift held.
    const key = ctrl && e.key.length === 1 ? e.key.toLowerCase() : e.key;

    switch (key) {
      case 'ArrowUp':
      case 'ArrowDown':
      case 'ArrowLeft':
      case 'ArrowRight': {
        e.preventDefault();
        const [dRow, dCol] = ARROW_DELTAS[key];
        if (ctrl) {
          // Jump to the edge of the data region (Ctrl/Cmd+Arrow); with Shift,
          // extend the selection from its head to that edge.
          const from = shift ? this._selection.head : this._selection.activeCell;
          const target = findDataEdge(
            from,
            dRow,
            dCol,
            this.rows,
            this.cols,
            (r, c) => this._isCellFilled(r, c),
            (r) => this._isRowHidden(r)
          );
          this._selection.moveTo(target.row, target.col, shift);
        } else {
          this._selection.move(dRow, dCol, shift);
        }
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;
      }

      case 'Enter':
        e.preventDefault();
        if (!this.readOnly) {
          this._startEditing();
        }
        break;

      case 'F2':
        e.preventDefault();
        if (!this.readOnly) {
          this._startEditing();
        }
        break;

      case 'Escape':
        e.preventDefault();
        this._selection.clearSelection();
        this._dispatchSelectionChange();
        break;

      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        if (!this.readOnly) {
          this._clearSelectedCells();
        }
        break;

      case 'Home':
        e.preventDefault();
        if (ctrl) {
          // With frozen panes, the top-left of the scrolling area (as in Excel)
          this._selection.moveTo(this._visibleRowToward(this._frozenRowCount, 1), this._frozenColCount, shift);
        } else {
          this._selection.moveTo(this._selection.activeCell.row, 0, shift);
        }
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'End':
        e.preventDefault();
        if (ctrl) {
          this._selection.moveTo(this._visibleRowToward(this.rows - 1, -1), this.cols - 1, shift);
        } else {
          this._selection.moveTo(this._selection.activeCell.row, this.cols - 1, shift);
        }
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'PageDown': {
        e.preventDefault();
        const cellHeight = this._cellSize.height;
        const gridEl = this._grid;
        const viewHeight = gridEl?.clientHeight || cellHeight * Y11nSpreadsheet.INITIAL_VIEW_ROWS;
        const pageRows = Math.max(1, Math.floor((viewHeight - this._frozenHeight) / cellHeight) - 1);
        this._selection.move(pageRows, 0, shift);
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;
      }

      case 'PageUp': {
        e.preventDefault();
        const cellHeight = this._cellSize.height;
        const gridEl = this._grid;
        const viewHeight = gridEl?.clientHeight || cellHeight * Y11nSpreadsheet.INITIAL_VIEW_ROWS;
        const pageRows = Math.max(1, Math.floor((viewHeight - this._frozenHeight) / cellHeight) - 1);
        this._selection.move(-pageRows, 0, shift);
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;
      }

      case 'a':
        if (ctrl) {
          e.preventDefault();
          this._selection.selectAll();
          this._dispatchSelectionChange();
        } else {
          handled = false;
        }
        break;

      case 'c':
        if (ctrl) {
          e.preventDefault();
          this._handleCopy();
        } else {
          handled = false;
        }
        break;

      case 'x':
        if (ctrl) {
          e.preventDefault();
          this._handleCut();
        } else {
          handled = false;
        }
        break;

      case 'v':
        if (ctrl) {
          e.preventDefault();
          // Ctrl+Shift+V: paste values only
          void this._handlePaste({ mode: shift ? 'values' : 'all', transpose: false });
        } else {
          handled = false;
        }
        break;

      case 'b':
        if (ctrl) {
          e.preventDefault();
          this.toggleFormat('bold');
        } else {
          handled = false;
        }
        break;

      case 'i':
        if (ctrl) {
          e.preventDefault();
          this.toggleFormat('italic');
        } else {
          handled = false;
        }
        break;

      case 'u':
        if (ctrl) {
          e.preventDefault();
          this.toggleFormat('underline');
        } else {
          handled = false;
        }
        break;

      case 'f':
        if (ctrl && !shift) {
          e.preventDefault();
          void this.openFind();
        } else {
          handled = false;
        }
        break;

      case 'd':
        if (ctrl && !shift && !e.altKey) {
          e.preventDefault();
          this._fillSelection('down');
        } else {
          handled = false;
        }
        break;

      case 'r':
        if (ctrl && !shift && !e.altKey) {
          e.preventDefault();
          this._fillSelection('right');
        } else {
          handled = false;
        }
        break;

      case 'y':
        if (ctrl) {
          e.preventDefault();
          this._redo();
        } else {
          handled = false;
        }
        break;

      case 'z':
        if (ctrl) {
          e.preventDefault();
          if (shift) {
            this._redo();
          } else {
            this._undo();
          }
        } else {
          handled = false;
        }
        break;

      default:
        handled = false;
        break;
    }

    if (handled) {
      e.stopPropagation();
    }
  }

  /**
   * Excel shortcuts matched by `e.code`, since Shift and Option change
   * `e.key` (Shift+4 is `$`, Option+= is `≠` on macOS). Returns whether the
   * key was handled.
   *
   * - Alt+=: AutoSum
   * - Ctrl+; / Ctrl+Shift+;: enter today's date / the current time
   * - Ctrl+' / Ctrl+Shift+': edit with the formula / value of the cell above
   * - Ctrl+Shift+1/3/4/5/`: number, date, currency, percent, general format
   * - Alt+Shift+S: announce the selection's Sum, Average and Count
   * - Ctrl+Alt+V: open Paste Special
   * - Alt+Shift+C: open the conditional-format rule manager
   */
  private _handleQuickShortcut(e: KeyboardEvent, ctrl: boolean, shift: boolean): boolean {
    const alt = e.altKey;
    if (alt && !ctrl && !shift && e.code === 'Equal') {
      this._autoSum();
      return true;
    }
    if (alt && shift && !ctrl && e.code === 'KeyS') {
      this._announceSelectionSummary();
      return true;
    }
    if (ctrl && alt && !shift && e.code === 'KeyV') {
      void this.openPasteSpecial();
      return true;
    }
    if (alt && shift && !ctrl && e.code === 'KeyC') {
      void this.openConditionalFormatManager();
      return true;
    }
    // Alt+Shift+V: the values of a cell that holds several (also where Alt+Down opens a filter menu)
    if (alt && shift && !ctrl && e.code === 'KeyV') {
      void this.showCellValues();
      return true;
    }
    // Ctrl+F3: the name manager
    if (ctrl && !alt && !shift && e.key === 'F3') {
      void this.openNameManager();
      return true;
    }
    // Ctrl+G: go to a cell, range or name through the Name Box
    if (ctrl && !alt && !shift && e.code === 'KeyG') {
      this.focusNameBox();
      return true;
    }
    if (!ctrl || alt) return false;
    if (e.code === 'Semicolon') {
      if (!this.readOnly) this._startEditing(shift ? currentTimeText() : todayText());
      return true;
    }
    if (e.code === 'Quote') {
      this._editFromCellAbove(shift ? 'value' : 'formula');
      return true;
    }
    if (!shift) return false;
    const formats: Record<string, [NumberFormatOptions | undefined, string]> = {
      Digit1: [{ type: 'number', decimals: 2, thousandsSep: true }, 'Number'],
      Digit3: [{ type: 'date', datePattern: 'YYYY-MM-DD' }, 'Date'],
      Digit4: [{ type: 'currency', decimals: 2 }, 'Currency'],
      Digit5: [{ type: 'percent', decimals: 0 }, 'Percent'],
      Backquote: [undefined, 'General'],
    };
    const entry = formats[e.code];
    if (!entry) return false;
    this._applyNumberFormatToSelection(entry[0], entry[1]);
    return true;
  }

  /** Ctrl+Shift+1/3/4/5/`: set (or, with undefined, remove) the selection's number format. */
  private _applyNumberFormatToSelection(numberFormat: NumberFormatOptions | undefined, label: string): void {
    if (this.readOnly || this._pasteInProgress) return;
    const { cellIds, locked } = this._selectionCellsToFormat();
    if (cellIds.length === 0) return;
    this._applyFormat(cellIds, { numberFormat }, 'user');
    if (locked === 0) this._modeAnnouncement = `${label} format`;
  }

  /**
   * Ctrl+' / Ctrl+Shift+': start editing with the formula (raw value) or the
   * displayed value of the nearest visible cell above, as in Excel.
   */
  private _editFromCellAbove(what: 'formula' | 'value'): void {
    if (this.readOnly) return;
    const { row, col } = this._selection.activeCell;
    let above = row - 1;
    while (above >= 0 && this._isRowHidden(above)) above--;
    if (above < 0) return;
    const cell = this._getCell(above, col);
    this._startEditing(what === 'formula' ? this._editTextFor(cell) : cell?.displayValue ?? '');
  }

  /**
   * Alt+=: AutoSum. On a single cell, start editing a SUM of the numbers
   * directly above it (or, if there are none, to its left). On a range, write
   * a SUM under each column: in the selection's last row if that row is
   * empty, otherwise in the row below. A single-row range gets its sum to
   * the right instead.
   */
  private _autoSum(): void {
    if (this.readOnly || this._pasteInProgress) return;
    const { start, end } = this._selection.range;
    const isNumber = (r: number, c: number) => this._getCell(r, c)?.type === 'number';

    if (start.row === end.row && start.col === end.col) {
      const { row, col } = start;
      let top = row;
      for (let r = row - 1; r >= 0; r--) {
        if (this._isRowHidden(r)) continue;
        if (!isNumber(r, col)) break;
        top = r;
      }
      let ref = '';
      if (top < row) {
        ref = `${coordToRef({ row: top, col })}:${coordToRef({ row: row - 1, col })}`;
      } else {
        let left = col;
        while (left > 0 && isNumber(row, left - 1)) left--;
        if (left < col) ref = `${coordToRef({ row, col: left })}:${coordToRef({ row, col: col - 1 })}`;
      }
      this._startEditing(`=SUM(${ref})`);
      return;
    }

    const updates: CellUpdate[] = [];
    if (start.row !== end.row) {
      const rows = this._visibleRowsBetween(start.row, end.row);
      const last = rows[rows.length - 1];
      let lastEmpty = rows.length > 1;
      for (let c = start.col; c <= end.col && lastEmpty; c++) lastEmpty = !this._isCellFilled(last, c);
      const target = lastEmpty ? last : this._visibleRowToward(end.row + 1, 1);
      const sumEnd = lastEmpty ? rows[rows.length - 2] : end.row;
      if (!lastEmpty && (end.row + 1 >= this.rows || target <= end.row)) {
        this._modeAnnouncement = 'No room below the selection for the sums';
        return;
      }
      for (let c = start.col; c <= end.col; c++) {
        updates.push({
          id: cellKey(target, c),
          value: `=SUM(${coordToRef({ row: start.row, col: c })}:${coordToRef({ row: sumEnd, col: c })})`,
        });
      }
    } else {
      const lastEmpty = !this._isCellFilled(start.row, end.col);
      const target = lastEmpty ? end.col : end.col + 1;
      if (target >= this.cols) {
        this._modeAnnouncement = 'No room right of the selection for the sum';
        return;
      }
      const sumEnd = lastEmpty ? end.col - 1 : end.col;
      updates.push({
        id: cellKey(start.row, target),
        value: `=SUM(${coordToRef({ row: start.row, col: start.col })}:${coordToRef({ row: start.row, col: sumEnd })})`,
      });
    }
    const batch = this._buildCellBatch(updates, 'edit', this._snapshotSelection());
    const vetted = batch && this._vetUserBatch(batch);
    if (!vetted) return;
    this._modeAnnouncement = `Inserted ${vetted.deltas.length === 1 ? 'a sum' : `${vetted.deltas.length} sums`}`;
    void this._runUserBatch(vetted);
  }

  /**
   * Sum, Average and Count of the selection, as the status bar shows them:
   * Count is the number of non-empty cells; Sum and Average cover the number
   * cells (formula results included) and use the first one's number format.
   * Rows hidden by the filter are left out. Null for a single cell, or when
   * fewer than two cells are filled. Cached until the data or selection changes.
   */
  private _selectionSummary(): SelectionSummary | null {
    const { start, end } = this._selection.range;
    const cacheKey = `${start.row}:${start.col}:${end.row}:${end.col}:${this._mutationCount}:${this._filterRevision}`;
    if (this._summaryCache?.key === cacheKey) return this._summaryCache.value;

    let count = 0;
    let numbers = 0;
    let sum = 0;
    let format: NumberFormatOptions | undefined;
    let firstNumberKey: string | null = null;
    const visit = (key: string, cell: CellData) => {
      if (!hasContent(cell)) return;
      count++;
      if (cell.type !== 'number') return;
      const n = this._sortNumber(key, cell);
      if (!Number.isFinite(n)) return;
      numbers++;
      sum += n;
      if (firstNumberKey === null || compareKeys(key, firstNumberKey) < 0) {
        firstNumberKey = key;
        format = cell.format?.numberFormat;
      }
    };
    const area = (end.row - start.row + 1) * (end.col - start.col + 1);
    if (area > this._internalData.size) {
      for (const [key, cell] of this._internalData) {
        const { row, col } = parseKey(key);
        if (row < start.row || row > end.row || col < start.col || col > end.col || this._isRowHidden(row)) continue;
        visit(key, cell);
      }
    } else {
      for (const r of this._visibleRowsBetween(start.row, end.row)) {
        for (let c = start.col; c <= end.col; c++) {
          const key = cellKey(r, c);
          const cell = this._internalData.get(key);
          if (cell) visit(key, cell);
        }
      }
    }

    let value: SelectionSummary | null = null;
    if (area > 1 && count > 1) {
      const show = (n: number) => (format ? formatNumber(n, format) : formatSeriesNumber(n));
      value = numbers > 0
        ? { count, sum: show(sum), average: show(sum / numbers) }
        : { count };
    }
    this._summaryCache = { key: cacheKey, value };
    return value;
  }

  /** Alt+Shift+S: announce the selection's Sum, Average and Count. */
  private _announceSelectionSummary(): void {
    const summary = this._selectionSummary();
    if (!summary) {
      this._modeAnnouncement = 'Select two or more filled cells for a sum, average and count';
      return;
    }
    this._modeAnnouncement = summary.sum !== undefined
      ? `Sum ${summary.sum}, Average ${summary.average}, Count ${summary.count}`
      : `Count ${summary.count}, no numbers`;
  }

  private _renderStatusBar() {
    const summary = this._selectionSummary();
    return html`
      <div class="ls-status" part="status-bar" title="Alt+Shift+S reads these out">
        ${summary
          ? html`${summary.sum !== undefined
              ? html`<span>Average: ${summary.average}</span><span>Count: ${summary.count}</span><span>Sum: ${summary.sum}</span>`
              : html`<span>Count: ${summary.count}</span>`}`
          : nothing}
      </div>
    `;
  }

  private _handleEditorKeydown(e: KeyboardEvent): void {
    e.stopPropagation();

    // Ctrl+; / Ctrl+Shift+;: insert today's date / the current time at the caret
    if ((e.ctrlKey || e.metaKey) && !e.altKey && e.code === 'Semicolon') {
      e.preventDefault();
      const input = e.target as HTMLInputElement;
      const text = e.shiftKey ? currentTimeText() : todayText();
      if (typeof input.setRangeText === 'function') {
        input.setRangeText(text, input.selectionStart ?? input.value.length, input.selectionEnd ?? input.value.length, 'end');
        this._editValue = input.value;
      } else {
        this._editValue += text;
      }
      return;
    }

    // Ctrl+PageDown / PageUp (or with Alt) in a formula: show another sheet to pick a reference there
    if (
      (e.ctrlKey || e.metaKey || e.altKey) &&
      !e.shiftKey &&
      (e.key === 'PageDown' || e.key === 'PageUp') &&
      this._sheets.length > 1
    ) {
      e.preventDefault();
      if (this._editValue.startsWith('=')) this._stepSheet(e.key === 'PageDown' ? 1 : -1);
      return;
    }

    // Ctrl+J: commit what was typed as a list (Ann, Bob, Cy → three values in the cell)
    if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && e.code === 'KeyJ') {
      e.preventDefault();
      this._exitRefMode();
      this._editListMode = true;
      this._commitEdit('keep');
      return;
    }

    switch (e.key) {
      case 'Enter':
        e.preventDefault();
        this._exitRefMode();
        if (e.ctrlKey || e.metaKey) {
          // Ctrl/Cmd+Enter: write the value into every selected cell, stay put
          this._commitEditToSelection();
          break;
        }
        if (!this._commitEdit('keep')) break;
        // Enter moves down after committing, Shift+Enter moves up
        this._selection.move(e.shiftKey ? -1 : 1, 0);
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'Tab':
        e.preventDefault();
        this._exitRefMode();
        if (!this._commitEdit('keep')) break;
        if (e.shiftKey) {
          this._selection.move(0, -1);
        } else {
          this._selection.move(0, 1);
        }
        this._dispatchSelectionChange();
        this.updateComplete.then(() => this._focusActiveCell());
        break;

      case 'Escape':
        e.preventDefault();
        this._exitRefMode();
        this._cancelEdit();
        break;

      case 'ArrowUp':
      case 'ArrowDown':
      case 'ArrowLeft':
      case 'ArrowRight':
        if (this._refMode || this._isRefPositionInFormula()) {
          e.preventDefault();
          this._handleRefArrow(e.key);
        }
        break;

      case 'f':
      case 'F':
        if (e.ctrlKey || e.metaKey) {
          e.preventDefault();
          this._exitRefMode();
          void this.openFind();
        }
        break;

      case 'F4':
        if (this._editValue.startsWith('=')) {
          e.preventDefault();
          this._cycleReferenceMode();
        }
        break;
    }
  }

  /**
   * Keyboard insertion/deletion of rows or columns. Acts on whole-row or
   * whole-column selections (Shift+Space / Ctrl+Space); with any other
   * selection it announces how to select rows or columns first.
   *
   * @returns true if the key was an insert/delete shortcut
   */
  private _handleStructureShortcut(e: KeyboardEvent): boolean {
    const insert =
      e.code === 'NumpadAdd' || e.key === '+' || ((e.shiftKey || e.altKey) && e.code === 'Equal');
    const remove =
      !e.shiftKey && (e.code === 'NumpadSubtract' || e.code === 'Minus' || e.key === '-');
    if (!insert && !remove) return false;
    e.preventDefault();
    e.stopPropagation();

    if (this.readOnly) {
      this._modeAnnouncement = 'The spreadsheet is read-only';
      return true;
    }
    const { start, end } = this._selection.range;
    const wholeRows = start.col === 0 && end.col === this.cols - 1;
    const wholeCols = start.row === 0 && end.row === this.rows - 1;
    if (!wholeRows && !wholeCols) {
      this._modeAnnouncement =
        'Select entire rows with Shift+Space or columns with Ctrl+Space first';
      return true;
    }
    // A whole-grid selection counts as rows
    const axis = wholeRows ? 'row' : 'col';
    const index = axis === 'row' ? start.row : start.col;
    const count = axis === 'row' ? end.row - start.row + 1 : end.col - start.col + 1;
    if (remove && count >= (axis === 'row' ? this.rows : this.cols)) {
      this._modeAnnouncement = `Cannot delete every ${axis === 'row' ? 'row' : 'column'}`;
      return true;
    }
    this._changeStructure({ axis, kind: insert ? 'insert' : 'delete', index, count }, 'user');
    this.updateComplete.then(() => this._focusActiveCell());
    return true;
  }

  private _isCellFilled(row: number, col: number): boolean {
    const cell = this._internalData.get(cellKey(row, col));
    return hasContent(cell);
  }

  /**
   * Ctrl+Space: select the full columns spanned by the selection.
   * Shift+Space: select the full rows. Both: select all.
   * The active cell stays where it is.
   */
  private _selectWholeRowsOrColumns(columns: boolean, rows: boolean): void {
    const sel = this._selection;
    const active = sel.activeCell;
    if (columns && rows) {
      sel.selectAll();
    } else if (columns) {
      sel.selectColumns(sel.anchor.col, sel.head.col, active);
    } else {
      sel.selectRows(sel.anchor.row, sel.head.row, active);
    }
    this._dispatchSelectionChange();
    this.updateComplete.then(() => this._focusActiveCell());
  }

  // ─── Row / Column Headers ───────────────────────────

  /**
   * Pointer selection on row/column headers: click selects the row or
   * column, Shift+click extends from the selection's anchor, and dragging
   * across headers selects a span. The active cell moves into the clicked
   * column (or row) but keeps its row (or column), so the view doesn't jump.
   */
  private _handleHeaderPointerDown(e: PointerEvent): void {
    if (e.button !== 0) return;
    if (this._endEditElsewhere()) {
      e.preventDefault();
      e.stopPropagation();
      return;
    }
    const headerEl = (e.target as HTMLElement).closest(
      '[data-header-col], [data-header-row]'
    ) as HTMLElement | null;
    if (!headerEl) return;
    e.preventDefault();

    if (this._isEditing) this._commitEdit();

    const isCol = headerEl.dataset.headerCol !== undefined;
    const attr = isCol ? 'headerCol' : 'headerRow';
    const index = parseInt(headerEl.dataset[attr]!, 10);
    const sel = this._selection;
    const from = e.shiftKey ? (isCol ? sel.anchor.col : sel.anchor.row) : index;

    const select = (to: number) => {
      const lo = Math.min(from, to);
      const hi = Math.max(from, to);
      const active = sel.activeCell;
      if (isCol) {
        const col = e.shiftKey && active.col >= lo && active.col <= hi ? active.col : from;
        sel.selectColumns(from, to, { row: active.row, col });
      } else {
        const row = e.shiftKey && active.row >= lo && active.row <= hi ? active.row : from;
        sel.selectRows(from, to, { row, col: active.col });
      }
      this._dispatchSelectionChange();
      // Keep DOM focus on the active cell, which a drag can move
      this.updateComplete.then(() => this._focusActiveCell());
    };

    select(index);

    const controller = new AbortController();
    const { signal } = controller;
    let last = index;
    const onPointerMove = (ev: PointerEvent) => {
      const el = this.shadowRoot?.elementFromPoint(ev.clientX, ev.clientY) as HTMLElement | null;
      const over = el?.closest?.(isCol ? '[data-header-col]' : '[data-header-row]') as HTMLElement | null;
      if (!over) return;
      const to = parseInt(over.dataset[attr]!, 10);
      if (to !== last) {
        last = to;
        select(to);
      }
    };
    const end = () => controller.abort();
    document.addEventListener('pointermove', onPointerMove, { signal });
    document.addEventListener('pointerup', end, { signal });
    document.addEventListener('pointercancel', end, { signal });
    window.addEventListener('blur', end, { signal });
  }

  // ─── Select All (Corner Header) ─────────────────────

  private _handleSelectAll(): void {
    if (this._endEditElsewhere()) return;
    this._selection.selectAll();
    this._dispatchSelectionChange();
    this.updateComplete.then(() => this._focusActiveCell());
  }

  private _handleCornerKeydown(e: KeyboardEvent): void {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      // Keep the grid handler from also treating this key as "start editing".
      e.stopPropagation();
      this._handleSelectAll();
    }
  }

  // ─── Mouse Handling ─────────────────────────────────

  private _handleCellPointerDown(e: PointerEvent): void {
    const target = e.target as HTMLElement;
    const cellEl = target.closest('[data-row]') as HTMLElement | null;
    if (!cellEl) return;

    const row = parseInt(cellEl.dataset.row!, 10);
    const col = parseInt(cellEl.dataset.col!, 10);

    if (this._isEditing) {
      // In a formula, where a reference fits, the click picks the cell
      if (e.button === 0 && this._editValue.startsWith('=') && (this._refMode || this._isRefPositionInFormula())) {
        this._pickReferenceWithPointer(row, col, e);
        return;
      }
      // Otherwise it commits the edit
      const shown = this._activeSheetId;
      this._commitEdit();
      // The edit was on another sheet, which is shown again: the click was not for it
      if (this._activeSheetId !== shown) return;
    }

    this._selection.startSelection(row, col, e.shiftKey);
    this._dispatchSelectionChange();
    cellEl.focus();

    // Set up drag tracking on the document with guaranteed cleanup
    const controller = new AbortController();
    const { signal } = controller;

    const cleanup = () => controller.abort();

    const onPointerMove = (ev: PointerEvent) => {
      // Find the cell under the pointer
      const el = this.shadowRoot?.elementFromPoint(ev.clientX, ev.clientY) as HTMLElement | null;
      const nearest = el?.closest?.('[data-row]') as HTMLElement | null;
      if (nearest) {
        const r = parseInt(nearest.dataset.row!, 10);
        const c = parseInt(nearest.dataset.col!, 10);
        this._selection.extendSelection(r, c);
        this._dispatchSelectionChange();
      }
    };

    const onPointerUp = () => {
      this._selection.endSelection();
      cleanup();
    };

    document.addEventListener('pointermove', onPointerMove, { signal });
    document.addEventListener('pointerup', onPointerUp, { signal });
    document.addEventListener('pointercancel', onPointerUp, { signal });
    window.addEventListener('blur', onPointerUp, { signal });
  }

  private _handleCellDblClick(e: MouseEvent): void {
    if (this.readOnly) return;
    const target = e.target as HTMLElement;
    const cellEl = target.closest('[data-row]') as HTMLElement | null;
    if (!cellEl) return;

    const row = parseInt(cellEl.dataset.row!, 10);
    const col = parseInt(cellEl.dataset.col!, 10);
    this._selection.moveTo(row, col);
    this._startEditing();
  }

  // ─── Clipboard ──────────────────────────────────────

  private async _handleCopy(): Promise<void> {
    // Rows hidden by the filter are not copied
    const range = this._selection.range;
    await this._clipboardManager.copy(this._internalData, range, this._visibleRowsIn(range), this._computedValue);
  }

  private async _handleCut(): Promise<void> {
    if (this.readOnly || this._pasteInProgress) return;

    const mutations = this._mutationCount;
    const range = this._selection.range;
    const keysToDelete = await this._clipboardManager.cut(
      this._internalData,
      range,
      this._visibleRowsIn(range),
      this._computedValue
    );
    // The grid changed while the clipboard was being written (rows moved, data
    // replaced, cells edited): clearing the saved coordinates could erase
    // something other than what was copied, so leave the data alone.
    if (this._changedSince(mutations)) {
      this._modeAnnouncement = 'Cut not applied because the data changed';
      return;
    }
    if (this._pasteInProgress) {
      this._modeAnnouncement = 'Cut not applied because another change is in progress';
      return;
    }

    const updates: Array<{ id: string; value: string }> = [];
    const selection = this._snapshotSelection();
    for (const key of keysToDelete) {
      if (this._internalData.has(key)) {
        updates.push({ id: key, value: '' });
      }
    }

    const batch = this._buildCommandBatch(updates, 'cut', selection, selection);
    const vetted = batch && this._vetUserBatch(batch);
    if (vetted) {
      this._executeUserBatch(vetted);
    }
  }

  /**
   * The cell a fill copies from. A value spilled into it by a formula is
   * filled as a plain value, unless the formula is part of what is filled
   * from: then the copy of the formula spills it again, and the cell is empty.
   */
  private _fillSource(key: string, inSource: (row: number, col: number) => boolean): CellData | undefined {
    const cell = this._internalData.get(key);
    if (!cell || cell.spilledFrom === undefined || cell.rawValue !== '') return cell;
    const { spilledFrom, array: _array, ...rest } = cell;
    const anchor = parseKey(spilledFrom);
    // As a copy would write it: text that reads like a number stays text, several values stay several
    return {
      ...rest,
      rawValue: inSource(anchor.row, anchor.col) ? '' : computedLiteral(key, cell, this._computedValue),
    };
  }

  /** A formula cell's unformatted result, which a values-only paste writes. */
  private _computedValue = (key: string, cell: CellData): string =>
    cell.array ? arrayToFormula(cell.array) : this._formulaEngine.resultFor(key, cell.rawValue).displayValue;

  private async _handlePaste(options: Required<PasteSpecialOptions> = { mode: 'all', transpose: false }): Promise<void> {
    if (this.readOnly || this._pasteInProgress) return;

    // Paste anchors at the top-left of the selection (not the head, which may
    // be the bottom-right corner after a Shift+arrow or drag selection).
    const { row, col } = this._selection.range.start;
    const revision = this._structureRevision;
    const generation = this._dataGeneration;
    const filterRevision = this._filterRevision;
    const dataRef = this._lastSyncedDataRef;
    // Pasted rows go to the visible rows from the anchor down, skipping
    // rows hidden by the filter (worked out when the paste is read)
    const targets = this._filter ? this._visibleRowsBetween(row, this.rows - 1) : null;
    const rowFor = targets ? (i: number) => targets[i] ?? -1 : undefined;
    const updates = await this._clipboardManager.paste(row, col, this.rows, this.cols, rowFor, {
      values: options.mode === 'values',
      transpose: options.transpose,
    });

    if (!updates || updates.length === 0) return;
    if (
      revision !== this._structureRevision ||
      generation !== this._dataGeneration ||
      filterRevision !== this._filterRevision ||
      dataRef !== this.data
    ) {
      this._modeAnnouncement = 'Paste not applied because rows, columns, the filter or data changed';
      return;
    }
    if (this._pasteInProgress) {
      this._modeAnnouncement = 'Paste not applied because another change is in progress';
      return;
    }

    const selection = this._snapshotSelection();
    let cellUpdates: CellUpdate[];
    if (options.mode === 'formats') {
      if (!updates.some((u) => u.format)) {
        this._modeAnnouncement = 'The clipboard has no formats to paste';
        return;
      }
      // Each target keeps its value and takes the copied cell's format (or none)
      cellUpdates = updates.map((u) => ({
        id: u.id,
        value: this._internalData.get(u.id)?.rawValue ?? '',
        format: u.format ?? null,
      }));
    } else if (options.mode === 'values') {
      // Targets keep their formats
      cellUpdates = updates.map((u) => ({ id: u.id, value: u.value }));
    } else {
      // Pasted formats replace the target's; cells pasted without one keep theirs.
      cellUpdates = updates.map((u) => ({ id: u.id, value: u.value, format: u.format }));
    }
    const batch = this._buildCellBatch(cellUpdates, 'paste', selection);
    const vetted = batch && this._vetUserBatch(batch);
    if (!vetted) return;
    if (options.mode !== 'all' || options.transpose) {
      const what = options.mode === 'values' ? 'values' : options.mode === 'formats' ? 'formats' : 'cells';
      this._modeAnnouncement = `Pasted ${what}${options.transpose ? ', transposed' : ''}`;
    }
    await this._runUserBatch(vetted);
  }

  /** Paste Special dialog: close it, optionally returning focus to the grid. */
  private _closePasteSpecial(focusCell: boolean): void {
    this._pasteSpecialCleanup?.();
    this._pasteSpecialCleanup = null;
    if (!this._pasteSpecial) return;
    this._pasteSpecial = null;
    if (focusCell) this._returnFocus();
  }

  private _applyPasteSpecial(): void {
    const options = this._pasteSpecial;
    if (!options) return;
    this._closePasteSpecial(true);
    void this._handlePaste(options);
  }

  private _handlePasteSpecialKeydown(e: KeyboardEvent): void {
    e.stopPropagation();
    const target = e.target as HTMLElement;
    if (e.key === 'Escape') {
      e.preventDefault();
      this._closePasteSpecial(true);
      this._modeAnnouncement = 'Paste Special closed';
    } else if (e.key === 'Enter' && target.tagName !== 'BUTTON') {
      e.preventDefault();
      this._applyPasteSpecial();
    } else if (e.key === 'Tab') {
      // Modal: keep focus inside
      const dialog = e.currentTarget as HTMLElement;
      const focusables = [...dialog.querySelectorAll<HTMLElement>('input:checked[type=radio], input[type=checkbox], button')];
      const first = focusables[0];
      const last = focusables[focusables.length - 1];
      const current = this.shadowRoot?.activeElement;
      if (e.shiftKey && current === first) {
        e.preventDefault();
        last?.focus();
      } else if (!e.shiftKey && current === last) {
        e.preventDefault();
        first?.focus();
      }
    }
  }

  private _renderPasteSpecial(options: Required<PasteSpecialOptions>) {
    const modes: Array<[PasteMode, string]> = [
      ['all', 'All'],
      ['values', 'Values'],
      ['formats', 'Formats'],
    ];
    return html`
      <div
        class="ls-paste-special"
        role="dialog"
        aria-modal="true"
        aria-labelledby="paste-special-title"
        @keydown="${this._handlePasteSpecialKeydown}"
      >
        <div id="paste-special-title" class="ls-paste-special-title">Paste Special</div>
        <fieldset class="ls-paste-special-group">
          <legend>Paste</legend>
          ${modes.map(
            ([mode, label]) => html`
              <label class="ls-filter-option">
                <input
                  type="radio"
                  name="paste-mode"
                  value="${mode}"
                  .checked=${options.mode === mode}
                  @change="${() => (this._pasteSpecial = { ...options, mode })}"
                />
                <span>${label}</span>
              </label>
            `
          )}
        </fieldset>
        <label class="ls-filter-option">
          <input
            type="checkbox"
            .checked=${options.transpose}
            @change="${(e: Event) =>
              (this._pasteSpecial = { ...options, transpose: (e.target as HTMLInputElement).checked })}"
          />
          <span>Transpose</span>
        </label>
        <div class="ls-filter-row ls-filter-buttons">
          <button type="button" class="ls-filter-ok" @click="${this._applyPasteSpecial}">OK</button>
          <button type="button" class="ls-filter-action" @click="${() => this._closePasteSpecial(true)}">Cancel</button>
        </div>
      </div>
    `;
  }

  /**
   * Apply a command batch progressively in chunks, yielding to the event loop
   * between chunks via requestAnimationFrame to keep the UI responsive.
   */
  private _applyBatchProgressive(batch: CommandBatch): Promise<boolean> {
    const generation = this._dataGeneration;
    return new Promise((resolve) => {
      const deltas = batch.deltas;
      let offset = 0;

      const applyChunk = () => {
        // The data was replaced between frames: the rest of this batch (and
        // its history entry) would apply to cells it was never built for.
        if (generation !== this._dataGeneration) {
          resolve(false);
          return;
        }
        this._mutationCount++;
        const end = Math.min(offset + Y11nSpreadsheet.PASTE_CHUNK_SIZE, deltas.length);

        const written: string[] = [];
        for (let i = offset; i < end; i++) {
          const delta = deltas[i];
          this._applyRawValueByString(delta.id, delta.after);
          if ('formatAfter' in delta) {
            this._applyCellFormat(delta.id, delta.formatAfter);
          }
          written.push(delta.id);
        }
        // Frames drawn before the batch is finalized must not show stale conditional formats
        this._cf.invalidate(written);

        offset = end;

        if (offset < deltas.length) {
          this.requestUpdate();
          requestAnimationFrame(applyChunk);
        } else {
          // Final: restore selection, recalc, dispatch events
          this._restoreSelection(batch.selectionAfter);
          this._finalizeBatch(batch, 'user', 'after');
          resolve(true);
        }
      };

      applyChunk();
    });
  }

  // ─── Cell Clearing ──────────────────────────────────

  private _clearSelectedCells(): void {
    if (this._pasteInProgress) return;
    const range = this._selection.range;
    const updates: Array<{ id: string; value: string }> = [];
    const selection = this._snapshotSelection();

    for (const r of this._visibleRowsBetween(range.start.row, range.end.row)) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        const key = cellKey(r, c);
        if (this._internalData.has(key)) {
          updates.push({ id: key, value: '' });
        }
      }
    }

    const batch = this._buildCommandBatch(updates, 'clear', selection, selection);
    const vetted = batch && this._vetUserBatch(batch);
    if (vetted) {
      this._executeUserBatch(vetted);
    }
  }

  // ─── Fill ───────────────────────────────────────────

  /**
   * Fill down (Ctrl+D) or right (Ctrl+R): copy the first row/column of the
   * selection into the rest of it, shifting relative formula references and
   * copying formats. With a single row/column selected, copies from the
   * row above / column to the left instead.
   */
  private _fillSelection(direction: 'down' | 'right'): void {
    if (this.readOnly || this._pasteInProgress) return;
    const { start, end } = this._selection.range;
    const down = direction === 'down';

    // Rows hidden by the filter are neither source nor target
    const rows = this._visibleRowsBetween(start.row, end.row);
    const cols = Array.from({ length: end.col - start.col + 1 }, (_, i) => start.col + i);
    const along = down ? rows : cols;
    if (along.length === 0) return;
    let source = along[0];
    let targets = along.slice(1);
    if (targets.length === 0) {
      // A single row (column): copy from the one above (to the left)
      const before = down ? this._selection.stepRows(source, -1) : source - 1;
      if (before < 0 || before === source) return;
      targets = [source];
      source = before;
    }

    const updates: CellUpdate[] = [];
    for (const target of targets) {
      for (const cross of down ? cols : rows) {
        const srcKey = down ? cellKey(source, cross) : cellKey(cross, source);
        const dstKey = down ? cellKey(target, cross) : cellKey(cross, target);
        const src = this._fillSource(srcKey, (r, c) =>
          down ? r === source && cols.includes(c) : c === source && rows.includes(r)
        );
        const offset = target - source;
        updates.push({
          id: dstKey,
          value: this._shiftFormula(src?.rawValue ?? '', down ? offset : 0, down ? 0 : offset),
          format: src?.format ?? null,
        });
      }
    }

    const batch = this._buildCellBatch(updates, 'fill', this._snapshotSelection());
    const vetted = batch && this._vetUserBatch(batch);
    if (vetted) void this._runUserBatch(vetted);
  }

  /**
   * Keyboard fill series (Ctrl+Alt+D / Ctrl+Alt+R): the leading rows (or
   * columns) of the selection that contain data are the source, and the rest
   * of the selection is filled from them exactly as a fill-handle drag would,
   * so `1, 2` followed by empty cells continues as `3, 4, ...`.
   */
  private _fillSeriesInSelection(direction: 'down' | 'right'): void {
    if (this.readOnly || this._pasteInProgress) return;
    const range = this._selection.range;
    const down = direction === 'down';
    const rows = this._visibleRowsBetween(range.start.row, range.end.row);
    const cols = Array.from({ length: range.end.col - range.start.col + 1 }, (_, i) => range.start.col + i);
    const along = down ? rows : cols;
    const cross = down ? cols : rows;

    // The leading lines with data are the seed (hidden rows don't count)
    let seeds = 0;
    while (
      seeds < along.length &&
      cross.some((x) => (down ? this._isCellFilled(along[seeds], x) : this._isCellFilled(x, along[seeds])))
    ) {
      seeds++;
    }
    const seedEnd = along[seeds - 1];
    if (down && seeds > 0 && seeds === along.length && this._fillDownToAdjacentData()) {
      this.updateComplete.then(() => this._focusActiveCell());
      return;
    }
    if (seeds === 0 || seeds === along.length) {
      this._modeAnnouncement = `Select the cells to continue and the empty cells ${down ? 'below' : 'to the right of'} them`;
      return;
    }
    const source: SelectionRange = {
      start: { ...range.start },
      end: down ? { row: seedEnd, col: range.end.col } : { row: range.end.row, col: seedEnd },
    };
    this._autoFill(source, { start: { ...range.start }, end: { ...range.end } });
    this.updateComplete.then(() => this._focusActiveCell());
  }

  /**
   * Fill-handle drag: dragging the square at the selection's bottom-right
   * corner previews the cells to fill (along whichever axis the pointer has
   * moved further past the selection) and fills them on release.
   */
  private _handleFillHandlePointerDown(e: PointerEvent): void {
    if (e.button !== 0 || this.readOnly || this._pasteInProgress) return;
    // Keep the cell underneath from starting a new selection
    e.preventDefault();
    e.stopPropagation();

    const { start, end } = this._selection.range;
    const source: SelectionRange = { start: { ...start }, end: { ...end } };
    let target: SelectionRange | null = null;

    const controller = new AbortController();
    const { signal } = controller;
    const onPointerMove = (ev: PointerEvent) => {
      const el = this.shadowRoot?.elementFromPoint(ev.clientX, ev.clientY) as HTMLElement | null;
      const cellEl = el?.closest?.('[data-row]') as HTMLElement | null;
      if (!cellEl) return;
      const next = this._fillTargetFor(
        source,
        parseInt(cellEl.dataset.row!, 10),
        parseInt(cellEl.dataset.col!, 10)
      );
      const same =
        next === target ||
        (next !== null &&
          target !== null &&
          next.start.row === target.start.row &&
          next.start.col === target.start.col &&
          next.end.row === target.end.row &&
          next.end.col === target.end.col);
      if (!same) {
        target = next;
        this._fillPreview = next;
      }
    };
    const finish = (apply: boolean) => {
      controller.abort();
      this._fillPreview = null;
      if (apply && target) this._autoFill(source, target);
      this.updateComplete.then(() => this._focusActiveCell());
    };
    document.addEventListener('pointermove', onPointerMove, { signal });
    document.addEventListener('pointerup', () => finish(true), { signal });
    document.addEventListener('pointercancel', () => finish(false), { signal });
    window.addEventListener('blur', () => finish(false), { signal });
  }

  /** Double-click on the fill handle: fill down as far as the data next to the selection goes. */
  private _handleFillHandleDblClick(e: MouseEvent): void {
    // Keep the cell underneath from opening the editor
    e.preventDefault();
    e.stopPropagation();
    if (this.readOnly || this._pasteInProgress) return;
    if (!this._fillDownToAdjacentData()) this._modeAnnouncement = 'No data next to the selection to fill along';
    this.updateComplete.then(() => this._focusActiveCell());
  }

  /**
   * Fill the selection down to the last row of the data in the column to its
   * left (or, if that is empty, to its right), as Excel does on a fill-handle
   * double-click, stopping above the first filled cell below the selection.
   * Also what Ctrl+Alt+D does when the selection has no empty cells. Returns
   * false if there is nothing to fill.
   */
  private _fillDownToAdjacentData(): boolean {
    const range = this._selection.range;
    const { start, end } = range;
    const filled = (r: number, c: number) => !this._isRowHidden(r) && this._isCellFilled(r, c);
    for (const adj of [start.col - 1, end.col + 1]) {
      if (adj < 0 || adj >= this.cols) continue;
      let last = end.row;
      for (let r = end.row + 1; r < this.rows; r++) {
        if (this._isRowHidden(r)) continue;
        if (!filled(r, adj)) break;
        let blocked = false;
        for (let c = start.col; c <= end.col && !blocked; c++) blocked = filled(r, c);
        if (blocked) break;
        last = r;
      }
      if (last > end.row) {
        this._autoFill({ start: { ...start }, end: { ...end } }, { start: { ...start }, end: { row: last, col: end.col } });
        return true;
      }
    }
    return false;
  }

  /**
   * The range a fill-handle drag from `source` to cell (row, col) covers, or
   * null while the pointer is still inside the source.
   */
  private _fillTargetFor(source: SelectionRange, row: number, col: number): SelectionRange | null {
    const below = row - source.end.row;
    const above = source.start.row - row;
    const right = col - source.end.col;
    const left = source.start.col - col;
    const vertical = Math.max(below, above, 0);
    const horizontal = Math.max(right, left, 0);
    if (vertical === 0 && horizontal === 0) return null;
    if (vertical >= horizontal) {
      return below > 0
        ? { start: { ...source.start }, end: { row, col: source.end.col } }
        : { start: { row, col: source.start.col }, end: { ...source.end } };
    }
    return right > 0
      ? { start: { ...source.start }, end: { row: source.end.row, col } }
      : { start: { row: source.start.row, col }, end: { ...source.end } };
  }

  /**
   * Fill `target` (which contains `source` and extends it along one axis)
   * from `source`, as the fill handle does. Each line of the source along
   * the fill direction is handled separately:
   *
   * - series continue as `seriesFor()` describes: numbers by their linear
   *   trend (1, 2 → 3, 4), single dates by a day, day and month names, and
   *   text ending in a number (Item 1 → Item 2);
   * - anything else repeats, with formulas' relative references shifted and
   *   formats copied.
   */
  private _autoFill(source: SelectionRange, target: SelectionRange): void {
    if (this.readOnly || this._pasteInProgress) return;
    const vertical = target.start.row < source.start.row || target.end.row > source.end.row;
    const along = vertical ? 'row' : 'col';
    const cross = vertical ? 'col' : 'row';
    // Positions along and across the fill, without rows hidden by the filter:
    // a series continues on the visible rows as if the hidden ones weren't there
    const positions = (axis: 'row' | 'col', from: number, to: number) =>
      axis === 'row' ? this._visibleRowsBetween(from, to) : Array.from({ length: to - from + 1 }, (_, i) => from + i);
    const span = positions(along, target.start[along], target.end[along]);
    const linePositions = positions(along, source.start[along], source.end[along]);
    const lineLength = linePositions.length;
    const lineStart = span.indexOf(linePositions[0]);

    const updates: CellUpdate[] = [];
    for (const x of lineLength > 0 ? positions(cross, source.start[cross], source.end[cross]) : []) {
      const keyAt = (p: number) => (vertical ? cellKey(p, x) : cellKey(x, p));
      const line = linePositions.map((p) =>
        this._fillSource(
          keyAt(p),
          // A formula on a row the filter hides is not part of what is filled from
          (r, c) =>
            r >= source.start.row &&
            r <= source.end.row &&
            c >= source.start.col &&
            c <= source.end.col &&
            !this._isRowHidden(r)
        )
      );
      const series = seriesFor(
        line.map((cell) => ({
          raw: cell?.rawValue ?? '',
          number: this._literalNumber(cell),
          isDate: cell?.format?.numberFormat?.type === 'date',
        }))
      );
      span.forEach((p, i) => {
        const offset = i - lineStart;
        if (offset >= 0 && offset < lineLength) return;
        const srcIndex = ((offset % lineLength) + lineLength) % lineLength;
        const src = line[srcIndex];
        let value: string;
        if (series) {
          value = series(offset);
        } else {
          // Shift references by how far the cell actually moves
          const shift = p - linePositions[srcIndex];
          value = this._shiftFormula(src?.rawValue ?? '', vertical ? shift : 0, vertical ? 0 : shift);
        }
        updates.push({ id: keyAt(p), value, format: src?.format ?? null });
      });
    }

    const selectionAfter: SelectionSnapshot = {
      anchor: { ...target.start },
      head: { ...target.end },
      active: { ...this._selection.activeCell },
    };
    this._modeAnnouncement = `Filled ${updates.length} cell${updates.length === 1 ? '' : 's'}`;
    const batch = this._buildCellBatch(updates, 'fill', this._snapshotSelection());
    const vetted = batch && this._vetUserBatch(batch);
    if (vetted) {
      vetted.selectionAfter = selectionAfter;
      void this._runUserBatch(vetted);
    } else {
      this._restoreSelection(selectionAfter);
    }
  }

  /** The number a cell holds as a literal (not a formula), or null. */
  private _literalNumber(cell: CellData | undefined): number | null {
    if (!cell || cell.rawValue === '' || cell.rawValue.startsWith('=')) return null;
    const evaluated = this._formulaEngine.evaluate(cell.rawValue);
    if (evaluated.type !== 'number') return null;
    const n = Number(evaluated.displayValue);
    return Number.isFinite(n) ? n : null;
  }

  /** Shift the relative references of a formula; non-formulas are returned as-is. */
  private _shiftFormula(raw: string, rowOffset: number, colOffset: number): string {
    if (!raw.startsWith('=') || (rowOffset === 0 && colOffset === 0)) return raw;
    return this._clipboardManager.adjustFormulaReferences(raw, rowOffset, colOffset);
  }

  /**
   * Ctrl/Cmd+Enter while editing: commit the edited value into every cell of
   * the selection. Formulas are written relative to each cell, as if the
   * edited cell were copied there.
   */
  private _commitEditToSelection(): void {
    if (!this._isEditing || !this._editingCellKey) return;
    this._returnToEditSheet();
    const origin = parseKey(this._editingCellKey);
    const typed = this._editValue;
    const { start, end } = this._selection.range;

    if (start.row === end.row && start.col === end.col) {
      this._commitEdit('keep');
      return;
    }
    const asList = this._editListMode;
    const value = this._interpretInput(this._editingCellKey, typed, asList).value;
    // Validate the edited cell's value up front so the editor can stay open;
    // the other cells are validated individually when the batch is vetted.
    if (!this._checkEditValue(this._editingCellKey, value, 'keep')) return;

    this._isEditing = false;
    this._editingCellKey = null;
    this._validationMessage = '';

    const updates: CellUpdate[] = [];
    for (const r of this._visibleRowsBetween(start.row, end.row)) {
      for (let c = start.col; c <= end.col; c++) {
        const id = cellKey(r, c);
        // Each cell reads the typed text itself: a percent cell turns 13 into 13%
        const own = this._interpretInput(id, typed, asList);
        const update: CellUpdate = {
          id,
          value: typed.startsWith('=') ? this._shiftFormula(value, r - origin.row, c - origin.col) : own.value,
        };
        const format = own.format;
        if (format) update.format = format;
        updates.push(update);
      }
    }
    const cellCount = updates.length;
    this._modeAnnouncement = `Filled ${cellCount} cells`;

    const batch = this._buildCellBatch(updates, 'fill', this._snapshotSelection());
    const vetted = batch && this._vetUserBatch(batch);
    if (vetted) void this._runUserBatch(vetted);
    this.updateComplete.then(() => this._focusActiveCell());
  }

  // ─── Scroll Handling ────────────────────────────────

  private _handleScroll(e: Event): void {
    const target = e.target as HTMLElement;
    this._scrollTop = target.scrollTop;
    this._scrollLeft = target.scrollLeft;
  }

  // ─── Find Bar ───────────────────────────────────────

  private _handleFindInput(e: Event): void {
    this._findQuery = (e.target as HTMLInputElement).value;
    // Like browser find: jump to the first match at or after the active cell
    this._findStep(1, true);
  }

  private _handleFindKeydown(e: KeyboardEvent): void {
    const ctrl = e.ctrlKey || e.metaKey;
    if (e.key === 'Enter' || e.key === 'F3') {
      e.preventDefault();
      this._findStep(e.shiftKey ? -1 : 1);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      this.closeFind();
    } else if (ctrl && e.key.toLowerCase() === 'f') {
      e.preventDefault();
      this._findInput?.select();
    } else {
      return;
    }
    e.stopPropagation();
  }

  /**
   * Move the active cell to the next (1) or previous (-1) match after the
   * active cell, wrapping around. With `inclusive`, the active cell itself
   * counts. Focus stays in the find input.
   */
  private _findStep(direction: 1 | -1, inclusive = false): void {
    if (!this._findQuery) {
      this._findStatus = '';
      return;
    }
    // Cells in rows hidden by the filter can't be selected, so they don't count
    const matches = this.findCells(this._findQuery).filter((key) => !this._isRowHidden(parseKey(key).row));
    if (matches.length === 0) {
      this._findStatus = 'No matches';
      return;
    }
    const active = this._selection.activeCell;
    const current = active.row * this.cols + active.col;
    const position = (key: string) => {
      const { row, col } = parseKey(key);
      return row * this.cols + col;
    };
    let index = -1;
    if (direction === 1) {
      index = matches.findIndex((k) => (inclusive ? position(k) >= current : position(k) > current));
      if (index < 0) index = 0;
    } else {
      for (let i = matches.length - 1; i >= 0; i--) {
        if (position(matches[i]) < current) {
          index = i;
          break;
        }
      }
      if (index < 0) index = matches.length - 1;
    }
    const { row, col } = parseKey(matches[index]);
    this._selection.moveTo(row, col);
    this._ensureCellVisible(row, col);
    this._dispatchSelectionChange();
    this._findStatus = `${index + 1} of ${matches.length}`;
  }

  // ─── Editor Input ───────────────────────────────────

  private _handleEditorInput(e: Event): void {
    this._editValue = (e.target as HTMLInputElement).value;
    this._validationMessage = '';
    this._exitRefMode();
  }

  private _handleEditorBlur(e: FocusEvent): void {
    if (!this._isEditing || this._refPickPointer) return;
    // With a formula open, going to the sheet tabs is how a reference on another sheet is picked
    const next = e.relatedTarget as Node | null;
    if (next && this._editValue.startsWith('=') && this.renderRoot.querySelector('y11n-sheet-tabs')?.contains(next)) {
      return;
    }
    this._commitEdit();
  }

  private _handleFormulaBarModeChange(
    e: CustomEvent<{ mode: FormulaBarMode }>
  ): void {
    this._formulaBarMode = e.detail.mode;
  }

  private _handleFormulaBarCommit(e: CustomEvent<{ value: string; cellRef: string }>): void {
    if (!this._applyFormulaBarCommit(e.detail)) {
      // Rejected: drop the draft so it can't be committed into another cell later
      this.shadowRoot?.querySelector<Y11nFormulaBar>('y11n-formula-bar')?.resetDraft();
    }
  }

  /** Apply a formula-bar commit; returns false if nothing was written. */
  private _applyFormulaBarCommit(detail: { value: string; cellRef: string }): boolean {
    if (this.readOnly) return false;

    // Use the cell key from the event (which cell the formula bar was showing)
    // rather than the current active cell, which may have changed due to clicking.
    let targetKey: string;
    if (this._editingCellKey) {
      if (this._isEditing) this._returnToEditSheet();
      targetKey = this._editingCellKey;
    } else {
      const coord = refToCoord(detail.cellRef);
      targetKey = coord ? cellKey(coord.row, coord.col) : this._getActiveCellKey();
    }

    if (this._isCellLocked(targetKey)) {
      this._modeAnnouncement = `${coordToRef(parseKey(targetKey))} is read-only`;
      return false;
    }
    const oldValue = this._internalData.get(targetKey)?.rawValue ?? '';
    // The formula bar shows a list as its items, so what comes back is a list too
    const { value, format } = this._interpretInput(targetKey, detail.value, listItemsOf(oldValue) !== null);
    const message = value === oldValue ? null : this._validate(targetKey, value);
    if (message) {
      this._dispatchValidationError({ errors: [{ cellId: targetKey, value, message }], operation: 'edit' });
      this._modeAnnouncement = `Invalid value: ${message}`;
      return false;
    }

    if (this._isEditing) {
      this._isEditing = false;
      this._editingCellKey = null;
      this._validationMessage = '';
    }

    // An unchanged value is a successful no-op, not a rejection
    return this._commitRawValue(targetKey, value, { validate: false, format }) || value === oldValue;
  }

  // ─── Event Dispatching ──────────────────────────────

  private _dispatchCellChange(detail: CellChangeDetail): void {
    this.dispatchEvent(new CustomEvent('cell-change', { detail, bubbles: true, composed: true }));
  }

  private _dispatchSelectionChange(): void {
    if (this._arrayCard && this._arrayCard.key !== this._getActiveCellKey()) this.closeCellValues(false);
    const detail: SelectionChangeDetail = { range: this._selection.range };
    this.dispatchEvent(
      new CustomEvent('selection-change', { detail, bubbles: true, composed: true })
    );
  }

  private _dispatchDataChange(detail: DataChangeDetail): void {
    this.dispatchEvent(
      new CustomEvent('data-change', { detail, bubbles: true, composed: true })
    );
  }

  private _dispatchFormatChange(detail: FormatChangeDetail): void {
    this.dispatchEvent(
      new CustomEvent('format-change', { detail, bubbles: true, composed: true })
    );
  }

  private _dispatchValidationError(detail: ValidationErrorDetail): void {
    this.dispatchEvent(
      new CustomEvent('validation-error', { detail, bubbles: true, composed: true })
    );
  }

  private _dispatchStructureChange(detail: StructureChangeDetail): void {
    this.dispatchEvent(
      new CustomEvent('structure-change', { detail, bubbles: true, composed: true })
    );
  }

  private _dispatchSortChange(detail: SortDetail): void {
    this.dispatchEvent(new CustomEvent('sort-change', { detail, bubbles: true, composed: true }));
  }

  // ─── Virtual Rendering ──────────────────────────────

  /**
   * Compute which rows and columns are visible based on scroll position.
   * We add a buffer of extra rows/cols for smooth scrolling.
   */
  private _getVisibleRange(): { rows: number[]; cols: number[] } {
    const cellHeight = this._cellSize.height;
    const cellWidth = this._cellSize.width;
    const gridEl = this._grid;
    const viewHeight =
      gridEl?.clientHeight || cellHeight * Y11nSpreadsheet.INITIAL_VIEW_ROWS;
    const viewWidth =
      gridEl?.clientWidth || cellWidth * Y11nSpreadsheet.INITIAL_VIEW_COLS;

    const buffer = 5;

    // Rows are laid out by visible index: hidden (filtered-out) rows take no space
    const visible = this._visibleRowList();
    const rowCount = visible ? visible.length : this.rows;
    const startRow = Math.max(0, Math.floor(this._scrollTop / cellHeight) - buffer);
    const endRow = Math.min(rowCount, Math.ceil((this._scrollTop + viewHeight) / cellHeight) + buffer);
    const startCol = Math.max(0, this._colAt(this._scrollLeft) - buffer);
    const endCol = Math.min(this.cols, this._colAt(this._scrollLeft + viewWidth) + 1 + buffer);

    // The active cell (roving tabindex target) and the reference cursor must
    // always be in the DOM. Render them individually rather than stretching
    // the window, which would render every row between the viewport and a
    // far-away active cell after the user scrolls with the mouse.
    const { row: activeRow, col: activeCol } = this._selection.activeCell;
    const extraRows = [activeRow];
    const extraCols = [activeCol];
    // Frozen rows and columns are always on screen
    for (let r = 0; r < this._frozenRowCount; r++) extraRows.push(r);
    for (let c = 0; c < this._frozenColCount; c++) extraCols.push(c);
    if (this._refMode) {
      extraRows.push(this._refCursorRow);
      extraCols.push(this._refCursorCol);
    }

    if (visible) {
      const rows = new Set<number>();
      for (let i = startRow; i < endRow; i++) rows.add(visible[i]);
      for (const r of extraRows) if (!this._isRowHidden(r)) rows.add(r);
      return { rows: [...rows].sort((a, b) => a - b), cols: windowWithExtras(startCol, endCol, extraCols) };
    }
    return {
      rows: windowWithExtras(startRow, endRow, extraRows),
      cols: windowWithExtras(startCol, endCol, extraCols),
    };
  }

  private _ensureCellVisible(row: number, col: number): void {
    const gridEl = this._grid;
    if (!gridEl) return;

    const cellHeight = this._cellSize.height;
    const cellWidth = this._cellSize.width;
    const viewHeight =
      gridEl.clientHeight || cellHeight * Y11nSpreadsheet.INITIAL_VIEW_ROWS;
    const viewWidth =
      gridEl.clientWidth || cellWidth * Y11nSpreadsheet.INITIAL_VIEW_COLS;
    const dataViewportWidth = Math.max(
      this._colWidth(col),
      viewWidth - Y11nSpreadsheet.ROW_HEADER_WIDTH
    );
    // Subtract the sticky column header height so cells aren't scrolled
    // behind it (mirrors how dataViewportWidth subtracts the row header).
    const dataViewportHeight = Math.max(cellHeight, viewHeight - cellHeight);

    const cellTop = this._rowTop(row);
    const cellBottom = cellTop + cellHeight;
    const cellLeft = this._colLeft(col);
    const cellRight = cellLeft + this._colWidth(col);

    // Frozen rows/columns never need scrolling; the others must clear the
    // frozen area, which covers the start of the viewport
    const frozenHeight = this._frozenHeight;
    const frozenWidth = this._frozenWidth;
    let nextScrollTop = this._scrollTop;
    if (row < this._frozenRowCount) {
      // stays put
    } else if (cellTop - frozenHeight < nextScrollTop) {
      nextScrollTop = Math.max(0, cellTop - frozenHeight);
    } else if (cellBottom > nextScrollTop + dataViewportHeight) {
      nextScrollTop = cellBottom - dataViewportHeight;
    }

    let nextScrollLeft = this._scrollLeft;
    if (col < this._frozenColCount) {
      // stays put
    } else if (cellLeft - frozenWidth < nextScrollLeft) {
      nextScrollLeft = Math.max(0, cellLeft - frozenWidth);
    } else if (cellRight > nextScrollLeft + dataViewportWidth) {
      nextScrollLeft = cellRight - dataViewportWidth;
    }

    if (nextScrollTop !== this._scrollTop || nextScrollLeft !== this._scrollLeft) {
      this._scrollTop = nextScrollTop;
      this._scrollLeft = nextScrollLeft;
      gridEl.scrollTop = nextScrollTop;
      gridEl.scrollLeft = nextScrollLeft;
      // The browser clamps to the scrollable size without firing a scroll
      // event when the position doesn't change: keep the state in sync
      this._scrollLeft = gridEl.scrollLeft;
    }
  }

  // ─── Cell Size ──────────────────────────────────────

  /**
   * Measure the rendered cell size from the hidden probe, which is sized with
   * the same `--_cell-width` / `--_cell-height` as the cells. This resolves
   * any CSS length (`7rem`, `2em`, `calc(...)`), not just px values.
   */
  private _measureCellSize(): void {
    const probe = this.renderRoot.querySelector<HTMLElement>('.ls-size-probe');
    if (!probe) return;
    const style = getComputedStyle(probe);
    this._setCellSize(parseFloat(style.width), parseFloat(style.height));
  }

  private _setCellSize(width: number, height: number): void {
    // A hidden (display: none) grid measures 0: keep the last good size
    if (!(width > 0) || !(height > 0)) return;
    if (width === this._cellSize.width && height === this._cellSize.height) return;
    this._cellSize = { width, height };
  }

  /** Re-measure whenever the probe resizes. Returns false where ResizeObserver is unavailable. */
  private _observeCellSize(): boolean {
    const probe = this.renderRoot.querySelector<HTMLElement>('.ls-size-probe');
    if (!probe || typeof ResizeObserver === 'undefined') return false;
    this._cellSizeObserver ??= new ResizeObserver((entries) => {
      const entry = entries[entries.length - 1];
      const box = entry.borderBoxSize?.[0];
      if (box) this._setCellSize(box.inlineSize, box.blockSize);
      else this._setCellSize(entry.contentRect.width, entry.contentRect.height);
      if (this._cellSize !== this._renderedCellSize) this.requestUpdate();
    });
    this._cellSizeObserver.observe(probe);
    return true;
  }

  // ─── Editor Positioning ─────────────────────────────

  private _getEditorStyle(): Record<string, string> {
    if (!this._isEditing) return { display: 'none' };
    if (this._editingElsewhere) {
      // The edited cell is on another sheet: keep the formula in view at the bottom of this one
      return {
        display: 'block',
        position: 'absolute',
        left: `${Y11nSpreadsheet.ROW_HEADER_WIDTH + 4}px`,
        bottom: '4px',
        width: 'min(420px, calc(100% - 70px))',
        height: `${this._cellSize.height}px`,
        'z-index': '10',
      };
    }

    const { row, col } = this._selection.activeCell;
    const cellHeight = this._cellSize.height;
    const cellWidth = this._colWidth(col);
    const headerWidth = Y11nSpreadsheet.ROW_HEADER_WIDTH;

    // Below the header row; frozen cells don't move with the scroll position
    const top = cellHeight + this._rowTop(row) - (row < this._frozenRowCount ? 0 : this._scrollTop);
    const left = headerWidth + this._colLeft(col) - (col < this._frozenColCount ? 0 : this._scrollLeft);

    const styles: Record<string, string> = {
      display: 'block',
      position: 'absolute',
      top: `${top}px`,
      left: `${left}px`,
      width: `${cellWidth}px`,
      height: `${cellHeight}px`,
      'z-index': '10',
    };
    // A scrolling cell can slide under the headers and frozen panes: clip the
    // editor to what's left of the cell, so it never covers them
    const clipTop = row < this._frozenRowCount ? 0 : Math.max(0, cellHeight + this._frozenHeight - top);
    const clipLeft = col < this._frozenColCount ? 0 : Math.max(0, headerWidth + this._frozenWidth - left);
    if (clipTop > 0 || clipLeft > 0) styles['clip-path'] = `inset(${clipTop}px 0 0 ${clipLeft}px)`;

    // Inherit cell formatting into the editor (via styleMap for safety)
    const fmtStyles = this._getCellStyles(row, col);
    for (const [k, v] of Object.entries(fmtStyles)) {
      if (!['justify-content', 'background-color', 'align-items', 'white-space', 'overflow-wrap'].includes(k)) {
        styles[k] = v;
      }
    }

    return styles;
  }

  /** Position the validation message just below the editor. */
  private _getValidationMessageStyle(): Record<string, string> {
    const editor = this._getEditorStyle();
    const cellHeight = this._cellSize.height;
    return {
      top: `${parseFloat(editor.top) + cellHeight}px`,
      left: editor.left,
    };
  }

  // ─── Render ─────────────────────────────────────────

  static styles = [
    conditionalIconStyles,
    css`
    :host {
      display: block;
      height: 100%;
      position: relative;
      font-family: var(--ls-font-family, system-ui, -apple-system, sans-serif);
      font-size: var(--ls-font-size, 13px);
      color: var(--ls-text-color, #333);
      --_cell-width: var(--ls-cell-width, 100px);
      --_cell-height: var(--ls-cell-height, 28px);
      --_header-bg: var(--ls-header-bg, #f3f3f3);
      --_border-color: var(--ls-border-color, #e0e0e0);
      --_selection-border: var(--ls-selection-border, 2px solid #1a73e8);
      --_selection-bg: var(--ls-selection-bg, rgba(26, 115, 232, 0.2));
      --_focus-ring: var(--ls-focus-ring, 2px solid #1a73e8);
      --_editor-bg: var(--ls-editor-bg, #fff);
      /* Background of frozen cells, which must be opaque over scrolled cells */
      --_cell-bg: var(--ls-cell-bg, #fff);
      --_frozen-line: var(--ls-frozen-line-color, #9aa0a6);
      --_editor-shadow: var(--ls-editor-shadow, 0 2px 6px rgba(0,0,0,0.2));
      --_row-header-width: 50px;
    }

    *,
    *::before,
    *::after {
      box-sizing: border-box;
    }

    .ls-wrapper {
      position: relative;
      display: flex;
      flex-direction: column;
      width: 100%;
      height: 100%;
      overflow: hidden;
    }

    .ls-status {
      flex: 0 0 auto;
      display: flex;
      justify-content: flex-end;
      gap: 16px;
      min-height: 22px;
      padding: 2px 12px;
      font-size: 12px;
      color: var(--ls-status-color, #555);
      background: var(--ls-status-bg, #f8fafc);
      border-top: 1px solid var(--_border-color);
      white-space: nowrap;
      overflow: hidden;
    }

    .ls-grid-shell {
      position: relative;
      flex: 1 1 auto;
      min-height: 0;
      overflow: hidden;
    }

    /* Sized like a cell so the component can read the resolved px size */
    .ls-size-probe {
      position: absolute;
      top: 0;
      left: 0;
      width: var(--_cell-width);
      height: var(--_cell-height);
      visibility: hidden;
      pointer-events: none;
    }

    .ls-grid {
      display: grid;
      grid-template-columns: var(--_row-header-width) repeat(var(--cols), var(--_cell-width));
      align-content: start;
      overflow: auto;
      position: relative;
      width: 100%;
      height: round(down, 100%, var(--_cell-height));
      scroll-snap-type: y mandatory;
      outline: none;
    }

    .ls-header-row {
      display: contents;
    }

    .ls-row {
      display: contents;
    }

    /*
     * Stacking, bottom to top: cells (focused/active: 1), frozen cells (2),
     * cells frozen on both axes (3), row headers (4), frozen row headers (5),
     * column headers (6), frozen column headers (7), the corner (8).
     */
    .ls-corner-header,
    .ls-col-header {
      position: sticky;
      top: 0;
      z-index: 6;
      display: flex;
      align-items: center;
      justify-content: center;
      height: var(--_cell-height);
      background: var(--_header-bg);
      border-right: 1px solid var(--_border-color);
      border-bottom: 2px solid var(--_border-color);
      font-weight: 600;
      user-select: none;
    }

    .ls-corner-header {
      position: sticky;
      left: 0;
      z-index: 8;
      scroll-snap-align: start;
    }

    .ls-col-header,
    .ls-row-header {
      cursor: pointer;
    }

    /* Drag handle on a column header's right border */
    .ls-col-resizer {
      position: absolute;
      top: 0;
      right: -1px;
      width: 7px;
      height: 100%;
      cursor: col-resize;
      touch-action: none;
      z-index: 1;
    }

    .ls-col-resizer:hover {
      background: linear-gradient(to right, transparent 2px, var(--ls-resizer-color, #1a73e8) 2px, var(--ls-resizer-color, #1a73e8) 4px, transparent 4px);
    }

    .ls-col-header.in-selection,
    .ls-row-header.in-selection {
      background: var(--ls-header-selected-bg, #dbe7fb);
    }

    .ls-row-header {
      position: sticky;
      left: 0;
      z-index: 4;
      display: flex;
      align-items: center;
      justify-content: center;
      height: var(--_cell-height);
      background: var(--_header-bg);
      border-right: 2px solid var(--_border-color);
      border-bottom: 1px solid var(--_border-color);
      font-weight: 600;
      user-select: none;
      scroll-snap-align: start;
    }

    .ls-cell {
      display: flex;
      align-items: center;
      padding: 0 4px;
      height: var(--_cell-height);
      border-right: 1px solid var(--_border-color);
      border-bottom: 1px solid var(--_border-color);
      overflow: hidden;
      white-space: nowrap;
      text-overflow: ellipsis;
      cursor: cell;
      outline: none;
      user-select: none;
      position: relative;
      /* Contain the fill handle's z-index, so it can't rise above frozen cells */
      isolation: isolate;
    }

    .ls-cell.read-only-cell:not([aria-selected='true']) {
      background: var(--ls-readonly-cell-bg, transparent);
    }

    .ls-cell.fill-preview {
      background: var(--ls-fill-preview-bg, rgba(26, 115, 232, 0.08));
      outline: 1px dashed var(--ls-fill-preview-border, #1a73e8);
      outline-offset: -1px;
    }

    .ls-fill-handle {
      position: absolute;
      right: 0;
      bottom: 0;
      width: 7px;
      height: 7px;
      background: var(--ls-fill-handle-color, #1a73e8);
      border: 1px solid #fff;
      cursor: crosshair;
      touch-action: none;
      z-index: 2;
    }

    .ls-cell:focus {
      outline: var(--_focus-ring);
      outline-offset: -2px;
      z-index: 1;
    }

    .ls-cell[aria-selected="true"] {
      background: var(--_selection-bg);
    }

    /*
     * A cell with a fill of its own (a format or a conditional format) sets
     * its background inline, which hides the selection tint: lay it over the fill
     */
    .ls-cell.has-fill[aria-selected='true'] {
      box-shadow: inset 0 0 0 100vmax var(--_selection-bg);
    }

    /* The range a formula's result spills over, outlined while the active cell is in it */
    .ls-cell.in-spill::after {
      content: '';
      position: absolute;
      inset: 0;
      pointer-events: none;
      border: 0 solid var(--ls-spill-border, #1a73e8);
    }
    .ls-cell.spill-top::after { border-top-width: 1px; }
    .ls-cell.spill-bottom::after { border-bottom-width: 1px; }
    .ls-cell.spill-left::after { border-left-width: 1px; }
    .ls-cell.spill-right::after { border-right-width: 1px; }

    /* A cell that holds several values: an icon that opens them */
    .ls-array-button {
      flex: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 16px;
      height: 16px;
      margin-right: 4px;
      border-radius: 3px;
      color: var(--ls-array-icon-color, #5f6368);
      cursor: pointer;
    }
    .ls-array-button:hover {
      background: var(--ls-array-icon-hover-bg, rgba(0, 0, 0, 0.08));
    }

    .ls-array-card {
      position: absolute;
      z-index: 20;
      display: flex;
      flex-direction: column;
      background: var(--ls-filter-menu-bg, #fff);
      border: 1px solid var(--_border-color);
      border-radius: 4px;
      box-shadow: var(--_editor-shadow);
      font-size: 13px;
      overflow: hidden;
    }
    .ls-array-card-head {
      display: flex;
      align-items: center;
      gap: 6px;
      padding: 6px 8px;
      border-bottom: 1px solid var(--_border-color);
    }
    .ls-array-card-title {
      flex: 1;
      min-width: 0;
      font-weight: 600;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .ls-array-back,
    .ls-array-close,
    .ls-array-nested {
      font: inherit;
      color: inherit;
      background: none;
      border: 1px solid var(--_border-color);
      border-radius: 3px;
      padding: 1px 6px;
      cursor: pointer;
    }
    .ls-array-close {
      border-color: transparent;
      font-size: 16px;
      line-height: 1;
    }
    .ls-array-nested {
      max-width: 100%;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      text-align: left;
    }
    .ls-array-back:focus-visible,
    .ls-array-close:focus-visible,
    .ls-array-nested:focus-visible,
    .ls-array-card-body:focus-visible {
      outline: var(--_focus-ring);
      outline-offset: -2px;
    }
    .ls-array-card-body {
      overflow: auto;
      padding: 6px 8px;
      min-height: 0;
    }
    .ls-array-list {
      margin: 0;
      padding-left: 2.2em;
    }
    .ls-array-list li {
      padding: 1px 0;
    }
    .ls-array-table {
      border-collapse: collapse;
    }
    .ls-array-table th,
    .ls-array-table td {
      border: 1px solid var(--_border-color);
      padding: 2px 6px;
      text-align: left;
      white-space: nowrap;
    }
    .ls-array-table th {
      background: var(--_header-bg);
      font-weight: 600;
    }
    .ls-array-blank,
    .ls-array-more {
      color: var(--ls-array-icon-color, #5f6368);
      font-style: italic;
    }
    .ls-array-more {
      margin: 6px 0 0;
    }

    .ls-cell.active-cell {
      outline: var(--_selection-border);
      outline-offset: -2px;
      z-index: 1;
    }

    .ls-cell.ref-highlight {
      outline: 2px dashed var(--ls-ref-highlight-border, #1a73e8);
      outline-offset: -2px;
      background: var(--ls-ref-highlight-bg, rgba(26, 115, 232, 0.15));
      z-index: 1;
    }

    .ls-cell .cell-text {
      overflow: hidden;
      text-overflow: ellipsis;
    }

    /* Frozen panes: sticky (top/left set inline), opaque, above scrolled cells */
    .ls-cell.frozen-row,
    .ls-cell.frozen-col {
      position: sticky;
      z-index: 2;
      background-color: var(--_cell-bg);
    }

    .ls-cell.frozen-row.frozen-col {
      z-index: 3;
    }

    /* Translucent states are layered over the opaque background */
    .ls-cell.frozen-row[aria-selected='true'],
    .ls-cell.frozen-col[aria-selected='true'] {
      background: linear-gradient(var(--_selection-bg), var(--_selection-bg)), var(--_cell-bg);
    }

    .ls-cell.frozen-row.read-only-cell:not([aria-selected='true']),
    .ls-cell.frozen-col.read-only-cell:not([aria-selected='true']) {
      background: linear-gradient(var(--ls-readonly-cell-bg, transparent), var(--ls-readonly-cell-bg, transparent)), var(--_cell-bg);
    }

    .ls-cell.frozen-row.fill-preview,
    .ls-cell.frozen-col.fill-preview {
      background: linear-gradient(var(--ls-fill-preview-bg, rgba(26, 115, 232, 0.08)), var(--ls-fill-preview-bg, rgba(26, 115, 232, 0.08))), var(--_cell-bg);
    }

    .ls-cell.frozen-row.ref-highlight,
    .ls-cell.frozen-col.ref-highlight {
      background: linear-gradient(var(--ls-ref-highlight-bg, rgba(26, 115, 232, 0.15)), var(--ls-ref-highlight-bg, rgba(26, 115, 232, 0.15))), var(--_cell-bg);
    }

    .ls-row-header.frozen-row {
      z-index: 5;
      /* Sticky snap areas would fight the scrolling rows' snap points */
      scroll-snap-align: none;
    }

    .ls-col-header.frozen-col {
      z-index: 7;
    }

    /* The freeze lines */
    .frozen-row-edge {
      border-bottom: 2px solid var(--_frozen-line);
    }

    .ls-cell.frozen-col-edge,
    .ls-col-header.frozen-col-edge {
      border-right: 2px solid var(--_frozen-line);
    }

    #editor {
      position: absolute;
      box-sizing: border-box;
      padding: 0 4px;
      border: var(--_selection-border);
      background: var(--_editor-bg);
      box-shadow: var(--_editor-shadow);
      font-family: inherit;
      font-size: inherit;
      color: inherit;
      outline: none;
      z-index: 10;
    }

    .ls-validation-message {
      position: absolute;
      z-index: 11;
      max-width: 280px;
      padding: 4px 8px;
      background: var(--ls-validation-bg, #fdecea);
      color: var(--ls-validation-color, #8a1c12);
      border: 1px solid var(--ls-validation-border, #d93025);
      font-size: 12px;
      white-space: normal;
      pointer-events: none;
    }

    .ls-find {
      position: absolute;
      top: 4px;
      right: 20px;
      z-index: 20;
      display: flex;
      align-items: center;
      gap: 4px;
      padding: 4px 6px;
      background: var(--ls-find-bg, #fff);
      border: 1px solid var(--_border-color);
      border-radius: 4px;
      box-shadow: var(--_editor-shadow);
    }

    .ls-find-input {
      width: 160px;
      padding: 2px 4px;
      font: inherit;
      border: 1px solid var(--_border-color);
      border-radius: 2px;
    }

    .ls-find-input:focus-visible,
    .ls-find-button:focus-visible {
      outline: var(--_focus-ring);
      outline-offset: 1px;
    }

    .ls-find-status {
      min-width: 5em;
      font-size: 12px;
      color: var(--ls-find-status-color, #555);
      text-align: center;
    }

    .ls-find-button {
      min-width: 24px;
      height: 24px;
      padding: 0 4px;
      font: inherit;
      color: inherit;
      background: transparent;
      border: 1px solid transparent;
      border-radius: 2px;
      cursor: pointer;
    }

    .ls-find-button:hover {
      background: var(--_header-bg);
      border-color: var(--_border-color);
    }

    /* Filters: a dropdown button in each header cell of the filtered table */
    .ls-cell.filter-header {
      padding-right: 26px;
    }

    .ls-filter-button {
      position: absolute;
      top: 50%;
      /* Clear of the fill handle in the corner */
      right: 7px;
      transform: translateY(-50%);
      display: flex;
      align-items: center;
      justify-content: center;
      width: 16px;
      height: 16px;
      color: var(--ls-filter-button-color, #555);
      background: var(--ls-filter-button-bg, #fff);
      border: 1px solid var(--_border-color);
      border-radius: 3px;
      cursor: pointer;
    }

    .ls-filter-button:hover {
      border-color: var(--ls-filter-active-color, #1a73e8);
    }

    .ls-filter-button.filtered {
      color: #fff;
      background: var(--ls-filter-active-color, #1a73e8);
      border-color: var(--ls-filter-active-color, #1a73e8);
    }

    /* Row numbers of a filtered table, which skip the hidden rows */
    .ls-row-header.filtered-row {
      color: var(--ls-filtered-row-header-color, #1a73e8);
    }

    .ls-filter-menu {
      position: absolute;
      z-index: 20;
      display: flex;
      flex-direction: column;
      gap: 6px;
      padding: 8px;
      overflow: auto;
      background: var(--ls-filter-menu-bg, #fff);
      border: 1px solid var(--_border-color);
      border-radius: 4px;
      box-shadow: var(--_editor-shadow);
      font-size: 13px;
    }

    .ls-paste-special {
      position: absolute;
      top: 8px;
      left: 50%;
      transform: translateX(-50%);
      z-index: 20;
      display: flex;
      flex-direction: column;
      gap: 6px;
      width: 220px;
      padding: 10px;
      background: var(--ls-filter-menu-bg, #fff);
      border: 1px solid var(--_border-color);
      border-radius: 4px;
      box-shadow: var(--_editor-shadow);
      font-size: 13px;
    }

    .ls-paste-special-title {
      font-weight: 600;
    }

    .ls-paste-special-group {
      margin: 0;
      padding: 4px 8px;
      border: 1px solid var(--_border-color);
      border-radius: 3px;
    }

    .ls-filter-row {
      display: flex;
      gap: 6px;
    }

    .ls-filter-row > button {
      flex: 1 1 0;
    }

    .ls-filter-buttons {
      justify-content: flex-end;
    }

    .ls-filter-action,
    .ls-filter-ok {
      min-height: 26px;
      padding: 2px 8px;
      font: inherit;
      color: inherit;
      text-align: left;
      background: transparent;
      border: 1px solid var(--_border-color);
      border-radius: 3px;
      cursor: pointer;
    }

    .ls-filter-row > button,
    .ls-filter-ok {
      text-align: center;
    }

    .ls-filter-action:hover:not(:disabled) {
      background: var(--_header-bg);
    }

    .ls-filter-ok {
      color: #fff;
      background: var(--ls-filter-active-color, #1a73e8);
      border-color: var(--ls-filter-active-color, #1a73e8);
    }

    .ls-filter-action:disabled,
    .ls-filter-ok:disabled {
      opacity: 0.5;
      cursor: default;
    }

    .ls-filter-search {
      padding: 3px 6px;
      font: inherit;
      border: 1px solid var(--_border-color);
      border-radius: 3px;
    }

    .ls-filter-list {
      display: flex;
      flex-direction: column;
      min-height: 60px;
      max-height: 220px;
      overflow: auto;
      padding: 2px 0;
      border: 1px solid var(--_border-color);
      border-radius: 3px;
    }

    .ls-filter-option {
      display: flex;
      align-items: center;
      gap: 6px;
      padding: 2px 6px;
      cursor: pointer;
      white-space: nowrap;
    }

    .ls-filter-option > span {
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .ls-filter-note {
      font-size: 12px;
      color: var(--ls-find-status-color, #555);
    }

    .ls-filter-menu button:focus-visible,
    .ls-filter-menu input:focus-visible {
      outline: var(--_focus-ring);
      outline-offset: 1px;
    }

    /*
     * Conditional formats. Bars and icons are decorative (the cell's hidden
     * label says what they mean); forced-colors mode drops fills, so there
     * highlighted cells get an edge marker and scales and bars a meter line.
     */
    .ls-cf-bar {
      position: absolute;
      top: 3px;
      bottom: 3px;
      z-index: -1;
      border-radius: 1px;
      background-color: var(--ls-data-bar-color, #8fb8ea);
      pointer-events: none;
    }

    .ls-cf-bar.negative {
      background-color: var(--ls-data-bar-negative-color, #f0a1a1);
    }

    .ls-cf-icon {
      display: inline-flex;
      flex: 0 0 auto;
      margin-right: 4px;
    }

    .ls-cf-meter {
      display: none;
    }

    @media (forced-colors: active) {
      .ls-cell.cf-highlight::before {
        content: '';
        position: absolute;
        top: 0;
        bottom: 0;
        left: 0;
        width: 3px;
        background: Highlight;
        forced-color-adjust: none;
      }

      .ls-cf-bar,
      .ls-cf-meter {
        display: block;
        position: absolute;
        top: auto;
        bottom: 1px;
        height: 3px;
        z-index: 0;
        border-radius: 0;
        background: CanvasText;
        forced-color-adjust: none;
        pointer-events: none;
      }

      .ls-cf-meter {
        left: 0;
      }

      /* Negative bars are dashed, so the sign survives without color */
      .ls-cf-bar.negative {
        background: repeating-linear-gradient(90deg, CanvasText 0 2px, transparent 2px 4px);
      }
    }

    /* Visually hidden live region for screen reader announcements */
    .sr-only {
      position: absolute;
      width: 1px;
      height: 1px;
      padding: 0;
      margin: -1px;
      overflow: hidden;
      clip: rect(0, 0, 0, 0);
      border: 0;
    }
  `,
  ];

  protected render() {
    this._renderedCellSize = this._cellSize;
    const { rows, cols } = this._getVisibleRange();
    const activeCell = this._selection.activeCell;
    // Percent and date cells show as typed (15%, 2026-01-05), like the inline editor
    const activeRawValue = this._editTextFor(this._getCell(activeCell.row, activeCell.col));
    const activeDisplayValue = this._getCellDisplay(activeCell.row, activeCell.col);
    this._activeSpill = this._spillRangeAt(activeCell.row, activeCell.col);
    // A spilled value shows the formula it comes from, greyed out, as in Excel
    const spilledFrom = this._formulaBarMode === 'raw' ? this._getCell(activeCell.row, activeCell.col)?.spilledFrom : undefined;

    return html`
      <div class="ls-wrapper">
        <y11n-formula-bar
          cell-ref="${this._getActiveCellRef()}"
          raw-value="${activeRawValue}"
          display-value="${activeDisplayValue}"
          placeholder="${spilledFrom !== undefined ? this._internalData.get(spilledFrom)?.rawValue ?? '' : ''}"
          selection-label="${this._selectionName()}"
          .names="${this._nameList()}"
          @name-box-commit="${this._handleNameBoxCommit}"
          @name-box-cancel="${() => this._focusActiveCell()}"
          .mode="${this._formulaBarMode}"
          .readOnly="${this.readOnly || this._isCellLocked(this._getActiveCellKey())}"
          @formula-bar-mode-change="${this._handleFormulaBarModeChange}"
          @formula-bar-commit="${this._handleFormulaBarCommit}"
        ></y11n-formula-bar>

        <div class="ls-grid-shell">
          <div class="ls-size-probe" aria-hidden="true"></div>
          <input
            id="editor"
            part="editor"
            style=${styleMap(this._getEditorStyle())}
            .value="${this._editValue}"
            @input="${this._handleEditorInput}"
            @keydown="${this._handleEditorKeydown}"
            @blur="${this._handleEditorBlur}"
            aria-label="Cell editor"
            aria-invalid="${this._validationMessage ? 'true' : nothing}"
            aria-describedby="${this._validationMessage ? 'editor-error' : nothing}"
          />
          ${this._isEditing && this._validationMessage
            ? html`<div
                id="editor-error"
                class="ls-validation-message"
                part="validation-message"
                style=${styleMap(this._getValidationMessageStyle())}
              >${this._validationMessage}</div>`
            : nothing}
          ${this._findOpen ? this._renderFindBar() : nothing}
          ${this._filterMenu && this._filter ? this._renderFilterMenu(this._filterMenu, this._filter) : nothing}
          ${this._pasteSpecial ? this._renderPasteSpecial(this._pasteSpecial) : nothing}
          ${this._arrayCard ? this._renderArrayCard(this._arrayCard) : nothing}
          ${this._nameManagerOpen
            ? html`<y11n-name-manager
                .names=${this.getNames()}
                .sheets=${this.sheets}
                active-sheet="${this._activeState.name}"
                selection-ref="${this._selectionReference()}"
                .readOnly=${this.readOnly}
                .status=${this._nameManagerStatus}
                @name-save="${this._handleNameSave}"
                @name-remove="${this._handleNameRemove}"
                @name-go="${this._handleNameGo}"
                @name-close="${() => this.closeNameManager(true)}"
              ></y11n-name-manager>`
            : nothing}
          ${this._cfManagerOpen
            ? html`<y11n-conditional-format-dialog
                .rules=${this._cfRules}
                .selection=${this._selection.range}
                .rows=${this.rows}
                .cols=${this.cols}
                .readOnly=${this.readOnly}
                .status=${this._cfManagerStatus}
                @cf-save="${this._handleCfSave}"
                @cf-remove="${this._handleCfRemove}"
                @cf-move="${this._handleCfMove}"
                @cf-close="${() => this.closeConditionalFormatManager(true)}"
              ></y11n-conditional-format-dialog>`
            : nothing}

          <div
            class="ls-grid"
            role="grid"
            aria-label="${this.gridLabel}"
            aria-rowcount="${this.rows + 1}"
            aria-colcount="${this.cols + 1}"
            aria-readonly="${this.readOnly}"
            aria-multiselectable="true"
            style="--cols: ${this.cols}; grid-template-columns: ${this._gridTemplateColumns()}; ${this._scrollPadding()}"
            tabindex="-1"
            @keydown="${this._handleGridKeydown}"
            @scroll="${this._handleScroll}"
          >
            ${this._renderHeaderRow(cols)}
            ${this._renderRows(rows, cols)}
          </div>
        </div>

        ${this.sheetTabs || this._sheets.length > 1
          ? html`<y11n-sheet-tabs
              .sheets=${this._sheets.map((sheet) => ({ id: sheet.id, name: sheet.name }))}
              .activeId=${this._activeSheetId}
              .readOnly=${this.readOnly}
              @sheet-select="${this._handleSheetSelect}"
              @sheet-add="${this._handleSheetAdd}"
              @sheet-rename="${this._handleSheetRename}"
              @sheet-remove="${this._handleSheetRemove}"
              @sheet-move="${this._handleSheetMove}"
            ></y11n-sheet-tabs>`
          : nothing}
        ${this.hideStatusBar ? nothing : this._renderStatusBar()}

        <div class="sr-only" role="status" aria-live="polite" aria-atomic="true">
          ${this._getAnnouncement(activeCell)}
        </div>
        <div class="sr-only" role="log" aria-live="assertive" aria-atomic="true">
          ${this._modeAnnouncement}
        </div>
      </div>
    `;
  }

  private _renderFindBar() {
    return html`
      <div class="ls-find" role="search" aria-label="Find in spreadsheet">
        <input
          id="find-input"
          class="ls-find-input"
          type="text"
          aria-label="Find"
          aria-describedby="find-status"
          placeholder="Find"
          autocomplete="off"
          spellcheck="false"
          .value="${this._findQuery}"
          @input="${this._handleFindInput}"
          @keydown="${this._handleFindKeydown}"
        />
        <span id="find-status" class="ls-find-status" role="status">${this._findStatus}</span>
        <button
          type="button"
          class="ls-find-button"
          aria-label="Previous match"
          title="Previous match (Shift+Enter)"
          @click="${() => this._findStep(-1)}"
        ><span aria-hidden="true">&#x2191;</span></button>
        <button
          type="button"
          class="ls-find-button"
          aria-label="Next match"
          title="Next match (Enter)"
          @click="${() => this._findStep(1)}"
        ><span aria-hidden="true">&#x2193;</span></button>
        <button
          type="button"
          class="ls-find-button"
          aria-label="Close find"
          title="Close (Escape)"
          @click="${() => this.closeFind()}"
        ><span aria-hidden="true">&#x2715;</span></button>
      </div>
    `;
  }

  /** Place the filter menu under its header cell, inside the grid area. */
  private _getFilterMenuStyle(col: number, headerRow: number): Record<string, string> {
    const h = this._cellSize.height;
    const width = 260;
    const gridEl = this._grid;
    const shellWidth = gridEl?.clientWidth ?? 0;
    const shellHeight = gridEl?.clientHeight ?? 0;
    // The bottom of the header cell; frozen cells don't move with the scroll position
    let top = h + this._rowTop(headerRow) + h - (headerRow < this._frozenRowCount ? 0 : this._scrollTop);
    let left =
      Y11nSpreadsheet.ROW_HEADER_WIDTH + this._colLeft(col) - (col < this._frozenColCount ? 0 : this._scrollLeft);
    if (shellWidth) left = Math.max(0, Math.min(left, shellWidth - width - 4));
    const minHeight = 200;
    if (shellHeight && top + minHeight > shellHeight) top = Math.max(0, shellHeight - minHeight - 4);
    const styles: Record<string, string> = { top: `${Math.max(0, top)}px`, left: `${Math.max(0, left)}px`, width: `${width}px` };
    if (shellHeight) styles['max-height'] = `${Math.max(minHeight, shellHeight - top - 4)}px`;
    return styles;
  }

  private _renderFilterMenu(menu: FilterMenuState, filter: FilterState) {
    const col = menu.col;
    const letter = colToLetter(col);
    const header = this._getCellDisplay(filter.range.start.row, col);
    const name = header ? `column ${letter}, ${header}` : `column ${letter}`;
    const matches = this._filterMenuMatches(menu);
    const shown = matches.slice(0, Y11nSpreadsheet.FILTER_MENU_MAX_VALUES);
    const checkedCount = matches.reduce((n, v) => n + (menu.checked.has(v) ? 1 : 0), 0);
    const allChecked = matches.length > 0 && checkedCount === matches.length;
    const searching = menu.search.trim() !== '';
    const item = (index: number) => ({
      'data-filter-item': String(index),
      tabindex: menu.focus === index ? 0 : -1,
    });
    return html`
      <div
        class="ls-filter-menu"
        role="dialog"
        aria-modal="true"
        aria-label="Filter ${name}"
        style=${styleMap(this._getFilterMenuStyle(col, filter.range.start.row))}
        @keydown="${this._handleFilterMenuKeydown}"
      >
        <div class="ls-filter-row">
          <button
            type="button"
            class="ls-filter-action"
            ?disabled=${this.readOnly}
            @click="${() => this._sortFromFilterMenu('ascending')}"
          >Sort A to Z</button>
          <button
            type="button"
            class="ls-filter-action"
            ?disabled=${this.readOnly}
            @click="${() => this._sortFromFilterMenu('descending')}"
          >Sort Z to A</button>
        </div>
        <button
          type="button"
          class="ls-filter-action"
          ?disabled=${!filter.criteria.has(col)}
          @click="${this._clearColumnFromFilterMenu}"
        >Clear filter from column ${letter}</button>
        <input
          id="filter-search"
          class="ls-filter-search"
          type="search"
          aria-label="Search values"
          aria-describedby="filter-status"
          placeholder="Search"
          autocomplete="off"
          spellcheck="false"
          .value="${menu.search}"
          @input="${this._handleFilterSearchInput}"
        />
        <div class="ls-filter-list" role="group" aria-label="Values to show">
          <label class="ls-filter-option">
            <input
              type="checkbox"
              data-filter-item="0"
              tabindex="${item(0).tabindex}"
              .checked=${allChecked}
              .indeterminate=${checkedCount > 0 && !allChecked}
              ?disabled=${matches.length === 0}
              @change="${this._toggleAllFilterValues}"
            />
            <span>${searching ? '(Select all search results)' : '(Select all)'}</span>
          </label>
          ${repeat(
            shown,
            (value) => value,
            (value, i) => html`
              <label class="ls-filter-option">
                <input
                  type="checkbox"
                  data-filter-item="${i + 1}"
                  tabindex="${item(i + 1).tabindex}"
                  .checked=${menu.checked.has(value)}
                  @change="${() => this._toggleFilterValue(value)}"
                />
                <span>${this._filterValueLabel(value)}</span>
              </label>
            `
          )}
        </div>
        <div id="filter-status" class="ls-filter-note" role="status">${
          matches.length === 0
            ? 'No matching values'
            : matches.length > shown.length
              ? `Showing ${shown.length} of ${matches.length} values. Search to narrow the list.`
              : `${checkedCount} of ${matches.length} values selected`
        }</div>
        <div class="ls-filter-row ls-filter-buttons">
          <button
            type="button"
            class="ls-filter-ok"
            ?disabled=${checkedCount === 0}
            @click="${this._applyFilterMenu}"
          >OK</button>
          <button type="button" class="ls-filter-action" @click="${() => this._closeFilterMenu(true)}">Cancel</button>
        </div>
      </div>
    `;
  }

  private _renderColumnSpacer(size: number) {
    return html`<div style="grid-column: span ${size};"></div>`;
  }

  private _renderHeaderRow(cols: number[]) {
    const range = this._selection.range;
    const fullColumns = range.start.row === 0 && range.end.row === this.rows - 1;
    return html`
      <div class="ls-header-row" role="row" aria-rowindex="1">
        <div
          class="ls-corner-header"
          role="columnheader"
          aria-colindex="1"
          aria-label="Select all"
          tabindex="-1"
          @click="${this._handleSelectAll}"
          @keydown="${this._handleCornerKeydown}"
        ><span class="sr-only">Select all</span></div>
        ${layoutWithGaps(
          cols,
          this.cols,
          (c) => html`
            <div
              class="ls-col-header ${c >= range.start.col && c <= range.end.col ? 'in-selection' : ''} ${this._frozenClasses(-1, c)}"
              style=${c < this._frozenColCount ? `left: ${this._stickyLeft(c)}` : nothing}
              role="columnheader"
              aria-colindex="${c + 2}"
              aria-label="Column ${colToLetter(c)}"
              aria-selected="${fullColumns && c >= range.start.col && c <= range.end.col ? 'true' : nothing}"
              data-header-col="${c}"
              @pointerdown="${this._handleHeaderPointerDown}"
            >
              ${colToLetter(c)}
              <div
                class="ls-col-resizer"
                aria-hidden="true"
                title="Drag to resize, double-click to fit (Alt+Shift+Left/Right, Alt+Shift+0)"
                data-resize-col="${c}"
                @pointerdown="${this._handleResizerPointerDown}"
                @dblclick="${this._handleResizerDblClick}"
              ></div>
            </div>
          `,
          (_start, size) => this._renderColumnSpacer(size)
        )}
      </div>
    `;
  }

  private _renderRows(rows: number[], cols: number[]) {
    type Segment = { key: string; row: number; gap: number };
    const segments = layoutWithGaps<Segment>(
      rows,
      this.rows,
      (row) => ({ key: `r${row}`, row, gap: 0 }),
      // A gap is as tall as the visible rows in it: hidden rows take no space
      (start, size) => ({ key: `g${start}`, row: start, gap: this._rowTop(start + size) - this._rowTop(start) })
    ).filter((seg) => seg.key.startsWith('r') || seg.gap > 0);

    return html`
      ${repeat(
        segments,
        (seg) => seg.key,
        (seg) =>
          seg.gap > 0
            ? html`<div style="grid-column: 1 / -1; height: ${seg.gap}px;"></div>`
            : this._renderRow(seg.row, cols)
      )}
    `;
  }

  private _renderCell(row: number, c: number) {
    const isSelected = this._selection.isCellSelected(row, c);
    const isActive = this._selection.isCellActive(row, c);
    const isRefTarget = this._refMode && row === this._refCursorRow && c === this._refCursorCol;
    const display = this._getCellDisplay(row, c);
    const key = cellKey(row, c);
    const locked = this._isCellLocked(key);
    const range = this._selection.range;
    const showFillHandle =
      !this.readOnly && !this._isEditing && row === range.end.row && c === range.end.col;
    const preview = this._fillPreview;
    const inFillPreview =
      preview !== null &&
      !isSelected &&
      row >= preview.start.row &&
      row <= preview.end.row &&
      c >= preview.start.col &&
      c <= preview.end.col;

    const cf = this._cf.resultFor(row, c);
    const cellStyles = this._getCellStyles(row, c, cf?.format);
    if (row < this._frozenRowCount) cellStyles.top = this._stickyTop(row);
    if (c < this._frozenColCount) cellStyles.left = this._stickyLeft(c);
    const filterHeader = this._isFilterHeader(row, c);
    const filtered = filterHeader && this._filter!.criteria.has(c);
    const array = this._internalData.get(key)?.array;
    const hasPopup = filterHeader || !!array;
    const expanded = filterHeader ? this._filterMenu?.col === c : this._arrayCard?.key === key;

    return html`
      <div
        class="ls-cell ${isActive ? 'active-cell' : ''} ${isRefTarget ? 'ref-highlight' : ''} ${locked ? 'read-only-cell' : ''} ${inFillPreview ? 'fill-preview' : ''} ${filterHeader ? 'filter-header' : ''} ${cf ? this._cfClasses(cf) : ''} ${cellStyles['background-color'] ? 'has-fill' : ''} ${array ? 'array-cell' : ''} ${this._spillClasses(row, c)} ${this._frozenClasses(row, c)}"
        role="gridcell"
        aria-colindex="${c + 2}"
        aria-selected="${isSelected}"
        aria-readonly="${this.readOnly || locked}"
        aria-current="${isActive ? 'true' : nothing}"
        aria-haspopup="${hasPopup ? 'dialog' : nothing}"
        aria-expanded="${hasPopup ? String(expanded) : nothing}"
        data-row="${row}"
        data-col="${c}"
        data-key="${key}"
        tabindex="${isActive ? 0 : -1}"
        style=${styleMap(cellStyles)}
        @pointerdown="${this._handleCellPointerDown}"
        @dblclick="${this._handleCellDblClick}"
      >
        ${cf ? this._renderCfGraphics(cf) : nothing}
        ${array
          ? html`<span
              class="ls-array-button"
              aria-hidden="true"
              title="Show values (Alt+Down)"
              @pointerdown="${this._handleArrayButtonPointerDown}"
              @dblclick="${(e: Event) => e.stopPropagation()}"
            >${array.cols === 1 || array.rows === 1 ? LIST_ICON : ARRAY_ICON}</span>`
          : nothing}
        <span class="cell-text ${cf && this._cfHidesValue(cf) ? 'sr-only' : ''}">${display}</span>
        ${array ? html`<span class="sr-only ls-array-label">, ${describeArray(array)}</span>` : nothing}
        ${cf && cf.labels.length > 0 ? html`<span class="sr-only ls-cf-label">, ${cf.labels.join(', ')}</span>` : nothing}
        ${filterHeader
          ? html`<span
              class="ls-filter-button ${filtered ? 'filtered' : ''}"
              aria-hidden="true"
              title="${filtered ? 'Filtered. ' : ''}Filter options (Alt+Down)"
              data-filter-col="${c}"
              @pointerdown="${this._handleFilterButtonPointerDown}"
              @dblclick="${(e: Event) => e.stopPropagation()}"
            >${filtered ? FUNNEL_ICON : CHEVRON_ICON}</span>`
          : nothing}
        ${showFillHandle
          ? html`<div
              class="ls-fill-handle"
              aria-hidden="true"
              @pointerdown="${this._handleFillHandlePointerDown}"
              @dblclick="${this._handleFillHandleDblClick}"
            ></div>`
          : nothing}
      </div>
    `;
  }

  private _renderRow(row: number, cols: number[]) {
    const range = this._selection.range;
    const inSelection = row >= range.start.row && row <= range.end.row;
    const fullRow = inSelection && range.start.col === 0 && range.end.col === this.cols - 1;
    // Row numbers of a filtered table are coloured, as in Excel, since they skip hidden rows
    const filter = this._filter;
    const filteredRow =
      filter !== null && filter.criteria.size > 0 && row > filter.range.start.row && row <= filter.range.end.row;
    return html`
      <div class="ls-row" role="row" aria-rowindex="${row + 2}">
        <div
          class="ls-row-header ${inSelection ? 'in-selection' : ''} ${filteredRow ? 'filtered-row' : ''} ${this._frozenClasses(row, -1)}"
          style=${row < this._frozenRowCount ? `top: ${this._stickyTop(row)}` : nothing}
          role="rowheader"
          aria-colindex="1"
          aria-label="Row ${row + 1}"
          aria-selected="${fullRow ? 'true' : nothing}"
          data-header-row="${row}"
          @pointerdown="${this._handleHeaderPointerDown}"
        >${row + 1}</div>
        ${repeat(
          layoutWithGaps<{ key: string; tpl: unknown }>(
            cols,
            this.cols,
            (c) => ({ key: `c${c}`, tpl: this._renderCell(row, c) }),
            (start, size) => ({ key: `g${start}`, tpl: this._renderColumnSpacer(size) })
          ),
          // Keyed so the focused cell's element is never reused for another column
          (seg) => seg.key,
          (seg) => seg.tpl
        )}
      </div>
    `;
  }

  private _getAnnouncement(activeCell: CellCoord): string {
    const ref = `${colToLetter(activeCell.col)}${activeCell.row + 1}`;
    const display = this._getCellDisplay(activeCell.row, activeCell.col);
    let base = display ? `${ref}: ${display}` : ref;
    const cell = this._getCell(activeCell.row, activeCell.col);
    if (cell?.array) {
      // On a filter header, Alt+Down belongs to the filter menu
      const keys = this._isFilterHeader(activeCell.row, activeCell.col) ? 'Alt+Shift+V' : 'Alt+Down';
      base += `, ${describeArray(cell.array)}. ${keys} to show the values`;
    }
    if (cell?.spilledFrom !== undefined) {
      base += `, spilled from ${coordToRef(parseKey(cell.spilledFrom))}`;
    } else if (cell?.spill) {
      const end = { row: activeCell.row + cell.spill.rows - 1, col: activeCell.col + cell.spill.cols - 1 };
      base += `, spills to ${coordToRef(end)}`;
    } else if (cell?.displayValue === '#SPILL!' && cell.rawValue.startsWith('=')) {
      const needs = this._formulaEngine.blockedSpill(cellKey(activeCell.row, activeCell.col));
      if (needs) {
        base +=
          activeCell.row + needs.rows > this.rows || activeCell.col + needs.cols > this.cols
            ? ', the result does not fit on the sheet'
            : ', the cells the result would fill are not empty';
      }
    }
    const cf = this._cf.resultFor(activeCell.row, activeCell.col);
    if (cf && cf.labels.length > 0) base += `, ${cf.labels.join(', ')}`;
    if (this._isFilterHeader(activeCell.row, activeCell.col)) {
      base += this._filter!.criteria.has(activeCell.col)
        ? ', filtered. Alt+Down for filter options'
        : ', Alt+Down for filter options';
    }

    const range = this._selection.range;
    const selRows = range.end.row - range.start.row + 1;
    const selCols = range.end.col - range.start.col + 1;
    if (selRows > 1 || selCols > 1) {
      const allRows = selRows === this.rows;
      const allCols = selCols === this.cols;
      if (allRows && !allCols) {
        const first = colToLetter(range.start.col);
        const cols = selCols > 1 ? `Columns ${first} to ${colToLetter(range.end.col)}` : `Column ${first}`;
        return `${base}, ${cols} selected`;
      }
      if (allCols && !allRows) {
        const rows = selRows > 1 ? `Rows ${range.start.row + 1} to ${range.end.row + 1}` : `Row ${range.start.row + 1}`;
        return `${base}, ${rows} selected`;
      }
      return `${base}, ${selRows} row${selRows > 1 ? 's' : ''}, ${selCols} column${selCols > 1 ? 's' : ''} selected`;
    }

    return base;
  }

  protected updated(): void {
    this._reportHistory();
    if (this._isEditing && this._editor) {
      this._editor.focus();
      if (this._refMode) {
        // Place cursor after the inserted reference
        this._editor.setSelectionRange(this._refInsertEnd, this._refInsertEnd);
      } else if (this._editorNeedsCursorReset) {
        // Place cursor at the end only when editing first starts
        this._editorNeedsCursorReset = false;
        this._editor.setSelectionRange(this._editValue.length, this._editValue.length);
      }
    }
  }
}

declare global {
  interface HTMLElementTagNameMap {
    'y11n-spreadsheet': Y11nSpreadsheet;
  }
}
