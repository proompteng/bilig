import {
  buildCellNumberFormatCode,
  getCellNumberFormatKind,
  type CellHorizontalAlignment,
  type CellVerticalAlignment,
  type CellBorderSideSnapshot,
  type CellNumberFormatRecord,
  type CellStyleRecord,
  MAX_COLS,
  MAX_ROWS,
  type CellRangeRef,
  type LiteralInput,
  type SheetFormatRangeSnapshot,
  type SheetStyleRangeSnapshot,
  type WorkbookAxisEntrySnapshot,
  type WorkbookCalculationSettingsSnapshot,
  type WorkbookDefinedNameValueSnapshot,
  type WorkbookPivotSnapshot,
  type WorkbookPivotValueSnapshot,
  type WorkbookTableSnapshot,
  type WorkbookVolatileContextSnapshot,
} from "@bilig/protocol";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import { SheetGrid } from "./sheet-grid.js";
import { CellStore } from "./cell-store.js";
import {
  canonicalWorkbookAddress,
  canonicalWorkbookRangeRef,
  cloneWorkbookRangeRecords,
  findWorkbookRangeRecord,
  overlayWorkbookRangeRecords,
  replaceWorkbookRangeRecords,
} from "./workbook-range-records.js";

const SHEET_STRIDE = MAX_ROWS * MAX_COLS;

export interface WorkbookDefinedNameRecord {
  name: string;
  value: WorkbookDefinedNameValueSnapshot;
}

export interface WorkbookPropertyRecord {
  key: string;
  value: LiteralInput;
}

export interface WorkbookSpillRecord {
  sheetName: string;
  address: string;
  rows: number;
  cols: number;
}

export interface WorkbookPivotRecord extends WorkbookPivotSnapshot {
  values: WorkbookPivotValueSnapshot[];
}

export interface WorkbookTableRecord extends WorkbookTableSnapshot {}

export interface WorkbookAxisMetadataRecord {
  sheetName: string;
  start: number;
  count: number;
  size: number | null;
  hidden: boolean | null;
}

export interface WorkbookAxisEntryRecord {
  id: string;
  size: number | null;
  hidden: boolean | null;
}

export interface WorkbookCellStyleRecord extends CellStyleRecord {}

export interface WorkbookStyleRangeRecord {
  range: CellRangeRef;
  styleId: string;
}

export interface WorkbookCellNumberFormatRecord extends CellNumberFormatRecord {}

export interface WorkbookFormatRangeRecord {
  range: CellRangeRef;
  formatId: string;
}

export interface WorkbookCalculationSettingsRecord extends WorkbookCalculationSettingsSnapshot {}

export interface WorkbookVolatileContextRecord extends WorkbookVolatileContextSnapshot {}

export interface WorkbookFreezePaneRecord {
  sheetName: string;
  rows: number;
  cols: number;
}

export interface WorkbookFilterRecord {
  sheetName: string;
  range: CellRangeRef;
}

export interface WorkbookSortKeyRecord {
  keyAddress: string;
  direction: "asc" | "desc";
}

export interface WorkbookSortRecord {
  sheetName: string;
  range: CellRangeRef;
  keys: WorkbookSortKeyRecord[];
}

export interface WorkbookMetadataRecord {
  properties: Map<string, WorkbookPropertyRecord>;
  definedNames: Map<string, WorkbookDefinedNameRecord>;
  tables: Map<string, WorkbookTableRecord>;
  spills: Map<string, WorkbookSpillRecord>;
  pivots: Map<string, WorkbookPivotRecord>;
  rowMetadata: Map<string, WorkbookAxisMetadataRecord>;
  columnMetadata: Map<string, WorkbookAxisMetadataRecord>;
  calculationSettings: WorkbookCalculationSettingsRecord;
  volatileContext: WorkbookVolatileContextRecord;
  freezePanes: Map<string, WorkbookFreezePaneRecord>;
  filters: Map<string, WorkbookFilterRecord>;
  sorts: Map<string, WorkbookSortRecord>;
}

export function normalizeDefinedName(name: string): string {
  return normalizeWorkbookObjectName(name, "Defined names");
}

export function normalizeWorkbookObjectName(name: string, label = "Workbook object"): string {
  const normalized = name.trim().toUpperCase();
  if (normalized.length === 0) {
    throw new Error(`${label} must be non-empty`);
  }
  return normalized;
}

export interface SheetRecord {
  id: number;
  name: string;
  order: number;
  grid: SheetGrid;
  rowAxis: Array<WorkbookAxisEntryRecord | undefined>;
  columnAxis: Array<WorkbookAxisEntryRecord | undefined>;
  styleRanges: WorkbookStyleRangeRecord[];
  formatRanges: WorkbookFormatRangeRecord[];
}

export interface EnsuredCell {
  cellIndex: number;
  created: boolean;
}

export class WorkbookStore {
  static readonly defaultStyleId = "style-0";
  static readonly defaultFormatId = "format-0";
  readonly cellStore = new CellStore();
  readonly sheetsByName = new Map<string, SheetRecord>();
  readonly sheetsById = new Map<number, SheetRecord>();
  readonly cellKeyToIndex = new Map<number, number>();
  readonly cellFormats = new Map<number, string>();
  readonly cellStyles = new Map<string, WorkbookCellStyleRecord>();
  readonly styleKeys = new Map<string, string>();
  readonly cellNumberFormats = new Map<string, WorkbookCellNumberFormatRecord>();
  readonly numberFormatKeys = new Map<string, string>();
  readonly metadata: WorkbookMetadataRecord = {
    properties: new Map(),
    definedNames: new Map(),
    tables: new Map(),
    spills: new Map(),
    pivots: new Map(),
    rowMetadata: new Map(),
    columnMetadata: new Map(),
    calculationSettings: { mode: "automatic", compatibilityMode: "excel-modern" },
    volatileContext: { recalcEpoch: 0 },
    freezePanes: new Map(),
    filters: new Map(),
    sorts: new Map(),
  };
  workbookName: string;
  private nextSheetId = 1;
  private nextRowAxisId = 1;
  private nextColumnAxisId = 1;
  private nextStyleId = 1;
  private nextFormatId = 1;

  constructor(workbookName = "Workbook") {
    this.workbookName = workbookName;
    this.ensureDefaultStyle();
    this.ensureDefaultNumberFormat();
  }

  createSheet(name: string, order = this.sheetsByName.size, id?: number): SheetRecord {
    const existing = this.sheetsByName.get(name);
    if (existing) {
      existing.order = order;
      if (id !== undefined && existing.id !== id) {
        this.sheetsById.delete(existing.id);
        existing.id = id;
        this.sheetsById.set(existing.id, existing);
        this.bumpSheetId(id);
      }
      return existing;
    }
    const sheet: SheetRecord = {
      id: id ?? this.nextSheetId++,
      name,
      order,
      grid: new SheetGrid(),
      rowAxis: [],
      columnAxis: [],
      styleRanges: [],
      formatRanges: [],
    };
    if (id !== undefined) {
      this.bumpSheetId(id);
    }
    this.sheetsByName.set(name, sheet);
    this.sheetsById.set(sheet.id, sheet);
    return sheet;
  }

  deleteSheet(name: string): void {
    const sheet = this.sheetsByName.get(name);
    if (!sheet) return;
    sheet.grid.forEachCell((cellIndex) => {
      const key = makeCellKey(
        sheet.id,
        this.cellStore.rows[cellIndex]!,
        this.cellStore.cols[cellIndex]!,
      );
      this.cellKeyToIndex.delete(key);
      this.cellFormats.delete(cellIndex);
    });
    this.deleteSheetMetadata(name);
    this.sheetsByName.delete(name);
    this.sheetsById.delete(sheet.id);
  }

  renameSheet(oldName: string, nextName: string): SheetRecord | undefined {
    const trimmedName = nextName.trim();
    if (trimmedName.length === 0) {
      throw new Error("Sheet name must be non-empty");
    }
    const sheet = this.sheetsByName.get(oldName);
    if (!sheet) {
      return undefined;
    }
    if (oldName === trimmedName) {
      return sheet;
    }
    if (this.sheetsByName.has(trimmedName)) {
      return undefined;
    }

    this.sheetsByName.delete(oldName);
    sheet.name = trimmedName;
    this.sheetsByName.set(trimmedName, sheet);

    rekeyRecords(this.metadata.freezePanes, (record) =>
      record.sheetName === oldName ? { ...record, sheetName: trimmedName } : record,
    );
    rekeyRecords(this.metadata.rowMetadata, (record) =>
      record.sheetName === oldName ? { ...record, sheetName: trimmedName } : record,
    );
    rekeyRecords(this.metadata.columnMetadata, (record) =>
      record.sheetName === oldName ? { ...record, sheetName: trimmedName } : record,
    );
    rekeyRecords(this.metadata.filters, (record) =>
      record.sheetName === oldName || record.range.sheetName === oldName
        ? {
            sheetName: record.sheetName === oldName ? trimmedName : record.sheetName,
            range: {
              ...record.range,
              sheetName: record.range.sheetName === oldName ? trimmedName : record.range.sheetName,
            },
          }
        : record,
    );
    rekeyRecords(this.metadata.sorts, (record) =>
      record.sheetName === oldName || record.range.sheetName === oldName
        ? {
            sheetName: record.sheetName === oldName ? trimmedName : record.sheetName,
            range: {
              ...record.range,
              sheetName: record.range.sheetName === oldName ? trimmedName : record.range.sheetName,
            },
            keys: record.keys.map((key) => ({ ...key })),
          }
        : record,
    );
    rekeyRecords(this.metadata.tables, (record) =>
      record.sheetName === oldName ? { ...record, sheetName: trimmedName } : record,
    );
    rekeyRecords(this.metadata.spills, (record) =>
      record.sheetName === oldName ? { ...record, sheetName: trimmedName } : record,
    );
    rekeyRecords(this.metadata.pivots, (record) =>
      record.sheetName === oldName || record.source.sheetName === oldName
        ? {
            ...record,
            sheetName: record.sheetName === oldName ? trimmedName : record.sheetName,
            source: {
              ...record.source,
              sheetName:
                record.source.sheetName === oldName ? trimmedName : record.source.sheetName,
            },
            groupBy: [...record.groupBy],
            values: record.values.map((value) => ({ ...value })),
          }
        : record,
    );

    sheet.styleRanges = sheet.styleRanges.map((record) =>
      record.range.sheetName === oldName
        ? { ...record, range: { ...record.range, sheetName: trimmedName } }
        : record,
    );
    sheet.formatRanges = sheet.formatRanges.map((record) =>
      record.range.sheetName === oldName
        ? { ...record, range: { ...record.range, sheetName: trimmedName } }
        : record,
    );

    return sheet;
  }

  getSheet(name: string): SheetRecord | undefined {
    return this.sheetsByName.get(name);
  }

  getSheetById(id: number): SheetRecord | undefined {
    return this.sheetsById.get(id);
  }

  getOrCreateSheet(name: string): SheetRecord {
    return this.getSheet(name) ?? this.createSheet(name);
  }

  ensureCell(sheetName: string, address: string): number {
    return this.ensureCellRecord(sheetName, address).cellIndex;
  }

  ensureCellRecord(sheetName: string, address: string): EnsuredCell {
    const sheet = this.getOrCreateSheet(sheetName);
    const parsed = parseCellAddress(address, sheetName);
    return this.ensureCellAt(sheet.id, parsed.row, parsed.col);
  }

  ensureCellAt(sheetId: number, row: number, col: number): EnsuredCell {
    const sheet = this.getSheetById(sheetId);
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${sheetId}`);
    }
    const key = makeCellKey(sheet.id, row, col);
    const existing = this.cellKeyToIndex.get(key);
    if (existing !== undefined) {
      return { cellIndex: existing, created: false };
    }
    const cellIndex = this.cellStore.allocate(sheet.id, row, col);
    this.cellKeyToIndex.set(key, cellIndex);
    sheet.grid.set(row, col, cellIndex);
    return { cellIndex, created: true };
  }

  getCellIndex(sheetName: string, address: string): number | undefined {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return undefined;
    const parsed = parseCellAddress(address, sheetName);
    return this.cellKeyToIndex.get(makeCellKey(sheet.id, parsed.row, parsed.col));
  }

  getSheetNameById(id: number): string {
    return this.sheetsById.get(id)?.name ?? "";
  }

  getAddress(index: number): string {
    return formatAddress(this.cellStore.rows[index]!, this.cellStore.cols[index]!);
  }

  getQualifiedAddress(index: number): string {
    return `${this.getSheetNameById(this.cellStore.sheetIds[index]!)}!${this.getAddress(index)}`;
  }

  setCellFormat(index: number, format: string | null | undefined): void {
    if (format === undefined || format === null || format === "") {
      this.cellFormats.delete(index);
      return;
    }
    this.internCellNumberFormat(format);
    this.cellFormats.set(index, format);
  }

  getCellFormat(index: number): string | undefined {
    return this.cellFormats.get(index);
  }

  upsertCellStyle(style: CellStyleRecord): WorkbookCellStyleRecord {
    const normalized = normalizeCellStyleRecord(style);
    const existing = this.cellStyles.get(normalized.id);
    if (existing) {
      this.styleKeys.delete(cellStyleKey(existing));
    }
    this.cellStyles.set(normalized.id, normalized);
    this.styleKeys.set(cellStyleKey(normalized), normalized.id);
    this.bumpStyleId(normalized.id);
    return normalized;
  }

  internCellStyle(style: Omit<WorkbookCellStyleRecord, "id">): WorkbookCellStyleRecord {
    const normalized = normalizeCellStyleRecord({
      id: WorkbookStore.defaultStyleId,
      ...style,
    });
    const key = cellStyleKey(normalized);
    const existingId = this.styleKeys.get(key);
    if (existingId) {
      return this.cellStyles.get(existingId)!;
    }
    return { ...normalized, id: cellStyleIdForKey(key) };
  }

  getCellStyle(id: string | undefined): WorkbookCellStyleRecord | undefined {
    if (!id) {
      return this.cellStyles.get(WorkbookStore.defaultStyleId);
    }
    return this.cellStyles.get(id) ?? this.cellStyles.get(WorkbookStore.defaultStyleId);
  }

  listCellStyles(): WorkbookCellStyleRecord[] {
    return [...this.cellStyles.values()].toSorted((left, right) => left.id.localeCompare(right.id));
  }

  upsertCellNumberFormat(format: CellNumberFormatRecord): WorkbookCellNumberFormatRecord {
    const normalized = normalizeCellNumberFormatRecord(format);
    const existing = this.cellNumberFormats.get(normalized.id);
    if (existing) {
      this.numberFormatKeys.delete(existing.code);
    }
    this.cellNumberFormats.set(normalized.id, normalized);
    this.numberFormatKeys.set(normalized.code, normalized.id);
    this.bumpFormatId(normalized.id);
    return normalized;
  }

  internCellNumberFormat(format: string | CellNumberFormatRecord): WorkbookCellNumberFormatRecord {
    const normalized =
      typeof format === "string"
        ? normalizeCellNumberFormatRecord({
            id: WorkbookStore.defaultFormatId,
            code: format,
            kind: getCellNumberFormatKind(format),
          })
        : normalizeCellNumberFormatRecord(format);
    const existingId = this.numberFormatKeys.get(normalized.code);
    if (existingId) {
      return this.cellNumberFormats.get(existingId)!;
    }
    return { ...normalized, id: cellNumberFormatIdForCode(normalized.code) };
  }

  getCellNumberFormat(id: string | undefined): WorkbookCellNumberFormatRecord | undefined {
    if (!id) {
      return this.cellNumberFormats.get(WorkbookStore.defaultFormatId);
    }
    return (
      this.cellNumberFormats.get(id) ?? this.cellNumberFormats.get(WorkbookStore.defaultFormatId)
    );
  }

  listCellNumberFormats(): WorkbookCellNumberFormatRecord[] {
    return [...this.cellNumberFormats.values()].toSorted((left, right) =>
      left.id.localeCompare(right.id),
    );
  }

  setStyleRange(range: CellRangeRef, styleId: string): WorkbookStyleRangeRecord {
    if (!this.cellStyles.has(styleId)) {
      throw new Error(`Unknown cell style: ${styleId}`);
    }
    const sheet = this.getOrCreateSheet(range.sheetName);
    const stored: WorkbookStyleRangeRecord = {
      range: { ...range },
      styleId,
    };
    sheet.styleRanges = overlayWorkbookRangeRecords(
      sheet.styleRanges,
      stored,
      (nextRange, record) => ({
        range: nextRange,
        styleId: record.styleId,
      }),
      (record) => record.styleId === WorkbookStore.defaultStyleId,
    );
    return stored;
  }

  listStyleRanges(sheetName: string): WorkbookStyleRangeRecord[] {
    return cloneWorkbookRangeRecords(this.getSheet(sheetName)?.styleRanges ?? [], (range, record) => ({
      range,
      styleId: record.styleId,
    }));
  }

  setStyleRanges(
    sheetName: string,
    ranges: readonly SheetStyleRangeSnapshot[],
  ): WorkbookStyleRangeRecord[] {
    const sheet = this.getOrCreateSheet(sheetName);
    const nextRanges = replaceWorkbookRangeRecords(
      ranges.map((entry) => ({
        range: { ...entry.range },
        styleId: entry.styleId,
      })),
      (range, record) => ({
        range,
        styleId: record.styleId,
      }),
      (entry) => {
        if (!this.cellStyles.has(entry.styleId)) {
          throw new Error(`Unknown cell style: ${entry.styleId}`);
        }
      },
    );
    sheet.styleRanges = nextRanges;
    return this.listStyleRanges(sheetName);
  }

  getStyleId(sheetName: string, row: number, col: number): string {
    const sheet = this.getSheet(sheetName);
    if (!sheet) {
      return WorkbookStore.defaultStyleId;
    }
    return (
      findWorkbookRangeRecord(sheet.styleRanges, row, col)?.styleId ??
      WorkbookStore.defaultStyleId
    );
  }

  setFormatRange(range: CellRangeRef, formatId: string): WorkbookFormatRangeRecord {
    if (!this.cellNumberFormats.has(formatId)) {
      throw new Error(`Unknown cell number format: ${formatId}`);
    }
    const sheet = this.getOrCreateSheet(range.sheetName);
    const stored: WorkbookFormatRangeRecord = {
      range: { ...range },
      formatId,
    };
    sheet.formatRanges = overlayWorkbookRangeRecords(
      sheet.formatRanges,
      stored,
      (nextRange, record) => ({
        range: nextRange,
        formatId: record.formatId,
      }),
      (record) => record.formatId === WorkbookStore.defaultFormatId,
    );
    return stored;
  }

  listFormatRanges(sheetName: string): WorkbookFormatRangeRecord[] {
    return cloneWorkbookRangeRecords(
      this.getSheet(sheetName)?.formatRanges ?? [],
      (range, record) => ({
        range,
        formatId: record.formatId,
      }),
    );
  }

  setFormatRanges(
    sheetName: string,
    ranges: readonly SheetFormatRangeSnapshot[],
  ): WorkbookFormatRangeRecord[] {
    const sheet = this.getOrCreateSheet(sheetName);
    const nextRanges = replaceWorkbookRangeRecords(
      ranges.map((entry) => ({
        range: { ...entry.range },
        formatId: entry.formatId,
      })),
      (range, record) => ({
        range,
        formatId: record.formatId,
      }),
      (entry) => {
        if (!this.cellNumberFormats.has(entry.formatId)) {
          throw new Error(`Unknown cell number format: ${entry.formatId}`);
        }
      },
    );
    sheet.formatRanges = nextRanges;
    return this.listFormatRanges(sheetName);
  }

  getRangeFormatId(sheetName: string, row: number, col: number): string {
    const sheet = this.getSheet(sheetName);
    if (!sheet) {
      return WorkbookStore.defaultFormatId;
    }
    return (
      findWorkbookRangeRecord(sheet.formatRanges, row, col)?.formatId ??
      WorkbookStore.defaultFormatId
    );
  }

  setWorkbookProperty(key: string, value: LiteralInput): WorkbookPropertyRecord | undefined {
    const trimmedKey = normalizeMetadataKey(key);
    if (value === null) {
      this.metadata.properties.delete(trimmedKey);
      return undefined;
    }
    const record: WorkbookPropertyRecord = { key: trimmedKey, value };
    this.metadata.properties.set(trimmedKey, record);
    return record;
  }

  getWorkbookProperty(key: string): WorkbookPropertyRecord | undefined {
    return this.metadata.properties.get(normalizeMetadataKey(key));
  }

  listWorkbookProperties(): WorkbookPropertyRecord[] {
    return [...this.metadata.properties.values()].toSorted((left, right) =>
      left.key.localeCompare(right.key),
    );
  }

  setCalculationSettings(
    settings: WorkbookCalculationSettingsSnapshot,
  ): WorkbookCalculationSettingsRecord {
    this.metadata.calculationSettings = {
      compatibilityMode: "excel-modern",
      ...settings,
    };
    return this.metadata.calculationSettings;
  }

  getCalculationSettings(): WorkbookCalculationSettingsRecord {
    return { ...this.metadata.calculationSettings };
  }

  setVolatileContext(context: WorkbookVolatileContextSnapshot): WorkbookVolatileContextRecord {
    this.metadata.volatileContext = { ...context };
    return this.metadata.volatileContext;
  }

  getVolatileContext(): WorkbookVolatileContextRecord {
    return { ...this.metadata.volatileContext };
  }

  setDefinedName(name: string, value: WorkbookDefinedNameValueSnapshot): WorkbookDefinedNameRecord {
    const trimmedName = name.trim();
    const record: WorkbookDefinedNameRecord = { name: trimmedName, value };
    this.metadata.definedNames.set(normalizeDefinedName(trimmedName), record);
    return record;
  }

  getDefinedName(name: string): WorkbookDefinedNameRecord | undefined {
    return this.metadata.definedNames.get(normalizeDefinedName(name));
  }

  deleteDefinedName(name: string): boolean {
    return this.metadata.definedNames.delete(normalizeDefinedName(name));
  }

  listDefinedNames(): WorkbookDefinedNameRecord[] {
    return [...this.metadata.definedNames.values()].toSorted((left, right) =>
      normalizeDefinedName(left.name).localeCompare(normalizeDefinedName(right.name)),
    );
  }

  setTable(record: WorkbookTableSnapshot): WorkbookTableRecord {
    const stored: WorkbookTableRecord = {
      name: record.name.trim(),
      sheetName: record.sheetName,
      startAddress: record.startAddress,
      endAddress: record.endAddress,
      columnNames: [...record.columnNames],
      headerRow: record.headerRow,
      totalsRow: record.totalsRow,
    };
    this.metadata.tables.set(tableKey(stored.name), stored);
    return stored;
  }

  getTable(name: string): WorkbookTableRecord | undefined {
    return this.metadata.tables.get(tableKey(name));
  }

  deleteTable(name: string): boolean {
    return this.metadata.tables.delete(tableKey(name));
  }

  listTables(): WorkbookTableRecord[] {
    return [...this.metadata.tables.values()].toSorted((left, right) =>
      tableKey(left.name).localeCompare(tableKey(right.name)),
    );
  }

  setRowMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
  ): WorkbookAxisMetadataRecord | undefined {
    return this.setAxisMetadata(
      this.getOrCreateSheet(sheetName),
      "row",
      this.metadata.rowMetadata,
      sheetName,
      start,
      count,
      size,
      hidden,
    );
  }

  getRowMetadata(
    sheetName: string,
    start: number,
    count: number,
  ): WorkbookAxisMetadataRecord | undefined {
    const sheet = this.getSheet(sheetName);
    return sheet ? this.getAxisMetadataRecord(sheet, "row", sheetName, start, count) : undefined;
  }

  listRowMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.listAxisMetadata(
      this.getSheet(sheetName),
      this.metadata.rowMetadata,
      sheetName,
      "row",
    );
  }

  setColumnMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
  ): WorkbookAxisMetadataRecord | undefined {
    return this.setAxisMetadata(
      this.getOrCreateSheet(sheetName),
      "column",
      this.metadata.columnMetadata,
      sheetName,
      start,
      count,
      size,
      hidden,
    );
  }

  getColumnMetadata(
    sheetName: string,
    start: number,
    count: number,
  ): WorkbookAxisMetadataRecord | undefined {
    const sheet = this.getSheet(sheetName);
    return sheet ? this.getAxisMetadataRecord(sheet, "column", sheetName, start, count) : undefined;
  }

  listColumnMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.listAxisMetadata(
      this.getSheet(sheetName),
      this.metadata.columnMetadata,
      sheetName,
      "column",
    );
  }

  listRowAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return this.listAxisEntries(this.getSheet(sheetName), "row");
  }

  listColumnAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return this.listAxisEntries(this.getSheet(sheetName), "column");
  }

  materializeRowAxisEntries(
    sheetName: string,
    start: number,
    count: number,
  ): WorkbookAxisEntrySnapshot[] {
    return this.materializeAxisEntries(this.getOrCreateSheet(sheetName), "row", start, count);
  }

  materializeColumnAxisEntries(
    sheetName: string,
    start: number,
    count: number,
  ): WorkbookAxisEntrySnapshot[] {
    return this.materializeAxisEntries(this.getOrCreateSheet(sheetName), "column", start, count);
  }

  insertRows(
    sheetName: string,
    start: number,
    count: number,
    entries?: readonly WorkbookAxisEntrySnapshot[],
  ): void {
    this.spliceAxisEntries(this.getOrCreateSheet(sheetName), "row", start, 0, count, entries);
  }

  deleteRows(sheetName: string, start: number, count: number): WorkbookAxisEntrySnapshot[] {
    return this.spliceAxisEntries(this.getOrCreateSheet(sheetName), "row", start, count, 0);
  }

  moveRows(sheetName: string, start: number, count: number, target: number): void {
    this.moveAxisEntries(this.getOrCreateSheet(sheetName), "row", start, count, target);
  }

  insertColumns(
    sheetName: string,
    start: number,
    count: number,
    entries?: readonly WorkbookAxisEntrySnapshot[],
  ): void {
    this.spliceAxisEntries(this.getOrCreateSheet(sheetName), "column", start, 0, count, entries);
  }

  deleteColumns(sheetName: string, start: number, count: number): WorkbookAxisEntrySnapshot[] {
    return this.spliceAxisEntries(this.getOrCreateSheet(sheetName), "column", start, count, 0);
  }

  moveColumns(sheetName: string, start: number, count: number, target: number): void {
    this.moveAxisEntries(this.getOrCreateSheet(sheetName), "column", start, count, target);
  }

  setFreezePane(sheetName: string, rows: number, cols: number): WorkbookFreezePaneRecord {
    const record: WorkbookFreezePaneRecord = { sheetName, rows, cols };
    this.metadata.freezePanes.set(sheetName, record);
    return record;
  }

  getFreezePane(sheetName: string): WorkbookFreezePaneRecord | undefined {
    return this.metadata.freezePanes.get(sheetName);
  }

  clearFreezePane(sheetName: string): boolean {
    return this.metadata.freezePanes.delete(sheetName);
  }

  setFilter(sheetName: string, range: CellRangeRef): WorkbookFilterRecord {
    const storedRange = canonicalWorkbookRangeRef(range);
    const record: WorkbookFilterRecord = { sheetName, range: storedRange };
    this.metadata.filters.set(filterKey(sheetName, storedRange), record);
    return record;
  }

  getFilter(sheetName: string, range: CellRangeRef): WorkbookFilterRecord | undefined {
    return this.metadata.filters.get(filterKey(sheetName, range));
  }

  deleteFilter(sheetName: string, range: CellRangeRef): boolean {
    return this.metadata.filters.delete(filterKey(sheetName, range));
  }

  listFilters(sheetName: string): WorkbookFilterRecord[] {
    return [...this.metadata.filters.values()]
      .filter((record) => record.sheetName === sheetName)
      .toSorted((left, right) =>
        filterKey(left.sheetName, left.range).localeCompare(
          filterKey(right.sheetName, right.range),
        ),
      );
  }

  setSort(
    sheetName: string,
    range: CellRangeRef,
    keys: readonly WorkbookSortKeyRecord[],
  ): WorkbookSortRecord {
    const storedRange = canonicalWorkbookRangeRef(range);
    const record: WorkbookSortRecord = {
      sheetName,
      range: storedRange,
      keys: keys.map((key) => Object.assign({}, key)),
    };
    this.metadata.sorts.set(sortKey(sheetName, storedRange), record);
    return record;
  }

  getSort(sheetName: string, range: CellRangeRef): WorkbookSortRecord | undefined {
    return this.metadata.sorts.get(sortKey(sheetName, range));
  }

  deleteSort(sheetName: string, range: CellRangeRef): boolean {
    return this.metadata.sorts.delete(sortKey(sheetName, range));
  }

  listSorts(sheetName: string): WorkbookSortRecord[] {
    return [...this.metadata.sorts.values()]
      .filter((record) => record.sheetName === sheetName)
      .toSorted((left, right) =>
        sortKey(left.sheetName, left.range).localeCompare(sortKey(right.sheetName, right.range)),
      );
  }

  setSpill(sheetName: string, address: string, rows: number, cols: number): WorkbookSpillRecord {
    const normalizedAddress = canonicalWorkbookAddress(sheetName, address);
    const record: WorkbookSpillRecord = { sheetName, address: normalizedAddress, rows, cols };
    this.metadata.spills.set(spillKey(sheetName, normalizedAddress), record);
    return record;
  }

  getSpill(sheetName: string, address: string): WorkbookSpillRecord | undefined {
    return this.metadata.spills.get(spillKey(sheetName, address));
  }

  deleteSpill(sheetName: string, address: string): boolean {
    return this.metadata.spills.delete(spillKey(sheetName, address));
  }

  listSpills(): WorkbookSpillRecord[] {
    return [...this.metadata.spills.values()].toSorted((left, right) =>
      `${left.sheetName}!${left.address}`.localeCompare(`${right.sheetName}!${right.address}`),
    );
  }

  setPivot(record: WorkbookPivotSnapshot): WorkbookPivotRecord {
    const normalizedAddress = canonicalWorkbookAddress(record.sheetName, record.address);
    const stored: WorkbookPivotRecord = {
      ...record,
      name: record.name.trim(),
      address: normalizedAddress,
      groupBy: [...record.groupBy],
      values: record.values.map((value) => Object.assign({}, value)),
      source: canonicalWorkbookRangeRef(record.source),
    };
    this.metadata.pivots.set(pivotKey(record.sheetName, normalizedAddress), stored);
    return stored;
  }

  getPivot(sheetName: string, address: string): WorkbookPivotRecord | undefined {
    return this.metadata.pivots.get(pivotKey(sheetName, address));
  }

  getPivotByKey(key: string): WorkbookPivotRecord | undefined {
    return this.metadata.pivots.get(key);
  }

  deletePivot(sheetName: string, address: string): boolean {
    return this.metadata.pivots.delete(pivotKey(sheetName, address));
  }

  listPivots(): WorkbookPivotRecord[] {
    return [...this.metadata.pivots.values()].toSorted((left, right) =>
      `${left.sheetName}!${left.address}`.localeCompare(`${right.sheetName}!${right.address}`),
    );
  }

  remapSheetCells(
    sheetName: string,
    axis: "row" | "column",
    remapIndex: (index: number) => number | undefined,
  ): { changedCellIndices: number[]; removedCellIndices: number[] } {
    const sheet = this.getSheet(sheetName);
    if (!sheet) {
      return { changedCellIndices: [], removedCellIndices: [] };
    }
    const entries: Array<{ cellIndex: number; row: number; col: number }> = [];
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      entries.push({ cellIndex, row, col });
    });
    entries.forEach(({ row, col }) => {
      this.cellKeyToIndex.delete(makeCellKey(sheet.id, row, col));
      sheet.grid.clear(row, col);
    });

    const changedCellIndices: number[] = [];
    const removedCellIndices: number[] = [];
    for (const { cellIndex, row, col } of entries) {
      const nextRow = axis === "row" ? remapIndex(row) : row;
      const nextCol = axis === "column" ? remapIndex(col) : col;
      if (nextRow === undefined || nextCol === undefined) {
        removedCellIndices.push(cellIndex);
        continue;
      }
      this.cellStore.rows[cellIndex] = nextRow;
      this.cellStore.cols[cellIndex] = nextCol;
      this.cellKeyToIndex.set(makeCellKey(sheet.id, nextRow, nextCol), cellIndex);
      sheet.grid.set(nextRow, nextCol, cellIndex);
      if (nextRow !== row || nextCol !== col) {
        changedCellIndices.push(cellIndex);
      }
    }

    return { changedCellIndices, removedCellIndices };
  }

  reset(workbookName = "Workbook"): void {
    this.workbookName = workbookName;
    this.sheetsByName.clear();
    this.sheetsById.clear();
    this.cellKeyToIndex.clear();
    this.cellFormats.clear();
    this.cellStyles.clear();
    this.styleKeys.clear();
    this.cellNumberFormats.clear();
    this.numberFormatKeys.clear();
    this.metadata.properties.clear();
    this.metadata.definedNames.clear();
    this.metadata.tables.clear();
    this.metadata.spills.clear();
    this.metadata.pivots.clear();
    this.metadata.rowMetadata.clear();
    this.metadata.columnMetadata.clear();
    this.metadata.calculationSettings = { mode: "automatic", compatibilityMode: "excel-modern" };
    this.metadata.volatileContext = { recalcEpoch: 0 };
    this.metadata.freezePanes.clear();
    this.metadata.filters.clear();
    this.metadata.sorts.clear();
    this.nextSheetId = 1;
    this.nextRowAxisId = 1;
    this.nextColumnAxisId = 1;
    this.nextStyleId = 1;
    this.nextFormatId = 1;
    this.cellStore.reset();
    this.ensureDefaultStyle();
    this.ensureDefaultNumberFormat();
  }

  private ensureDefaultStyle(): void {
    const defaultStyle: WorkbookCellStyleRecord = { id: WorkbookStore.defaultStyleId };
    this.cellStyles.set(defaultStyle.id, defaultStyle);
    this.styleKeys.set(cellStyleKey(defaultStyle), defaultStyle.id);
  }

  private ensureDefaultNumberFormat(): void {
    const defaultFormat: WorkbookCellNumberFormatRecord = {
      id: WorkbookStore.defaultFormatId,
      code: "general",
      kind: "general",
    };
    this.cellNumberFormats.set(defaultFormat.id, defaultFormat);
    this.numberFormatKeys.set(defaultFormat.code, defaultFormat.id);
  }

  private bumpStyleId(id: string): void {
    const match = /^style-(\d+)$/.exec(id);
    if (!match) {
      return;
    }
    const numericId = Number.parseInt(match[1]!, 10);
    if (Number.isFinite(numericId)) {
      this.nextStyleId = Math.max(this.nextStyleId, numericId + 1);
    }
  }

  private bumpSheetId(id: number): void {
    if (Number.isInteger(id) && id >= this.nextSheetId) {
      this.nextSheetId = id + 1;
    }
  }

  private bumpFormatId(id: string): void {
    const match = /^format-(\d+)$/.exec(id);
    if (!match) {
      return;
    }
    const numericId = Number.parseInt(match[1]!, 10);
    if (Number.isFinite(numericId)) {
      this.nextFormatId = Math.max(this.nextFormatId, numericId + 1);
    }
  }

  private setAxisMetadata(
    sheet: SheetRecord,
    axis: "row" | "column",
    bucket: Map<string, WorkbookAxisMetadataRecord>,
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
  ): WorkbookAxisMetadataRecord | undefined {
    const entries = this.materializeAxisEntryRecords(sheet, axis, start, count);
    entries.forEach((entry) => {
      entry.size = size;
      entry.hidden = hidden;
    });
    this.syncAxisMetadataBucket(sheetName, sheet, axis, bucket);
    const record = this.getAxisMetadataRecord(sheet, axis, sheetName, start, count);
    if (!record) {
      bucket.delete(axisMetadataKey(sheetName, start, count));
    }
    return record;
  }

  private listAxisMetadata(
    sheet: SheetRecord | undefined,
    bucket: Map<string, WorkbookAxisMetadataRecord>,
    sheetName: string,
    axis: "row" | "column",
  ): WorkbookAxisMetadataRecord[] {
    if (!sheet) {
      return [];
    }
    this.syncAxisMetadataBucket(sheetName, sheet, axis, bucket);
    return [...bucket.values()]
      .filter((record) => record.sheetName === sheetName)
      .toSorted((left, right) => left.start - right.start || left.count - right.count);
  }

  private deleteSheetMetadata(sheetName: string): void {
    const sheet = this.getSheet(sheetName);
    deleteRecordsBySheet(this.metadata.tables, sheetName, (record) => record.sheetName);
    deleteRecordsBySheet(this.metadata.spills, sheetName, (record) => record.sheetName);
    deleteRecordsBySheet(this.metadata.pivots, sheetName, (record) => record.sheetName);
    deleteRecordsBySheet(this.metadata.rowMetadata, sheetName, (record) => record.sheetName);
    deleteRecordsBySheet(this.metadata.columnMetadata, sheetName, (record) => record.sheetName);
    deleteRecordsBySheet(this.metadata.filters, sheetName, (record) => record.sheetName);
    deleteRecordsBySheet(this.metadata.sorts, sheetName, (record) => record.sheetName);
    this.metadata.freezePanes.delete(sheetName);
    if (sheet) {
      sheet.rowAxis.length = 0;
      sheet.columnAxis.length = 0;
      sheet.styleRanges.length = 0;
      sheet.formatRanges.length = 0;
    }
  }

  private listAxisEntries(
    sheet: SheetRecord | undefined,
    axis: "row" | "column",
  ): WorkbookAxisEntrySnapshot[] {
    if (!sheet) {
      return [];
    }
    const entries = axis === "row" ? sheet.rowAxis : sheet.columnAxis;
    const result: WorkbookAxisEntrySnapshot[] = [];
    entries.forEach((entry, index) => {
      if (!entry) {
        return;
      }
      const snapshot: WorkbookAxisEntrySnapshot = { id: entry.id, index };
      if (entry.size !== null) {
        snapshot.size = entry.size;
      }
      if (entry.hidden !== null) {
        snapshot.hidden = entry.hidden;
      }
      result.push(snapshot);
    });
    return result;
  }

  private materializeAxisEntries(
    sheet: SheetRecord,
    axis: "row" | "column",
    start: number,
    count: number,
  ): WorkbookAxisEntrySnapshot[] {
    return this.materializeAxisEntryRecords(sheet, axis, start, count).map((entry, offset) => {
      const snapshot: WorkbookAxisEntrySnapshot = { id: entry.id, index: start + offset };
      if (entry.size !== null) {
        snapshot.size = entry.size;
      }
      if (entry.hidden !== null) {
        snapshot.hidden = entry.hidden;
      }
      return snapshot;
    });
  }

  private materializeAxisEntryRecords(
    sheet: SheetRecord,
    axis: "row" | "column",
    start: number,
    count: number,
  ): WorkbookAxisEntryRecord[] {
    const entries = axis === "row" ? sheet.rowAxis : sheet.columnAxis;
    const materialized: WorkbookAxisEntryRecord[] = [];
    for (let index = 0; index < count; index += 1) {
      const position = start + index;
      let entry = entries[position];
      if (!entry) {
        entry = {
          id: axis === "row" ? `row-${this.nextRowAxisId++}` : `column-${this.nextColumnAxisId++}`,
          size: null,
          hidden: null,
        };
        entries[position] = entry;
      }
      materialized.push(entry);
    }
    return materialized;
  }

  private spliceAxisEntries(
    sheet: SheetRecord,
    axis: "row" | "column",
    start: number,
    deleteCount: number,
    insertCount: number,
    entries?: readonly WorkbookAxisEntrySnapshot[],
  ): WorkbookAxisEntrySnapshot[] {
    const axisEntries = axis === "row" ? sheet.rowAxis : sheet.columnAxis;
    if (axisEntries.length < start) {
      axisEntries.length = start;
    }
    if (deleteCount > 0) {
      this.materializeAxisEntryRecords(sheet, axis, start, deleteCount);
    }
    const removed = axisEntries.splice(
      start,
      deleteCount,
      ...Array.from({ length: insertCount }, (_, index) => {
        const provided = entries?.[index];
        return provided
          ? { id: provided.id, size: provided.size ?? null, hidden: provided.hidden ?? null }
          : {
              id:
                axis === "row"
                  ? `row-${this.nextRowAxisId++}`
                  : `column-${this.nextColumnAxisId++}`,
              size: null,
              hidden: null,
            };
      }),
    );
    return removed.flatMap((entry, index) => {
      if (!entry) {
        return [];
      }
      const snapshot: WorkbookAxisEntrySnapshot = { id: entry.id, index: start + index };
      if (entry.size !== null) {
        snapshot.size = entry.size;
      }
      if (entry.hidden !== null) {
        snapshot.hidden = entry.hidden;
      }
      return [snapshot];
    });
  }

  private moveAxisEntries(
    sheet: SheetRecord,
    axis: "row" | "column",
    start: number,
    count: number,
    target: number,
  ): void {
    if (count <= 0 || start === target) {
      return;
    }
    const axisEntries = axis === "row" ? sheet.rowAxis : sheet.columnAxis;
    this.materializeAxisEntryRecords(sheet, axis, start, count);
    const moved = axisEntries.splice(start, count);
    axisEntries.splice(target, 0, ...moved);
  }

  private getAxisMetadataRecord(
    sheet: SheetRecord,
    axis: "row" | "column",
    sheetName: string,
    start: number,
    count: number,
  ): WorkbookAxisMetadataRecord | undefined {
    const entries = axis === "row" ? sheet.rowAxis : sheet.columnAxis;
    let size: number | null | undefined;
    let hidden: boolean | null | undefined;
    let sawMaterialized = false;
    for (let index = start; index < start + count; index += 1) {
      const entry = entries[index];
      if (!entry) {
        if (size === undefined) {
          size = null;
        }
        if (hidden === undefined) {
          hidden = null;
        }
        continue;
      }
      sawMaterialized = true;
      size ??= entry.size;
      hidden ??= entry.hidden;
      if (size !== entry.size || hidden !== entry.hidden) {
        return undefined;
      }
    }
    if (!sawMaterialized || ((size ?? null) === null && (hidden ?? null) === null)) {
      return undefined;
    }
    return { sheetName, start, count, size: size ?? null, hidden: hidden ?? null };
  }

  private syncAxisMetadataBucket(
    sheetName: string,
    sheet: SheetRecord,
    axis: "row" | "column",
    bucket: Map<string, WorkbookAxisMetadataRecord>,
  ): void {
    deleteRecordsBySheet(bucket, sheetName, (record) => record.sheetName);
    const entries = axis === "row" ? sheet.rowAxis : sheet.columnAxis;
    let cursor = 0;
    while (cursor < entries.length) {
      const entry = entries[cursor];
      if (!entry || (entry.size === null && entry.hidden === null)) {
        cursor += 1;
        continue;
      }
      const start = cursor;
      const size = entry.size;
      const hidden = entry.hidden;
      cursor += 1;
      while (cursor < entries.length) {
        const next = entries[cursor];
        if (!next || next.size !== size || next.hidden !== hidden) {
          break;
        }
        cursor += 1;
      }
      const record: WorkbookAxisMetadataRecord = {
        sheetName,
        start,
        count: cursor - start,
        size,
        hidden,
      };
      bucket.set(axisMetadataKey(sheetName, start, record.count), record);
    }
  }
}

function deleteRecordsBySheet<T>(
  bucket: Map<string, T>,
  sheetName: string,
  readSheetName: (record: T) => string,
): void {
  for (const [key, record] of bucket.entries()) {
    if (readSheetName(record) === sheetName) {
      bucket.delete(key);
    }
  }
}

function rekeyRecords<T>(bucket: Map<string, T>, rewrite: (record: T) => T): void {
  const rewritten = [...bucket.values()].map((record) => rewrite(record));
  bucket.clear();
  rewritten.forEach((record) => {
    bucket.set(recordKey(record), record);
  });
}

function recordKey(record: unknown): string {
  if (isFreezePaneRecord(record)) {
    return record.sheetName;
  }
  if (isAxisMetadataRecord(record)) {
    return axisMetadataKey(record.sheetName, record.start, record.count);
  }
  if (isFilterRecord(record)) {
    return filterKey(record.sheetName, record.range);
  }
  if (isSortRecord(record)) {
    return sortKey(record.sheetName, record.range);
  }
  if (isTableRecord(record)) {
    return tableKey(record.name);
  }
  if (isSpillRecord(record)) {
    return spillKey(record.sheetName, record.address);
  }
  if (isPivotRecord(record)) {
    return pivotKey(record.sheetName, record.address);
  }
  throw new Error("Unsupported workbook metadata record");
}

function isFreezePaneRecord(record: unknown): record is WorkbookFreezePaneRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "sheetName" in record &&
    "rows" in record &&
    "cols" in record &&
    !("address" in record)
  );
}

function isAxisMetadataRecord(record: unknown): record is WorkbookAxisMetadataRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "sheetName" in record &&
    "start" in record &&
    "count" in record &&
    "size" in record &&
    "hidden" in record
  );
}

function isFilterRecord(record: unknown): record is WorkbookFilterRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "sheetName" in record &&
    "range" in record &&
    !("keys" in record)
  );
}

function isSortRecord(record: unknown): record is WorkbookSortRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "sheetName" in record &&
    "range" in record &&
    "keys" in record
  );
}

function isTableRecord(record: unknown): record is WorkbookTableRecord {
  return (
    typeof record === "object" && record !== null && "name" in record && "columnNames" in record
  );
}

function isSpillRecord(record: unknown): record is WorkbookSpillRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "sheetName" in record &&
    "address" in record &&
    "rows" in record &&
    "cols" in record &&
    !("source" in record)
  );
}

function isPivotRecord(record: unknown): record is WorkbookPivotRecord {
  return (
    typeof record === "object" && record !== null && "sheetName" in record && "source" in record
  );
}

function normalizeCellStyleRecord(style: CellStyleRecord): WorkbookCellStyleRecord {
  const record: WorkbookCellStyleRecord = { id: style.id.trim() };
  if (record.id.length === 0) {
    throw new Error("Cell style id must be non-empty");
  }
  const fillColor = normalizeBackgroundColor(style.fill?.backgroundColor);
  if (fillColor) {
    record.fill = { backgroundColor: fillColor };
  }
  const fontFamily = normalizeFontFamily(style.font?.family);
  const fontColor = normalizeColor(style.font?.color);
  const fontSize =
    typeof style.font?.size === "number" && Number.isFinite(style.font.size)
      ? Math.max(1, Math.min(144, style.font.size))
      : undefined;
  if (
    fontFamily ||
    fontColor ||
    fontSize !== undefined ||
    style.font?.bold === true ||
    style.font?.italic === true ||
    style.font?.underline === true
  ) {
    record.font = {
      ...(fontFamily ? { family: fontFamily } : {}),
      ...(fontColor ? { color: fontColor } : {}),
      ...(fontSize !== undefined ? { size: fontSize } : {}),
      ...(style.font?.bold ? { bold: true } : {}),
      ...(style.font?.italic ? { italic: true } : {}),
      ...(style.font?.underline ? { underline: true } : {}),
    };
  }
  const horizontal = normalizeHorizontalAlignment(style.alignment?.horizontal);
  const vertical = normalizeVerticalAlignment(style.alignment?.vertical);
  const wrap = style.alignment?.wrap === true ? true : undefined;
  const indent =
    typeof style.alignment?.indent === "number" && Number.isFinite(style.alignment.indent)
      ? Math.max(0, Math.min(16, Math.trunc(style.alignment.indent)))
      : undefined;
  if (horizontal || vertical || wrap || indent !== undefined) {
    record.alignment = {
      ...(horizontal ? { horizontal } : {}),
      ...(vertical ? { vertical } : {}),
      ...(wrap ? { wrap: true } : {}),
      ...(indent !== undefined ? { indent } : {}),
    };
  }
  const borders = normalizeBorders(style.borders);
  if (borders) {
    record.borders = borders;
  }
  return record;
}

function normalizeBackgroundColor(color: string | undefined): string | undefined {
  return normalizeColor(color);
}

function normalizeColor(color: string | undefined): string | undefined {
  if (!color) {
    return undefined;
  }
  const trimmed = color.trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(trimmed)) {
    return trimmed;
  }
  if (/^#[0-9a-f]{3}$/.test(trimmed)) {
    return `#${trimmed[1]}${trimmed[1]}${trimmed[2]}${trimmed[2]}${trimmed[3]}${trimmed[3]}`;
  }
  throw new Error(`Unsupported background color: ${color}`);
}

function normalizeFontFamily(family: string | undefined): string | undefined {
  if (!family) {
    return undefined;
  }
  const trimmed = family.trim();
  return trimmed.length === 0 ? undefined : trimmed;
}

function normalizeHorizontalAlignment(
  value: CellHorizontalAlignment | undefined,
): CellHorizontalAlignment | undefined {
  switch (value) {
    case undefined:
      return undefined;
    case "general":
    case "left":
    case "center":
    case "right":
      return value;
    default:
      return undefined;
  }
}

function normalizeVerticalAlignment(
  value: CellVerticalAlignment | undefined,
): CellVerticalAlignment | undefined {
  switch (value) {
    case undefined:
      return undefined;
    case "top":
    case "middle":
    case "bottom":
      return value;
    default:
      return undefined;
  }
}

function normalizeBorderStyle(
  value: CellBorderSideSnapshot["style"] | undefined,
): CellBorderSideSnapshot["style"] | undefined {
  switch (value) {
    case undefined:
      return undefined;
    case "solid":
    case "dashed":
    case "dotted":
    case "double":
      return value;
    default:
      return undefined;
  }
}

function normalizeBorderWeight(
  value: CellBorderSideSnapshot["weight"] | undefined,
): CellBorderSideSnapshot["weight"] | undefined {
  switch (value) {
    case undefined:
      return undefined;
    case "thin":
    case "medium":
    case "thick":
      return value;
    default:
      return undefined;
  }
}

function normalizeBorderSide(
  side: CellBorderSideSnapshot | undefined,
): CellBorderSideSnapshot | undefined {
  if (!side) {
    return undefined;
  }
  const color = normalizeColor(side.color);
  const style = normalizeBorderStyle(side.style);
  const weight = normalizeBorderWeight(side.weight);
  if (!color || !style || !weight) {
    return undefined;
  }
  return { color, style, weight };
}

function normalizeBorders(
  style: CellStyleRecord["borders"],
): CellStyleRecord["borders"] | undefined {
  if (!style) {
    return undefined;
  }
  const top = normalizeBorderSide(style.top);
  const right = normalizeBorderSide(style.right);
  const bottom = normalizeBorderSide(style.bottom);
  const left = normalizeBorderSide(style.left);
  if (!top && !right && !bottom && !left) {
    return undefined;
  }
  return {
    ...(top ? { top } : {}),
    ...(right ? { right } : {}),
    ...(bottom ? { bottom } : {}),
    ...(left ? { left } : {}),
  };
}

function normalizeCellNumberFormatRecord(
  format: CellNumberFormatRecord,
): WorkbookCellNumberFormatRecord {
  const id = format.id.trim();
  if (id.length === 0) {
    throw new Error("Cell number format id must be non-empty");
  }
  const code = buildCellNumberFormatCode(format.code);
  return {
    id,
    code,
    kind: format.kind ?? getCellNumberFormatKind(code),
  };
}

function cellStyleKey(style: CellStyleRecord): string {
  return JSON.stringify({
    fill: style.fill?.backgroundColor ?? null,
    font: style.font ?? null,
    alignment: style.alignment ?? null,
    borders: style.borders ?? null,
  });
}

function cellStyleIdForKey(key: string): string {
  let hash = 2166136261;
  for (let index = 0; index < key.length; index += 1) {
    hash ^= key.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return `style-${(hash >>> 0).toString(16)}`;
}

function cellNumberFormatIdForCode(code: string): string {
  let hash = 2166136261;
  for (let index = 0; index < code.length; index += 1) {
    hash ^= code.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return `format-${(hash >>> 0).toString(16)}`;
}

function normalizeMetadataKey(key: string): string {
  const trimmed = key.trim();
  if (trimmed.length === 0) {
    throw new Error("Workbook metadata keys must be non-empty");
  }
  return trimmed;
}

function axisMetadataKey(sheetName: string, start: number, count: number): string {
  return `${sheetName}:${start}:${count}`;
}

function filterKey(sheetName: string, range: CellRangeRef): string {
  const normalized = canonicalWorkbookRangeRef(range);
  return `${sheetName}:${normalized.startAddress}:${normalized.endAddress}`;
}

function sortKey(sheetName: string, range: CellRangeRef): string {
  const normalized = canonicalWorkbookRangeRef(range);
  return `${sheetName}:${normalized.startAddress}:${normalized.endAddress}`;
}

function tableKey(name: string): string {
  return normalizeWorkbookObjectName(name, "Tables");
}

function spillKey(sheetName: string, address: string): string {
  return `${sheetName}!${canonicalWorkbookAddress(sheetName, address)}`;
}

export function pivotKey(sheetName: string, address: string): string {
  return `${sheetName}!${canonicalWorkbookAddress(sheetName, address)}`;
}

export function makeCellKey(sheetId: number, row: number, col: number): number {
  return sheetId * SHEET_STRIDE + row * MAX_COLS + col;
}
