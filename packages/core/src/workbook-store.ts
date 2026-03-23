import {
  MAX_COLS,
  MAX_ROWS,
  type CellRangeRef,
  type LiteralInput,
  type WorkbookAxisEntrySnapshot,
  type WorkbookCalculationSettingsSnapshot,
  type WorkbookPivotSnapshot,
  type WorkbookPivotValueSnapshot,
  type WorkbookTableSnapshot,
  type WorkbookVolatileContextSnapshot
} from "@bilig/protocol";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import { SheetGrid } from "./sheet-grid.js";
import { CellStore } from "./cell-store.js";

const SHEET_STRIDE = MAX_ROWS * MAX_COLS;

export interface WorkbookDefinedNameRecord {
  name: string;
  value: LiteralInput;
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
}

export interface EnsuredCell {
  cellIndex: number;
  created: boolean;
}

export class WorkbookStore {
  readonly cellStore = new CellStore();
  readonly sheetsByName = new Map<string, SheetRecord>();
  readonly sheetsById = new Map<number, SheetRecord>();
  readonly cellKeyToIndex = new Map<number, number>();
  readonly cellFormats = new Map<number, string>();
  readonly metadata: WorkbookMetadataRecord = {
    properties: new Map(),
    definedNames: new Map(),
    tables: new Map(),
    spills: new Map(),
    pivots: new Map(),
    rowMetadata: new Map(),
    columnMetadata: new Map(),
    calculationSettings: { mode: "automatic" },
    volatileContext: { recalcEpoch: 0 },
    freezePanes: new Map(),
    filters: new Map(),
    sorts: new Map()
  };
  workbookName: string;
  private nextSheetId = 1;
  private nextRowAxisId = 1;
  private nextColumnAxisId = 1;

  constructor(workbookName = "Workbook") {
    this.workbookName = workbookName;
  }

  createSheet(name: string, order = this.sheetsByName.size): SheetRecord {
    const existing = this.sheetsByName.get(name);
    if (existing) {
      existing.order = order;
      return existing;
    }
    const sheet: SheetRecord = {
      id: this.nextSheetId++,
      name,
      order,
      grid: new SheetGrid(),
      rowAxis: [],
      columnAxis: []
    };
    this.sheetsByName.set(name, sheet);
    this.sheetsById.set(sheet.id, sheet);
    return sheet;
  }

  deleteSheet(name: string): void {
    const sheet = this.sheetsByName.get(name);
    if (!sheet) return;
    sheet.grid.forEachCell((cellIndex) => {
      const key = makeCellKey(sheet.id, this.cellStore.rows[cellIndex]!, this.cellStore.cols[cellIndex]!);
      this.cellKeyToIndex.delete(key);
      this.cellFormats.delete(cellIndex);
    });
    this.deleteSheetMetadata(name);
    this.sheetsByName.delete(name);
    this.sheetsById.delete(sheet.id);
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
    this.cellFormats.set(index, format);
  }

  getCellFormat(index: number): string | undefined {
    return this.cellFormats.get(index);
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
    return [...this.metadata.properties.values()].toSorted((left, right) => left.key.localeCompare(right.key));
  }

  setCalculationSettings(settings: WorkbookCalculationSettingsSnapshot): WorkbookCalculationSettingsRecord {
    this.metadata.calculationSettings = { ...settings };
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

  setDefinedName(name: string, value: LiteralInput): WorkbookDefinedNameRecord {
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
      normalizeDefinedName(left.name).localeCompare(normalizeDefinedName(right.name))
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
      totalsRow: record.totalsRow
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
    return [...this.metadata.tables.values()].toSorted((left, right) => tableKey(left.name).localeCompare(tableKey(right.name)));
  }

  setRowMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null
  ): WorkbookAxisMetadataRecord | undefined {
    return this.setAxisMetadata(this.getOrCreateSheet(sheetName), "row", this.metadata.rowMetadata, sheetName, start, count, size, hidden);
  }

  getRowMetadata(sheetName: string, start: number, count: number): WorkbookAxisMetadataRecord | undefined {
    const sheet = this.getSheet(sheetName);
    return sheet ? this.getAxisMetadataRecord(sheet, "row", sheetName, start, count) : undefined;
  }

  listRowMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.listAxisMetadata(this.getSheet(sheetName), this.metadata.rowMetadata, sheetName, "row");
  }

  setColumnMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null
  ): WorkbookAxisMetadataRecord | undefined {
    return this.setAxisMetadata(
      this.getOrCreateSheet(sheetName),
      "column",
      this.metadata.columnMetadata,
      sheetName,
      start,
      count,
      size,
      hidden
    );
  }

  getColumnMetadata(sheetName: string, start: number, count: number): WorkbookAxisMetadataRecord | undefined {
    const sheet = this.getSheet(sheetName);
    return sheet ? this.getAxisMetadataRecord(sheet, "column", sheetName, start, count) : undefined;
  }

  listColumnMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.listAxisMetadata(this.getSheet(sheetName), this.metadata.columnMetadata, sheetName, "column");
  }

  listRowAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return this.listAxisEntries(this.getSheet(sheetName), "row");
  }

  listColumnAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return this.listAxisEntries(this.getSheet(sheetName), "column");
  }

  materializeRowAxisEntries(sheetName: string, start: number, count: number): WorkbookAxisEntrySnapshot[] {
    return this.materializeAxisEntries(this.getOrCreateSheet(sheetName), "row", start, count);
  }

  materializeColumnAxisEntries(sheetName: string, start: number, count: number): WorkbookAxisEntrySnapshot[] {
    return this.materializeAxisEntries(this.getOrCreateSheet(sheetName), "column", start, count);
  }

  insertRows(sheetName: string, start: number, count: number, entries?: readonly WorkbookAxisEntrySnapshot[]): void {
    this.spliceAxisEntries(this.getOrCreateSheet(sheetName), "row", start, 0, count, entries);
  }

  deleteRows(sheetName: string, start: number, count: number): WorkbookAxisEntrySnapshot[] {
    return this.spliceAxisEntries(this.getOrCreateSheet(sheetName), "row", start, count, 0);
  }

  moveRows(sheetName: string, start: number, count: number, target: number): void {
    this.moveAxisEntries(this.getOrCreateSheet(sheetName), "row", start, count, target);
  }

  insertColumns(sheetName: string, start: number, count: number, entries?: readonly WorkbookAxisEntrySnapshot[]): void {
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
    const storedRange = { ...range };
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
      .toSorted((left, right) => filterKey(left.sheetName, left.range).localeCompare(filterKey(right.sheetName, right.range)));
  }

  setSort(sheetName: string, range: CellRangeRef, keys: readonly WorkbookSortKeyRecord[]): WorkbookSortRecord {
    const storedRange = { ...range };
    const record: WorkbookSortRecord = {
      sheetName,
      range: storedRange,
      keys: keys.map((key) => Object.assign({}, key))
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
      .toSorted((left, right) => sortKey(left.sheetName, left.range).localeCompare(sortKey(right.sheetName, right.range)));
  }

  setSpill(sheetName: string, address: string, rows: number, cols: number): WorkbookSpillRecord {
    const record: WorkbookSpillRecord = { sheetName, address, rows, cols };
    this.metadata.spills.set(spillKey(sheetName, address), record);
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
      `${left.sheetName}!${left.address}`.localeCompare(`${right.sheetName}!${right.address}`)
    );
  }

  setPivot(record: WorkbookPivotSnapshot): WorkbookPivotRecord {
    const stored: WorkbookPivotRecord = {
      ...record,
      name: record.name.trim(),
      groupBy: [...record.groupBy],
      values: record.values.map((value) => Object.assign({}, value)),
      source: { ...record.source }
    };
    this.metadata.pivots.set(pivotKey(record.sheetName, record.address), stored);
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
      `${left.sheetName}!${left.address}`.localeCompare(`${right.sheetName}!${right.address}`)
    );
  }

  remapSheetCells(
    sheetName: string,
    axis: "row" | "column",
    remapIndex: (index: number) => number | undefined
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
    this.metadata.properties.clear();
    this.metadata.definedNames.clear();
    this.metadata.tables.clear();
    this.metadata.spills.clear();
    this.metadata.pivots.clear();
    this.metadata.rowMetadata.clear();
    this.metadata.columnMetadata.clear();
    this.metadata.calculationSettings = { mode: "automatic" };
    this.metadata.volatileContext = { recalcEpoch: 0 };
    this.metadata.freezePanes.clear();
    this.metadata.filters.clear();
    this.metadata.sorts.clear();
    this.nextSheetId = 1;
    this.nextRowAxisId = 1;
    this.nextColumnAxisId = 1;
    this.cellStore.reset();
  }

  private setAxisMetadata(
    sheet: SheetRecord,
    axis: "row" | "column",
    bucket: Map<string, WorkbookAxisMetadataRecord>,
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null
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
    axis: "row" | "column"
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
    }
  }

  private listAxisEntries(sheet: SheetRecord | undefined, axis: "row" | "column"): WorkbookAxisEntrySnapshot[] {
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
    count: number
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
    count: number
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
          hidden: null
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
    entries?: readonly WorkbookAxisEntrySnapshot[]
  ): WorkbookAxisEntrySnapshot[] {
    const axisEntries = axis === "row" ? sheet.rowAxis : sheet.columnAxis;
    if (axisEntries.length < start) {
      axisEntries.length = start;
    }
    if (deleteCount > 0) {
      this.materializeAxisEntryRecords(sheet, axis, start, deleteCount);
    }
    const removed = axisEntries.splice(start, deleteCount, ...Array.from({ length: insertCount }, (_, index) => {
      const provided = entries?.[index];
      return provided
        ? { id: provided.id, size: provided.size ?? null, hidden: provided.hidden ?? null }
        : {
            id: axis === "row" ? `row-${this.nextRowAxisId++}` : `column-${this.nextColumnAxisId++}`,
            size: null,
            hidden: null
          };
    }));
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
    target: number
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
    count: number
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
    if (!sawMaterialized || (size ?? null) === null && (hidden ?? null) === null) {
      return undefined;
    }
    return { sheetName, start, count, size: size ?? null, hidden: hidden ?? null };
  }

  private syncAxisMetadataBucket(
    sheetName: string,
    sheet: SheetRecord,
    axis: "row" | "column",
    bucket: Map<string, WorkbookAxisMetadataRecord>
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
      const record: WorkbookAxisMetadataRecord = { sheetName, start, count: cursor - start, size, hidden };
      bucket.set(axisMetadataKey(sheetName, start, record.count), record);
    }
  }
}

function deleteRecordsBySheet<T>(
  bucket: Map<string, T>,
  sheetName: string,
  readSheetName: (record: T) => string
): void {
  for (const [key, record] of bucket.entries()) {
    if (readSheetName(record) === sheetName) {
      bucket.delete(key);
    }
  }
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
  return `${sheetName}:${range.startAddress}:${range.endAddress}`;
}

function sortKey(sheetName: string, range: CellRangeRef): string {
  return `${sheetName}:${range.startAddress}:${range.endAddress}`;
}

function tableKey(name: string): string {
  return normalizeWorkbookObjectName(name, "Tables");
}

function spillKey(sheetName: string, address: string): string {
  return `${sheetName}!${address}`;
}

export function pivotKey(sheetName: string, address: string): string {
  return `${sheetName}!${address}`;
}

export function makeCellKey(sheetId: number, row: number, col: number): number {
  return sheetId * SHEET_STRIDE + row * MAX_COLS + col;
}
