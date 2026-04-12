import {
  ValueTag,
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
  type WorkbookCommentThreadSnapshot,
  type WorkbookConditionalFormatSnapshot,
  type WorkbookDataValidationSnapshot,
  type WorkbookDefinedNameValueSnapshot,
  type WorkbookNoteSnapshot,
  type WorkbookRangeProtectionSnapshot,
  type WorkbookSheetProtectionSnapshot,
  type WorkbookPivotSnapshot,
  type WorkbookTableSnapshot,
  type WorkbookVolatileContextSnapshot,
} from "@bilig/protocol";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import { SheetGrid } from "./sheet-grid.js";
import { CellFlags, CellStore } from "./cell-store.js";
import {
  createWorkbookMetadataService,
  runWorkbookMetadataEffect,
} from "./workbook-metadata-service.js";
import {
  createWorkbookMetadataRecord,
  type WorkbookAxisEntryRecord,
  type WorkbookAxisMetadataRecord,
  type WorkbookCalculationSettingsRecord,
  type WorkbookCommentThreadRecord,
  type WorkbookConditionalFormatRecord,
  type WorkbookDataValidationRecord,
  type WorkbookCellNumberFormatRecord,
  type WorkbookCellStyleRecord,
  type WorkbookDefinedNameRecord,
  type WorkbookFilterRecord,
  type WorkbookFormatRangeRecord,
  type WorkbookFreezePaneRecord,
  type WorkbookMetadataRecord,
  type WorkbookPivotRecord,
  type WorkbookPropertyRecord,
  type WorkbookRangeProtectionRecord,
  type WorkbookSheetProtectionRecord,
  type WorkbookSortKeyRecord,
  type WorkbookSortRecord,
  type WorkbookSpillRecord,
  type WorkbookStyleRangeRecord,
  type WorkbookTableRecord,
  type WorkbookVolatileContextRecord,
  type WorkbookNoteRecord,
} from "./workbook-metadata-types.js";
import {
  getAxisMetadataRecord,
  listAxisEntries,
  materializeAxisEntries,
  materializeAxisEntryRecords,
  moveAxisEntries,
  snapshotAxisEntriesInRange,
  spliceAxisEntries,
  syncAxisMetadataBucket,
} from "./workbook-axis-records.js";
import { cellStyleKey, axisMetadataKey } from "./workbook-store-records.js";
import {
  getCellStyle as readCellStyle,
  getCellNumberFormat as readCellNumberFormat,
  getRangeFormatId as readRangeFormatId,
  getStyleId as readStyleId,
  internCellNumberFormat as internWorkbookCellNumberFormat,
  internCellStyle as internWorkbookCellStyle,
  listCellNumberFormats as listWorkbookCellNumberFormats,
  listCellStyles as listWorkbookCellStyles,
  listFormatRanges as listWorkbookFormatRanges,
  listStyleRanges as listWorkbookStyleRanges,
  setFormatRange as storeFormatRange,
  setFormatRanges as replaceFormatRanges,
  setStyleRange as storeStyleRange,
  setStyleRanges as replaceStyleRanges,
  upsertCellNumberFormat as storeCellNumberFormat,
  upsertCellStyle as storeCellStyle,
} from "./workbook-style-format-store.js";

const SHEET_STRIDE = MAX_ROWS * MAX_COLS;
export {
  normalizeDefinedName,
  normalizeWorkbookObjectName,
  pivotKey,
} from "./workbook-metadata-types.js";
export type {
  WorkbookAxisEntryRecord,
  WorkbookAxisMetadataRecord,
  WorkbookCalculationSettingsRecord,
  WorkbookCommentThreadRecord,
  WorkbookConditionalFormatRecord,
  WorkbookDataValidationRecord,
  WorkbookCellNumberFormatRecord,
  WorkbookCellStyleRecord,
  WorkbookDefinedNameRecord,
  WorkbookFilterRecord,
  WorkbookFormatRangeRecord,
  WorkbookFreezePaneRecord,
  WorkbookMetadataRecord,
  WorkbookPivotRecord,
  WorkbookPropertyRecord,
  WorkbookRangeProtectionRecord,
  WorkbookSheetProtectionRecord,
  WorkbookSortKeyRecord,
  WorkbookSortRecord,
  WorkbookSpillRecord,
  WorkbookStyleRangeRecord,
  WorkbookTableRecord,
  WorkbookVolatileContextRecord,
  WorkbookNoteRecord,
} from "./workbook-metadata-types.js";

export interface SheetRecord {
  id: number;
  name: string;
  order: number;
  grid: SheetGrid;
  columnVersions: Uint32Array;
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
  readonly metadata: WorkbookMetadataRecord = createWorkbookMetadataRecord();
  private readonly metadataService = createWorkbookMetadataService(this.metadata);
  workbookName: string;
  private batchedColumnVersionUpdates: Map<number, Set<number>> | null = null;
  private nextSheetId = 1;
  private nextRowAxisId = 1;
  private nextColumnAxisId = 1;
  private nextStyleId = 1;
  private nextFormatId = 1;

  constructor(workbookName = "Workbook") {
    this.workbookName = workbookName;
    this.cellStore.onSetValue = (index) => {
      this.bumpColumnVersionByCellIndex(index);
    };
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
      columnVersions: new Uint32Array(MAX_COLS),
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
    runWorkbookMetadataEffect(this.metadataService.deleteSheetRecords(name));
    sheet.rowAxis.length = 0;
    sheet.columnAxis.length = 0;
    sheet.styleRanges.length = 0;
    sheet.formatRanges.length = 0;
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
    runWorkbookMetadataEffect(this.metadataService.renameSheet(oldName, trimmedName));

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

  getSheetColumnVersion(sheetName: string, col: number): number {
    return this.sheetsByName.get(sheetName)?.columnVersions[col] ?? 0;
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

  withBatchedColumnVersionUpdates<T>(execute: () => T): T {
    if (this.batchedColumnVersionUpdates) {
      return execute();
    }
    const pending = new Map<number, Set<number>>();
    this.batchedColumnVersionUpdates = pending;
    try {
      return execute();
    } finally {
      this.batchedColumnVersionUpdates = null;
      pending.forEach((columns, sheetId) => {
        columns.forEach((col) => {
          this.bumpColumnVersion(sheetId, col);
        });
      });
    }
  }

  private bumpColumnVersionByCellIndex(cellIndex: number): void {
    const sheetId = this.cellStore.sheetIds[cellIndex]!;
    const col = this.cellStore.cols[cellIndex]!;
    const pending = this.batchedColumnVersionUpdates;
    if (pending) {
      let columns = pending.get(sheetId);
      if (!columns) {
        columns = new Set<number>();
        pending.set(sheetId, columns);
      }
      columns.add(col);
      return;
    }
    this.bumpColumnVersion(sheetId, col);
  }

  private bumpColumnVersion(sheetId: number, col: number): void {
    const sheet = this.getSheetById(sheetId);
    if (!sheet) {
      return;
    }
    sheet.columnVersions[col] = (sheet.columnVersions[col] ?? 0) + 1;
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

  pruneCellIfEmpty(index: number): boolean {
    const sheetId = this.cellStore.sheetIds[index];
    if (!sheetId) {
      return false;
    }
    const sheet = this.getSheetById(sheetId);
    if (!sheet) {
      return false;
    }
    const row = this.cellStore.rows[index];
    const col = this.cellStore.cols[index];
    if (row === undefined || col === undefined) {
      return false;
    }
    const value = this.cellStore.getValue(index, () => "");
    const flags = this.cellStore.flags[index] ?? 0;
    if (
      value.tag !== ValueTag.Empty ||
      this.cellFormats.has(index) ||
      (flags &
        (CellFlags.HasFormula |
          CellFlags.SpillChild |
          CellFlags.PivotOutput |
          CellFlags.PendingDelete)) !==
        0
    ) {
      return false;
    }
    this.cellKeyToIndex.delete(makeCellKey(sheet.id, row, col));
    sheet.grid.clear(row, col);
    return true;
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
    return storeCellStyle(this, style, (id) => this.bumpStyleId(id));
  }

  internCellStyle(style: Omit<WorkbookCellStyleRecord, "id">): WorkbookCellStyleRecord {
    return internWorkbookCellStyle(this, style, WorkbookStore.defaultStyleId);
  }

  getCellStyle(id: string | undefined): WorkbookCellStyleRecord | undefined {
    return readCellStyle(this, id, WorkbookStore.defaultStyleId);
  }

  listCellStyles(): WorkbookCellStyleRecord[] {
    return listWorkbookCellStyles(this);
  }

  upsertCellNumberFormat(format: CellNumberFormatRecord): WorkbookCellNumberFormatRecord {
    return storeCellNumberFormat(this, format, (id) => this.bumpFormatId(id));
  }

  internCellNumberFormat(format: string | CellNumberFormatRecord): WorkbookCellNumberFormatRecord {
    return internWorkbookCellNumberFormat(this, format, WorkbookStore.defaultFormatId);
  }

  getCellNumberFormat(id: string | undefined): WorkbookCellNumberFormatRecord | undefined {
    return readCellNumberFormat(this, id, WorkbookStore.defaultFormatId);
  }

  listCellNumberFormats(): WorkbookCellNumberFormatRecord[] {
    return listWorkbookCellNumberFormats(this);
  }

  setStyleRange(range: CellRangeRef, styleId: string): WorkbookStyleRangeRecord {
    return storeStyleRange(
      this,
      this.getOrCreateSheet(range.sheetName),
      range,
      styleId,
      WorkbookStore.defaultStyleId,
    );
  }

  listStyleRanges(sheetName: string): WorkbookStyleRangeRecord[] {
    return listWorkbookStyleRanges(this.getSheet(sheetName));
  }

  setStyleRanges(
    sheetName: string,
    ranges: readonly SheetStyleRangeSnapshot[],
  ): WorkbookStyleRangeRecord[] {
    return replaceStyleRanges(this, this.getOrCreateSheet(sheetName), ranges);
  }

  getStyleId(sheetName: string, row: number, col: number): string {
    return readStyleId(this.getSheet(sheetName), row, col, WorkbookStore.defaultStyleId);
  }

  setFormatRange(range: CellRangeRef, formatId: string): WorkbookFormatRangeRecord {
    return storeFormatRange(
      this,
      this.getOrCreateSheet(range.sheetName),
      range,
      formatId,
      WorkbookStore.defaultFormatId,
    );
  }

  listFormatRanges(sheetName: string): WorkbookFormatRangeRecord[] {
    return listWorkbookFormatRanges(this.getSheet(sheetName));
  }

  setFormatRanges(
    sheetName: string,
    ranges: readonly SheetFormatRangeSnapshot[],
  ): WorkbookFormatRangeRecord[] {
    return replaceFormatRanges(this, this.getOrCreateSheet(sheetName), ranges);
  }

  getRangeFormatId(sheetName: string, row: number, col: number): string {
    return readRangeFormatId(this.getSheet(sheetName), row, col, WorkbookStore.defaultFormatId);
  }

  setWorkbookProperty(key: string, value: LiteralInput): WorkbookPropertyRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.setWorkbookProperty(key, value));
  }

  getWorkbookProperty(key: string): WorkbookPropertyRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getWorkbookProperty(key));
  }

  listWorkbookProperties(): WorkbookPropertyRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listWorkbookProperties());
  }

  setCalculationSettings(
    settings: WorkbookCalculationSettingsSnapshot,
  ): WorkbookCalculationSettingsRecord {
    return runWorkbookMetadataEffect(this.metadataService.setCalculationSettings(settings));
  }

  getCalculationSettings(): WorkbookCalculationSettingsRecord {
    return runWorkbookMetadataEffect(this.metadataService.getCalculationSettings());
  }

  setVolatileContext(context: WorkbookVolatileContextSnapshot): WorkbookVolatileContextRecord {
    return runWorkbookMetadataEffect(this.metadataService.setVolatileContext(context));
  }

  getVolatileContext(): WorkbookVolatileContextRecord {
    return runWorkbookMetadataEffect(this.metadataService.getVolatileContext());
  }

  setDefinedName(name: string, value: WorkbookDefinedNameValueSnapshot): WorkbookDefinedNameRecord {
    return runWorkbookMetadataEffect(this.metadataService.setDefinedName(name, value));
  }

  getDefinedName(name: string): WorkbookDefinedNameRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getDefinedName(name));
  }

  deleteDefinedName(name: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteDefinedName(name));
  }

  listDefinedNames(): WorkbookDefinedNameRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listDefinedNames());
  }

  setTable(record: WorkbookTableSnapshot): WorkbookTableRecord {
    return runWorkbookMetadataEffect(this.metadataService.setTable(record));
  }

  getTable(name: string): WorkbookTableRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getTable(name));
  }

  deleteTable(name: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteTable(name));
  }

  listTables(): WorkbookTableRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listTables());
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

  snapshotRowAxisEntries(
    sheetName: string,
    start: number,
    count: number,
  ): WorkbookAxisEntrySnapshot[] {
    return this.snapshotAxisEntriesInRange(this.getSheet(sheetName), "row", start, count);
  }

  snapshotColumnAxisEntries(
    sheetName: string,
    start: number,
    count: number,
  ): WorkbookAxisEntrySnapshot[] {
    return this.snapshotAxisEntriesInRange(this.getSheet(sheetName), "column", start, count);
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
    return runWorkbookMetadataEffect(this.metadataService.setFreezePane(sheetName, rows, cols));
  }

  getFreezePane(sheetName: string): WorkbookFreezePaneRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getFreezePane(sheetName));
  }

  clearFreezePane(sheetName: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.clearFreezePane(sheetName));
  }

  setSheetProtection(record: WorkbookSheetProtectionSnapshot): WorkbookSheetProtectionRecord {
    return runWorkbookMetadataEffect(this.metadataService.setSheetProtection(record));
  }

  getSheetProtection(sheetName: string): WorkbookSheetProtectionRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getSheetProtection(sheetName));
  }

  clearSheetProtection(sheetName: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.clearSheetProtection(sheetName));
  }

  setFilter(sheetName: string, range: CellRangeRef): WorkbookFilterRecord {
    return runWorkbookMetadataEffect(this.metadataService.setFilter(sheetName, range));
  }

  getFilter(sheetName: string, range: CellRangeRef): WorkbookFilterRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getFilter(sheetName, range));
  }

  deleteFilter(sheetName: string, range: CellRangeRef): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteFilter(sheetName, range));
  }

  listFilters(sheetName: string): WorkbookFilterRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listFilters(sheetName));
  }

  setSort(
    sheetName: string,
    range: CellRangeRef,
    keys: readonly WorkbookSortKeyRecord[],
  ): WorkbookSortRecord {
    return runWorkbookMetadataEffect(this.metadataService.setSort(sheetName, range, keys));
  }

  getSort(sheetName: string, range: CellRangeRef): WorkbookSortRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getSort(sheetName, range));
  }

  deleteSort(sheetName: string, range: CellRangeRef): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteSort(sheetName, range));
  }

  listSorts(sheetName: string): WorkbookSortRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listSorts(sheetName));
  }

  setDataValidation(record: WorkbookDataValidationSnapshot): WorkbookDataValidationRecord {
    return runWorkbookMetadataEffect(this.metadataService.setDataValidation(record));
  }

  getDataValidation(
    sheetName: string,
    range: CellRangeRef,
  ): WorkbookDataValidationRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getDataValidation(sheetName, range));
  }

  deleteDataValidation(sheetName: string, range: CellRangeRef): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteDataValidation(sheetName, range));
  }

  listDataValidations(sheetName: string): WorkbookDataValidationRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listDataValidations(sheetName));
  }

  setConditionalFormat(record: WorkbookConditionalFormatSnapshot): WorkbookConditionalFormatRecord {
    return runWorkbookMetadataEffect(this.metadataService.setConditionalFormat(record));
  }

  getConditionalFormat(id: string): WorkbookConditionalFormatRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getConditionalFormat(id));
  }

  deleteConditionalFormat(id: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteConditionalFormat(id));
  }

  listConditionalFormats(sheetName: string): WorkbookConditionalFormatRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listConditionalFormats(sheetName));
  }

  setRangeProtection(record: WorkbookRangeProtectionSnapshot): WorkbookRangeProtectionRecord {
    return runWorkbookMetadataEffect(this.metadataService.setRangeProtection(record));
  }

  getRangeProtection(id: string): WorkbookRangeProtectionRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getRangeProtection(id));
  }

  deleteRangeProtection(id: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteRangeProtection(id));
  }

  listRangeProtections(sheetName: string): WorkbookRangeProtectionRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listRangeProtections(sheetName));
  }

  setCommentThread(record: WorkbookCommentThreadSnapshot): WorkbookCommentThreadRecord {
    return runWorkbookMetadataEffect(this.metadataService.setCommentThread(record));
  }

  getCommentThread(sheetName: string, address: string): WorkbookCommentThreadRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getCommentThread(sheetName, address));
  }

  deleteCommentThread(sheetName: string, address: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteCommentThread(sheetName, address));
  }

  listCommentThreads(sheetName: string): WorkbookCommentThreadRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listCommentThreads(sheetName));
  }

  setNote(record: WorkbookNoteSnapshot): WorkbookNoteRecord {
    return runWorkbookMetadataEffect(this.metadataService.setNote(record));
  }

  getNote(sheetName: string, address: string): WorkbookNoteRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getNote(sheetName, address));
  }

  deleteNote(sheetName: string, address: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteNote(sheetName, address));
  }

  listNotes(sheetName: string): WorkbookNoteRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listNotes(sheetName));
  }

  setSpill(sheetName: string, address: string, rows: number, cols: number): WorkbookSpillRecord {
    return runWorkbookMetadataEffect(this.metadataService.setSpill(sheetName, address, rows, cols));
  }

  getSpill(sheetName: string, address: string): WorkbookSpillRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getSpill(sheetName, address));
  }

  deleteSpill(sheetName: string, address: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteSpill(sheetName, address));
  }

  listSpills(): WorkbookSpillRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listSpills());
  }

  setPivot(record: WorkbookPivotSnapshot): WorkbookPivotRecord {
    return runWorkbookMetadataEffect(this.metadataService.setPivot(record));
  }

  getPivot(sheetName: string, address: string): WorkbookPivotRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getPivot(sheetName, address));
  }

  getPivotByKey(key: string): WorkbookPivotRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getPivotByKey(key));
  }

  deletePivot(sheetName: string, address: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deletePivot(sheetName, address));
  }

  listPivots(): WorkbookPivotRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listPivots());
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
    runWorkbookMetadataEffect(this.metadataService.reset());
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

  private createAxisEntry(axis: "row" | "column"): WorkbookAxisEntryRecord {
    return {
      id: axis === "row" ? `row-${this.nextRowAxisId++}` : `column-${this.nextColumnAxisId++}`,
      size: null,
      hidden: null,
    };
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

  private listAxisEntries(
    sheet: SheetRecord | undefined,
    axis: "row" | "column",
  ): WorkbookAxisEntrySnapshot[] {
    if (!sheet) {
      return [];
    }
    return listAxisEntries(axis === "row" ? sheet.rowAxis : sheet.columnAxis);
  }

  private materializeAxisEntries(
    sheet: SheetRecord,
    axis: "row" | "column",
    start: number,
    count: number,
  ): WorkbookAxisEntrySnapshot[] {
    return materializeAxisEntries(
      axis === "row" ? sheet.rowAxis : sheet.columnAxis,
      start,
      count,
      () => this.createAxisEntry(axis),
    );
  }

  private snapshotAxisEntriesInRange(
    sheet: SheetRecord | undefined,
    axis: "row" | "column",
    start: number,
    count: number,
  ): WorkbookAxisEntrySnapshot[] {
    if (!sheet) {
      return [];
    }
    return snapshotAxisEntriesInRange(
      axis === "row" ? sheet.rowAxis : sheet.columnAxis,
      start,
      count,
    );
  }

  private materializeAxisEntryRecords(
    sheet: SheetRecord,
    axis: "row" | "column",
    start: number,
    count: number,
  ): WorkbookAxisEntryRecord[] {
    return materializeAxisEntryRecords(
      axis === "row" ? sheet.rowAxis : sheet.columnAxis,
      start,
      count,
      () => this.createAxisEntry(axis),
    );
  }

  private spliceAxisEntries(
    sheet: SheetRecord,
    axis: "row" | "column",
    start: number,
    deleteCount: number,
    insertCount: number,
    entries?: readonly WorkbookAxisEntrySnapshot[],
  ): WorkbookAxisEntrySnapshot[] {
    return spliceAxisEntries(
      axis === "row" ? sheet.rowAxis : sheet.columnAxis,
      start,
      deleteCount,
      insertCount,
      () => this.createAxisEntry(axis),
      entries,
    );
  }

  private moveAxisEntries(
    sheet: SheetRecord,
    axis: "row" | "column",
    start: number,
    count: number,
    target: number,
  ): void {
    moveAxisEntries(axis === "row" ? sheet.rowAxis : sheet.columnAxis, start, count, target, () =>
      this.createAxisEntry(axis),
    );
  }

  private getAxisMetadataRecord(
    sheet: SheetRecord,
    axis: "row" | "column",
    sheetName: string,
    start: number,
    count: number,
  ): WorkbookAxisMetadataRecord | undefined {
    return getAxisMetadataRecord(
      axis === "row" ? sheet.rowAxis : sheet.columnAxis,
      sheetName,
      start,
      count,
    );
  }

  private syncAxisMetadataBucket(
    sheetName: string,
    sheet: SheetRecord,
    axis: "row" | "column",
    bucket: Map<string, WorkbookAxisMetadataRecord>,
  ): void {
    syncAxisMetadataBucket(bucket, sheetName, axis === "row" ? sheet.rowAxis : sheet.columnAxis);
  }
}

export function makeCellKey(sheetId: number, row: number, col: number): number {
  return sheetId * SHEET_STRIDE + row * MAX_COLS + col;
}
