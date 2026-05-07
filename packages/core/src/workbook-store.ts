import type {
  CellNumberFormatRecord,
  CellStyleRecord,
  CellRangeRef,
  LiteralInput,
  SheetFormatRangeSnapshot,
  SheetStyleRangeSnapshot,
  WorkbookAutoFilterSnapshot,
  WorkbookAxisEntrySnapshot,
  WorkbookCalculationSettingsSnapshot,
  WorkbookCommentThreadSnapshot,
  WorkbookChartSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookFreezePaneSnapshot,
  WorkbookImageSnapshot,
  WorkbookMacroPayloadSnapshot,
  WorkbookNoteSnapshot,
  WorkbookRangeProtectionSnapshot,
  WorkbookSheetProtectionSnapshot,
  WorkbookPivotSnapshot,
  WorkbookShapeSnapshot,
  WorkbookTableSnapshot,
  WorkbookVolatileContextSnapshot,
} from '@bilig/protocol'
import type { StructuralAxisTransform } from '@bilig/formula'
import type { SheetGridAxisRemapScope } from './sheet-grid.js'
import { CellStore } from './cell-store.js'
import type { StructuralTransaction } from './engine/structural-transaction.js'
import type { EngineCounters } from './perf/engine-counters.js'
import { createWorkbookMetadataService, runWorkbookMetadataEffect } from './workbook-metadata-service.js'
import {
  createWorkbookMetadataRecord,
  type WorkbookAxisMetadataRecord,
  type WorkbookCalculationSettingsRecord,
  type WorkbookCommentThreadRecord,
  type WorkbookChartRecord,
  type WorkbookConditionalFormatRecord,
  type WorkbookDataValidationRecord,
  type WorkbookCellNumberFormatRecord,
  type WorkbookCellStyleRecord,
  type WorkbookDefinedNameRecord,
  type WorkbookFilterRecord,
  type WorkbookFormatRangeRecord,
  type WorkbookFreezePaneRecord,
  type WorkbookImageRecord,
  type WorkbookMacroPayloadRecord,
  type WorkbookMergeRangeRecord,
  type WorkbookMetadataRecord,
  type WorkbookPivotRecord,
  type WorkbookPropertyRecord,
  type WorkbookRangeProtectionRecord,
  type WorkbookSheetProtectionRecord,
  type WorkbookShapeRecord,
  type WorkbookSortKeyRecord,
  type WorkbookSortRecord,
  type WorkbookSpillRecord,
  type WorkbookStyleRangeRecord,
  type WorkbookTableRecord,
  type WorkbookVolatileContextRecord,
  type WorkbookNoteRecord,
} from './workbook-metadata-types.js'
import {
  coalesceStyleRanges as coalesceWorkbookStyleRanges,
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
} from './workbook-style-format-store.js'
import { WORKBOOK_DEFAULT_FORMAT_ID, WORKBOOK_DEFAULT_STYLE_ID, ensureWorkbookDefaultStyleFormat } from './workbook-default-style-format.js'
import { createCellKeyIndexMap } from './workbook-cell-key-index.js'
import { WorkbookCellRecordStore, type EnsuredCell } from './workbook-cell-record-store.js'
import { WorkbookAxisEntryStore } from './workbook-axis-entry-store.js'
import { WorkbookColumnVersionStore } from './workbook-column-version-store.js'
import { WorkbookIdAllocator } from './workbook-id-allocator.js'
import type { SheetRecord } from './workbook-sheet-record.js'
import { WorkbookSheetRegistryStore } from './workbook-sheet-registry-store.js'
import { WorkbookStructuralCellStore } from './workbook-structural-cell-store.js'

export { makeCellKey, makeLogicalCellKey } from './workbook-cell-key-index.js'
export { normalizeDefinedName, normalizeWorkbookObjectName, imageKey, pivotKey, shapeKey } from './workbook-metadata-types.js'
export type { EnsuredCell } from './workbook-cell-record-store.js'
export type { SheetRecord } from './workbook-sheet-record.js'
export type {
  WorkbookAxisEntryRecord,
  WorkbookAxisMetadataRecord,
  WorkbookCalculationSettingsRecord,
  WorkbookCommentThreadRecord,
  WorkbookChartRecord,
  WorkbookConditionalFormatRecord,
  WorkbookDataValidationRecord,
  WorkbookCellNumberFormatRecord,
  WorkbookCellStyleRecord,
  WorkbookDefinedNameRecord,
  WorkbookFilterRecord,
  WorkbookFormatRangeRecord,
  WorkbookFreezePaneRecord,
  WorkbookImageRecord,
  WorkbookMacroPayloadRecord,
  WorkbookMergeRangeRecord,
  WorkbookMetadataRecord,
  WorkbookPivotRecord,
  WorkbookPropertyRecord,
  WorkbookRangeProtectionRecord,
  WorkbookSheetProtectionRecord,
  WorkbookShapeRecord,
  WorkbookSortKeyRecord,
  WorkbookSortRecord,
  WorkbookSpillRecord,
  WorkbookStyleRangeRecord,
  WorkbookTableRecord,
  WorkbookVolatileContextRecord,
  WorkbookNoteRecord,
} from './workbook-metadata-types.js'

export class WorkbookStore {
  static readonly defaultStyleId = WORKBOOK_DEFAULT_STYLE_ID
  static readonly defaultFormatId = WORKBOOK_DEFAULT_FORMAT_ID
  readonly cellStore = new CellStore()
  readonly sheetsByName = new Map<string, SheetRecord>()
  readonly sheetsById = new Map<number, SheetRecord>()
  readonly cellKeyToIndex: Map<number, number>
  readonly cellFormats = new Map<number, string>()
  readonly cellStyles = new Map<string, WorkbookCellStyleRecord>()
  readonly styleKeys = new Map<string, string>()
  readonly cellNumberFormats = new Map<string, WorkbookCellNumberFormatRecord>()
  readonly numberFormatKeys = new Map<string, string>()
  readonly metadata: WorkbookMetadataRecord = createWorkbookMetadataRecord()
  private readonly idAllocator = new WorkbookIdAllocator()
  private readonly metadataService = createWorkbookMetadataService(this.metadata)
  private readonly sheetRegistry: WorkbookSheetRegistryStore
  private readonly cellRecordStore: WorkbookCellRecordStore
  private readonly axisEntryStore: WorkbookAxisEntryStore
  private readonly columnVersionStore: WorkbookColumnVersionStore
  private readonly structuralCellStore: WorkbookStructuralCellStore
  workbookName: string

  constructor(
    workbookName = 'Workbook',
    private readonly counters?: EngineCounters,
  ) {
    this.workbookName = workbookName
    this.cellKeyToIndex = createCellKeyIndexMap((sheetId, row, col) => this.getCellIndexAt(sheetId, row, col))
    this.sheetRegistry = new WorkbookSheetRegistryStore({
      sheetsByName: this.sheetsByName,
      sheetsById: this.sheetsById,
      metadata: this.metadata,
      counters: this.counters,
      cellKeyToIndex: this.cellKeyToIndex,
      cellFormats: this.cellFormats,
      getCellPosition: (cellIndex) => this.getCellPosition(cellIndex),
      deleteSheetRecords: (sheetName) => {
        runWorkbookMetadataEffect(this.metadataService.deleteSheetRecords(sheetName))
      },
      renameSheetRecords: (oldName, nextName) => {
        runWorkbookMetadataEffect(this.metadataService.renameSheet(oldName, nextName))
      },
    })
    this.cellRecordStore = new WorkbookCellRecordStore({
      cellStore: this.cellStore,
      cellKeyToIndex: this.cellKeyToIndex,
      cellFormats: this.cellFormats,
      getSheet: (sheetName) => this.getSheet(sheetName),
      getOrCreateSheet: (sheetName) => this.getOrCreateSheet(sheetName),
      getSheetById: (sheetId) => this.getSheetById(sheetId),
      getSheetNameById: (sheetId) => this.getSheetNameById(sheetId),
      createLogicalAxisId: (axis) => this.createLogicalAxisId(axis),
    })
    this.axisEntryStore = new WorkbookAxisEntryStore({
      counters: this.counters,
      createAxisEntry: (axis) => this.idAllocator.createAxisEntry(axis),
    })
    this.columnVersionStore = new WorkbookColumnVersionStore({
      cellStore: this.cellStore,
      getSheetById: (sheetId) => this.getSheetById(sheetId),
    })
    this.structuralCellStore = new WorkbookStructuralCellStore({
      counters: this.counters,
      cellStore: this.cellStore,
      cellKeyToIndex: this.cellKeyToIndex,
      getSheet: (sheetName) => this.getSheet(sheetName),
      createLogicalAxisId: (axis) => this.createLogicalAxisId(axis),
    })
    this.cellStore.onSetValue = (index) => {
      this.notifyCellValueWritten(index)
    }
    ensureWorkbookDefaultStyleFormat(this)
  }

  createSheet(name: string, order = this.sheetsByName.size, id?: number): SheetRecord {
    return this.sheetRegistry.createSheet(name, order, id)
  }

  deleteSheet(name: string): void {
    this.sheetRegistry.deleteSheet(name)
  }

  renameSheet(oldName: string, nextName: string): SheetRecord | undefined {
    return this.sheetRegistry.renameSheet(oldName, nextName)
  }

  getSheet(name: string): SheetRecord | undefined {
    return this.sheetRegistry.getSheet(name)
  }

  getSheetColumnVersion(sheetName: string, col: number): number {
    return this.sheetRegistry.getSheetColumnVersion(sheetName, col)
  }

  getSheetStructureVersion(sheetName: string): number {
    return this.sheetRegistry.getSheetStructureVersion(sheetName)
  }

  getSheetById(id: number): SheetRecord | undefined {
    return this.sheetRegistry.getSheetById(id)
  }

  getOrCreateSheet(name: string): SheetRecord {
    return this.getSheet(name) ?? this.createSheet(name)
  }

  ensureCell(sheetName: string, address: string): number {
    return this.cellRecordStore.ensureCell(sheetName, address)
  }

  ensureCellRecord(sheetName: string, address: string): EnsuredCell {
    return this.cellRecordStore.ensureCellRecord(sheetName, address)
  }

  ensureCellAt(sheetId: number, row: number, col: number): EnsuredCell {
    return this.cellRecordStore.ensureCellAt(sheetId, row, col)
  }

  attachAllocatedCell(sheetId: number, row: number, col: number, cellIndex: number): void {
    this.cellRecordStore.attachAllocatedCell(sheetId, row, col, cellIndex)
  }

  ensureLogicalAxisId(sheetId: number, axis: 'row' | 'column', index: number): string {
    return this.cellRecordStore.ensureLogicalAxisId(sheetId, axis, index)
  }

  createLogicalAxisIdEnsurer(sheetId: number, axis: 'row' | 'column'): (index: number) => string {
    return this.cellRecordStore.createLogicalAxisIdEnsurer(sheetId, axis)
  }

  attachAllocatedCellWithLogicalAxisIds(sheetId: number, row: number, col: number, cellIndex: number, rowId: string, colId: string): void {
    this.cellRecordStore.attachAllocatedCellWithLogicalAxisIds(sheetId, row, col, cellIndex, rowId, colId)
  }

  withBatchedColumnVersionUpdates<T>(execute: () => T): T {
    return this.columnVersionStore.withBatchedColumnVersionUpdates(execute)
  }

  notifyCellValueWritten(cellIndex: number): void {
    this.columnVersionStore.notifyCellValueWritten(cellIndex)
  }

  notifyColumnsWritten(sheetId: number, columns: readonly number[] | Uint32Array): void {
    this.columnVersionStore.notifyColumnsWritten(sheetId, columns)
  }

  getCellIndex(sheetName: string, address: string): number | undefined {
    return this.cellRecordStore.getCellIndex(sheetName, address)
  }

  getCellIndexAt(sheetId: number, row: number, col: number): number | undefined {
    return this.cellRecordStore.getCellIndexAt(sheetId, row, col)
  }

  getSheetNameById(id: number): string {
    return this.sheetRegistry.getSheetNameById(id)
  }

  getAddress(index: number): string {
    return this.cellRecordStore.getAddress(index)
  }

  getQualifiedAddress(index: number): string {
    return this.cellRecordStore.getQualifiedAddress(index)
  }

  getCellPosition(index: number): { sheetId: number; row: number; col: number } | undefined {
    return this.cellRecordStore.getCellPosition(index)
  }

  getCellAxisIndex(index: number, axis: 'row' | 'column'): number | undefined {
    return this.cellRecordStore.getCellAxisIndex(index, axis)
  }

  detachCellIndex(index: number): boolean {
    return this.cellRecordStore.detachCellIndex(index)
  }

  pruneCellIfEmpty(index: number): boolean {
    return this.cellRecordStore.pruneCellIfEmpty(index)
  }

  setCellFormat(index: number, format: string | null | undefined): void {
    if (format === undefined || format === null || format === '') {
      this.cellFormats.delete(index)
      return
    }
    this.internCellNumberFormat(format)
    this.cellFormats.set(index, format)
  }

  getCellFormat(index: number): string | undefined {
    return this.cellFormats.get(index)
  }

  upsertCellStyle(style: CellStyleRecord): WorkbookCellStyleRecord {
    return storeCellStyle(this, style, (id) => this.idAllocator.bumpStyleId(id))
  }

  internCellStyle(style: Omit<WorkbookCellStyleRecord, 'id'>): WorkbookCellStyleRecord {
    return internWorkbookCellStyle(this, style, WorkbookStore.defaultStyleId)
  }

  getCellStyle(id: string | undefined): WorkbookCellStyleRecord | undefined {
    return readCellStyle(this, id, WorkbookStore.defaultStyleId)
  }

  listCellStyles(): WorkbookCellStyleRecord[] {
    return listWorkbookCellStyles(this)
  }

  upsertCellNumberFormat(format: CellNumberFormatRecord): WorkbookCellNumberFormatRecord {
    return storeCellNumberFormat(this, format, (id) => this.idAllocator.bumpFormatId(id))
  }

  internCellNumberFormat(format: string | CellNumberFormatRecord): WorkbookCellNumberFormatRecord {
    return internWorkbookCellNumberFormat(this, format, WorkbookStore.defaultFormatId)
  }

  getCellNumberFormat(id: string | undefined): WorkbookCellNumberFormatRecord | undefined {
    return readCellNumberFormat(this, id, WorkbookStore.defaultFormatId)
  }

  listCellNumberFormats(): WorkbookCellNumberFormatRecord[] {
    return listWorkbookCellNumberFormats(this)
  }

  setStyleRange(range: CellRangeRef, styleId: string): WorkbookStyleRangeRecord {
    return storeStyleRange(this, this.getOrCreateSheet(range.sheetName), range, styleId, WorkbookStore.defaultStyleId)
  }

  coalesceStyleRanges(sheetName: string): WorkbookStyleRangeRecord[] {
    const sheet = this.getSheet(sheetName)
    if (!sheet) {
      return []
    }
    return coalesceWorkbookStyleRanges(sheet)
  }

  listStyleRanges(sheetName: string): WorkbookStyleRangeRecord[] {
    return listWorkbookStyleRanges(this.getSheet(sheetName))
  }

  setStyleRanges(sheetName: string, ranges: readonly SheetStyleRangeSnapshot[]): WorkbookStyleRangeRecord[] {
    return replaceStyleRanges(this, this.getOrCreateSheet(sheetName), ranges)
  }

  getStyleId(sheetName: string, row: number, col: number): string {
    return readStyleId(this.getSheet(sheetName), row, col, WorkbookStore.defaultStyleId)
  }

  setFormatRange(range: CellRangeRef, formatId: string): WorkbookFormatRangeRecord {
    return storeFormatRange(this, this.getOrCreateSheet(range.sheetName), range, formatId, WorkbookStore.defaultFormatId)
  }

  listFormatRanges(sheetName: string): WorkbookFormatRangeRecord[] {
    return listWorkbookFormatRanges(this.getSheet(sheetName))
  }

  setFormatRanges(sheetName: string, ranges: readonly SheetFormatRangeSnapshot[]): WorkbookFormatRangeRecord[] {
    return replaceFormatRanges(this, this.getOrCreateSheet(sheetName), ranges)
  }

  getRangeFormatId(sheetName: string, row: number, col: number): string {
    return readRangeFormatId(this.getSheet(sheetName), row, col, WorkbookStore.defaultFormatId)
  }

  setWorkbookProperty(key: string, value: LiteralInput): WorkbookPropertyRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.setWorkbookProperty(key, value))
  }

  getWorkbookProperty(key: string): WorkbookPropertyRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getWorkbookProperty(key))
  }

  listWorkbookProperties(): WorkbookPropertyRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listWorkbookProperties())
  }

  setMacroPayload(record: WorkbookMacroPayloadSnapshot): WorkbookMacroPayloadRecord {
    return runWorkbookMetadataEffect(this.metadataService.setMacroPayload(record))
  }

  listMacroPayloads(): WorkbookMacroPayloadRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listMacroPayloads())
  }

  setCalculationSettings(settings: WorkbookCalculationSettingsSnapshot): WorkbookCalculationSettingsRecord {
    return runWorkbookMetadataEffect(this.metadataService.setCalculationSettings(settings))
  }

  getCalculationSettings(): WorkbookCalculationSettingsRecord {
    return runWorkbookMetadataEffect(this.metadataService.getCalculationSettings())
  }

  setVolatileContext(context: WorkbookVolatileContextSnapshot): WorkbookVolatileContextRecord {
    return runWorkbookMetadataEffect(this.metadataService.setVolatileContext(context))
  }

  getVolatileContext(): WorkbookVolatileContextRecord {
    return runWorkbookMetadataEffect(this.metadataService.getVolatileContext())
  }

  setDefinedName(name: string, value: WorkbookDefinedNameValueSnapshot, scopeSheetName?: string): WorkbookDefinedNameRecord {
    return runWorkbookMetadataEffect(this.metadataService.setDefinedName(name, value, scopeSheetName))
  }

  getDefinedName(name: string, scopeSheetName?: string): WorkbookDefinedNameRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getDefinedName(name, scopeSheetName))
  }

  deleteDefinedName(name: string, scopeSheetName?: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteDefinedName(name, scopeSheetName))
  }

  listDefinedNames(): WorkbookDefinedNameRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listDefinedNames())
  }

  setTable(record: WorkbookTableSnapshot): WorkbookTableRecord {
    return runWorkbookMetadataEffect(this.metadataService.setTable(record))
  }

  getTable(name: string): WorkbookTableRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getTable(name))
  }

  deleteTable(name: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteTable(name))
  }

  listTables(): WorkbookTableRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listTables())
  }

  setRowMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
  ): WorkbookAxisMetadataRecord | undefined {
    return this.axisEntryStore.setAxisMetadata(
      this.getOrCreateSheet(sheetName),
      'row',
      this.metadata.rowMetadata,
      sheetName,
      start,
      count,
      size,
      hidden,
    )
  }

  getRowMetadata(sheetName: string, start: number, count: number): WorkbookAxisMetadataRecord | undefined {
    const sheet = this.getSheet(sheetName)
    return sheet ? this.axisEntryStore.getAxisMetadataRecord(sheet, 'row', sheetName, start, count) : undefined
  }

  listRowMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.axisEntryStore.listAxisMetadata(this.getSheet(sheetName), this.metadata.rowMetadata, sheetName, 'row')
  }

  setColumnMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
  ): WorkbookAxisMetadataRecord | undefined {
    return this.axisEntryStore.setAxisMetadata(
      this.getOrCreateSheet(sheetName),
      'column',
      this.metadata.columnMetadata,
      sheetName,
      start,
      count,
      size,
      hidden,
    )
  }

  getColumnMetadata(sheetName: string, start: number, count: number): WorkbookAxisMetadataRecord | undefined {
    const sheet = this.getSheet(sheetName)
    return sheet ? this.axisEntryStore.getAxisMetadataRecord(sheet, 'column', sheetName, start, count) : undefined
  }

  listColumnMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.axisEntryStore.listAxisMetadata(this.getSheet(sheetName), this.metadata.columnMetadata, sheetName, 'column')
  }

  listRowAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return this.axisEntryStore.listAxisEntries(this.getSheet(sheetName), 'row')
  }

  listColumnAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return this.axisEntryStore.listAxisEntries(this.getSheet(sheetName), 'column')
  }

  snapshotRowAxisEntries(sheetName: string, start: number, count: number): WorkbookAxisEntrySnapshot[] {
    return this.axisEntryStore.snapshotAxisEntriesInRange(this.getSheet(sheetName), 'row', start, count)
  }

  snapshotColumnAxisEntries(sheetName: string, start: number, count: number): WorkbookAxisEntrySnapshot[] {
    return this.axisEntryStore.snapshotAxisEntriesInRange(this.getSheet(sheetName), 'column', start, count)
  }

  materializeRowAxisEntries(sheetName: string, start: number, count: number): WorkbookAxisEntrySnapshot[] {
    return this.axisEntryStore.materializeAxisEntries(this.getOrCreateSheet(sheetName), 'row', start, count)
  }

  materializeColumnAxisEntries(sheetName: string, start: number, count: number): WorkbookAxisEntrySnapshot[] {
    return this.axisEntryStore.materializeAxisEntries(this.getOrCreateSheet(sheetName), 'column', start, count)
  }

  insertRows(sheetName: string, start: number, count: number, entries?: readonly WorkbookAxisEntrySnapshot[]): void {
    const sheet = this.getOrCreateSheet(sheetName)
    this.axisEntryStore.spliceAxisEntries(sheet, 'row', start, 0, count, entries)
    this.bumpSheetStructureVersion(sheet)
  }

  deleteRows(sheetName: string, start: number, count: number): WorkbookAxisEntrySnapshot[] {
    const sheet = this.getOrCreateSheet(sheetName)
    const deleted = this.axisEntryStore.spliceAxisEntries(sheet, 'row', start, count, 0)
    this.bumpSheetStructureVersion(sheet)
    return deleted
  }

  moveRows(sheetName: string, start: number, count: number, target: number): void {
    const sheet = this.getOrCreateSheet(sheetName)
    this.axisEntryStore.moveAxisEntries(sheet, 'row', start, count, target)
    this.bumpSheetStructureVersion(sheet)
  }

  insertColumns(sheetName: string, start: number, count: number, entries?: readonly WorkbookAxisEntrySnapshot[]): void {
    const sheet = this.getOrCreateSheet(sheetName)
    this.axisEntryStore.spliceAxisEntries(sheet, 'column', start, 0, count, entries)
    this.bumpSheetStructureVersion(sheet)
  }

  deleteColumns(sheetName: string, start: number, count: number): WorkbookAxisEntrySnapshot[] {
    const sheet = this.getOrCreateSheet(sheetName)
    const deleted = this.axisEntryStore.spliceAxisEntries(sheet, 'column', start, count, 0)
    this.bumpSheetStructureVersion(sheet)
    return deleted
  }

  moveColumns(sheetName: string, start: number, count: number, target: number): void {
    const sheet = this.getOrCreateSheet(sheetName)
    this.axisEntryStore.moveAxisEntries(sheet, 'column', start, count, target)
    this.bumpSheetStructureVersion(sheet)
  }

  setFreezePane(
    sheetName: string,
    rows: number,
    cols: number,
    options?: Pick<WorkbookFreezePaneSnapshot, 'topLeftCell' | 'activePane'>,
  ): WorkbookFreezePaneRecord {
    return runWorkbookMetadataEffect(this.metadataService.setFreezePane(sheetName, rows, cols, options))
  }

  getFreezePane(sheetName: string): WorkbookFreezePaneRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getFreezePane(sheetName))
  }

  clearFreezePane(sheetName: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.clearFreezePane(sheetName))
  }

  setMergeRange(range: CellRangeRef): WorkbookMergeRangeRecord {
    return runWorkbookMetadataEffect(this.metadataService.setMergeRange(range))
  }

  setMergeRanges(sheetName: string, ranges: readonly CellRangeRef[]): WorkbookMergeRangeRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.setMergeRanges(sheetName, ranges))
  }

  getMergeRange(sheetName: string, address: string): WorkbookMergeRangeRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getMergeRange(sheetName, address))
  }

  getMergeRangeByRange(range: CellRangeRef): WorkbookMergeRangeRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getMergeRangeByRange(range))
  }

  clearMergeRanges(range: CellRangeRef): WorkbookMergeRangeRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.clearMergeRanges(range))
  }

  listMergeRanges(sheetName: string): WorkbookMergeRangeRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listMergeRanges(sheetName))
  }

  setSheetProtection(record: WorkbookSheetProtectionSnapshot): WorkbookSheetProtectionRecord {
    return runWorkbookMetadataEffect(this.metadataService.setSheetProtection(record))
  }

  getSheetProtection(sheetName: string): WorkbookSheetProtectionRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getSheetProtection(sheetName))
  }

  clearSheetProtection(sheetName: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.clearSheetProtection(sheetName))
  }

  setFilter(sheetName: string, range: WorkbookAutoFilterSnapshot): WorkbookFilterRecord {
    return runWorkbookMetadataEffect(this.metadataService.setFilter(sheetName, range))
  }

  getFilter(sheetName: string, range: CellRangeRef): WorkbookFilterRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getFilter(sheetName, range))
  }

  deleteFilter(sheetName: string, range: CellRangeRef): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteFilter(sheetName, range))
  }

  listFilters(sheetName: string): WorkbookFilterRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listFilters(sheetName))
  }

  setSort(sheetName: string, range: CellRangeRef, keys: readonly WorkbookSortKeyRecord[]): WorkbookSortRecord {
    return runWorkbookMetadataEffect(this.metadataService.setSort(sheetName, range, keys))
  }

  getSort(sheetName: string, range: CellRangeRef): WorkbookSortRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getSort(sheetName, range))
  }

  deleteSort(sheetName: string, range: CellRangeRef): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteSort(sheetName, range))
  }

  listSorts(sheetName: string): WorkbookSortRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listSorts(sheetName))
  }

  setDataValidation(record: WorkbookDataValidationSnapshot): WorkbookDataValidationRecord {
    return runWorkbookMetadataEffect(this.metadataService.setDataValidation(record))
  }

  getDataValidation(sheetName: string, range: CellRangeRef): WorkbookDataValidationRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getDataValidation(sheetName, range))
  }

  deleteDataValidation(sheetName: string, range: CellRangeRef): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteDataValidation(sheetName, range))
  }

  listDataValidations(sheetName: string): WorkbookDataValidationRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listDataValidations(sheetName))
  }

  setConditionalFormat(record: WorkbookConditionalFormatSnapshot): WorkbookConditionalFormatRecord {
    return runWorkbookMetadataEffect(this.metadataService.setConditionalFormat(record))
  }

  getConditionalFormat(id: string): WorkbookConditionalFormatRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getConditionalFormat(id))
  }

  deleteConditionalFormat(id: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteConditionalFormat(id))
  }

  listConditionalFormats(sheetName: string): WorkbookConditionalFormatRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listConditionalFormats(sheetName))
  }

  setRangeProtection(record: WorkbookRangeProtectionSnapshot): WorkbookRangeProtectionRecord {
    return runWorkbookMetadataEffect(this.metadataService.setRangeProtection(record))
  }

  getRangeProtection(id: string): WorkbookRangeProtectionRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getRangeProtection(id))
  }

  deleteRangeProtection(id: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteRangeProtection(id))
  }

  listRangeProtections(sheetName: string): WorkbookRangeProtectionRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listRangeProtections(sheetName))
  }

  setCommentThread(record: WorkbookCommentThreadSnapshot): WorkbookCommentThreadRecord {
    return runWorkbookMetadataEffect(this.metadataService.setCommentThread(record))
  }

  getCommentThread(sheetName: string, address: string): WorkbookCommentThreadRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getCommentThread(sheetName, address))
  }

  deleteCommentThread(sheetName: string, address: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteCommentThread(sheetName, address))
  }

  listCommentThreads(sheetName: string): WorkbookCommentThreadRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listCommentThreads(sheetName))
  }

  setNote(record: WorkbookNoteSnapshot): WorkbookNoteRecord {
    return runWorkbookMetadataEffect(this.metadataService.setNote(record))
  }

  getNote(sheetName: string, address: string): WorkbookNoteRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getNote(sheetName, address))
  }

  deleteNote(sheetName: string, address: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteNote(sheetName, address))
  }

  listNotes(sheetName: string): WorkbookNoteRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listNotes(sheetName))
  }

  setSpill(sheetName: string, address: string, rows: number, cols: number): WorkbookSpillRecord {
    return runWorkbookMetadataEffect(this.metadataService.setSpill(sheetName, address, rows, cols))
  }

  getSpill(sheetName: string, address: string): WorkbookSpillRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getSpill(sheetName, address))
  }

  deleteSpill(sheetName: string, address: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteSpill(sheetName, address))
  }

  listSpills(): WorkbookSpillRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listSpills())
  }

  setPivot(record: WorkbookPivotSnapshot): WorkbookPivotRecord {
    return runWorkbookMetadataEffect(this.metadataService.setPivot(record))
  }

  getPivot(sheetName: string, address: string): WorkbookPivotRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getPivot(sheetName, address))
  }

  getPivotByKey(key: string): WorkbookPivotRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getPivotByKey(key))
  }

  deletePivot(sheetName: string, address: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deletePivot(sheetName, address))
  }

  hasPivots(): boolean {
    return this.metadata.pivots.size > 0
  }

  listPivots(): WorkbookPivotRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listPivots())
  }

  setChart(record: WorkbookChartSnapshot): WorkbookChartRecord {
    return runWorkbookMetadataEffect(this.metadataService.setChart(record))
  }

  getChart(id: string): WorkbookChartRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getChart(id))
  }

  deleteChart(id: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteChart(id))
  }

  listCharts(): WorkbookChartRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listCharts())
  }

  setImage(record: WorkbookImageSnapshot): WorkbookImageRecord {
    return runWorkbookMetadataEffect(this.metadataService.setImage(record))
  }

  getImage(id: string): WorkbookImageRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getImage(id))
  }

  deleteImage(id: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteImage(id))
  }

  listImages(): WorkbookImageRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listImages())
  }

  setShape(record: WorkbookShapeSnapshot): WorkbookShapeRecord {
    return runWorkbookMetadataEffect(this.metadataService.setShape(record))
  }

  getShape(id: string): WorkbookShapeRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getShape(id))
  }

  deleteShape(id: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteShape(id))
  }

  listShapes(): WorkbookShapeRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listShapes())
  }

  remapSheetCells(
    sheetName: string,
    axis: 'row' | 'column',
    remapIndex: (index: number) => number | undefined,
    scope?: SheetGridAxisRemapScope,
  ): { changedCellIndices: number[]; removedCellIndices: number[] } {
    return this.structuralCellStore.remapSheetCells(sheetName, axis, remapIndex, scope)
  }

  planStructuralAxisTransform(sheetName: string, transform: StructuralAxisTransform): StructuralTransaction | undefined {
    return this.structuralCellStore.planStructuralAxisTransform(sheetName, transform)
  }

  applyPlannedStructuralTransaction(transaction: StructuralTransaction): StructuralTransaction | undefined {
    return this.structuralCellStore.applyPlannedStructuralTransaction(transaction)
  }

  applyStructuralAxisTransform(sheetName: string, transform: StructuralAxisTransform): StructuralTransaction | undefined {
    return this.structuralCellStore.applyStructuralAxisTransform(sheetName, transform)
  }

  reset(workbookName = 'Workbook'): void {
    this.workbookName = workbookName
    this.sheetRegistry.reset()
    this.cellKeyToIndex.clear()
    this.cellFormats.clear()
    this.cellStyles.clear()
    this.styleKeys.clear()
    this.cellNumberFormats.clear()
    this.numberFormatKeys.clear()
    runWorkbookMetadataEffect(this.metadataService.reset())
    this.idAllocator.reset()
    this.cellStore.reset()
    ensureWorkbookDefaultStyleFormat(this)
  }

  private bumpSheetStructureVersion(sheet: SheetRecord): void {
    sheet.structureVersion += 1
  }

  private createLogicalAxisId(axis: 'row' | 'column'): string {
    return this.idAllocator.createLogicalAxisId(axis)
  }
}
