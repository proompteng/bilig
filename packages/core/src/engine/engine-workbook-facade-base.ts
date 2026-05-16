import type {
  CellNumberFormatInput,
  CellNumberFormatRecord,
  CellRangeRef,
  CellSnapshot,
  CellStyleField,
  CellStylePatch,
  CellStyleRecord,
  LiteralInput,
  WorkbookAutoFilterSnapshot,
  WorkbookAxisEntrySnapshot,
  WorkbookCalculationSettingsSnapshot,
  WorkbookChartSnapshot,
  WorkbookCommentThreadSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookFreezePaneSnapshot,
  WorkbookImageSnapshot,
  WorkbookMergeRangeSnapshot,
  WorkbookNoteSnapshot,
  WorkbookPivotSnapshot,
  WorkbookRangeProtectionSnapshot,
  WorkbookSheetProtectionSnapshot,
  WorkbookSortSnapshot,
  WorkbookShapeSnapshot,
} from '@bilig/protocol'
import { calculationSettingsEqual, definedNameValuesEqual, normalizeWorkbookCalculationSettings } from '../engine-metadata-utils.js'
import { buildFormatClearOps, buildFormatPatchOps, buildStyleClearOps, buildStylePatchOps } from '../engine-range-format-ops.js'
import { hasEngineStructuralDeleteImpact } from './engine-structural-delete-impact.js'
import { upsertNumericDefinedNameFast as upsertNumericDefinedNameFastPath } from './engine-numeric-defined-name-fast-path.js'
import {
  buildSetConditionalFormatOps,
  buildSetDataValidationOps,
  buildSetFilterOps,
  buildSetPivotTableOps,
  buildSetRangeProtectionOps,
  buildSetSheetProtectionOps,
  buildSetSortOps,
} from './engine-workbook-metadata-ops.js'
import { normalizeEngineCommentThread, normalizeEngineNote, workbookObjectRecordEqual } from './engine-workbook-object-helpers.js'
import {
  buildDeleteChartOps,
  buildDeleteImageOps,
  buildDeleteShapeOps,
  buildDeleteTableOps,
  buildSetChartOps,
  buildSetImageOps,
  buildSetShapeOps,
  buildSetTableOps,
} from './engine-workbook-object-metadata-ops.js'
import { SpreadsheetEngineRuntimeBase } from './engine-runtime-base.js'
import type { PivotTableInput } from './runtime-state.js'
import { rangesIntersect } from '../workbook-merge-records.js'
import { canonicalWorkbookRangeRef } from '../workbook-range-records.js'
import {
  normalizeDefinedName,
  type WorkbookAxisMetadataRecord,
  type WorkbookCalculationSettingsRecord,
  type WorkbookCommentThreadRecord,
  type WorkbookConditionalFormatRecord,
  type WorkbookDataValidationRecord,
  type WorkbookDefinedNameRecord,
  type WorkbookFilterRecord,
  type WorkbookMergeRangeRecord,
  type WorkbookNoteRecord,
  type WorkbookPropertyRecord,
  type WorkbookRangeProtectionRecord,
  type WorkbookSheetProtectionRecord,
  type WorkbookSortRecord,
  type WorkbookSpillRecord,
  type WorkbookTableRecord,
  type WorkbookVolatileContextRecord,
} from '../workbook-store.js'

export abstract class SpreadsheetEngineWorkbookFacadeBase extends SpreadsheetEngineRuntimeBase {
  abstract override getCellByIndex(cellIndex: number): CellSnapshot

  setRangeNumberFormat(range: CellRangeRef, format: CellNumberFormatInput): void {
    const ops = buildFormatPatchOps(this.workbook, range, format)
    this.executeLocalTransaction(ops)
  }

  clearRangeNumberFormat(range: CellRangeRef): void {
    const ops = buildFormatClearOps(this.workbook, range)
    this.executeLocalTransaction(ops)
  }

  setRangeStyle(range: CellRangeRef, patch: CellStylePatch): void {
    const ops = buildStylePatchOps(this.workbook, range, patch)
    this.executeLocalTransaction(ops)
  }

  clearRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): void {
    const ops = buildStyleClearOps(this.workbook, range, fields)
    this.executeLocalTransaction(ops)
  }

  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined {
    return this.workbook.getCellStyle(styleId)
  }

  getCellNumberFormat(id: string | undefined): CellNumberFormatRecord | undefined {
    return this.workbook.getCellNumberFormat(id)
  }

  setDefinedName(name: string, value: WorkbookDefinedNameValueSnapshot): void {
    const normalizedName = normalizeDefinedName(name)
    const previous = this.workbook.getDefinedName(normalizedName)
    const trimmedName = name.trim()
    if (previous?.name === trimmedName && definedNameValuesEqual(previous.value, value)) {
      return
    }
    this.executeLocalTransaction([{ kind: 'upsertDefinedName', name: trimmedName, value }])
  }

  collectDefinedNameDependentFormulaCells(name: string): readonly number[] {
    return this.runtime.binding.collectFormulaCellsForDefinedNamesNow([normalizeDefinedName(name)])
  }

  upsertNumericDefinedNameFast(name: string, value: WorkbookDefinedNameValueSnapshot, numericValue: number): readonly number[] | null {
    return upsertNumericDefinedNameFastPath({
      workbook: this.workbook,
      formulas: this.formulas,
      replicaState: this.replicaState,
      entityVersions: this.entityVersions,
      undoStack: this.undoStack,
      redoStack: this.redoStack,
      collectDependentFormulaCells: (normalizedName) => this.runtime.binding.collectFormulaCellsForDefinedNamesNow([normalizedName]),
      name,
      value,
      numericValue,
    })
  }

  deleteDefinedName(name: string): boolean {
    if (!this.workbook.getDefinedName(name)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'deleteDefinedName', name }])
    return true
  }

  getDefinedName(name: string, scopeSheetName?: string): WorkbookDefinedNameRecord | undefined {
    return this.workbook.getDefinedName(name, scopeSheetName)
  }

  getDefinedNames(): WorkbookDefinedNameRecord[] {
    return this.workbook.listDefinedNames()
  }

  setWorkbookMetadata(key: string, value: LiteralInput): void {
    const existing = this.workbook.getWorkbookProperty(key)
    if (existing?.value === value || (existing === undefined && value === null)) {
      return
    }
    this.executeLocalTransaction([{ kind: 'setWorkbookMetadata', key, value }])
  }

  getWorkbookMetadata(key: string): WorkbookPropertyRecord | undefined {
    return this.workbook.getWorkbookProperty(key)
  }

  getWorkbookMetadataEntries(): WorkbookPropertyRecord[] {
    return this.workbook.listWorkbookProperties()
  }

  setCalculationSettings(settings: WorkbookCalculationSettingsSnapshot): void {
    const current = this.workbook.getCalculationSettings()
    const nextSettings = normalizeWorkbookCalculationSettings(settings, current)
    if (calculationSettingsEqual(current, nextSettings)) {
      return
    }
    this.executeLocalTransaction([{ kind: 'setCalculationSettings', settings: nextSettings }])
  }

  getCalculationSettings(): WorkbookCalculationSettingsRecord {
    return this.workbook.getCalculationSettings()
  }

  getVolatileContext(): WorkbookVolatileContextRecord {
    return this.workbook.getVolatileContext()
  }

  updateRowMetadata(sheetName: string, start: number, count: number, size: number | null, hidden: boolean | null): void {
    const existing = this.workbook.getRowMetadata(sheetName, start, count)
    if (existing?.size === size && existing.hidden === hidden) {
      return
    }
    if (existing === undefined && size === null && hidden === null) {
      return
    }
    this.executeLocalTransaction([{ kind: 'updateRowMetadata', sheetName, start, count, size, hidden }])
  }

  getRowMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.workbook.listRowMetadata(sheetName)
  }

  getRowAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return this.workbook.listRowAxisEntries(sheetName)
  }

  insertRows(sheetName: string, start: number, count: number): void {
    if (count <= 0) {
      return
    }
    this.executeLocalTransaction([{ kind: 'insertRows', sheetName, start, count }])
  }

  deleteRows(sheetName: string, start: number, count: number): void {
    if (count <= 0 || !this.hasStructuralDeleteImpact(sheetName, 'row', start, count)) {
      return
    }
    this.executeLocalTransaction([{ kind: 'deleteRows', sheetName, start, count }])
  }

  moveRows(sheetName: string, start: number, count: number, target: number): void {
    if (count <= 0 || start === target) {
      return
    }
    this.executeLocalTransaction([{ kind: 'moveRows', sheetName, start, count, target }])
  }

  updateColumnMetadata(sheetName: string, start: number, count: number, size: number | null, hidden: boolean | null): void {
    const existing = this.workbook.getColumnMetadata(sheetName, start, count)
    if (existing?.size === size && existing.hidden === hidden) {
      return
    }
    if (existing === undefined && size === null && hidden === null) {
      return
    }
    this.executeLocalTransaction([{ kind: 'updateColumnMetadata', sheetName, start, count, size, hidden }])
  }

  getColumnMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.workbook.listColumnMetadata(sheetName)
  }

  getColumnAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return this.workbook.listColumnAxisEntries(sheetName)
  }

  insertColumns(sheetName: string, start: number, count: number): void {
    if (count <= 0) {
      return
    }
    this.executeLocalTransaction([{ kind: 'insertColumns', sheetName, start, count }])
  }

  deleteColumns(sheetName: string, start: number, count: number): void {
    if (count <= 0 || !this.hasStructuralDeleteImpact(sheetName, 'column', start, count)) {
      return
    }
    this.executeLocalTransaction([{ kind: 'deleteColumns', sheetName, start, count }])
  }

  moveColumns(sheetName: string, start: number, count: number, target: number): void {
    if (count <= 0 || start === target) {
      return
    }
    this.executeLocalTransaction([{ kind: 'moveColumns', sheetName, start, count, target }])
  }

  setFreezePane(sheetName: string, rows: number, cols: number): void {
    const existing = this.workbook.getFreezePane(sheetName)
    if (existing?.rows === rows && existing.cols === cols) {
      return
    }
    this.executeLocalTransaction([{ kind: 'setFreezePane', sheetName, rows, cols }])
  }

  clearFreezePane(sheetName: string): boolean {
    if (!this.workbook.getFreezePane(sheetName)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'clearFreezePane', sheetName }])
    return true
  }

  getFreezePane(sheetName: string): WorkbookFreezePaneSnapshot | undefined {
    return this.workbook.getFreezePane(sheetName)
  }

  mergeCells(range: CellRangeRef): void {
    const normalized = canonicalWorkbookRangeRef(range)
    const existing = this.workbook.getMergeRangeByRange(normalized)
    if (existing) {
      return
    }
    this.executeLocalTransaction([{ kind: 'mergeCells', range: normalized }])
  }

  unmergeCells(range: CellRangeRef): boolean {
    const overlaps = this.workbook.listMergeRanges(range.sheetName).filter((record) => rangesIntersect(record, range))
    if (overlaps.length === 0) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'unmergeCells', range: canonicalWorkbookRangeRef(range) }])
    return true
  }

  getMergeRange(sheetName: string, address: string): WorkbookMergeRangeSnapshot | undefined {
    return this.workbook.getMergeRange(sheetName, address)
  }

  getMergeRangeByRange(range: CellRangeRef): WorkbookMergeRangeSnapshot | undefined {
    return this.workbook.getMergeRangeByRange(range)
  }

  listMergeRanges(sheetName: string): WorkbookMergeRangeRecord[] {
    return this.workbook.listMergeRanges(sheetName)
  }

  setSheetProtection(protection: WorkbookSheetProtectionSnapshot): void {
    this.executeLocalTransaction(buildSetSheetProtectionOps(this.workbook, protection) ?? [])
  }

  clearSheetProtection(sheetName: string): boolean {
    if (!this.workbook.getSheetProtection(sheetName)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'clearSheetProtection', sheetName }])
    return true
  }

  getSheetProtection(sheetName: string): WorkbookSheetProtectionRecord | undefined {
    return this.workbook.getSheetProtection(sheetName)
  }

  setFilter(sheetName: string, range: WorkbookAutoFilterSnapshot): void {
    this.executeLocalTransaction(buildSetFilterOps(this.workbook, sheetName, range) ?? [])
  }

  clearFilter(sheetName: string, range: CellRangeRef): boolean {
    if (!this.workbook.getFilter(sheetName, range)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'clearFilter', sheetName, range: { ...range } }])
    return true
  }

  getFilters(sheetName: string): WorkbookFilterRecord[] {
    return this.workbook.listFilters(sheetName)
  }

  setSort(sheetName: string, range: CellRangeRef, keys: WorkbookSortSnapshot['keys']): void {
    this.executeLocalTransaction(buildSetSortOps(this.workbook, sheetName, range, keys) ?? [])
  }

  clearSort(sheetName: string, range: CellRangeRef): boolean {
    if (!this.workbook.getSort(sheetName, range)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'clearSort', sheetName, range: { ...range } }])
    return true
  }

  getSorts(sheetName: string): WorkbookSortRecord[] {
    return this.workbook.listSorts(sheetName)
  }

  setDataValidation(validation: WorkbookDataValidationSnapshot): void {
    this.executeLocalTransaction(buildSetDataValidationOps(this.workbook, validation) ?? [])
  }

  clearDataValidation(sheetName: string, range: CellRangeRef): boolean {
    if (!this.workbook.getDataValidation(sheetName, range)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'clearDataValidation', sheetName, range: { ...range } }])
    return true
  }

  getDataValidation(sheetName: string, range: CellRangeRef): WorkbookDataValidationRecord | undefined {
    return this.workbook.getDataValidation(sheetName, range)
  }

  getDataValidations(sheetName: string): WorkbookDataValidationRecord[] {
    return this.workbook.listDataValidations(sheetName)
  }

  setConditionalFormat(format: WorkbookConditionalFormatSnapshot): void {
    this.executeLocalTransaction(buildSetConditionalFormatOps(this.workbook, format) ?? [])
  }

  deleteConditionalFormat(id: string): boolean {
    const existing = this.workbook.getConditionalFormat(id)
    if (!existing) {
      return false
    }
    this.executeLocalTransaction([
      {
        kind: 'deleteConditionalFormat',
        id: existing.id,
        sheetName: existing.range.sheetName,
      },
    ])
    return true
  }

  getConditionalFormat(id: string): WorkbookConditionalFormatRecord | undefined {
    return this.workbook.getConditionalFormat(id)
  }

  getConditionalFormats(sheetName: string): WorkbookConditionalFormatRecord[] {
    return this.workbook.listConditionalFormats(sheetName)
  }

  setRangeProtection(protection: WorkbookRangeProtectionSnapshot): void {
    this.executeLocalTransaction(buildSetRangeProtectionOps(this.workbook, protection) ?? [])
  }

  deleteRangeProtection(id: string): boolean {
    const existing = this.workbook.getRangeProtection(id)
    if (!existing) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'deleteRangeProtection', id: existing.id, sheetName: existing.range.sheetName }])
    return true
  }

  getRangeProtection(id: string): WorkbookRangeProtectionRecord | undefined {
    return this.workbook.getRangeProtection(id)
  }

  getRangeProtections(sheetName: string): WorkbookRangeProtectionRecord[] {
    return this.workbook.listRangeProtections(sheetName)
  }

  setCommentThread(thread: WorkbookCommentThreadSnapshot): void {
    const normalized = normalizeEngineCommentThread(thread)
    const existing = this.workbook.getCommentThread(normalized.sheetName, normalized.address)
    if (workbookObjectRecordEqual(existing, normalized)) {
      return
    }
    this.executeLocalTransaction([{ kind: 'upsertCommentThread', thread: normalized }])
  }

  deleteCommentThread(sheetName: string, address: string): boolean {
    if (!this.workbook.getCommentThread(sheetName, address)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'deleteCommentThread', sheetName, address }])
    return true
  }

  getCommentThread(sheetName: string, address: string): WorkbookCommentThreadRecord | undefined {
    return this.workbook.getCommentThread(sheetName, address)
  }

  getCommentThreads(sheetName: string): WorkbookCommentThreadRecord[] {
    return this.workbook.listCommentThreads(sheetName)
  }

  setNote(note: WorkbookNoteSnapshot): void {
    const normalized = normalizeEngineNote(note)
    const existing = this.workbook.getNote(normalized.sheetName, normalized.address)
    if (workbookObjectRecordEqual(existing, normalized)) {
      return
    }
    this.executeLocalTransaction([{ kind: 'upsertNote', note: normalized }])
  }

  deleteNote(sheetName: string, address: string): boolean {
    if (!this.workbook.getNote(sheetName, address)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'deleteNote', sheetName, address }])
    return true
  }

  getNote(sheetName: string, address: string): WorkbookNoteRecord | undefined {
    return this.workbook.getNote(sheetName, address)
  }

  getNotes(sheetName: string): WorkbookNoteRecord[] {
    return this.workbook.listNotes(sheetName)
  }

  setTable(table: WorkbookTableRecord): void {
    this.executeLocalTransaction(buildSetTableOps(this.workbook, table))
  }

  deleteTable(name: string): boolean {
    const ops = buildDeleteTableOps(this.workbook, name)
    if (!ops) {
      return false
    }
    this.executeLocalTransaction(ops)
    return true
  }

  getTable(name: string): WorkbookTableRecord | undefined {
    return this.workbook.getTable(name)
  }

  getTables(): WorkbookTableRecord[] {
    return this.workbook.listTables()
  }

  setSpillRange(sheetName: string, address: string, rows: number, cols: number): void {
    const existing = this.workbook.getSpill(sheetName, address)
    if (existing?.rows === rows && existing.cols === cols) {
      return
    }
    this.executeLocalTransaction([{ kind: 'upsertSpillRange', sheetName, address, rows, cols }])
  }

  deleteSpillRange(sheetName: string, address: string): boolean {
    if (!this.workbook.getSpill(sheetName, address)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'deleteSpillRange', sheetName, address }])
    return true
  }

  getSpillRanges(): WorkbookSpillRecord[] {
    return this.workbook.listSpills()
  }

  setPivotTable(sheetName: string, address: string, definition: PivotTableInput): void {
    this.executeLocalTransaction(buildSetPivotTableOps(sheetName, address, definition))
  }

  deletePivotTable(sheetName: string, address: string): boolean {
    if (!this.workbook.getPivot(sheetName, address)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'deletePivotTable', sheetName, address }])
    return true
  }

  getPivotTable(sheetName: string, address: string): WorkbookPivotSnapshot | undefined {
    return this.workbook.getPivot(sheetName, address)
  }

  getPivotTables(): WorkbookPivotSnapshot[] {
    return this.workbook.listPivots()
  }

  setChart(chart: WorkbookChartSnapshot): void {
    this.executeLocalTransaction(buildSetChartOps(this.workbook, chart))
  }

  deleteChart(id: string): boolean {
    const ops = buildDeleteChartOps(this.workbook, id)
    if (!ops) {
      return false
    }
    this.executeLocalTransaction(ops)
    return true
  }

  getChart(id: string): WorkbookChartSnapshot | undefined {
    return this.workbook.getChart(id)
  }

  getCharts(): WorkbookChartSnapshot[] {
    return this.workbook.listCharts()
  }

  setImage(image: WorkbookImageSnapshot): void {
    this.executeLocalTransaction(buildSetImageOps(this.workbook, image))
  }

  deleteImage(id: string): boolean {
    const ops = buildDeleteImageOps(this.workbook, id)
    if (!ops) {
      return false
    }
    this.executeLocalTransaction(ops)
    return true
  }

  getImage(id: string): WorkbookImageSnapshot | undefined {
    return this.workbook.getImage(id)
  }

  getImages(): WorkbookImageSnapshot[] {
    return this.workbook.listImages()
  }

  setShape(shape: WorkbookShapeSnapshot): void {
    this.executeLocalTransaction(buildSetShapeOps(this.workbook, shape))
  }

  deleteShape(id: string): boolean {
    const ops = buildDeleteShapeOps(this.workbook, id)
    if (!ops) {
      return false
    }
    this.executeLocalTransaction(ops)
    return true
  }

  getShape(id: string): WorkbookShapeSnapshot | undefined {
    return this.workbook.getShape(id)
  }

  getShapes(): WorkbookShapeSnapshot[] {
    return this.workbook.listShapes()
  }

  private hasStructuralDeleteImpact(sheetName: string, axis: 'row' | 'column', start: number, count: number): boolean {
    return hasEngineStructuralDeleteImpact({
      workbook: this.workbook,
      getCellByIndex: (cellIndex) => this.getCellByIndex(cellIndex),
      sheetName,
      axis,
      start,
      count,
    })
  }
}
