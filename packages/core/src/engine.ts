import type {
  CellNumberFormatInput,
  CellNumberFormatRecord,
  CellRangeRef,
  CellStyleField,
  CellStylePatch,
  CellStyleRecord,
  CellSnapshot,
  CellValue,
  DependencySnapshot,
  EngineEvent,
  ExplainCellSnapshot,
  LiteralInput,
  RecalcMetrics,
  SelectionState,
  SyncState,
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
  WorkbookSnapshot,
} from '@bilig/protocol'
import type { CsvParseOptions } from './csv.js'
import { formatAddress } from '@bilig/formula'
import type { EngineOp, EngineOpBatch } from '@bilig/workbook-domain'
import type {
  EngineCellMutationRef,
  EngineExistingLiteralCellMutationRef,
  EngineExistingNumericCellMutationRef,
  EngineExistingNumericCellMutationResult,
  EngineFormulaSourceRef,
} from './cell-mutations-at.js'
import { calculationSettingsEqual, definedNameValuesEqual, normalizeWorkbookCalculationSettings } from './engine-metadata-utils.js'
import { buildFormatClearOps, buildFormatPatchOps, buildStyleClearOps, buildStylePatchOps } from './engine-range-format-ops.js'
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
  type WorkbookPropertyRecord,
  type WorkbookRangeProtectionRecord,
  type WorkbookSheetProtectionRecord,
  type WorkbookSortRecord,
  type WorkbookSpillRecord,
  type WorkbookTableRecord,
  type WorkbookVolatileContextRecord,
  type WorkbookNoteRecord,
} from './workbook-store.js'
import { cloneEngineCounters, resetEngineCounters, type EngineCounters } from './perf/engine-counters.js'
import { canonicalWorkbookRangeRef } from './workbook-range-records.js'
import { rangesIntersect } from './workbook-merge-records.js'
import type { CommitOp, EngineReplicaSnapshot, EngineSyncClient, PivotTableInput } from './engine/runtime-state.js'
import { runEngineEffect, runEngineEffectPromise } from './engine/live.js'
import { SpreadsheetEngineRuntimeBase } from './engine/engine-runtime-base.js'
import { hasEngineStructuralDeleteImpact } from './engine/engine-structural-delete-impact.js'
import { upsertNumericDefinedNameFast as upsertNumericDefinedNameFastPath } from './engine/engine-numeric-defined-name-fast-path.js'
import {
  buildSetConditionalFormatOps,
  buildSetDataValidationOps,
  buildSetFilterOps,
  buildSetPivotTableOps,
  buildSetRangeProtectionOps,
  buildSetSheetProtectionOps,
  buildSetSortOps,
} from './engine/engine-workbook-metadata-ops.js'
import {
  cloneEngineTableRecord,
  normalizeEngineCommentThread,
  normalizeEngineNote,
  workbookChartsEqual,
  workbookImagesEqual,
  workbookObjectRecordEqual,
  workbookShapesEqual,
  workbookTablesEqual,
} from './engine/engine-workbook-object-helpers.js'

export type {
  CommitOp,
  EngineReplicaSnapshot,
  EngineSyncClient,
  EngineSyncClientConnection,
  SpreadsheetEngineOptions,
} from './engine/runtime-state.js'
export { selectors } from './engine-selectors.js'

export class SpreadsheetEngine extends SpreadsheetEngineRuntimeBase {
  async ready(): Promise<void> {
    await this.wasm.init()
  }

  subscribe(listener: (event: EngineEvent) => void): () => void {
    return runEngineEffect(this.runtime.events.subscribe(listener))
  }

  subscribeCell(sheetName: string, address: string, listener: () => void): () => void {
    return runEngineEffect(this.runtime.events.subscribeCell(sheetName, address, listener))
  }

  subscribeCells(sheetName: string, addresses: readonly string[], listener: () => void): () => void {
    return runEngineEffect(this.runtime.events.subscribeCells(sheetName, addresses, listener))
  }

  subscribeBatches(listener: (batch: EngineOpBatch) => void): () => void {
    return runEngineEffect(this.runtime.events.subscribeBatches(listener))
  }

  subscribeSelection(listener: () => void): () => void {
    return runEngineEffect(this.runtime.selection.subscribe(listener))
  }

  getSelectionState(): SelectionState {
    return runEngineEffect(this.runtime.selection.getSelectionState())
  }

  setSelection(
    sheetName: string,
    address: string | null,
    options: {
      anchorAddress?: string | null
      range?: { startAddress: string; endAddress: string } | null
      editMode?: SelectionState['editMode']
    } = {},
  ): void {
    runEngineEffect(this.runtime.selection.setSelection(sheetName, address, options))
  }

  getLastMetrics(): RecalcMetrics {
    return this.lastMetrics
  }

  getPerformanceCounters(): EngineCounters {
    return cloneEngineCounters(this.performanceCounters)
  }

  resetPerformanceCounters(): void {
    resetEngineCounters(this.performanceCounters)
  }

  setUseColumnIndexEnabled(enabled: boolean): void {
    this.state.setUseColumnIndex(enabled)
  }

  getSyncState(): SyncState {
    return this.syncState
  }

  async connectSyncClient(client: EngineSyncClient): Promise<void> {
    if (!this.state.trackReplicaVersions) {
      throw new Error('Sync is unavailable when trackReplicaVersions is disabled; construct the engine with trackReplicaVersions enabled.')
    }
    await runEngineEffectPromise(this.runtime.sync.connectClient(client))
  }

  async disconnectSyncClient(): Promise<void> {
    await runEngineEffectPromise(this.runtime.sync.disconnectClient())
  }

  createSheet(name: string): void {
    this.executeLocalTransaction([{ kind: 'upsertSheet', name, order: this.workbook.sheetsByName.size }])
  }

  createSheetForInitialization(name: string): number {
    return this.workbook.createSheet(name, this.workbook.sheetsByName.size).id
  }

  renameSheet(oldName: string, newName: string): void {
    const trimmedName = newName.trim()
    if (trimmedName.length === 0 || oldName === trimmedName) {
      return
    }
    if (this.workbook.getSheet(trimmedName)) {
      return
    }
    this.executeLocalTransaction([{ kind: 'renameSheet', oldName, newName: trimmedName }])
  }

  deleteSheet(name: string): void {
    this.executeLocalTransaction([{ kind: 'deleteSheet', name }])
  }

  setCellValue(sheetName: string, address: string, value: LiteralInput): CellValue {
    this.executeLocalTransaction([{ kind: 'setCellValue', sheetName, address, value }])
    return this.getCellValue(sheetName, address)
  }

  setCellValueAt(sheetId: number, row: number, col: number, value: LiteralInput): CellValue {
    const sheetName = this.workbook.getSheetById(sheetId)?.name
    if (!sheetName) {
      throw new Error(`Unknown sheet id: ${sheetId}`)
    }
    const address = formatAddress(row, col)
    this.runtime.mutation.executeLocalCellMutationsAtNow([{ sheetId, mutation: { kind: 'setCellValue', row, col, value } }], 1, {
      returnUndoOps: false,
    })
    return this.getCellValue(sheetName, address)
  }

  setCellFormula(sheetName: string, address: string, formula: string): CellValue {
    if (this.getCell(sheetName, address).formula === formula) {
      return this.getCellValue(sheetName, address)
    }
    this.executeLocalTransaction([{ kind: 'setCellFormula', sheetName, address, formula }])
    return this.getCellValue(sheetName, address)
  }

  setCellFormulaAt(sheetId: number, row: number, col: number, formula: string): CellValue {
    const sheetName = this.workbook.getSheetById(sheetId)?.name
    if (!sheetName) {
      throw new Error(`Unknown sheet id: ${sheetId}`)
    }
    const address = formatAddress(row, col)
    if (this.getCell(sheetName, address).formula === formula) {
      return this.getCellValue(sheetName, address)
    }
    this.runtime.mutation.executeLocalCellMutationsAtNow([{ sheetId, mutation: { kind: 'setCellFormula', row, col, formula } }], 1, {
      returnUndoOps: false,
    })
    return this.getCellValue(sheetName, address)
  }

  setCellFormat(sheetName: string, address: string, format: string | null): void {
    this.executeLocalTransaction([{ kind: 'setCellFormat', sheetName, address, format }])
  }

  clearCellAt(sheetId: number, row: number, col: number): void {
    this.runtime.mutation.executeLocalCellMutationsAtNow([{ sheetId, mutation: { kind: 'clearCell', row, col } }], 0, {
      returnUndoOps: false,
    })
  }

  applyCellMutationsAt(refs: readonly EngineCellMutationRef[], potentialNewCells?: number): readonly EngineOp[] | null {
    return this.runtime.mutation.executeLocalCellMutationsAtNow(refs, potentialNewCells)
  }

  applyCellMutationsAtWithOptions(
    refs: readonly EngineCellMutationRef[],
    options: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
      returnUndoOps?: boolean
      reuseRefs?: boolean
    } = {},
  ): readonly EngineOp[] | null {
    return this.runtime.mutation.applyCellMutationsAtNow(refs, options)
  }

  tryApplyExistingNumericCellMutationAt(request: EngineExistingNumericCellMutationRef): EngineExistingNumericCellMutationResult | null {
    return this.runtime.mutation.executeLocalExistingNumericCellMutationAtNow(request, { returnUndoOps: false })
  }

  tryApplyExistingLiteralCellMutationAt(request: EngineExistingLiteralCellMutationRef): EngineExistingNumericCellMutationResult | null {
    return this.runtime.mutation.executeLocalExistingLiteralCellMutationAtNow(request, { returnUndoOps: false })
  }

  initializeCellFormulasAt(refs: readonly EngineCellMutationRef[], potentialNewCells?: number): void {
    runEngineEffect(this.runtime.formulaInitialization.initializeCellFormulasAt(refs, potentialNewCells))
  }

  initializeCellFormulasAtNow(refs: readonly EngineCellMutationRef[], potentialNewCells?: number): void {
    this.runtime.formulaInitialization.initializeCellFormulasAtNow(refs, potentialNewCells)
  }

  initializeFormulaSourcesAtNow(refs: readonly EngineFormulaSourceRef[], potentialNewCells?: number): void {
    this.runtime.formulaInitialization.initializeFormulaSourcesAtNow(refs, potentialNewCells)
  }

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

  recalculateNow(): number[] {
    return runEngineEffect(this.runtime.recalc.recalculateNow())
  }

  recalculateDifferential(): { js: CellSnapshot[]; wasm: CellSnapshot[]; drift: string[] } {
    return runEngineEffect(this.runtime.recalc.recalculateDifferential())
  }

  recalculateDirty(
    dirtyRegions: Array<{
      sheetName: string
      rowStart: number
      rowEnd: number
      colStart: number
      colEnd: number
    }>,
  ): number[] {
    return runEngineEffect(this.runtime.recalc.recalculateDirty(dirtyRegions))
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
    const existing = this.workbook.getTable(table.name)
    if (workbookTablesEqual(existing, table)) {
      return
    }
    this.executeLocalTransaction([
      {
        kind: 'upsertTable',
        table: cloneEngineTableRecord(table),
      },
    ])
  }

  deleteTable(name: string): boolean {
    if (!this.workbook.getTable(name)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'deleteTable', name }])
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
    const existing = this.workbook.getChart(chart.id)
    if (workbookChartsEqual(existing, chart)) {
      return
    }
    this.executeLocalTransaction([{ kind: 'upsertChart', chart: structuredClone(chart) }])
  }

  deleteChart(id: string): boolean {
    if (!this.workbook.getChart(id)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'deleteChart', id }])
    return true
  }

  getChart(id: string): WorkbookChartSnapshot | undefined {
    return this.workbook.getChart(id)
  }

  getCharts(): WorkbookChartSnapshot[] {
    return this.workbook.listCharts()
  }

  setImage(image: WorkbookImageSnapshot): void {
    const existing = this.workbook.getImage(image.id)
    if (workbookImagesEqual(existing, image)) {
      return
    }
    this.executeLocalTransaction([{ kind: 'upsertImage', image: structuredClone(image) }])
  }

  deleteImage(id: string): boolean {
    if (!this.workbook.getImage(id)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'deleteImage', id }])
    return true
  }

  getImage(id: string): WorkbookImageSnapshot | undefined {
    return this.workbook.getImage(id)
  }

  getImages(): WorkbookImageSnapshot[] {
    return this.workbook.listImages()
  }

  setShape(shape: WorkbookShapeSnapshot): void {
    const existing = this.workbook.getShape(shape.id)
    if (workbookShapesEqual(existing, shape)) {
      return
    }
    this.executeLocalTransaction([{ kind: 'upsertShape', shape: structuredClone(shape) }])
  }

  deleteShape(id: string): boolean {
    if (!this.workbook.getShape(id)) {
      return false
    }
    this.executeLocalTransaction([{ kind: 'deleteShape', id }])
    return true
  }

  getShape(id: string): WorkbookShapeSnapshot | undefined {
    return this.workbook.getShape(id)
  }

  getShapes(): WorkbookShapeSnapshot[] {
    return this.workbook.listShapes()
  }

  clearCell(sheetName: string, address: string): void {
    this.executeLocalTransaction([{ kind: 'clearCell', sheetName, address }])
  }

  setRangeValues(range: CellRangeRef, values: readonly (readonly LiteralInput[])[]): void {
    runEngineEffect(this.runtime.mutation.setRangeValues(range, values))
  }

  setRangeFormulas(range: CellRangeRef, formulas: readonly (readonly string[])[]): void {
    runEngineEffect(this.runtime.mutation.setRangeFormulas(range, formulas))
  }

  clearRange(range: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.clearRange(range))
  }

  fillRange(source: CellRangeRef, target: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.fillRange(source, target))
  }

  copyRange(source: CellRangeRef, target: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.copyRange(source, target))
  }

  moveRange(source: CellRangeRef, target: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.moveRange(source, target))
  }

  pasteRange(source: CellRangeRef, target: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.copyRange(source, target))
  }

  undo(): boolean {
    return runEngineEffect(this.runtime.history.undo())
  }

  redo(): boolean {
    return runEngineEffect(this.runtime.history.redo())
  }

  canUndo(): boolean {
    return this.undoStack.length > 0
  }

  canRedo(): boolean {
    return this.redoStack.length > 0
  }

  clearHistory(): void {
    this.undoStack.length = 0
    this.redoStack.length = 0
  }

  exportSheetCsv(sheetName: string): string {
    return runEngineEffect(this.runtime.read.exportSheetCsv(sheetName))
  }

  importSheetCsv(sheetName: string, csv: string, options?: CsvParseOptions): void {
    runEngineEffect(this.runtime.mutation.importSheetCsv(sheetName, csv, options))
    if (csv.includes('=')) {
      // CSV import applies one bulk mutation batch. A second full recalc settles
      // formulas whose imported ranges include other formulas introduced later in the same batch.
      this.recalculateNow()
      this.recalculateNow()
    }
  }

  getCellValue(sheetName: string, address: string): CellValue {
    return runEngineEffect(this.runtime.read.getCellValue(sheetName, address))
  }

  getRangeValues(range: CellRangeRef): CellValue[][] {
    return runEngineEffect(this.runtime.read.getRangeValues(range))
  }

  getCell(sheetName: string, address: string): CellSnapshot {
    return runEngineEffect(this.runtime.read.getCell(sheetName, address))
  }

  getCellByIndex(cellIndex: number): CellSnapshot {
    return runEngineEffect(this.runtime.read.getCellByIndex(cellIndex))
  }

  getDependencies(sheetName: string, address: string): DependencySnapshot {
    return runEngineEffect(this.runtime.read.getDependencies(sheetName, address))
  }

  getDependents(sheetName: string, address: string): DependencySnapshot {
    return runEngineEffect(this.runtime.read.getDependents(sheetName, address))
  }

  explainCell(sheetName: string, address: string): ExplainCellSnapshot {
    return runEngineEffect(this.runtime.read.explainCell(sheetName, address))
  }

  exportSnapshot(): WorkbookSnapshot {
    return runEngineEffect(this.runtime.snapshot.exportWorkbook())
  }

  importSnapshot(snapshot: WorkbookSnapshot): void {
    runEngineEffect(this.runtime.snapshot.importWorkbook(snapshot))
  }

  exportReplicaSnapshot(): EngineReplicaSnapshot {
    return runEngineEffect(this.runtime.snapshot.exportReplica())
  }

  importReplicaSnapshot(snapshot: EngineReplicaSnapshot): void {
    runEngineEffect(this.runtime.snapshot.importReplica(snapshot))
  }

  renderCommit(ops: CommitOp[]): void {
    runEngineEffect(this.runtime.mutation.renderCommit(ops))
  }

  applyRemoteBatch(batch: EngineOpBatch): boolean {
    return runEngineEffect(this.runtime.sync.applyRemoteBatch(batch))
  }

  captureUndoOps<T>(mutate: () => T): {
    result: T
    undoOps: readonly EngineOp[] | null
  } {
    return runEngineEffect(this.runtime.mutation.captureUndoOps(mutate))
  }

  applyOps(
    ops: readonly EngineOp[],
    options: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
      trusted?: boolean
    } = {},
  ): readonly EngineOp[] | null {
    return this.runtime.mutation.applyOpsNow(ops, options)
  }
}
