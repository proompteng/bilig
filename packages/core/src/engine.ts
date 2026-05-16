import type {
  CellRangeRef,
  CellSnapshot,
  CellValue,
  DependencySnapshot,
  EngineEvent,
  ExplainCellSnapshot,
  LiteralInput,
  RecalcMetrics,
  SelectionState,
  SyncState,
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
  EngineFormulaSourceRefs,
} from './cell-mutations-at.js'
import { cloneEngineCounters, resetEngineCounters, type EngineCounters } from './perf/engine-counters.js'
import type { CommitOp, EngineReplicaSnapshot, EngineSyncClient } from './engine/runtime-state.js'
import { runEngineEffect, runEngineEffectPromise } from './engine/live.js'
import { SpreadsheetEngineWorkbookFacadeBase } from './engine/engine-workbook-facade-base.js'

export type {
  CommitOp,
  EngineReplicaSnapshot,
  EngineSyncClient,
  EngineSyncClientConnection,
  SpreadsheetEngineOptions,
} from './engine/runtime-state.js'
export { selectors } from './engine-selectors.js'

export class SpreadsheetEngine extends SpreadsheetEngineWorkbookFacadeBase {
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

  initializeFormulaSourcesAtNow(refs: EngineFormulaSourceRefs, potentialNewCells?: number): void {
    this.runtime.formulaInitialization.initializeFormulaSourcesAtNow(refs, potentialNewCells)
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
