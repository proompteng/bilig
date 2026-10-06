import {
  createWorkbookFacade,
  WORKBOOK_TABLE_COMMAND_IDS,
  WORKBOOK_TABLES_FEATURE_ID,
  type CommitOp,
  type SpreadsheetEngine,
} from '@bilig/core'
import {
  createWorkbookAgentCommandBundle,
  buildWorkbookAgentPreview,
  isWorkbookAgentCommandBundle,
  type WorkbookAgentCommand,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentPreviewSummary,
} from '@bilig/agent-api'
import type { EngineOp, EngineOpBatch, WorkbookCommandReceipt } from '@bilig/workbook'
import type { AuthoritativeWorkbookEventRecord } from '@bilig/zero-sync'
import {
  type CellRangeRef,
  type CellNumberFormatInput,
  type CellStyleField,
  type CellStylePatch,
  type CellSnapshot,
  type EngineEvent,
  type RecalcMetrics,
  type SyncState,
  type WorkbookSnapshot,
  formatCellDisplayValue,
  isCellRangeRef,
} from '@bilig/protocol'
import {
  buildUnavailableWorkbookViewWindow,
  buildWorkbookViewWindowFromViewportPatch,
  type RenderTileDeltaSubscription,
  type ViewportPatch,
  type ViewportPatchSubscription,
  type WorkbookViewWindow,
  type WorkbookViewWindowSheetIdentity,
  type WorkbookViewWindowSubscription,
} from '@bilig/worker-transport'
import { isPendingWorkbookMutationList, type PendingWorkbookMutation, type PendingWorkbookMutationInput } from './workbook-sync.js'
import { WorkerRuntimeMutationJournal } from './worker-runtime-mutation-journal.js'
import { acquireProjectionEngine } from './worker-runtime-projection-engine.js'
import { WorkerRuntimeProjectionCommands } from './worker-runtime-projection-commands.js'
import { createProjectionEngineFromState } from './worker-runtime-engine-state.js'
import type {
  InstallAuthoritativeSnapshotInput,
  InstallBenchmarkCorpusResult,
  WorkbookWorkerBootstrapOptions,
  WorkbookWorkerBootstrapResult,
  WorkbookWorkerStateSnapshot,
} from './worker-runtime-state.js'
import { WorkerRuntimeStateCoordinator } from './worker-runtime-state-coordinator.js'
import { WorkerRuntimeSnapshotCaches } from './worker-runtime-snapshot-caches.js'
import { WorkerRuntimeAuthoritativeStateCoordinator } from './worker-runtime-authoritative-state-coordinator.js'
import { WorkerRuntimeWorkbookDeltaPublisher } from './worker-runtime-delta-publisher.js'
import { applyWorkerRuntimeLocalHistoryChange } from './worker-runtime-local-history.js'
import { WorkerRuntimeRenderTileDeltaPublisher } from './worker-runtime-render-tile-subscription.js'
import type { WorkerRuntimeRenderTileDiagnostics } from './worker-runtime-render-tile-subscription.js'
import {
  collectSheetViewportImpacts,
  normalizeViewport,
  type SheetViewportImpact,
  type ViewportSubscriptionState,
  type WorkerEngine,
} from './worker-runtime-support.js'
import { AUTOFIT_CHAR_WIDTH, AUTOFIT_PADDING, MAX_COLUMN_WIDTH, MIN_COLUMN_WIDTH } from './worker-runtime-viewport.js'
import {
  WorkerViewportPatchPublisher,
  createEmptyCellSnapshot,
  type ViewportPatchBroadcastReason,
} from './worker-runtime-viewport-publisher.js'
import {
  collectProjectionOverlayScopeFromEngineEvents,
  mergeProjectionOverlayScopes,
  type ProjectionOverlayScope,
} from './worker-local-overlay.js'
import { exportWorkerRuntimeSnapshot } from './worker-runtime-export-snapshot.js'
import { applyAuthoritativeWorkbookEvents } from './worker-runtime-authoritative-events.js'
import { prepareAuthoritativeSnapshotProjection } from './worker-runtime-authoritative-snapshot.js'
import { collectWorkbookSelectionSummary } from './workbook-selection-summary.js'
export type {
  InstallAuthoritativeSnapshotInput,
  InstallBenchmarkCorpusResult,
  WorkbookFailedPendingMutationSnapshot,
  WorkbookPendingMutationSummarySnapshot,
  WorkbookWorkerBootstrapOptions,
  WorkbookWorkerBootstrapResult,
  WorkbookWorkerStateSnapshot,
} from './worker-runtime-state.js'

function tableAgentCommandFromOp(op: EngineOp): WorkbookAgentCommand[] {
  if (op.kind === 'upsertTable') {
    return [{ kind: 'upsertTable', table: structuredClone(op.table) }]
  }
  if (op.kind === 'deleteTable') {
    return [{ kind: 'deleteTable', name: op.name }]
  }
  return []
}

function createTransientViewportSubscriptionState(subscription: ViewportPatchSubscription): ViewportSubscriptionState {
  return {
    subscription: normalizeViewport(subscription),
    listener: () => undefined,
    nextVersion: 1,
    knownStyleIds: new Set(),
    lastStyleSignatures: new Map(),
    lastCellSignatures: new Map(),
    lastColumnSignatures: new Map(),
    lastRowSignatures: new Map(),
    lastMergeSignatures: new Map(),
  }
}

export class WorkbookWorkerRuntime {
  [method: string]: unknown
  private engine: (SpreadsheetEngine & WorkerEngine) | null = null
  private projectionEnginePromise: Promise<SpreadsheetEngine & WorkerEngine> | null = null
  private projectionBuildVersion = 0
  private bootstrapOptions: WorkbookWorkerBootstrapOptions | null = null
  private engineSubscription: (() => void) | null = null
  private projectionOverlayScope: ProjectionOverlayScope | null = null
  private readonly workbookDeltaPublisher = new WorkerRuntimeWorkbookDeltaPublisher()
  private readonly renderTileDeltaPublisher = new WorkerRuntimeRenderTileDeltaPublisher()
  private readonly snapshotCaches = new WorkerRuntimeSnapshotCaches()
  private readonly authoritativeState = new WorkerRuntimeAuthoritativeStateCoordinator({
    snapshotCaches: this.snapshotCaches,
    getDocumentId: () => this.requireBootstrapOptions().documentId,
    getReplicaId: () => this.requireBootstrapOptions().replicaId,
    listPendingMutations: () => this.mutationJournal.listPendingMutations(),
  })
  private readonly viewportPatchPublisher = new WorkerViewportPatchPublisher({
    buildPatch: (state, event, metrics, authoritativeRevision, sheetImpact) =>
      this.buildViewportPatch(state, event, metrics, authoritativeRevision, sheetImpact),
    getAuthoritativeRevision: () => this.authoritativeState.getRevision(),
    getCurrentMetrics: () => this.getCurrentMetrics(),
    getProjectionEngine: () => this.requireEngine(),
    hasProjectionEngine: () => this.engine !== null,
  })
  private readonly projectionCommands = new WorkerRuntimeProjectionCommands({
    invalidateProjectionCache: () => this.invalidateProjectionCache(),
    getProjectionEngine: () => this.getProjectionEngine(),
    getCell: (sheetName, address) => this.getCell(sheetName, address),
    minColumnWidth: MIN_COLUMN_WIDTH,
    maxColumnWidth: MAX_COLUMN_WIDTH,
    autofitCharWidth: AUTOFIT_CHAR_WIDTH,
    autofitPadding: AUTOFIT_PADDING,
    formatCellDisplayValue,
  })
  private readonly mutationJournal: WorkerRuntimeMutationJournal = new WorkerRuntimeMutationJournal({
    getDocumentId: () => this.requireBootstrapOptions().documentId,
    getClientMutationScope: () => this.requireBootstrapOptions().replicaId,
    getAuthoritativeRevision: () => this.authoritativeState.getRevision(),
    getProjectionEngine: () => this.getProjectionEngine(),
    invalidateProjectionCache: () => this.invalidateProjectionCache(),
  })
  private readonly stateCoordinator = new WorkerRuntimeStateCoordinator({
    getEngine: () => this.engine,
    getAuthoritativeRevision: () => this.authoritativeState.getRevision(),
    buildPendingMutationSummary: () => this.mutationJournal.buildPendingMutationSummary(),
  })

  async ready(): Promise<void> {
    await (await this.getProjectionEngine()).ready()
  }

  dispose(): void {
    this.cleanup()
  }

  async bootstrap(options: WorkbookWorkerBootstrapOptions): Promise<WorkbookWorkerBootstrapResult> {
    this.cleanup()
    this.bootstrapOptions = options
    this.snapshotCaches.reset()
    this.mutationJournal.reset()
    this.projectionOverlayScope = null
    this.authoritativeState.reset()
    const restoredJournalEntries = isPendingWorkbookMutationList(options.mutationJournalEntries) ? options.mutationJournalEntries : []
    if (
      restoredJournalEntries.length > 0 ||
      (typeof options.nextPendingMutationSeq === 'number' &&
        Number.isSafeInteger(options.nextPendingMutationSeq) &&
        options.nextPendingMutationSeq > 0)
    ) {
      this.mutationJournal.restoreFromBootstrap({
        mutationJournalEntries: restoredJournalEntries,
        nextPendingMutationSeq:
          typeof options.nextPendingMutationSeq === 'number' &&
          Number.isSafeInteger(options.nextPendingMutationSeq) &&
          options.nextPendingMutationSeq > 0
            ? options.nextPendingMutationSeq
            : Math.max(...restoredJournalEntries.map((mutation) => mutation.localSeq), 0) + 1,
      })
    }
    const { engine, overlayScope } = await createProjectionEngineFromState({
      workbookName: options.documentId,
      replicaId: options.replicaId,
      snapshot: null,
      replica: null,
      pendingMutations: this.mutationJournal.listPendingMutations(),
    })
    this.projectionOverlayScope = overlayScope
    this.installEngine(engine)

    return {
      runtimeState: this.getRuntimeState(),
      restoredFromPersistence: false,
      requiresAuthoritativeHydrate: true,
      localPersistenceMode: this.stateCoordinator.getLocalPersistenceMode(),
    }
  }

  getRuntimeState(): WorkbookWorkerStateSnapshot {
    return this.stateCoordinator.getRuntimeState(() => this.requireEngine())
  }

  getAuthoritativeRevision(): number {
    return this.authoritativeState.getRevision()
  }

  async installAuthoritativeSnapshot(input: InstallAuthoritativeSnapshotInput): Promise<WorkbookWorkerStateSnapshot> {
    const { snapshot, authoritativeRevision, mode } = input
    if (mode !== 'bootstrap' && mode !== 'reconcile') {
      throw new Error('Invalid authoritative snapshot install mode')
    }
    return this.installAuthoritativeSnapshotInternal(snapshot, authoritativeRevision, mode)
  }

  async installBenchmarkCorpus(corpusId: string): Promise<InstallBenchmarkCorpusResult> {
    const benchmarks = await import('./worker-runtime-benchmark-corpus.js')
    if (!benchmarks.isWorkbookBenchmarkCorpusId(corpusId)) {
      throw new Error(`Unknown benchmark corpus ${corpusId}`)
    }
    const corpus = benchmarks.buildWorkbookBenchmarkCorpus(corpusId)
    await this.installAuthoritativeSnapshotInternal(corpus.snapshot, 0, 'bootstrap')
    return {
      id: corpus.id,
      materializedCellCount: corpus.materializedCellCount,
      primaryViewport: corpus.primaryViewport,
    }
  }

  private async installAuthoritativeSnapshotInternal(
    snapshot: WorkbookSnapshot,
    authoritativeRevision: number,
    mode: InstallAuthoritativeSnapshotInput['mode'],
  ): Promise<WorkbookWorkerStateSnapshot> {
    this.projectionBuildVersion += 1
    this.projectionEnginePromise = null
    const options = this.requireBootstrapOptions()
    this.authoritativeState.acceptSnapshotRevision(authoritativeRevision, mode)

    const prepared = await prepareAuthoritativeSnapshotProjection({
      documentId: options.documentId,
      replicaId: options.replicaId,
      snapshot,
      mode,
      pendingMutations: this.mutationJournal.listPendingMutations(),
    })
    this.authoritativeState.installPreparedState(prepared)
    if (prepared.shouldMarkPendingMutationsRebased) {
      await this.markRemainingJournalMutationsRebased()
    }
    return await this.installAuthoritativeProjectionEngine(prepared.projectionEngine, prepared.projectionOverlayScope)
  }

  private async installAuthoritativeProjectionEngine(
    engine: SpreadsheetEngine & WorkerEngine,
    overlayScope: ProjectionOverlayScope | null,
  ): Promise<WorkbookWorkerStateSnapshot> {
    this.projectionOverlayScope = overlayScope
    this.installEngine(engine)
    this.snapshotCaches.invalidateProjectionSnapshot()
    this.broadcastViewportPatches(null, engine.getLastMetrics(), 'authoritative-snapshot')
    return this.getRuntimeState()
  }

  async applyAuthoritativeEvents(
    events: readonly AuthoritativeWorkbookEventRecord[],
    authoritativeRevision: number,
  ): Promise<WorkbookWorkerStateSnapshot> {
    this.projectionBuildVersion += 1
    this.projectionEnginePromise = null
    const authoritativeEngine = await this.getAuthoritativeEngine()
    const { absorbedMutationIds } = applyAuthoritativeWorkbookEvents(authoritativeEngine, events)
    this.mutationJournal.ackAbsorbedMutations(absorbedMutationIds)
    this.authoritativeState.acceptEventRevision(authoritativeRevision)
    this.authoritativeState.invalidateCachedState()
    const { engine, overlayScope } = await this.rebuildProjectionEngine()
    if (events.length > 0 && this.mutationJournal.getPendingMutationCount() > 0) {
      await this.markRemainingJournalMutationsRebased()
    }
    this.projectionOverlayScope = overlayScope
    this.installEngine(engine)
    this.snapshotCaches.invalidateProjectionSnapshot()
    this.broadcastViewportPatches(null, engine.getLastMetrics(), 'authoritative-events')
    return this.getRuntimeState()
  }

  setExternalSyncState(syncState: SyncState | null): WorkbookWorkerStateSnapshot {
    return this.stateCoordinator.setExternalSyncState(syncState)
  }

  exportSnapshot(): WorkbookSnapshot {
    return exportWorkerRuntimeSnapshot({
      exportProjectionSnapshot: this.engine
        ? () => this.snapshotCaches.getProjectionSnapshot(() => this.requireEngine().exportSnapshot())
        : null,
      getReadyAuthoritativeSnapshot: () => this.snapshotCaches.getReadyAuthoritativeSnapshot(),
      pendingMutationCount: this.mutationJournal.getPendingMutationCount(),
    })
  }

  async previewAgentCommandBundle(bundle: WorkbookAgentCommandBundle): Promise<WorkbookAgentPreviewSummary> {
    if (!isWorkbookAgentCommandBundle(bundle)) {
      throw new Error('Invalid workbook agent command bundle')
    }
    return await buildWorkbookAgentPreview({
      snapshot: this.engine ? this.exportSnapshot() : (await this.getProjectionEngine()).exportSnapshot(),
      replicaId: this.requireBootstrapOptions().replicaId,
      bundle,
    })
  }

  async getSelectionSummary(range: CellRangeRef) {
    if (!isCellRangeRef(range)) throw new Error('Selection summary requires a workbook range')
    return collectWorkbookSelectionSummary(await this.getProjectionEngine(), range)
  }

  async buildCreateTableCommandBundle(range: CellRangeRef): Promise<WorkbookAgentCommandBundle> {
    if (!isCellRangeRef(range)) {
      throw new Error('Create table requires a workbook range')
    }
    const engine = await this.getProjectionEngine()
    const facade = createWorkbookFacade(engine)
    try {
      const receipt = await facade.selection(range).createTable().preview()
      return this.buildTableFeatureAgentBundle(receipt, 'toolbar')
    } finally {
      facade.dispose()
    }
  }

  async buildResizeTableCommandBundle(name: string, range: CellRangeRef): Promise<WorkbookAgentCommandBundle> {
    if (typeof name !== 'string' || name.trim() === '') {
      throw new Error('Resize table requires a table name')
    }
    if (!isCellRangeRef(range)) {
      throw new Error('Resize table requires a workbook range')
    }
    const engine = await this.getProjectionEngine()
    const facade = createWorkbookFacade(engine)
    try {
      const receipt = await facade
        .command({
          featureId: WORKBOOK_TABLES_FEATURE_ID,
          commandId: WORKBOOK_TABLE_COMMAND_IDS.resize,
          category: 'operation',
          input: {
            name,
            range: {
              sheetName: range.sheetName,
              startAddress: range.startAddress,
              endAddress: range.endAddress,
            },
          },
        })
        .preview()
      return this.buildTableFeatureAgentBundle(receipt, 'tables-panel')
    } finally {
      facade.dispose()
    }
  }

  async buildDeleteTableCommandBundle(name: string): Promise<WorkbookAgentCommandBundle> {
    if (typeof name !== 'string' || name.trim() === '') {
      throw new Error('Delete table requires a table name')
    }
    const engine = await this.getProjectionEngine()
    const facade = createWorkbookFacade(engine)
    try {
      const receipt = await facade
        .command({
          featureId: WORKBOOK_TABLES_FEATURE_ID,
          commandId: WORKBOOK_TABLE_COMMAND_IDS.delete,
          category: 'operation',
          input: { name },
        })
        .preview()
      return this.buildTableFeatureAgentBundle(receipt, 'tables-panel')
    } finally {
      facade.dispose()
    }
  }

  async materializeProjectionEngine(): Promise<void> {
    const hadInstalledEngine = this.engine !== null
    await this.getProjectionEngine()
    if (!hadInstalledEngine && this.engine) {
      this.broadcastViewportPatches(null, this.getCurrentMetrics(), 'projection-materialized')
    }
  }

  listPendingMutations(): PendingWorkbookMutation[] {
    return this.mutationJournal.listPendingMutations()
  }

  listMutationJournalEntries(): PendingWorkbookMutation[] {
    return this.mutationJournal.listMutationJournalEntries()
  }

  async enqueuePendingMutation(input: PendingWorkbookMutationInput): Promise<PendingWorkbookMutation> {
    const mutation = await this.mutationJournal.enqueuePendingMutation(input)
    return mutation
  }

  async markPendingMutationSubmitted(id: string): Promise<void> {
    await this.mutationJournal.markPendingMutationSubmitted(id)
  }

  async ackPendingMutation(id: string): Promise<void> {
    await this.mutationJournal.ackPendingMutation(id)
  }

  async recordPendingMutationAttempt(id: string): Promise<void> {
    await this.mutationJournal.recordPendingMutationAttempt(id)
  }

  async markPendingMutationFailed(id: string, failureMessage: string): Promise<void> {
    await this.mutationJournal.markPendingMutationFailed(id, failureMessage)
  }

  async retryPendingMutation(id: string): Promise<void> {
    await this.mutationJournal.retryPendingMutation(id)
  }

  getCell(sheetName: string, address: string): CellSnapshot {
    const engine = this.engine
    if (!engine) {
      return createEmptyCellSnapshot(sheetName, address)
    }
    if (!engine.workbook.getSheet(sheetName)) {
      return createEmptyCellSnapshot(sheetName, address)
    }
    return engine.getCell(sheetName, address)
  }

  async setCellValue(sheetName: string, address: string, value: CellSnapshot['input']): Promise<CellSnapshot> {
    return await this.projectionCommands.setCellValue(sheetName, address, value)
  }

  async setCellFormula(sheetName: string, address: string, formula: string): Promise<CellSnapshot> {
    return await this.projectionCommands.setCellFormula(sheetName, address, formula)
  }

  async setRangeStyle(range: CellRangeRef, patch: CellStylePatch): Promise<void> {
    await this.projectionCommands.setRangeStyle(range, patch)
  }

  async clearRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): Promise<void> {
    await this.projectionCommands.clearRangeStyle(range, fields)
  }

  async setRangeNumberFormat(range: CellRangeRef, format: CellNumberFormatInput): Promise<void> {
    await this.projectionCommands.setRangeNumberFormat(range, format)
  }

  async clearRangeNumberFormat(range: CellRangeRef): Promise<void> {
    await this.projectionCommands.clearRangeNumberFormat(range)
  }

  async clearRange(range: CellRangeRef): Promise<void> {
    await this.projectionCommands.clearRange(range)
  }

  async clearCell(sheetName: string, address: string): Promise<CellSnapshot> {
    return await this.projectionCommands.clearCell(sheetName, address)
  }

  async undoLocalChange(): Promise<boolean> {
    return await this.applyLocalHistoryChange('undo')
  }

  async redoLocalChange(): Promise<boolean> {
    return await this.applyLocalHistoryChange('redo')
  }

  async renderCommit(ops: CommitOp[]): Promise<void> {
    await this.projectionCommands.renderCommit(ops)
  }

  async fillRange(source: CellRangeRef, target: CellRangeRef): Promise<void> {
    await this.projectionCommands.fillRange(source, target)
  }

  async copyRange(source: CellRangeRef, target: CellRangeRef): Promise<void> {
    await this.projectionCommands.copyRange(source, target)
  }

  async moveRange(source: CellRangeRef, target: CellRangeRef): Promise<void> {
    await this.projectionCommands.moveRange(source, target)
  }

  async updateRowMetadata(
    sheetName: string,
    startRow: number,
    count: number,
    height: number | null,
    hidden: boolean | null,
  ): Promise<void> {
    await this.projectionCommands.updateRowMetadata(sheetName, startRow, count, height, hidden)
  }

  async updateColumnMetadata(
    sheetName: string,
    startCol: number,
    count: number,
    width: number | null,
    hidden: boolean | null,
  ): Promise<number | null> {
    return await this.projectionCommands.updateColumnMetadata(sheetName, startCol, count, width, hidden)
  }

  async updateColumnWidth(sheetName: string, columnIndex: number, width: number): Promise<number> {
    return await this.projectionCommands.updateColumnWidth(sheetName, columnIndex, width)
  }

  async setFreezePane(sheetName: string, rows: number, cols: number): Promise<void> {
    await this.projectionCommands.setFreezePane(sheetName, rows, cols)
  }

  async mergeCells(range: CellRangeRef): Promise<void> {
    await this.projectionCommands.mergeCells(range)
  }

  async unmergeCells(range: CellRangeRef): Promise<void> {
    await this.projectionCommands.unmergeCells(range)
  }

  async autofitColumn(sheetName: string, columnIndex: number): Promise<number> {
    return await this.projectionCommands.autofitColumn(sheetName, columnIndex)
  }

  subscribe(listener: (event: EngineEvent) => void): () => void {
    return this.requireEngine().subscribe(listener)
  }

  subscribeBatches(listener: (batch: EngineOpBatch) => void): () => void {
    return this.requireEngine().subscribeBatches(listener)
  }

  subscribeViewportPatches(subscription: ViewportPatchSubscription, listener: (patch: Uint8Array) => void): () => void {
    return this.viewportPatchPublisher.subscribe(subscription, listener)
  }

  getWorkbookViewWindow(subscription: WorkbookViewWindowSubscription): WorkbookViewWindow {
    const engine = this.requireEngine()
    const authoritativeRevision = this.authoritativeState.getRevision()
    const sheetName = engine.workbook.getSheetNameById(subscription.sheetId)
    if (sheetName.length === 0) {
      return buildUnavailableWorkbookViewWindow({
        status: 'missing-sheet',
        reason: 'sheet-id-not-found',
        request: subscription,
        authoritativeRevision,
      })
    }
    const sheet = engine.workbook.getSheet(sheetName)
    if (!sheet) {
      return buildUnavailableWorkbookViewWindow({
        status: 'missing-sheet',
        reason: 'sheet-name-not-found',
        request: subscription,
        authoritativeRevision,
      })
    }
    const identity: WorkbookViewWindowSheetIdentity = {
      sheetId: sheet.id ?? subscription.sheetId,
      sheetOrdinal: sheet.order,
      sheetName: sheet.name,
    }
    if (subscription.sheetName !== undefined && subscription.sheetName !== sheet.name) {
      return buildUnavailableWorkbookViewWindow({
        status: 'identity-mismatch',
        reason: 'sheet-name-does-not-match-sheet-id',
        request: subscription,
        sheet: identity,
        authoritativeRevision,
      })
    }
    if (sheet.id !== undefined && sheet.id !== subscription.sheetId) {
      return buildUnavailableWorkbookViewWindow({
        status: 'identity-mismatch',
        reason: 'resolved-sheet-id-does-not-match-request',
        request: subscription,
        sheet: identity,
        authoritativeRevision,
      })
    }

    const state = createTransientViewportSubscriptionState({
      sheetName: sheet.name,
      rowStart: subscription.rowStart,
      rowEnd: subscription.rowEnd,
      colStart: subscription.colStart,
      colEnd: subscription.colEnd,
      initialPatch: 'full',
    })
    const patch = this.viewportPatchPublisher.buildPatch(state, null, this.getCurrentMetrics(), authoritativeRevision, null)
    return buildWorkbookViewWindowFromViewportPatch({
      patch,
      request: {
        ...subscription,
        sheetName: subscription.sheetName ?? sheet.name,
        sheetOrdinal: subscription.sheetOrdinal ?? sheet.order,
      },
      sheet: identity,
    })
  }

  subscribeWorkbookDeltas(listener: (delta: Uint8Array) => void): () => void {
    this.requireEngine()
    return this.workbookDeltaPublisher.subscribe(listener)
  }

  subscribeRenderTileDeltas(subscription: RenderTileDeltaSubscription, listener: (delta: Uint8Array) => void): () => void {
    return this.renderTileDeltaPublisher.subscribe({
      getProjectionEngine: () => this.getProjectionEngine(),
      listener,
      subscription,
    })
  }

  getRenderTileDiagnostics(): WorkerRuntimeRenderTileDiagnostics {
    return this.renderTileDeltaPublisher.getDiagnostics()
  }

  private cleanup(): void {
    this.projectionBuildVersion += 1
    this.engineSubscription?.()
    this.engineSubscription = null
    this.projectionEnginePromise = null
    this.mutationJournal.reset()
    this.viewportPatchPublisher.reset()
    this.stateCoordinator.reset()
    this.snapshotCaches.reset()
    this.workbookDeltaPublisher.reset()
    this.renderTileDeltaPublisher.reset()
    this.authoritativeState.reset()
    this.projectionOverlayScope = null
    this.engine = null
  }

  private requireBootstrapOptions(): WorkbookWorkerBootstrapOptions {
    if (!this.bootstrapOptions) {
      throw new Error('Workbook worker runtime has not been bootstrapped')
    }
    return this.bootstrapOptions
  }

  private requireEngine(): SpreadsheetEngine & WorkerEngine {
    if (!this.engine) {
      throw new Error('Workbook worker runtime has not been bootstrapped')
    }
    return this.engine
  }

  private invalidateProjectionCache(): void {
    this.projectionBuildVersion += 1
    this.projectionEnginePromise = null
  }

  private async markRemainingJournalMutationsRebased(rebasedAtUnixMs = Date.now()): Promise<void> {
    await this.mutationJournal.markRemainingJournalMutationsRebased(rebasedAtUnixMs)
  }

  private updateRuntimeStateFromEngine(engine: SpreadsheetEngine & WorkerEngine = this.requireEngine()): WorkbookWorkerStateSnapshot {
    return this.stateCoordinator.updateRuntimeStateFromEngine(engine)
  }

  private getCurrentMetrics(): RecalcMetrics {
    return this.stateCoordinator.getCurrentMetrics()
  }

  private async applyLocalHistoryChange(direction: 'undo' | 'redo'): Promise<boolean> {
    return await applyWorkerRuntimeLocalHistoryChange(
      {
        getProjectionEngine: () => this.getProjectionEngine(),
        invalidateProjectionCache: () => this.invalidateProjectionCache(),
        updateRuntimeStateFromEngine: (engine) => {
          this.updateRuntimeStateFromEngine(engine)
        },
      },
      direction,
    )
  }

  private buildTableFeatureAgentBundle(receipt: WorkbookCommandReceipt, source: 'toolbar' | 'tables-panel'): WorkbookAgentCommandBundle {
    if (receipt.status === 'rejected') {
      throw new Error(receipt.message ?? 'Table command preview was rejected')
    }
    const commands = (receipt.previewOps ?? []).flatMap((op) => tableAgentCommandFromOp(op))
    if (commands.length === 0) {
      throw new Error(receipt.message ?? 'Table command preview did not produce an applicable table mutation')
    }
    const options = this.requireBootstrapOptions()
    const now = Date.now()
    return createWorkbookAgentCommandBundle({
      documentId: options.documentId,
      threadId: `workbook-${source}`,
      turnId: `${source}:${String(now)}`,
      goalText: receipt.message ?? 'Apply workbook table command',
      baseRevision: this.authoritativeState.getRevision(),
      context: null,
      commands,
      now,
    })
  }

  private async getAuthoritativeEngine(): Promise<SpreadsheetEngine> {
    return await this.authoritativeState.getEngine()
  }

  private async rebuildProjectionEngine(): Promise<{
    engine: SpreadsheetEngine
    overlayScope: ProjectionOverlayScope | null
  }> {
    return await this.authoritativeState.rebuildProjectionEngine()
  }

  private async getProjectionEngine(): Promise<SpreadsheetEngine & WorkerEngine> {
    return await acquireProjectionEngine({
      getInstalledEngine: () => this.engine,
      getProjectionEnginePromise: () => this.projectionEnginePromise,
      getProjectionBuildVersion: () => this.projectionBuildVersion,
      rebuildProjectionEngine: () => this.rebuildProjectionEngine(),
      setProjectionOverlayScope: (overlayScope) => {
        this.projectionOverlayScope = overlayScope
      },
      installEngine: (engine) => {
        this.installEngine(engine)
      },
      setProjectionEnginePromise: (promise) => {
        this.projectionEnginePromise = promise
      },
      requireInstalledEngine: () => this.requireEngine(),
    })
  }

  private installEngine(engine: SpreadsheetEngine & WorkerEngine): void {
    this.engineSubscription?.()
    this.engine = engine
    this.updateRuntimeStateFromEngine(engine)
    this.engineSubscription = engine.subscribe((event) => {
      this.snapshotCaches.invalidateProjectionSnapshot()
      if (this.mutationJournal.getPendingMutationCount() > 0) {
        this.projectionOverlayScope = mergeProjectionOverlayScopes(
          this.projectionOverlayScope,
          collectProjectionOverlayScopeFromEngineEvents(engine, [event]),
        )
      }
      this.updateRuntimeStateFromEngine(engine)
      this.broadcastViewportPatches(event)
      this.workbookDeltaPublisher.publish(engine, event)
    })
  }

  private broadcastViewportPatches(
    event: EngineEvent | null,
    metrics: RecalcMetrics = this.getCurrentMetrics(),
    reason: ViewportPatchBroadcastReason = 'state-change',
  ): void {
    this.viewportPatchPublisher.broadcast({
      event,
      metrics,
      reason,
      impactsBySheet: event === null ? null : collectSheetViewportImpacts(this.requireEngine(), event),
    })
  }

  private buildViewportPatch(
    state: ViewportSubscriptionState,
    event: EngineEvent | null,
    metrics: RecalcMetrics = this.getCurrentMetrics(),
    authoritativeRevision: number = this.authoritativeState.getRevision(),
    sheetImpact: SheetViewportImpact | null = null,
  ): ViewportPatch {
    return this.viewportPatchPublisher.buildPatch(state, event, metrics, authoritativeRevision, sheetImpact)
  }
}
