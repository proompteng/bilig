import type { SpreadsheetEngine } from '@bilig/core'
import type { RecalcMetrics, SyncState, WorkbookDefinedNameSnapshot } from '@bilig/protocol'
import type { WorkerEngine } from './worker-runtime-support.js'

interface WorkbookFailedPendingMutationLike {
  readonly id: string
  readonly method: string
  readonly failureMessage: string
  readonly attemptCount: number
}

interface WorkbookPendingMutationSummaryLike {
  readonly activeCount: number
  readonly failedCount: number
  readonly firstFailed: WorkbookFailedPendingMutationLike | null
}

export const EMPTY_RUNTIME_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

export interface WorkbookRuntimeSheetSnapshot {
  readonly id: number
  readonly name: string
  readonly order: number
}

interface WorkbookSheetLike {
  id: number
  name: string
  order: number
}

interface WorkbookLike {
  sheetsByName: Map<string, WorkbookSheetLike>
}

export function listOrderedSheetNames(workbook: WorkbookLike): string[] {
  return listOrderedSheets(workbook).map((sheet) => sheet.name)
}

export function listOrderedSheets(workbook: WorkbookLike): WorkbookRuntimeSheetSnapshot[] {
  return [...workbook.sheetsByName.values()]
    .toSorted((left, right) => left.order - right.order)
    .map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      order: sheet.order,
    }))
}

export function cloneRuntimeMetrics(metrics: RecalcMetrics = EMPTY_RUNTIME_METRICS): RecalcMetrics {
  return { ...metrics }
}

function clonePendingMutationSummary(
  summary: WorkbookPendingMutationSummaryLike | undefined,
): WorkbookPendingMutationSummaryLike | undefined {
  if (!summary) {
    return undefined
  }
  return {
    activeCount: summary.activeCount,
    failedCount: summary.failedCount,
    firstFailed: summary.firstFailed
      ? {
          id: summary.firstFailed.id,
          method: summary.firstFailed.method,
          failureMessage: summary.firstFailed.failureMessage,
          attemptCount: summary.firstFailed.attemptCount,
        }
      : null,
  }
}

export function cloneWorkerRuntimeState(input: {
  workbookName: string
  sheets?: readonly WorkbookRuntimeSheetSnapshot[] | undefined
  sheetNames: readonly string[]
  definedNames: readonly WorkbookDefinedNameSnapshot[]
  metrics: RecalcMetrics
  syncState: SyncState
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
}): {
  workbookName: string
  sheets: WorkbookRuntimeSheetSnapshot[]
  sheetNames: string[]
  definedNames: WorkbookDefinedNameSnapshot[]
  metrics: RecalcMetrics
  syncState: SyncState
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
} {
  const pendingMutationSummary = clonePendingMutationSummary(input.pendingMutationSummary)
  const sheets = cloneRuntimeSheets(input.sheets, input.sheetNames)
  return {
    workbookName: input.workbookName,
    sheets,
    sheetNames: [...input.sheetNames],
    definedNames: input.definedNames.map((entry) => structuredClone(entry)),
    metrics: cloneRuntimeMetrics(input.metrics),
    syncState: input.syncState,
    ...(pendingMutationSummary ? { pendingMutationSummary } : {}),
    ...(input.localPersistenceMode ? { localPersistenceMode: input.localPersistenceMode } : {}),
  }
}

export function withExternalSyncState(
  state: {
    workbookName: string
    sheets?: readonly WorkbookRuntimeSheetSnapshot[] | undefined
    sheetNames: readonly string[]
    definedNames: readonly WorkbookDefinedNameSnapshot[]
    metrics: RecalcMetrics
    syncState: SyncState
    pendingMutationSummary?: WorkbookPendingMutationSummaryLike
    localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
  },
  externalSyncState: SyncState | null,
): {
  workbookName: string
  sheets: WorkbookRuntimeSheetSnapshot[]
  sheetNames: string[]
  definedNames: WorkbookDefinedNameSnapshot[]
  metrics: RecalcMetrics
  syncState: SyncState
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
} {
  const nextState = cloneWorkerRuntimeState(state)
  nextState.syncState = externalSyncState ?? state.syncState
  return nextState
}

export function buildWorkerRuntimeStateFromBootstrap(input: {
  workbookName: string
  sheets?: readonly WorkbookRuntimeSheetSnapshot[] | undefined
  sheetNames: readonly string[]
  definedNames?: readonly WorkbookDefinedNameSnapshot[]
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
}): {
  workbookName: string
  sheets: WorkbookRuntimeSheetSnapshot[]
  sheetNames: string[]
  definedNames: WorkbookDefinedNameSnapshot[]
  metrics: RecalcMetrics
  syncState: SyncState
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
} {
  const sheets = cloneRuntimeSheets(input.sheets, input.sheetNames)
  return {
    workbookName: input.workbookName,
    sheets,
    sheetNames: [...input.sheetNames],
    definedNames: (input.definedNames ?? []).map((entry) => structuredClone(entry)),
    metrics: cloneRuntimeMetrics(),
    syncState: 'syncing',
    ...(input.localPersistenceMode ? { localPersistenceMode: input.localPersistenceMode } : {}),
  }
}

export function buildWorkerRuntimeStateFromEngine(engine: SpreadsheetEngine & WorkerEngine): {
  workbookName: string
  sheets: WorkbookRuntimeSheetSnapshot[]
  sheetNames: string[]
  definedNames: WorkbookDefinedNameSnapshot[]
  metrics: RecalcMetrics
  syncState: SyncState
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
} {
  const sheets = listOrderedSheets(engine.workbook)
  return {
    workbookName: engine.workbook.workbookName,
    sheets,
    sheetNames: sheets.map((sheet) => sheet.name),
    definedNames: engine.getDefinedNames().map((entry) => structuredClone(entry)),
    metrics: cloneRuntimeMetrics(engine.getLastMetrics()),
    syncState: engine.getSyncState(),
  }
}

function cloneRuntimeSheets(
  sheets: readonly WorkbookRuntimeSheetSnapshot[] | undefined,
  sheetNames: readonly string[],
): WorkbookRuntimeSheetSnapshot[] {
  if (sheets && sheets.length > 0) {
    return sheets
      .map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        order: sheet.order,
      }))
      .toSorted((left, right) => left.order - right.order)
  }
  return sheetNames.map((name, index) => ({
    id: index + 1,
    name,
    order: index,
  }))
}
