import type { SpreadsheetEngine } from '@bilig/core'
import type { RecalcMetrics, SyncState, WorkbookDefinedNameSnapshot, WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'
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

interface WorkbookLocalHistoryStateLike {
  readonly canUndo: boolean
  readonly canRedo: boolean
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0
}

function hasUniqueStrings(values: readonly string[]): boolean {
  return values.length === new Set(values).size
}

function isSyncState(value: unknown): value is SyncState {
  return value === 'local-only' || value === 'syncing' || value === 'live' || value === 'behind' || value === 'reconnecting'
}

function isFailedPendingMutationSnapshot(value: unknown): value is WorkbookFailedPendingMutationLike {
  return (
    isRecord(value) &&
    typeof value['id'] === 'string' &&
    typeof value['method'] === 'string' &&
    typeof value['failureMessage'] === 'string' &&
    isNonNegativeInteger(value['attemptCount'])
  )
}

function isPendingMutationSummarySnapshot(value: unknown): value is WorkbookPendingMutationSummaryLike {
  if (!isRecord(value) || !isNonNegativeInteger(value['activeCount']) || !isNonNegativeInteger(value['failedCount'])) {
    return false
  }
  const activeCount = value['activeCount']
  const failedCount = value['failedCount']
  const firstFailed = value['firstFailed']
  return failedCount <= activeCount && (firstFailed === null || (failedCount > 0 && isFailedPendingMutationSnapshot(firstFailed)))
}

function isWorkbookRuntimeSheetSnapshot(value: unknown): value is WorkbookRuntimeSheetSnapshot {
  return isRecord(value) && isNonNegativeInteger(value['id']) && typeof value['name'] === 'string' && isNonNegativeInteger(value['order'])
}

function hasUniqueSheetIdentities(sheets: readonly WorkbookRuntimeSheetSnapshot[]): boolean {
  const ids = new Set<number>()
  const names = new Set<string>()
  for (const sheet of sheets) {
    if (ids.has(sheet.id) || names.has(sheet.name)) {
      return false
    }
    ids.add(sheet.id)
    names.add(sheet.name)
  }
  return true
}

function isWorkbookRuntimeSheetList(value: unknown): value is WorkbookRuntimeSheetSnapshot[] {
  return Array.isArray(value) && value.every(isWorkbookRuntimeSheetSnapshot) && hasUniqueSheetIdentities(value)
}

function isRuntimeSheetNameList(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((sheetName) => typeof sheetName === 'string') && hasUniqueStrings(value)
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}

function isLiteralInput(value: unknown): value is number | string | boolean | null {
  return value === null || typeof value === 'string' || typeof value === 'boolean' || isFiniteNumber(value)
}

function isDefinedNameValueSnapshot(value: unknown): value is WorkbookDefinedNameValueSnapshot {
  if (isLiteralInput(value)) {
    return true
  }
  if (!isRecord(value)) {
    return false
  }
  switch (value['kind']) {
    case 'scalar':
      return isLiteralInput(value['value'])
    case 'cell-ref':
      return typeof value['sheetName'] === 'string' && typeof value['address'] === 'string'
    case 'range-ref':
      return typeof value['sheetName'] === 'string' && typeof value['startAddress'] === 'string' && typeof value['endAddress'] === 'string'
    case 'structured-ref':
      return typeof value['tableName'] === 'string' && typeof value['columnName'] === 'string'
    case 'formula':
      return typeof value['formula'] === 'string'
    default:
      return false
  }
}

function definedNameSnapshotKey(entry: WorkbookDefinedNameSnapshot): string {
  const scope = entry.scopeSheetName?.trim()
  return `${scope && scope.length > 0 ? scope : '<workbook>'}\u0000${entry.name.trim().toUpperCase()}`
}

function hasUniqueDefinedNameKeys(entries: readonly WorkbookDefinedNameSnapshot[]): boolean {
  const keys = new Set<string>()
  for (const entry of entries) {
    const key = definedNameSnapshotKey(entry)
    if (keys.has(key)) {
      return false
    }
    keys.add(key)
  }
  return true
}

function isWorkbookDefinedNameSnapshot(value: unknown): value is WorkbookDefinedNameSnapshot {
  return (
    isRecord(value) &&
    typeof value['name'] === 'string' &&
    value['name'].trim().length > 0 &&
    (value['scopeSheetName'] === undefined || typeof value['scopeSheetName'] === 'string') &&
    isDefinedNameValueSnapshot(value['value'])
  )
}

function isWorkbookDefinedNameList(value: unknown): value is WorkbookDefinedNameSnapshot[] {
  return Array.isArray(value) && value.every(isWorkbookDefinedNameSnapshot) && hasUniqueDefinedNameKeys(value)
}

function isRecalcMetricsSnapshot(value: unknown): value is RecalcMetrics {
  return (
    isRecord(value) &&
    isFiniteNumber(value['batchId']) &&
    isFiniteNumber(value['changedInputCount']) &&
    isFiniteNumber(value['dirtyFormulaCount']) &&
    isFiniteNumber(value['wasmFormulaCount']) &&
    isFiniteNumber(value['jsFormulaCount']) &&
    isFiniteNumber(value['rangeNodeVisits']) &&
    isFiniteNumber(value['recalcMs']) &&
    isFiniteNumber(value['compileMs'])
  )
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
  id?: number
  name: string
  order: number
}

interface WorkbookLike {
  sheetsByName: Map<string, WorkbookSheetLike>
}

export function listOrderedSheetNames(workbook: WorkbookLike): string[] {
  return listOrderedSheets(workbook).map((sheet) => sheet.name)
}

function listOrderedSheets(workbook: WorkbookLike): WorkbookRuntimeSheetSnapshot[] {
  return [...workbook.sheetsByName.values()]
    .toSorted((left, right) => left.order - right.order)
    .map((sheet, index) => ({
      id: typeof sheet.id === 'number' ? sheet.id : index + 1,
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
  localHistoryState?: WorkbookLocalHistoryStateLike | undefined
  authoritativeRevision?: number | undefined
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
}): {
  workbookName: string
  sheets: WorkbookRuntimeSheetSnapshot[]
  sheetNames: string[]
  definedNames: WorkbookDefinedNameSnapshot[]
  metrics: RecalcMetrics
  syncState: SyncState
  localHistoryState: WorkbookLocalHistoryStateLike
  authoritativeRevision?: number | undefined
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
} {
  const pendingMutationSummary = clonePendingMutationSummary(input.pendingMutationSummary)
  const sheets = cloneRuntimeSheets(input.sheets, input.sheetNames)
  return {
    workbookName: input.workbookName,
    sheets,
    sheetNames: sheets.map((sheet) => sheet.name),
    definedNames: input.definedNames.map((entry) => structuredClone(entry)),
    metrics: cloneRuntimeMetrics(input.metrics),
    syncState: input.syncState,
    localHistoryState: {
      canUndo: input.localHistoryState?.canUndo === true,
      canRedo: input.localHistoryState?.canRedo === true,
    },
    ...(typeof input.authoritativeRevision === 'number' ? { authoritativeRevision: input.authoritativeRevision } : {}),
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
    localHistoryState?: WorkbookLocalHistoryStateLike | undefined
    authoritativeRevision?: number | undefined
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
  localHistoryState: WorkbookLocalHistoryStateLike
  authoritativeRevision?: number | undefined
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
} {
  const nextState = cloneWorkerRuntimeState(state)
  nextState.syncState = externalSyncState ?? state.syncState
  return nextState
}

export function isWorkerRuntimeStateSnapshot(value: unknown): value is ReturnType<typeof cloneWorkerRuntimeState> {
  return (
    isRecord(value) &&
    typeof value['workbookName'] === 'string' &&
    isRuntimeSheetNameList(value['sheetNames']) &&
    (value['sheets'] === undefined || isWorkbookRuntimeSheetList(value['sheets'])) &&
    isWorkbookDefinedNameList(value['definedNames']) &&
    isRecalcMetricsSnapshot(value['metrics']) &&
    isSyncState(value['syncState']) &&
    isRecord(value['localHistoryState']) &&
    typeof value['localHistoryState']['canUndo'] === 'boolean' &&
    typeof value['localHistoryState']['canRedo'] === 'boolean' &&
    (value['authoritativeRevision'] === undefined || isNonNegativeInteger(value['authoritativeRevision'])) &&
    (value['pendingMutationSummary'] === undefined || isPendingMutationSummarySnapshot(value['pendingMutationSummary'])) &&
    (value['localPersistenceMode'] === undefined ||
      value['localPersistenceMode'] === 'persistent' ||
      value['localPersistenceMode'] === 'ephemeral' ||
      value['localPersistenceMode'] === 'follower')
  )
}

export function normalizeWorkerRuntimeStateSnapshot(value: unknown): ReturnType<typeof cloneWorkerRuntimeState> | null {
  return isWorkerRuntimeStateSnapshot(value) ? cloneWorkerRuntimeState(value) : null
}

export function buildCachedWorkerRuntimeState(input: {
  cachedState: {
    workbookName: string
    sheets?: readonly WorkbookRuntimeSheetSnapshot[] | undefined
    sheetNames: readonly string[]
    definedNames: readonly WorkbookDefinedNameSnapshot[]
    metrics: RecalcMetrics
    syncState: SyncState
  }
  externalSyncState: SyncState | null
  localHistoryState: WorkbookLocalHistoryStateLike
  authoritativeRevision: number
  pendingMutationSummary: WorkbookPendingMutationSummaryLike
  localPersistenceMode: 'persistent' | 'ephemeral' | 'follower'
}) {
  return withExternalSyncState(
    {
      ...input.cachedState,
      localHistoryState: input.localHistoryState,
      authoritativeRevision: input.authoritativeRevision,
      pendingMutationSummary: input.pendingMutationSummary,
      localPersistenceMode: input.localPersistenceMode,
    },
    input.externalSyncState,
  )
}

export function buildWorkerRuntimeStateFromBootstrap(input: {
  workbookName: string
  sheets?: readonly WorkbookRuntimeSheetSnapshot[] | undefined
  sheetNames: readonly string[]
  definedNames?: readonly WorkbookDefinedNameSnapshot[]
  authoritativeRevision?: number | undefined
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
}): {
  workbookName: string
  sheets: WorkbookRuntimeSheetSnapshot[]
  sheetNames: string[]
  definedNames: WorkbookDefinedNameSnapshot[]
  metrics: RecalcMetrics
  syncState: SyncState
  localHistoryState: WorkbookLocalHistoryStateLike
  authoritativeRevision?: number | undefined
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
} {
  const sheets = cloneRuntimeSheets(input.sheets, input.sheetNames)
  return {
    workbookName: input.workbookName,
    sheets,
    sheetNames: sheets.map((sheet) => sheet.name),
    definedNames: (input.definedNames ?? []).map((entry) => structuredClone(entry)),
    metrics: cloneRuntimeMetrics(),
    syncState: 'syncing',
    localHistoryState: { canUndo: false, canRedo: false },
    ...(typeof input.authoritativeRevision === 'number' ? { authoritativeRevision: input.authoritativeRevision } : {}),
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
  localHistoryState: WorkbookLocalHistoryStateLike
  authoritativeRevision?: number | undefined
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
    localHistoryState: {
      canUndo: engine.canUndo(),
      canRedo: engine.canRedo(),
    },
  }
}

function cloneRuntimeSheets(
  sheets: readonly WorkbookRuntimeSheetSnapshot[] | undefined,
  sheetNames: readonly string[],
): WorkbookRuntimeSheetSnapshot[] {
  if (sheets && sheets.length > 0) {
    return sheets
      .map((sheet, index) => ({
        id: typeof sheet.id === 'number' ? sheet.id : index + 1,
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
