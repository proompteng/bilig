import { normalizeRestoredPendingWorkbookMutation } from './workbook-mutation-journal.js'
import { isPendingWorkbookMutationList, type PendingWorkbookMutation } from './workbook-sync.js'

const STORAGE_VERSION = 1
const STORAGE_KEY_PREFIX = 'bilig:workbook-local-mutation-journal:'

export interface WorkbookMutationJournalPersistenceScope {
  readonly documentId: string
  readonly replicaId: string
}

export interface PersistedWorkbookMutationJournal {
  readonly mutationJournalEntries: readonly PendingWorkbookMutation[]
  readonly nextPendingMutationSeq: number
}

interface StoredWorkbookMutationJournal extends PersistedWorkbookMutationJournal {
  readonly version: typeof STORAGE_VERSION
  readonly documentId: string
  readonly replicaId: string
  readonly savedAtUnixMs: number
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isSafePositiveInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value > 0
}

function storageKey(scope: WorkbookMutationJournalPersistenceScope): string {
  return `${STORAGE_KEY_PREFIX}${encodeURIComponent(scope.documentId)}:${encodeURIComponent(scope.replicaId)}`
}

function resolveLocalStorage(): Storage | null {
  const candidate = (globalThis as { localStorage?: Storage | undefined }).localStorage
  return candidate ?? null
}

function belongsToReplica(scope: WorkbookMutationJournalPersistenceScope, mutation: PendingWorkbookMutation): boolean {
  return mutation.id.startsWith(`${scope.documentId}:${scope.replicaId}:pending:`)
}

function replicaJournalEntries(
  scope: WorkbookMutationJournalPersistenceScope,
  entries: readonly PendingWorkbookMutation[],
): PendingWorkbookMutation[] {
  return entries.filter((mutation) => belongsToReplica(scope, mutation)).map(normalizeRestoredPendingWorkbookMutation)
}

function activeJournalEntries(
  scope: WorkbookMutationJournalPersistenceScope,
  entries: readonly PendingWorkbookMutation[],
): PendingWorkbookMutation[] {
  return replicaJournalEntries(scope, entries).filter((mutation) => mutation.status !== 'acked')
}

function nextMutationSeq(entries: readonly PendingWorkbookMutation[]): number {
  const maxSeq = entries.reduce((max, mutation) => Math.max(max, mutation.localSeq), 0)
  return maxSeq + 1
}

function parseStoredJournal(scope: WorkbookMutationJournalPersistenceScope, value: unknown): PersistedWorkbookMutationJournal | null {
  if (
    !isRecord(value) ||
    value['version'] !== STORAGE_VERSION ||
    value['documentId'] !== scope.documentId ||
    value['replicaId'] !== scope.replicaId ||
    !isPendingWorkbookMutationList(value['mutationJournalEntries'])
  ) {
    return null
  }
  const entries = activeJournalEntries(scope, value['mutationJournalEntries'])
  const restoredNextSeq = isSafePositiveInteger(value['nextPendingMutationSeq']) ? value['nextPendingMutationSeq'] : 1
  if (entries.length === 0 && restoredNextSeq <= 1) {
    return null
  }
  return {
    mutationJournalEntries: entries,
    nextPendingMutationSeq: Math.max(restoredNextSeq, nextMutationSeq(entries)),
  }
}

function removeStorageItem(storage: Storage, key: string): void {
  try {
    storage.removeItem(key)
  } catch {
    // Storage may be unavailable or quota-restricted; persistence is best effort.
  }
}

export function loadPersistedWorkbookMutationJournal(
  scope: WorkbookMutationJournalPersistenceScope,
): PersistedWorkbookMutationJournal | null {
  const storage = resolveLocalStorage()
  if (!storage) {
    return null
  }
  const key = storageKey(scope)
  try {
    const raw = storage.getItem(key)
    if (!raw) {
      return null
    }
    const parsed = parseStoredJournal(scope, JSON.parse(raw) as unknown)
    if (!parsed) {
      removeStorageItem(storage, key)
      return null
    }
    return parsed
  } catch {
    removeStorageItem(storage, key)
    return null
  }
}

export function persistWorkbookMutationJournal(
  scope: WorkbookMutationJournalPersistenceScope,
  entries: readonly PendingWorkbookMutation[],
): void {
  const storage = resolveLocalStorage()
  if (!storage) {
    return
  }
  const key = storageKey(scope)
  const scopedEntries = replicaJournalEntries(scope, entries)
  const activeEntries = scopedEntries.filter((mutation) => mutation.status !== 'acked')
  const nextPendingMutationSeq = Math.max(
    nextMutationSeq(scopedEntries),
    loadPersistedWorkbookMutationJournal(scope)?.nextPendingMutationSeq ?? 1,
  )
  try {
    if (activeEntries.length === 0 && nextPendingMutationSeq <= 1) {
      storage.removeItem(key)
      return
    }
    const stored: StoredWorkbookMutationJournal = {
      version: STORAGE_VERSION,
      documentId: scope.documentId,
      replicaId: scope.replicaId,
      savedAtUnixMs: Date.now(),
      mutationJournalEntries: activeEntries,
      nextPendingMutationSeq,
    }
    storage.setItem(key, JSON.stringify(stored))
  } catch {
    // A full or disabled localStorage must not block workbook edits.
  }
}
