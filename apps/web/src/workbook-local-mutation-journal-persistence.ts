import { normalizeRestoredPendingWorkbookMutation } from './workbook-mutation-journal.js'
import { isPendingWorkbookMutationList, type PendingWorkbookMutation } from './workbook-sync.js'

const STORAGE_VERSION = 1
const STORAGE_KEY_PREFIX = 'bilig:workbook-local-mutation-journal:'

export interface PersistedWorkbookMutationJournal {
  readonly mutationJournalEntries: readonly PendingWorkbookMutation[]
  readonly nextPendingMutationSeq: number
}

interface StoredWorkbookMutationJournal extends PersistedWorkbookMutationJournal {
  readonly version: typeof STORAGE_VERSION
  readonly documentId: string
  readonly savedAtUnixMs: number
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isSafePositiveInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value > 0
}

function storageKey(documentId: string): string {
  return `${STORAGE_KEY_PREFIX}${encodeURIComponent(documentId)}`
}

function resolveLocalStorage(): Storage | null {
  const candidate = (globalThis as { localStorage?: Storage | undefined }).localStorage
  return candidate ?? null
}

function activeJournalEntries(entries: readonly PendingWorkbookMutation[]): PendingWorkbookMutation[] {
  return entries.filter((mutation) => mutation.status !== 'acked').map(normalizeRestoredPendingWorkbookMutation)
}

function nextMutationSeq(entries: readonly PendingWorkbookMutation[]): number {
  const maxSeq = entries.reduce((max, mutation) => Math.max(max, mutation.localSeq), 0)
  return maxSeq + 1
}

function parseStoredJournal(documentId: string, value: unknown): PersistedWorkbookMutationJournal | null {
  if (
    !isRecord(value) ||
    value['version'] !== STORAGE_VERSION ||
    value['documentId'] !== documentId ||
    !isPendingWorkbookMutationList(value['mutationJournalEntries'])
  ) {
    return null
  }
  const entries = activeJournalEntries(value['mutationJournalEntries'])
  if (entries.length === 0) {
    return null
  }
  const restoredNextSeq = isSafePositiveInteger(value['nextPendingMutationSeq']) ? value['nextPendingMutationSeq'] : 1
  return {
    mutationJournalEntries: entries,
    nextPendingMutationSeq: Math.max(restoredNextSeq, nextMutationSeq(entries)),
  }
}

export function loadPersistedWorkbookMutationJournal(documentId: string): PersistedWorkbookMutationJournal | null {
  const storage = resolveLocalStorage()
  if (!storage) {
    return null
  }
  const key = storageKey(documentId)
  try {
    const raw = storage.getItem(key)
    if (!raw) {
      return null
    }
    const parsed = parseStoredJournal(documentId, JSON.parse(raw) as unknown)
    if (!parsed) {
      storage.removeItem(key)
      return null
    }
    return parsed
  } catch {
    try {
      storage.removeItem(key)
    } catch {
      // Storage may be unavailable or quota-restricted; persistence is best effort.
    }
    return null
  }
}

export function persistWorkbookMutationJournal(documentId: string, entries: readonly PendingWorkbookMutation[]): void {
  const storage = resolveLocalStorage()
  if (!storage) {
    return
  }
  const key = storageKey(documentId)
  const activeEntries = activeJournalEntries(entries)
  try {
    if (activeEntries.length === 0) {
      storage.removeItem(key)
      return
    }
    const stored: StoredWorkbookMutationJournal = {
      version: STORAGE_VERSION,
      documentId,
      savedAtUnixMs: Date.now(),
      mutationJournalEntries: activeEntries,
      nextPendingMutationSeq: nextMutationSeq(entries),
    }
    storage.setItem(key, JSON.stringify(stored))
  } catch {
    // A full or disabled localStorage must not block workbook edits.
  }
}
