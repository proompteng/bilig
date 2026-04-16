import {
  WorkbookLocalStoreLockedError,
  type WorkbookBootstrapState,
  type WorkbookLocalStore,
  type WorkbookLocalStoreFactory,
} from '@bilig/storage-browser'
import { isPendingWorkbookMutation, type PendingWorkbookMutation } from './workbook-sync.js'
import { syncPendingMutationsFromJournal } from './worker-runtime-pending-mutations.js'

export interface WorkerRuntimeBootstrapPersistenceResult {
  localStore: WorkbookLocalStore | null
  restoredFromPersistence: boolean
  restoredBootstrapState: WorkbookBootstrapState | null
  authoritativeRevision: number
  appliedPendingLocalSeq: number
  mutationJournalEntries: PendingWorkbookMutation[]
  pendingMutations: PendingWorkbookMutation[]
  nextPendingMutationSeq: number
  localPersistenceMode: 'persistent' | 'ephemeral' | 'follower'
}

const DEFAULT_LOCK_RETRY_COUNT = 6
const DEFAULT_LOCK_RETRY_DELAY_MS = 100

async function defaultSleep(delayMs: number): Promise<void> {
  await new Promise((resolve) => {
    setTimeout(resolve, delayMs)
  })
}

async function openLocalStoreWithRetry(args: {
  documentId: string
  localStoreFactory: WorkbookLocalStoreFactory
  lockRetryCount: number
  lockRetryDelayMs: number
  sleep: (delayMs: number) => Promise<void>
}): Promise<WorkbookLocalStore> {
  const attemptOpen = async (attempt: number): Promise<WorkbookLocalStore> => {
    try {
      return await args.localStoreFactory.open(args.documentId)
    } catch (error) {
      if (!(error instanceof WorkbookLocalStoreLockedError) || attempt >= args.lockRetryCount) {
        throw error
      }
      await args.sleep(args.lockRetryDelayMs)
      return attemptOpen(attempt + 1)
    }
  }

  return attemptOpen(0)
}

export async function restoreBootstrapPersistence(args: {
  persistState: boolean
  documentId: string
  localStoreFactory: WorkbookLocalStoreFactory
  lockRetryCount?: number
  lockRetryDelayMs?: number
  sleep?: (delayMs: number) => Promise<void>
}): Promise<WorkerRuntimeBootstrapPersistenceResult> {
  if (!args.persistState) {
    return {
      localStore: null,
      restoredFromPersistence: false,
      restoredBootstrapState: null,
      authoritativeRevision: 0,
      appliedPendingLocalSeq: 0,
      mutationJournalEntries: [],
      pendingMutations: [],
      nextPendingMutationSeq: 1,
      localPersistenceMode: 'ephemeral',
    }
  }

  let localStore: WorkbookLocalStore | null = null
  let restoredBootstrapState: WorkbookBootstrapState | null = null

  try {
    localStore = await openLocalStoreWithRetry({
      documentId: args.documentId,
      localStoreFactory: args.localStoreFactory,
      lockRetryCount: args.lockRetryCount ?? DEFAULT_LOCK_RETRY_COUNT,
      lockRetryDelayMs: args.lockRetryDelayMs ?? DEFAULT_LOCK_RETRY_DELAY_MS,
      sleep: args.sleep ?? defaultSleep,
    })
    restoredBootstrapState = await localStore.loadBootstrapState()
  } catch (error) {
    if (!(error instanceof WorkbookLocalStoreLockedError)) {
      throw error
    }
    return {
      localStore: null,
      restoredFromPersistence: false,
      restoredBootstrapState: null,
      authoritativeRevision: 0,
      appliedPendingLocalSeq: 0,
      mutationJournalEntries: [],
      pendingMutations: [],
      nextPendingMutationSeq: 1,
      localPersistenceMode: 'follower',
    }
  }

  const persistedPendingMutations = localStore ? await localStore.listMutationJournalEntries() : []
  const mutationJournalEntries = persistedPendingMutations.flatMap((mutation) => (isPendingWorkbookMutation(mutation) ? [mutation] : []))
  const pendingMutations = syncPendingMutationsFromJournal(mutationJournalEntries)
  const appliedPendingLocalSeq = restoredBootstrapState?.appliedPendingLocalSeq ?? 0

  return {
    localStore,
    restoredFromPersistence: restoredBootstrapState !== null,
    restoredBootstrapState,
    authoritativeRevision: restoredBootstrapState?.authoritativeRevision ?? 0,
    appliedPendingLocalSeq,
    mutationJournalEntries,
    pendingMutations,
    nextPendingMutationSeq:
      Math.max(
        appliedPendingLocalSeq,
        persistedPendingMutations.reduce((max, mutation) => Math.max(max, mutation.localSeq), 0),
      ) + 1,
    localPersistenceMode: localStore ? 'persistent' : 'ephemeral',
  }
}
