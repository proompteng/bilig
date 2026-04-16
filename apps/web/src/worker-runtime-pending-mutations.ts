import { Effect } from 'effect'
import type { PendingWorkbookMutation } from './workbook-sync.js'
import {
  clonePendingWorkbookMutation,
  isActivePendingWorkbookMutation,
  markPendingWorkbookMutationRebased,
} from './workbook-mutation-journal.js'

export interface WorkerRuntimeFailedPendingMutationSnapshot {
  readonly id: string
  readonly method: string
  readonly failureMessage: string
  readonly attemptCount: number
}

export interface WorkerRuntimePendingMutationSummarySnapshot {
  readonly activeCount: number
  readonly failedCount: number
  readonly firstFailed: WorkerRuntimeFailedPendingMutationSnapshot | null
}

export function syncPendingMutationsFromJournal(mutationJournalEntries: readonly PendingWorkbookMutation[]): PendingWorkbookMutation[] {
  return mutationJournalEntries.filter((mutation) => isActivePendingWorkbookMutation(mutation)).map(clonePendingWorkbookMutation)
}

export function replaceJournalMutation(
  mutationJournalEntries: readonly PendingWorkbookMutation[],
  nextMutation: PendingWorkbookMutation,
): {
  mutationJournalEntries: PendingWorkbookMutation[]
  pendingMutations: PendingWorkbookMutation[]
} {
  const nextEntries = mutationJournalEntries.map((mutation) =>
    mutation.id === nextMutation.id ? clonePendingWorkbookMutation(nextMutation) : mutation,
  )
  return {
    mutationJournalEntries: nextEntries,
    pendingMutations: syncPendingMutationsFromJournal(nextEntries),
  }
}

export function markJournalMutationsRebased(
  mutationJournalEntries: readonly PendingWorkbookMutation[],
  rebasedAtUnixMs = Date.now(),
): {
  mutationJournalEntries: PendingWorkbookMutation[]
  pendingMutations: PendingWorkbookMutation[]
  updatedMutations: PendingWorkbookMutation[]
} {
  const updatedMutations: PendingWorkbookMutation[] = []
  const nextEntries = mutationJournalEntries.map((mutation) => {
    if (!isActivePendingWorkbookMutation(mutation)) {
      return mutation
    }
    const updated = Effect.runSync(markPendingWorkbookMutationRebased(mutation, rebasedAtUnixMs))
    updatedMutations.push(updated)
    return updated
  })
  return {
    mutationJournalEntries: nextEntries,
    pendingMutations: syncPendingMutationsFromJournal(nextEntries),
    updatedMutations,
  }
}

export function buildPendingMutationSummary(
  mutationJournalEntries: readonly PendingWorkbookMutation[],
  pendingMutations: readonly PendingWorkbookMutation[],
): WorkerRuntimePendingMutationSummarySnapshot {
  const firstFailed = mutationJournalEntries.find((mutation) => mutation.status === 'failed' && mutation.failureMessage !== null)
  return {
    activeCount: pendingMutations.length,
    failedCount: mutationJournalEntries.filter((mutation) => mutation.status === 'failed').length,
    firstFailed: firstFailed
      ? {
          id: firstFailed.id,
          method: firstFailed.method,
          failureMessage: firstFailed.failureMessage ?? 'Mutation failed',
          attemptCount: firstFailed.attemptCount,
        }
      : null,
  }
}
