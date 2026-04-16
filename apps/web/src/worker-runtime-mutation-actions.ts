import { Effect } from 'effect'
import type { PendingWorkbookMutation, PendingWorkbookMutationInput } from './workbook-sync.js'
import {
  markPendingWorkbookMutationAcked,
  markPendingWorkbookMutationFailed,
  markPendingWorkbookMutationSubmitted,
  queuePendingWorkbookMutationRetry,
  recordPendingWorkbookMutationAttempt,
} from './workbook-mutation-journal.js'
import { replaceJournalMutation, syncPendingMutationsFromJournal } from './worker-runtime-pending-mutations.js'

export function createRuntimePendingMutation(args: {
  documentId: string
  localSeq: number
  authoritativeRevision: number
  input: PendingWorkbookMutationInput
  enqueuedAtUnixMs: number
}): PendingWorkbookMutation {
  return {
    id: `${args.documentId}:pending:${args.localSeq}`,
    localSeq: args.localSeq,
    baseRevision: args.authoritativeRevision,
    method: args.input.method,
    args: [...args.input.args],
    enqueuedAtUnixMs: args.enqueuedAtUnixMs,
    submittedAtUnixMs: null,
    lastAttemptedAtUnixMs: null,
    ackedAtUnixMs: null,
    rebasedAtUnixMs: null,
    failedAtUnixMs: null,
    attemptCount: 0,
    failureMessage: null,
    status: 'local',
  }
}

function findMutation(mutationJournalEntries: readonly PendingWorkbookMutation[], id: string): PendingWorkbookMutation | null {
  return mutationJournalEntries.find((mutation) => mutation.id === id) ?? null
}

function updateMutation(args: {
  mutationJournalEntries: readonly PendingWorkbookMutation[]
  id: string
  update: (mutation: PendingWorkbookMutation) => PendingWorkbookMutation
}): {
  mutationJournalEntries: PendingWorkbookMutation[]
  pendingMutations: PendingWorkbookMutation[]
  updatedMutation: PendingWorkbookMutation
} | null {
  const mutation = findMutation(args.mutationJournalEntries, args.id)
  if (!mutation) {
    return null
  }
  const updatedMutation = args.update(mutation)
  const result = replaceJournalMutation([...args.mutationJournalEntries], updatedMutation)
  return {
    ...result,
    updatedMutation,
  }
}

export function ackAbsorbedMutations(args: {
  mutationJournalEntries: readonly PendingWorkbookMutation[]
  absorbedMutationIds: ReadonlySet<string>
  ackedAtUnixMs: number
}): {
  mutationJournalEntries: PendingWorkbookMutation[]
  pendingMutations: PendingWorkbookMutation[]
} {
  const mutationJournalEntries = args.mutationJournalEntries.map((mutation) => {
    if (!args.absorbedMutationIds.has(mutation.id)) {
      return mutation
    }
    return Effect.runSync(markPendingWorkbookMutationAcked(mutation, args.ackedAtUnixMs))
  })
  return {
    mutationJournalEntries,
    pendingMutations: syncPendingMutationsFromJournal(mutationJournalEntries),
  }
}

export function markMutationSubmittedInJournal(args: {
  mutationJournalEntries: readonly PendingWorkbookMutation[]
  id: string
  submittedAtUnixMs: number
}): {
  mutationJournalEntries: PendingWorkbookMutation[]
  pendingMutations: PendingWorkbookMutation[]
  updatedMutation: PendingWorkbookMutation
} | null {
  return updateMutation({
    mutationJournalEntries: args.mutationJournalEntries,
    id: args.id,
    update: (mutation) => {
      if (mutation.status === 'submitted') {
        return mutation
      }
      return Effect.runSync(markPendingWorkbookMutationSubmitted(mutation, args.submittedAtUnixMs))
    },
  })
}

export function recordMutationAttemptInJournal(args: {
  mutationJournalEntries: readonly PendingWorkbookMutation[]
  id: string
  attemptedAtUnixMs: number
}): {
  mutationJournalEntries: PendingWorkbookMutation[]
  pendingMutations: PendingWorkbookMutation[]
  updatedMutation: PendingWorkbookMutation
} | null {
  return updateMutation({
    mutationJournalEntries: args.mutationJournalEntries,
    id: args.id,
    update: (mutation) => Effect.runSync(recordPendingWorkbookMutationAttempt(mutation, args.attemptedAtUnixMs)),
  })
}

export function ackMutationInJournal(args: {
  mutationJournalEntries: readonly PendingWorkbookMutation[]
  id: string
  ackedAtUnixMs: number
}): {
  mutationJournalEntries: PendingWorkbookMutation[]
  pendingMutations: PendingWorkbookMutation[]
  updatedMutation: PendingWorkbookMutation
} | null {
  return updateMutation({
    mutationJournalEntries: args.mutationJournalEntries,
    id: args.id,
    update: (mutation) => Effect.runSync(markPendingWorkbookMutationAcked(mutation, args.ackedAtUnixMs)),
  })
}

export function failMutationInJournal(args: {
  mutationJournalEntries: readonly PendingWorkbookMutation[]
  id: string
  failureMessage: string
  failedAtUnixMs: number
}): {
  mutationJournalEntries: PendingWorkbookMutation[]
  pendingMutations: PendingWorkbookMutation[]
  updatedMutation: PendingWorkbookMutation
} | null {
  return updateMutation({
    mutationJournalEntries: args.mutationJournalEntries,
    id: args.id,
    update: (mutation) => Effect.runSync(markPendingWorkbookMutationFailed(mutation, args.failureMessage, args.failedAtUnixMs)),
  })
}

export function retryMutationInJournal(args: { mutationJournalEntries: readonly PendingWorkbookMutation[]; id: string }): {
  mutationJournalEntries: PendingWorkbookMutation[]
  pendingMutations: PendingWorkbookMutation[]
  updatedMutation: PendingWorkbookMutation
} | null {
  return updateMutation({
    mutationJournalEntries: args.mutationJournalEntries,
    id: args.id,
    update: (mutation) => Effect.runSync(queuePendingWorkbookMutationRetry(mutation)),
  })
}
