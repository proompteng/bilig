import { describe, expect, it } from 'vitest'
import type { PendingWorkbookMutation, PendingWorkbookMutationInput } from '../workbook-sync.js'
import {
  ackAbsorbedMutations,
  ackMutationInJournal,
  createRuntimePendingMutation,
  failMutationInJournal,
  markMutationSubmittedInJournal,
  recordMutationAttemptInJournal,
  retryMutationInJournal,
} from '../worker-runtime-mutation-actions.js'

function createMutation(overrides: Partial<PendingWorkbookMutation> = {}): PendingWorkbookMutation {
  return {
    id: 'worker-doc:pending:1',
    localSeq: 1,
    baseRevision: 0,
    method: 'setCellValue',
    args: ['Sheet1', 'A1', 17],
    enqueuedAtUnixMs: 100,
    submittedAtUnixMs: null,
    lastAttemptedAtUnixMs: null,
    ackedAtUnixMs: null,
    rebasedAtUnixMs: null,
    failedAtUnixMs: null,
    attemptCount: 0,
    failureMessage: null,
    status: 'local',
    ...overrides,
  }
}

describe('worker runtime mutation actions', () => {
  it('creates runtime pending mutations with cloned args', () => {
    const input: PendingWorkbookMutationInput = {
      method: 'setCellValue',
      args: ['Sheet1', 'B2', 23],
    }

    const mutation = createRuntimePendingMutation({
      documentId: 'doc-1',
      localSeq: 4,
      authoritativeRevision: 9,
      input,
      enqueuedAtUnixMs: 250,
    })

    expect(mutation).toEqual({
      id: 'doc-1:pending:4',
      localSeq: 4,
      baseRevision: 9,
      method: 'setCellValue',
      args: ['Sheet1', 'B2', 23],
      enqueuedAtUnixMs: 250,
      submittedAtUnixMs: null,
      lastAttemptedAtUnixMs: null,
      ackedAtUnixMs: null,
      rebasedAtUnixMs: null,
      failedAtUnixMs: null,
      attemptCount: 0,
      failureMessage: null,
      status: 'local',
    })
    expect(mutation.args).not.toBe(input.args)
  })

  it('acks absorbed mutations and removes them from the active pending list', () => {
    const local = createMutation()
    const rebased = createMutation({
      id: 'worker-doc:pending:2',
      localSeq: 2,
      status: 'rebased',
      rebasedAtUnixMs: 130,
    })

    const result = ackAbsorbedMutations({
      mutationJournalEntries: [local, rebased],
      absorbedMutationIds: new Set([rebased.id]),
      ackedAtUnixMs: 200,
    })

    expect(result.mutationJournalEntries).toEqual([
      local,
      expect.objectContaining({
        id: rebased.id,
        status: 'acked',
        ackedAtUnixMs: 200,
      }),
    ])
    expect(result.pendingMutations).toEqual([local])
  })

  it('ignores absorbed mutations that were already acked', () => {
    const local = createMutation()
    const acked = createMutation({
      id: 'worker-doc:pending:2',
      localSeq: 2,
      status: 'acked',
      ackedAtUnixMs: 180,
    })

    const result = ackAbsorbedMutations({
      mutationJournalEntries: [local, acked],
      absorbedMutationIds: new Set([acked.id]),
      ackedAtUnixMs: 220,
    })

    expect(result.mutationJournalEntries).toEqual([local, acked])
    expect(result.pendingMutations).toEqual([local])
  })

  it('updates submitted, attempted, failed, acked, and retried mutations', () => {
    const local = createMutation()
    const submitted = markMutationSubmittedInJournal({
      mutationJournalEntries: [local],
      id: local.id,
      submittedAtUnixMs: 150,
    })
    expect(submitted?.updatedMutation.status).toBe('submitted')

    const failed = failMutationInJournal({
      mutationJournalEntries: submitted?.mutationJournalEntries ?? [],
      id: local.id,
      failureMessage: 'mutation rejected',
      failedAtUnixMs: 190,
    })
    expect(failed?.updatedMutation.status).toBe('failed')

    const retried = retryMutationInJournal({
      mutationJournalEntries: failed?.mutationJournalEntries ?? [],
      id: local.id,
    })
    expect(retried?.updatedMutation.status).toBe('local')

    const attempted = recordMutationAttemptInJournal({
      mutationJournalEntries: retried?.mutationJournalEntries ?? [],
      id: local.id,
      attemptedAtUnixMs: 195,
    })
    expect(attempted?.updatedMutation.attemptCount).toBe(1)

    const acked = ackMutationInJournal({
      mutationJournalEntries: attempted?.mutationJournalEntries ?? [],
      id: local.id,
      ackedAtUnixMs: 220,
    })
    expect(acked?.updatedMutation.status).toBe('acked')
    expect(acked?.pendingMutations).toEqual([])
  })
})
