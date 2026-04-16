import { describe, expect, it } from 'vitest'
import { Effect } from 'effect'
import type { PendingWorkbookMutation } from '../workbook-sync.js'
import {
  markPendingWorkbookMutationAcked,
  markPendingWorkbookMutationFailed,
  markPendingWorkbookMutationRebased,
  markPendingWorkbookMutationSubmitted,
  queuePendingWorkbookMutationRetry,
  recordPendingWorkbookMutationAttempt,
} from '../workbook-mutation-journal.js'

function createMutation(overrides: Partial<PendingWorkbookMutation> = {}): PendingWorkbookMutation {
  return {
    id: 'journal-doc:pending:1',
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

describe('workbook mutation journal', () => {
  it('marks unsent local mutations as rebased', () => {
    const rebased = Effect.runSync(markPendingWorkbookMutationRebased(createMutation(), 250))

    expect(rebased).toMatchObject({
      status: 'rebased',
      rebasedAtUnixMs: 250,
      submittedAtUnixMs: null,
      failedAtUnixMs: null,
    })
  })

  it('preserves submitted state while recording rebase time', () => {
    const rebased = Effect.runSync(
      markPendingWorkbookMutationRebased(
        createMutation({
          status: 'submitted',
          submittedAtUnixMs: 200,
          attemptCount: 1,
          lastAttemptedAtUnixMs: 200,
        }),
        300,
      ),
    )

    expect(rebased).toMatchObject({
      status: 'submitted',
      submittedAtUnixMs: 200,
      rebasedAtUnixMs: 300,
      attemptCount: 1,
    })
  })

  it('records attempt metadata before a send and clears failed state on retry', () => {
    const attempted = Effect.runSync(
      recordPendingWorkbookMutationAttempt(
        createMutation({
          status: 'failed',
          rebasedAtUnixMs: 180,
          failedAtUnixMs: 190,
          failureMessage: 'permission denied',
          attemptCount: 2,
        }),
        220,
      ),
    )

    expect(attempted).toMatchObject({
      status: 'rebased',
      attemptCount: 3,
      lastAttemptedAtUnixMs: 220,
      failedAtUnixMs: null,
      failureMessage: null,
      rebasedAtUnixMs: 180,
    })
  })

  it('moves an active mutation through submitted, failed, retry, and acked states', () => {
    const submitted = Effect.runSync(
      markPendingWorkbookMutationSubmitted(Effect.runSync(recordPendingWorkbookMutationAttempt(createMutation(), 150)), 150),
    )
    const failed = Effect.runSync(markPendingWorkbookMutationFailed(submitted, 'mutation rejected', 175))
    const retried = Effect.runSync(queuePendingWorkbookMutationRetry(failed))
    const acked = Effect.runSync(
      markPendingWorkbookMutationAcked(
        Effect.runSync(markPendingWorkbookMutationSubmitted(Effect.runSync(recordPendingWorkbookMutationAttempt(retried, 210)), 210)),
        260,
      ),
    )

    expect(failed).toMatchObject({
      status: 'failed',
      failedAtUnixMs: 175,
      failureMessage: 'mutation rejected',
    })
    expect(retried).toMatchObject({
      status: 'local',
      submittedAtUnixMs: null,
      failedAtUnixMs: null,
      failureMessage: null,
    })
    expect(acked).toMatchObject({
      status: 'acked',
      ackedAtUnixMs: 260,
      attemptCount: 2,
      submittedAtUnixMs: 210,
    })
  })

  it('rejects invalid ack transitions', () => {
    expect(() =>
      Effect.runSync(
        markPendingWorkbookMutationAcked(
          createMutation({
            status: 'acked',
            ackedAtUnixMs: 400,
          }),
          450,
        ),
      ),
    ).toThrowError(/already been acked/i)
  })
})
