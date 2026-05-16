import { describe, expect, it, vi } from 'vitest'
import type { PendingWorkbookMutation } from '../workbook-sync.js'
import { WorkerRuntimeMutationJournal } from '../worker-runtime-mutation-journal.js'

function buildMutation(overrides: Partial<PendingWorkbookMutation> = {}): PendingWorkbookMutation {
  return {
    id: 'doc:pending:1',
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

describe('worker runtime mutation journal', () => {
  it('restores cloned bootstrap entries and recomputes active pending mutations', () => {
    const local = buildMutation()
    const acked = buildMutation({
      id: 'doc:pending:2',
      localSeq: 2,
      status: 'acked',
      ackedAtUnixMs: 300,
    })
    const journal = new WorkerRuntimeMutationJournal({
      getDocumentId: () => 'doc',
      getAuthoritativeRevision: () => 0,
      getProjectionEngine: vi.fn(),
      invalidateProjectionCache: vi.fn(),
      now: () => 500,
    })

    journal.restoreFromBootstrap({
      mutationJournalEntries: [local, acked],
      nextPendingMutationSeq: 7,
    })

    local.args[2] = 99
    expect(journal.listMutationJournalEntries()).toEqual([
      buildMutation(),
      buildMutation({
        id: 'doc:pending:2',
        localSeq: 2,
        status: 'acked',
        ackedAtUnixMs: 300,
      }),
    ])
    expect(journal.listPendingMutations()).toEqual([buildMutation()])
    expect(journal.getAppliedPendingLocalSeq()).toBe(1)
  })

  it('restores submitted bootstrap entries as retryable pending mutations', () => {
    const journal = new WorkerRuntimeMutationJournal({
      getDocumentId: () => 'doc',
      getAuthoritativeRevision: () => 0,
      getProjectionEngine: vi.fn(),
      invalidateProjectionCache: vi.fn(),
      now: () => 500,
    })

    journal.restoreFromBootstrap({
      mutationJournalEntries: [
        buildMutation({
          status: 'submitted',
          submittedAtUnixMs: 400,
          lastAttemptedAtUnixMs: 390,
          attemptCount: 1,
        }),
        buildMutation({
          id: 'doc:pending:2',
          localSeq: 2,
          status: 'submitted',
          submittedAtUnixMs: 420,
          lastAttemptedAtUnixMs: 410,
          rebasedAtUnixMs: 405,
          attemptCount: 1,
        }),
      ],
      nextPendingMutationSeq: 3,
    })

    expect(journal.listPendingMutations()).toEqual([
      buildMutation({
        status: 'local',
        submittedAtUnixMs: null,
        lastAttemptedAtUnixMs: 390,
        attemptCount: 1,
      }),
      buildMutation({
        id: 'doc:pending:2',
        localSeq: 2,
        status: 'rebased',
        submittedAtUnixMs: null,
        lastAttemptedAtUnixMs: 410,
        rebasedAtUnixMs: 405,
        attemptCount: 1,
      }),
    ])
  })

  it('enqueues mutations and replays them into the projection engine', async () => {
    const invalidateProjectionCache = vi.fn()
    const engine = {
      setCellValue: vi.fn(),
    }
    const journal = new WorkerRuntimeMutationJournal({
      getDocumentId: () => 'doc',
      getAuthoritativeRevision: () => 4,
      getProjectionEngine: vi.fn(async () => engine),
      invalidateProjectionCache,
      now: () => 250,
    })

    const mutation = await journal.enqueuePendingMutation({
      method: 'setCellValue',
      args: ['Sheet1', 'A1', 17],
    })

    expect(mutation).toEqual({
      id: 'doc:pending:1',
      localSeq: 1,
      baseRevision: 4,
      method: 'setCellValue',
      args: ['Sheet1', 'A1', 17],
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
    expect(invalidateProjectionCache).toHaveBeenCalledTimes(1)
    expect(engine.setCellValue).toHaveBeenCalledWith('Sheet1', 'A1', 17)
  })

  it('removes acknowledged mutations from the active queue', async () => {
    const invalidateProjectionCache = vi.fn()
    const journal = new WorkerRuntimeMutationJournal({
      getDocumentId: () => 'doc',
      getAuthoritativeRevision: () => 0,
      getProjectionEngine: vi.fn(),
      invalidateProjectionCache,
      now: () => 800,
    })
    journal.restoreFromBootstrap({
      mutationJournalEntries: [buildMutation({ status: 'submitted', submittedAtUnixMs: 700 })],
      nextPendingMutationSeq: 2,
    })

    await journal.ackPendingMutation('doc:pending:1')

    expect(invalidateProjectionCache).toHaveBeenCalledTimes(1)
    expect(journal.listPendingMutations()).toEqual([])
  })

  it('rebases only active mutations in memory', async () => {
    const journal = new WorkerRuntimeMutationJournal({
      getDocumentId: () => 'doc',
      getAuthoritativeRevision: () => 0,
      getProjectionEngine: vi.fn(),
      invalidateProjectionCache: vi.fn(),
      now: () => 900,
    })
    journal.restoreFromBootstrap({
      mutationJournalEntries: [
        buildMutation(),
        buildMutation({
          id: 'doc:pending:2',
          localSeq: 2,
          status: 'acked',
          ackedAtUnixMs: 400,
        }),
      ],
      nextPendingMutationSeq: 3,
    })

    await journal.markRemainingJournalMutationsRebased(901)

    expect(journal.listPendingMutations()).toEqual([
      buildMutation({
        status: 'rebased',
        rebasedAtUnixMs: 901,
      }),
    ])
  })
})
