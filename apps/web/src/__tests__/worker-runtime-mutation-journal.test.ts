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
      getLocalStore: () => null,
      getProjectionEngine: vi.fn(),
      markProjectionDivergedFromLocalStore: vi.fn(),
      queuePersist: vi.fn(async () => {}),
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

  it('enqueues mutations, persists them, and replays them into the projection engine', async () => {
    const appendPendingMutation = vi.fn(async () => {})
    const queuePersist = vi.fn(async () => {})
    const markProjectionDivergedFromLocalStore = vi.fn()
    const engine = {
      setCellValue: vi.fn(),
    }
    const journal = new WorkerRuntimeMutationJournal({
      getDocumentId: () => 'doc',
      getAuthoritativeRevision: () => 4,
      getLocalStore: () => ({ appendPendingMutation, updatePendingMutation: vi.fn() }),
      getProjectionEngine: vi.fn(async () => engine),
      markProjectionDivergedFromLocalStore,
      queuePersist,
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
    expect(markProjectionDivergedFromLocalStore).toHaveBeenCalledTimes(1)
    expect(appendPendingMutation).toHaveBeenCalledWith(mutation)
    expect(engine.setCellValue).toHaveBeenCalledWith('Sheet1', 'A1', 17)
    expect(queuePersist).toHaveBeenCalledTimes(1)
  })

  it('queues persistence after acknowledging a persisted mutation', async () => {
    const updatePendingMutation = vi.fn(async () => {})
    const queuePersist = vi.fn(async () => {})
    const markProjectionDivergedFromLocalStore = vi.fn()
    const journal = new WorkerRuntimeMutationJournal({
      getDocumentId: () => 'doc',
      getAuthoritativeRevision: () => 0,
      getLocalStore: () => ({ appendPendingMutation: vi.fn(), updatePendingMutation }),
      getProjectionEngine: vi.fn(),
      markProjectionDivergedFromLocalStore,
      queuePersist,
      now: () => 800,
    })
    journal.restoreFromBootstrap({
      mutationJournalEntries: [buildMutation({ status: 'submitted', submittedAtUnixMs: 700 })],
      nextPendingMutationSeq: 2,
    })

    await journal.ackPendingMutation('doc:pending:1')

    expect(markProjectionDivergedFromLocalStore).toHaveBeenCalledTimes(1)
    expect(updatePendingMutation).toHaveBeenCalledWith(
      buildMutation({
        status: 'acked',
        submittedAtUnixMs: 700,
        ackedAtUnixMs: 800,
      }),
    )
    expect(queuePersist).toHaveBeenCalledTimes(1)
    expect(journal.listPendingMutations()).toEqual([])
  })

  it('rebases only active mutations back into local persistence', async () => {
    const updatePendingMutation = vi.fn(async () => {})
    const journal = new WorkerRuntimeMutationJournal({
      getDocumentId: () => 'doc',
      getAuthoritativeRevision: () => 0,
      getLocalStore: () => ({ appendPendingMutation: vi.fn(), updatePendingMutation }),
      getProjectionEngine: vi.fn(),
      markProjectionDivergedFromLocalStore: vi.fn(),
      queuePersist: vi.fn(async () => {}),
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

    expect(updatePendingMutation).toHaveBeenCalledTimes(1)
    expect(updatePendingMutation).toHaveBeenCalledWith(
      buildMutation({
        status: 'rebased',
        rebasedAtUnixMs: 901,
      }),
    )
    expect(journal.listPendingMutations()).toEqual([
      buildMutation({
        status: 'rebased',
        rebasedAtUnixMs: 901,
      }),
    ])
  })
})
