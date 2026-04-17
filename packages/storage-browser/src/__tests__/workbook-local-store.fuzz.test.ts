import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { createMemoryWorkbookLocalStoreFactory, type WorkbookLocalMutationRecord } from '../index.js'
import { runScheduledProperty } from '@bilig/test-fuzz'

type StoreMutationSlot = 'alpha' | 'beta' | 'gamma'
type StoreMutationPhase = 'local' | 'submitted' | 'acked' | 'rebased' | 'failed'
type StoreAction =
  | { kind: 'append'; mutation: WorkbookLocalMutationRecord }
  | { kind: 'update'; mutation: WorkbookLocalMutationRecord }
  | { kind: 'remove'; id: string }

describe('memory workbook local store fuzz', () => {
  it('should preserve the pending journal and active view across scheduled append, update, remove, and reopen races', async () => {
    await runScheduledProperty({
      suite: 'storage-browser/local-store/journal-races',
      arbitrary: fc.array(storeActionArbitrary, { minLength: 4, maxLength: 16 }),
      predicate: async ({ scheduler, value: actions }) => {
        const factory = createMemoryWorkbookLocalStoreFactory()
        const store = await factory.open('fuzz-storage')
        const executed: Array<{ action: StoreAction; ok: boolean }> = []
        const applyAction = scheduler.scheduleFunction(async (action: StoreAction): Promise<void> => {
          try {
            await applyStoreAction(store, action)
            executed.push({ action, ok: true })
          } catch {
            executed.push({ action, ok: false })
          }
        })

        const actionPromises = actions.map((action) => applyAction(action))
        await scheduler.waitFor(Promise.all(actionPromises))

        const reopened = await factory.open('fuzz-storage')
        const expected = replayStoreActions(executed)

        await expect(reopened.listMutationJournalEntries()).resolves.toEqual(expected.journal)
        await expect(reopened.listPendingMutations()).resolves.toEqual(expected.pending)
        store.close()
        reopened.close()
      },
    })
  })
})

// Helpers

const storeMutationSlots = ['alpha', 'beta', 'gamma'] as const
const storeMutationPhases = ['local', 'submitted', 'acked', 'rebased', 'failed'] as const

const storeActionArbitrary = fc.oneof<StoreAction>(
  mutationRecordArbitrary().map((mutation) => ({ kind: 'append', mutation })),
  mutationRecordArbitrary().map((mutation) => ({ kind: 'update', mutation })),
  fc.constantFrom<StoreMutationSlot>(...storeMutationSlots).map((slot) => ({
    kind: 'remove' as const,
    id: slotConfig(slot).id,
  })),
)

function mutationRecordArbitrary(): fc.Arbitrary<WorkbookLocalMutationRecord> {
  return fc
    .record({
      slot: fc.constantFrom<StoreMutationSlot>(...storeMutationSlots),
      phase: fc.constantFrom<StoreMutationPhase>(...storeMutationPhases),
      value: fc.integer({ min: -50, max: 50 }),
      attemptCount: fc.integer({ min: 0, max: 3 }),
    })
    .map((record) => createMutationRecord(record.slot, record.phase, record.value, record.attemptCount))
}

function slotConfig(slot: StoreMutationSlot): {
  id: string
  localSeq: number
  address: string
} {
  switch (slot) {
    case 'alpha':
      return { id: 'fuzz-storage:pending:1', localSeq: 1, address: 'A1' }
    case 'beta':
      return { id: 'fuzz-storage:pending:2', localSeq: 2, address: 'B2' }
    case 'gamma':
      return { id: 'fuzz-storage:pending:3', localSeq: 3, address: 'C3' }
  }
}

function createMutationRecord(
  slot: StoreMutationSlot,
  phase: StoreMutationPhase,
  value: number,
  attemptCount: number,
): WorkbookLocalMutationRecord {
  const config = slotConfig(slot)
  const enqueuedAtUnixMs = 100 + config.localSeq
  const submittedAtUnixMs = phase === 'submitted' || phase === 'acked' ? enqueuedAtUnixMs + 10 : null
  const ackedAtUnixMs = phase === 'acked' ? enqueuedAtUnixMs + 20 : null
  const rebasedAtUnixMs = phase === 'rebased' ? enqueuedAtUnixMs + 15 : null
  const failedAtUnixMs = phase === 'failed' ? enqueuedAtUnixMs + 18 : null

  return {
    id: config.id,
    localSeq: config.localSeq,
    baseRevision: config.localSeq - 1,
    method: 'setCellValue',
    args: ['Sheet1', config.address, value],
    enqueuedAtUnixMs,
    submittedAtUnixMs,
    lastAttemptedAtUnixMs: attemptCount > 0 ? enqueuedAtUnixMs + 5 : null,
    ackedAtUnixMs,
    rebasedAtUnixMs,
    failedAtUnixMs,
    attemptCount,
    failureMessage: phase === 'failed' ? `mutation ${slot} failed` : null,
    status: phase,
  }
}

async function applyStoreAction(
  store: Awaited<ReturnType<ReturnType<typeof createMemoryWorkbookLocalStoreFactory>['open']>>,
  action: StoreAction,
): Promise<void> {
  switch (action.kind) {
    case 'append':
      await store.appendPendingMutation(action.mutation)
      return
    case 'update':
      await store.updatePendingMutation(action.mutation)
      return
    case 'remove':
      await store.removePendingMutation(action.id)
      return
  }
}

function replayStoreActions(executed: ReadonlyArray<{ action: StoreAction; ok: boolean }>): {
  journal: WorkbookLocalMutationRecord[]
  pending: WorkbookLocalMutationRecord[]
} {
  const entries = new Map<string, WorkbookLocalMutationRecord>()
  executed.forEach(({ action, ok }) => {
    if (!ok) {
      return
    }
    switch (action.kind) {
      case 'append':
        entries.set(action.mutation.id, action.mutation)
        return
      case 'update':
        if (entries.has(action.mutation.id)) {
          entries.set(action.mutation.id, action.mutation)
        }
        return
      case 'remove':
        entries.delete(action.id)
        return
    }
  })
  const journal = [...entries.values()].toSorted((left, right) => left.localSeq - right.localSeq)
  return {
    journal,
    pending: journal.filter((mutation) => mutation.status !== 'acked'),
  }
}
