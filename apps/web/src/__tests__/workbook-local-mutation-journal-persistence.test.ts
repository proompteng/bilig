import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { PendingWorkbookMutation } from '../workbook-sync.js'
import { loadPersistedWorkbookMutationJournal, persistWorkbookMutationJournal } from '../workbook-local-mutation-journal-persistence.js'

function createStorage() {
  const values = new Map<string, string>()
  const removeItem = vi.fn((key: string) => {
    values.delete(key)
  })
  const storage = {
    getItem: vi.fn((key: string) => values.get(key) ?? null),
    setItem: vi.fn((key: string, value: string) => {
      values.set(key, value)
    }),
    removeItem,
    clear: vi.fn(() => {
      values.clear()
    }),
    key: vi.fn((index: number) => [...values.keys()][index] ?? null),
    get length() {
      return values.size
    },
  } satisfies Storage
  return { removeItem, storage }
}

function mutation(overrides: Partial<PendingWorkbookMutation> = {}): PendingWorkbookMutation {
  return {
    id: 'doc-1:pending:1',
    localSeq: 1,
    baseRevision: 0,
    method: 'clearCell',
    args: ['Sheet1', 'D10'],
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

describe('workbook local mutation journal persistence', () => {
  let storage: Storage
  let removeItem: ReturnType<typeof vi.fn>

  beforeEach(() => {
    const created = createStorage()
    storage = created.storage
    removeItem = created.removeItem
    vi.stubGlobal('localStorage', storage)
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('persists active mutations and restores the next local sequence', () => {
    persistWorkbookMutationJournal('doc-1', [
      mutation({ id: 'doc-1:pending:7', localSeq: 7 }),
      mutation({ id: 'doc-1:pending:8', localSeq: 8, status: 'acked', ackedAtUnixMs: 300 }),
    ])

    const restored = loadPersistedWorkbookMutationJournal('doc-1')

    expect(restored).toEqual({
      mutationJournalEntries: [mutation({ id: 'doc-1:pending:7', localSeq: 7 })],
      nextPendingMutationSeq: 9,
    })
  })

  it('clears the stored journal after every mutation is acknowledged', () => {
    persistWorkbookMutationJournal('doc-1', [mutation({ status: 'acked', ackedAtUnixMs: 300 })])

    expect(loadPersistedWorkbookMutationJournal('doc-1')).toBeNull()
    expect(removeItem).toHaveBeenCalledWith('bilig:workbook-local-mutation-journal:doc-1')
  })

  it('drops corrupt stored journals instead of replaying unknown edits', () => {
    storage.setItem('bilig:workbook-local-mutation-journal:doc-1', '{"version":1,"documentId":"doc-1","mutationJournalEntries":[{}]}')

    expect(loadPersistedWorkbookMutationJournal('doc-1')).toBeNull()
    expect(removeItem).toHaveBeenCalledWith('bilig:workbook-local-mutation-journal:doc-1')
  })
})
