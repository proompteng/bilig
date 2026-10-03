import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { PendingWorkbookMutation } from '../workbook-sync.js'
import { loadPersistedWorkbookMutationJournal, persistWorkbookMutationJournal } from '../workbook-local-mutation-journal-persistence.js'

const scope = {
  documentId: 'doc-1',
  replicaId: 'browser:test',
}

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
    id: 'doc-1:browser:test:pending:1',
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
    persistWorkbookMutationJournal(scope, [
      mutation({ id: 'doc-1:browser:test:pending:7', localSeq: 7 }),
      mutation({ id: 'doc-1:browser:test:pending:8', localSeq: 8, status: 'acked', ackedAtUnixMs: 300 }),
    ])

    const restored = loadPersistedWorkbookMutationJournal(scope)

    expect(restored).toEqual({
      mutationJournalEntries: [mutation({ id: 'doc-1:browser:test:pending:7', localSeq: 7 })],
      nextPendingMutationSeq: 9,
    })
  })

  it('does not restore another replica journal for the same document', () => {
    persistWorkbookMutationJournal(scope, [mutation({ id: 'doc-1:browser:test:pending:7', localSeq: 7 })])

    expect(loadPersistedWorkbookMutationJournal({ documentId: 'doc-1', replicaId: 'browser:other' })).toBeNull()
  })

  it('does not let another replica advance this replica mutation sequence', () => {
    persistWorkbookMutationJournal(scope, [
      mutation({ id: 'doc-1:browser:test:pending:3', localSeq: 3, status: 'acked', ackedAtUnixMs: 300 }),
      mutation({ id: 'doc-1:browser:other:pending:99', localSeq: 99 }),
    ])

    expect(loadPersistedWorkbookMutationJournal(scope)).toEqual({
      mutationJournalEntries: [],
      nextPendingMutationSeq: 4,
    })
  })

  it('keeps the next local sequence after every mutation is acknowledged', () => {
    persistWorkbookMutationJournal(scope, [mutation({ status: 'acked', ackedAtUnixMs: 300 })])

    expect(loadPersistedWorkbookMutationJournal(scope)).toEqual({
      mutationJournalEntries: [],
      nextPendingMutationSeq: 2,
    })
    expect(removeItem).not.toHaveBeenCalled()
  })

  it('clears empty journals that have no mutation high-water mark', () => {
    persistWorkbookMutationJournal(scope, [])

    expect(loadPersistedWorkbookMutationJournal(scope)).toBeNull()
    expect(removeItem).toHaveBeenCalledWith('bilig:workbook-local-mutation-journal:doc-1:browser%3Atest')
  })

  it('restores in-flight submitted mutations as retryable local mutations after reload', () => {
    persistWorkbookMutationJournal(scope, [
      mutation({
        status: 'submitted',
        submittedAtUnixMs: 220,
        lastAttemptedAtUnixMs: 210,
        attemptCount: 1,
      }),
    ])

    expect(loadPersistedWorkbookMutationJournal(scope)).toEqual({
      mutationJournalEntries: [
        mutation({
          status: 'local',
          submittedAtUnixMs: null,
          lastAttemptedAtUnixMs: 210,
          attemptCount: 1,
        }),
      ],
      nextPendingMutationSeq: 2,
    })
  })

  it('restores in-flight submitted rebased mutations as retryable rebased mutations after reload', () => {
    persistWorkbookMutationJournal(scope, [
      mutation({
        status: 'submitted',
        submittedAtUnixMs: 220,
        lastAttemptedAtUnixMs: 210,
        rebasedAtUnixMs: 205,
        attemptCount: 1,
      }),
    ])

    expect(loadPersistedWorkbookMutationJournal(scope)).toEqual({
      mutationJournalEntries: [
        mutation({
          status: 'rebased',
          submittedAtUnixMs: null,
          lastAttemptedAtUnixMs: 210,
          rebasedAtUnixMs: 205,
          attemptCount: 1,
        }),
      ],
      nextPendingMutationSeq: 2,
    })
  })

  it('drops corrupt stored journals instead of replaying unknown edits', () => {
    storage.setItem(
      'bilig:workbook-local-mutation-journal:doc-1:browser%3Atest',
      '{"version":1,"documentId":"doc-1","replicaId":"browser:test","mutationJournalEntries":[{}]}',
    )

    expect(loadPersistedWorkbookMutationJournal(scope)).toBeNull()
    expect(removeItem).toHaveBeenCalledWith('bilig:workbook-local-mutation-journal:doc-1:browser%3Atest')
  })

  it('ignores obsolete document-only journals', () => {
    storage.setItem(
      'bilig:workbook-local-mutation-journal:doc-1',
      JSON.stringify({
        version: 1,
        documentId: 'doc-1',
        savedAtUnixMs: 100,
        mutationJournalEntries: [mutation()],
        nextPendingMutationSeq: 2,
      }),
    )

    expect(loadPersistedWorkbookMutationJournal(scope)).toBeNull()
    expect(removeItem).not.toHaveBeenCalled()
  })
})
