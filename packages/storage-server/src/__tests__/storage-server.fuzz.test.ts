import { describe, expect, it, vi } from 'vitest'
import * as fc from 'fast-check'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import { runProperty } from '@bilig/test-fuzz'
import { createInMemoryDocumentPersistence } from '../index.js'

type StorageAction =
  | { kind: 'append'; documentId: string; counter: number; now: number }
  | { kind: 'reset'; documentId: string; cursor: number; now: number }
  | { kind: 'claim'; documentId: string; ownerId: string; leaseExpiresAtUnixMs: number; now: number }
  | { kind: 'release'; documentId: string; ownerId: string; now: number }
  | { kind: 'join'; documentId: string; sessionId: string; now: number }
  | { kind: 'leave'; documentId: string; sessionId: string; now: number }

interface BatchRecordModel {
  documentId: string
  cursor: number
  batch: EngineOpBatch
  receivedAtUnixMs: number
}

interface DocumentModel {
  baseCursor: number
  batches: BatchRecordModel[]
  owner: { ownerId: string; leaseExpiresAtUnixMs: number } | null
  sessions: Set<string>
}

const documentIds = ['book-alpha', 'book-beta'] as const
const ownerIds = ['svc-a', 'svc-b', 'svc-c'] as const
const sessionIds = ['sess-1', 'sess-2', 'sess-3'] as const

describe('storage-server fuzz', () => {
  it('should keep in-memory batch, ownership, and presence stores aligned with a sequential model', async () => {
    await runProperty({
      suite: 'storage-server/in-memory-persistence-model',
      arbitrary: fc.array(storageActionArbitrary, { minLength: 4, maxLength: 24 }),
      predicate: async (actions) => {
        vi.useFakeTimers()
        try {
          const persistence = createInMemoryDocumentPersistence()
          const model = createModel()
          let currentNow = 0
          await actions.reduce<Promise<void>>(async (previous, action) => {
            await previous
            currentNow = Math.max(currentNow, action.now)
            vi.setSystemTime(new Date(currentNow))
            await applyStorageAction(persistence, action)
            applyStorageActionToModel(model, action, currentNow)
            await expectStorageState(persistence, model, currentNow)
          }, Promise.resolve())
        } finally {
          vi.useRealTimers()
        }
      },
    })
  })
})

// Helpers

const storageActionArbitrary = fc.oneof<StorageAction>(
  fc
    .record({
      documentId: fc.constantFrom(...documentIds),
      counter: fc.integer({ min: 1, max: 1_000 }),
      now: fc.integer({ min: 1_000, max: 50_000 }),
    })
    .map(({ documentId, counter, now }) => ({ kind: 'append' as const, documentId, counter, now })),
  fc
    .record({
      documentId: fc.constantFrom(...documentIds),
      cursor: fc.integer({ min: 0, max: 20 }),
      now: fc.integer({ min: 1_000, max: 50_000 }),
    })
    .map(({ documentId, cursor, now }) => ({ kind: 'reset' as const, documentId, cursor, now })),
  fc
    .record({
      documentId: fc.constantFrom(...documentIds),
      ownerId: fc.constantFrom(...ownerIds),
      leaseExpiresAtUnixMs: fc.integer({ min: 1_000, max: 60_000 }),
      now: fc.integer({ min: 1_000, max: 50_000 }),
    })
    .map(({ documentId, ownerId, leaseExpiresAtUnixMs, now }) => ({
      kind: 'claim' as const,
      documentId,
      ownerId,
      leaseExpiresAtUnixMs,
      now,
    })),
  fc
    .record({
      documentId: fc.constantFrom(...documentIds),
      ownerId: fc.constantFrom(...ownerIds),
      now: fc.integer({ min: 1_000, max: 50_000 }),
    })
    .map(({ documentId, ownerId, now }) => ({ kind: 'release' as const, documentId, ownerId, now })),
  fc
    .record({
      documentId: fc.constantFrom(...documentIds),
      sessionId: fc.constantFrom(...sessionIds),
      now: fc.integer({ min: 1_000, max: 50_000 }),
    })
    .map(({ documentId, sessionId, now }) => ({ kind: 'join' as const, documentId, sessionId, now })),
  fc
    .record({
      documentId: fc.constantFrom(...documentIds),
      sessionId: fc.constantFrom(...sessionIds),
      now: fc.integer({ min: 1_000, max: 50_000 }),
    })
    .map(({ documentId, sessionId, now }) => ({ kind: 'leave' as const, documentId, sessionId, now })),
)

function createBatch(documentId: string, counter: number): EngineOpBatch {
  return {
    id: `${documentId}:${counter}`,
    replicaId: documentId,
    clock: { counter },
    ops: [],
  }
}

function createModel(): Map<string, DocumentModel> {
  return new Map(
    documentIds.map((documentId) => [
      documentId,
      {
        baseCursor: 0,
        batches: [],
        owner: null,
        sessions: new Set<string>(),
      },
    ]),
  )
}

async function applyStorageAction(persistence: ReturnType<typeof createInMemoryDocumentPersistence>, action: StorageAction): Promise<void> {
  switch (action.kind) {
    case 'append':
      await persistence.batches.append(action.documentId, createBatch(action.documentId, action.counter), action.now)
      return
    case 'reset':
      await persistence.batches.reset(action.documentId, action.cursor)
      return
    case 'claim':
      await persistence.ownership.claim(action.documentId, action.ownerId, action.leaseExpiresAtUnixMs)
      return
    case 'release':
      await persistence.ownership.release(action.documentId, action.ownerId)
      return
    case 'join':
      await persistence.presence.join(action.documentId, action.sessionId)
      return
    case 'leave':
      await persistence.presence.leave(action.documentId, action.sessionId)
      return
  }
}

function applyStorageActionToModel(model: Map<string, DocumentModel>, action: StorageAction, now: number): void {
  const document = model.get(action.documentId)
  if (!document) {
    throw new Error(`Missing document model for ${action.documentId}`)
  }
  if (document.owner && document.owner.leaseExpiresAtUnixMs <= now) {
    document.owner = null
  }

  switch (action.kind) {
    case 'append': {
      const previousCursor = document.batches.at(-1)?.cursor ?? document.baseCursor
      document.batches.push({
        cursor: previousCursor + 1,
        batch: createBatch(action.documentId, action.counter),
        documentId: action.documentId,
        receivedAtUnixMs: action.now,
      })
      return
    }
    case 'reset':
      document.baseCursor = action.cursor
      document.batches = []
      return
    case 'claim':
      if (document.owner && document.owner.ownerId !== action.ownerId && document.owner.leaseExpiresAtUnixMs > now) {
        return
      }
      document.owner = {
        ownerId: action.ownerId,
        leaseExpiresAtUnixMs: action.leaseExpiresAtUnixMs,
      }
      return
    case 'release':
      if (document.owner?.ownerId === action.ownerId) {
        document.owner = null
      }
      return
    case 'join':
      document.sessions.add(action.sessionId)
      return
    case 'leave':
      document.sessions.delete(action.sessionId)
      return
  }
}

function expectedOwner(document: DocumentModel, now: number): string | null {
  if (!document.owner) {
    return null
  }
  return document.owner.leaseExpiresAtUnixMs > now ? document.owner.ownerId : null
}

async function expectStorageState(
  persistence: ReturnType<typeof createInMemoryDocumentPersistence>,
  model: Map<string, DocumentModel>,
  now: number,
): Promise<void> {
  await Promise.all(
    documentIds.map(async (documentId) => {
      const document = model.get(documentId)
      if (!document) {
        throw new Error(`Missing document model for ${documentId}`)
      }
      const latestCursor = document.batches.at(-1)?.cursor ?? document.baseCursor
      await Promise.all([
        expect(persistence.batches.latestCursor(documentId)).resolves.toBe(latestCursor),
        expect(persistence.ownership.owner(documentId)).resolves.toBe(expectedOwner(document, now)),
        expect(persistence.presence.sessions(documentId)).resolves.toEqual([...document.sessions].toSorted()),
        ...Array.from({ length: latestCursor + 2 }, (_, cursor) => {
          const expected = document.batches.filter((entry) => entry.cursor > cursor)
          return expect(persistence.batches.listAfter(documentId, cursor, 256)).resolves.toEqual(expected)
        }),
      ])
    }),
  )
}
