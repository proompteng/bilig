import { describe, expect, it, vi } from 'vitest'
import { createEmptyWorkbookSnapshot } from '../store-support.js'

const storeFns = vi.hoisted(() => ({
  loadWorkbookEventRecordsAfter: vi.fn(),
}))

vi.mock('../store.js', async () => ({
  loadWorkbookEventRecordsAfter: storeFns.loadWorkbookEventRecordsAfter,
}))

import { acquireWorkbookMutationLock, loadWorkbookRuntimeMetadata, loadWorkbookState } from '../workbook-runtime-store.js'
import type { Queryable } from '../store.js'

describe('workbook runtime store', () => {
  it('returns inline workbook state without replay', async () => {
    const snapshot = createEmptyWorkbookSnapshot('book-1')
    const db: Queryable = {
      query: vi.fn().mockResolvedValue({
        rows: [
          {
            snapshot,
            replica_snapshot: null,
            head_revision: '4',
            calculated_revision: '3',
            owner_user_id: 'owner-1',
          },
        ],
      }),
    }

    await expect(loadWorkbookState(db, 'book-1')).resolves.toEqual({
      snapshot,
      replicaSnapshot: null,
      headRevision: 4,
      calculatedRevision: 3,
      ownerUserId: 'owner-1',
    })
    expect(storeFns.loadWorkbookEventRecordsAfter).not.toHaveBeenCalled()
  })

  it('replays events from the latest checkpoint when inline state is absent', async () => {
    storeFns.loadWorkbookEventRecordsAfter.mockResolvedValueOnce([
      {
        revision: 2,
        payload: {
          kind: 'setCellValue',
          sheetName: 'Sheet1',
          address: 'A1',
          value: 123,
        },
      },
    ])
    const checkpoint = createEmptyWorkbookSnapshot('book-2')
    const db: Queryable = {
      query: vi
        .fn()
        .mockResolvedValueOnce({
          rows: [
            {
              snapshot: null,
              replica_snapshot: null,
              head_revision: '2',
              calculated_revision: '1',
              owner_user_id: 'owner-2',
            },
          ],
        })
        .mockResolvedValueOnce({
          rows: [
            {
              revision: '1',
              payload: checkpoint,
              replica_snapshot: null,
            },
          ],
        }),
    }

    const loaded = await loadWorkbookState(db, 'book-2')

    expect(loaded.headRevision).toBe(2)
    expect(loaded.snapshot.sheets[0]?.cells.length).toBeGreaterThan(0)
    expect(storeFns.loadWorkbookEventRecordsAfter).toHaveBeenCalledWith(db, 'book-2', 1)
  })

  it('loads metadata and acquires advisory locks', async () => {
    const query = vi
      .fn()
      .mockResolvedValueOnce({
        rows: [{ head_revision: '6', calculated_revision: '5', owner_user_id: 'owner-3' }],
      })
      .mockResolvedValueOnce({ rows: [] })
    const db: Queryable = { query }

    await expect(loadWorkbookRuntimeMetadata(db, 'book-3')).resolves.toEqual({
      headRevision: 6,
      calculatedRevision: 5,
      ownerUserId: 'owner-3',
    })

    await acquireWorkbookMutationLock(db, 'book-3')

    expect(query).toHaveBeenLastCalledWith(`SELECT pg_advisory_xact_lock(hashtext($1))`, ['book-3'])
  })
})
