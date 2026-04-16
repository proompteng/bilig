import { describe, expect, it, vi } from 'vitest'
import { ensureZeroSyncSchema } from '../zero-schema-store.js'
import type { Queryable } from '../store.js'

describe('zero schema store', () => {
  it('creates schema tables without running data-repair migrations', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await ensureZeroSyncSchema(db)

    expect(query.mock.calls.some(([text]) => String(text).includes('CREATE TABLE IF NOT EXISTS workbook_snapshot'))).toBe(true)
    expect(
      query.mock.calls.some(([text]) => String(text).includes('UPDATE sheets SET sheet_id = sort_order + 1 WHERE sheet_id IS NULL')),
    ).toBe(false)
    expect(query.mock.calls.some(([text]) => String(text).includes('INSERT INTO workbook_snapshot'))).toBe(false)
  })
})
