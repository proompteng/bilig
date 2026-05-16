import { describe, expect, it, vi } from 'vitest'
import { backfillWorkbookSnapshotsFromInlineState, persistCellEvalDiff } from '../workbook-calculation-store.js'
import type { Queryable } from '../store.js'

describe('workbook calculation store', () => {
  it('rejects malformed existing cell_eval rows before diffing projection output', async () => {
    const query = vi
      .fn()
      .mockResolvedValueOnce({
        rows: [
          {
            workbook_id: 'book-1',
            sheet_name: 'Sheet1',
            address: 'A1',
            row_num: '-1',
            col_num: 0,
            value: { tag: 0 },
            flags: 0,
            version: 1,
            style_id: null,
            style_json: null,
            format_id: null,
            format_code: null,
            calc_revision: 1,
            updated_at: '2026-05-16T00:00:00.000Z',
          },
        ],
      })
      .mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await expect(persistCellEvalDiff(db, 'book-1', [])).rejects.toThrow('Invalid cell_eval projection row for workbook book-1')

    expect(query).toHaveBeenCalledTimes(1)
  })

  it('backfills json-v1 workbook snapshots from inline state', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await backfillWorkbookSnapshotsFromInlineState(db)

    expect(query).toHaveBeenCalledWith(expect.stringContaining('INSERT INTO workbook_snapshot'), ['json-v1'])
  })
})
