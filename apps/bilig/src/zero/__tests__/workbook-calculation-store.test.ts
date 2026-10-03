import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { persistCellEvalDiff, persistCellEvalIncremental } from '../workbook-calculation-store.js'
import type { CellEvalRow } from '../projection.js'
import type { QueryResultRow, Queryable } from '../store.js'

class ConcurrentDetectingQueryable implements Queryable {
  readonly calls: string[] = []
  maxActiveQueries = 0
  private activeQueries = 0

  async query<T extends QueryResultRow = QueryResultRow>(text: string): Promise<{ rows: T[] }> {
    this.calls.push(text)
    this.activeQueries += 1
    this.maxActiveQueries = Math.max(this.maxActiveQueries, this.activeQueries)
    await new Promise((resolve) => setTimeout(resolve, 0))
    this.activeQueries -= 1
    return { rows: [] }
  }
}

function createCellEvalRow(address: string, value: CellValue): CellEvalRow {
  const [, colLabel, rowLabel] = /^([A-Z]+)(\d+)$/u.exec(address) ?? []
  const colNum = colLabel ? colLabel.charCodeAt(0) - 'A'.charCodeAt(0) : 0
  const rowNum = rowLabel ? Number(rowLabel) - 1 : 0
  return {
    workbookId: 'book-1',
    sheetName: 'Sheet1',
    address,
    rowNum,
    colNum,
    value,
    flags: 0,
    version: 1,
    styleId: null,
    styleJson: null,
    formatId: null,
    formatCode: null,
    calcRevision: 2,
    updatedAt: '2026-05-16T00:00:00.000Z',
  }
}

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

  it('serializes full cell_eval diff writes so rendered values do not leave in-flight transaction batches', async () => {
    const db = new ConcurrentDetectingQueryable()

    await persistCellEvalDiff(db, 'book-1', [
      createCellEvalRow('A1', { tag: ValueTag.Number, number: 1 }),
      createCellEvalRow('B1', { tag: ValueTag.Number, number: 2 }),
    ])

    expect(db.maxActiveQueries).toBe(1)
    expect(db.calls).toHaveLength(2)
    expect(db.calls[0]).toContain('FROM cell_eval')
    expect(db.calls[1]).toContain('INSERT INTO cell_eval')
    expect(db.calls[1]).toContain('VALUES ($1')
    expect(db.calls[1]).toContain('($15')
  })

  it('serializes incremental cell_eval writes from recalc jobs', async () => {
    const db = new ConcurrentDetectingQueryable()

    await persistCellEvalIncremental(db, 'book-1', [
      createCellEvalRow('A1', { tag: ValueTag.Number, number: 3 }),
      createCellEvalRow('B1', { tag: ValueTag.Number, number: 4 }),
    ])

    expect(db.maxActiveQueries).toBe(1)
    expect(db.calls).toHaveLength(1)
    expect(db.calls.every((call) => call.includes('INSERT INTO cell_eval'))).toBe(true)
    expect(db.calls[0]).toContain('VALUES ($1')
    expect(db.calls[0]).toContain('($15')
  })
})
