import { describe, expect, it, vi } from 'vitest'

const storeFns = vi.hoisted(() => ({
  applyAxisMetadataDiff: vi.fn(),
  applyCalculationSettings: vi.fn(),
  applyCellDiff: vi.fn(),
  applyDefinedNameDiff: vi.fn(),
  applyNumberFormatDiff: vi.fn(),
  applySheetDiff: vi.fn(),
  applyStyleDiff: vi.fn(),
  applyWorkbookMetadataDiff: vi.fn(),
  insertWorkbookHeaderIfMissing: vi.fn(),
  persistCellEvalRows: vi.fn(),
  repairWorkbookSheetIds: vi.fn(),
  upsertWorkbookHeader: vi.fn(),
}))

vi.mock('../store.js', () => ({
  applyAxisMetadataDiff: storeFns.applyAxisMetadataDiff,
  applyCalculationSettings: storeFns.applyCalculationSettings,
  applyCellDiff: storeFns.applyCellDiff,
  applyDefinedNameDiff: storeFns.applyDefinedNameDiff,
  applyNumberFormatDiff: storeFns.applyNumberFormatDiff,
  applySheetDiff: storeFns.applySheetDiff,
  applyStyleDiff: storeFns.applyStyleDiff,
  applyWorkbookMetadataDiff: storeFns.applyWorkbookMetadataDiff,
  insertWorkbookHeaderIfMissing: storeFns.insertWorkbookHeaderIfMissing,
  upsertWorkbookHeader: storeFns.upsertWorkbookHeader,
}))

vi.mock('../workbook-calculation-store.js', () => ({
  persistCellEvalRows: storeFns.persistCellEvalRows,
}))

vi.mock('../sheet-id-repair.js', () => ({
  repairWorkbookSheetIds: storeFns.repairWorkbookSheetIds,
}))

import {
  backfillCellEvalStyleJson,
  backfillWorkbookSourceProjectionVersion,
  dropLegacyZeroSyncSchemaObjects,
  ensureWorkbookDocumentExists,
  repairWorkbookSheetIdsForMigration,
} from '../workbook-migration-store.js'
import type { Queryable } from '../store.js'

describe('workbook migration store', () => {
  it('skips projection replacement when the workbook already exists', async () => {
    storeFns.insertWorkbookHeaderIfMissing.mockResolvedValueOnce(false)
    const query = vi.fn()
    const db: Queryable = { query }

    await ensureWorkbookDocumentExists(db, 'book-1', 'owner-1')

    expect(storeFns.insertWorkbookHeaderIfMissing).toHaveBeenCalledOnce()
    expect(storeFns.applySheetDiff).not.toHaveBeenCalled()
    expect(query).not.toHaveBeenCalled()
  })

  it('drops the legacy zero-sync schema objects', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await dropLegacyZeroSyncSchemaObjects(db)

    expect(query.mock.calls).toEqual([
      [`DROP INDEX IF EXISTS sheet_style_ranges_workbook_sheet_idx`],
      [`DROP INDEX IF EXISTS sheet_format_ranges_workbook_sheet_idx`],
      [`DROP TABLE IF EXISTS sheet_style_ranges`],
      [`DROP TABLE IF EXISTS sheet_format_ranges`],
    ])
  })

  it('repairs sheet ids after backfilling missing sort-order ids', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await repairWorkbookSheetIdsForMigration(db)

    expect(query).toHaveBeenCalledWith(`UPDATE sheets SET sheet_id = sort_order + 1 WHERE sheet_id IS NULL`)
    expect(storeFns.repairWorkbookSheetIds).toHaveBeenCalledWith(db)
  })

  it('returns early from projection backfill when no legacy workbook ids are found', async () => {
    const query = vi
      .fn()
      .mockResolvedValueOnce({ rows: [{ relation: null }] })
      .mockResolvedValueOnce({ rows: [{ relation: null }] })
      .mockResolvedValueOnce({ rows: [] })
    const db: Queryable = { query }

    await backfillWorkbookSourceProjectionVersion(db)

    expect(query).toHaveBeenCalledTimes(3)
    expect(storeFns.upsertWorkbookHeader).not.toHaveBeenCalled()
    expect(storeFns.persistCellEvalRows).not.toHaveBeenCalled()
  })

  it('returns early from cell-eval backfill when no stale style_json rows are found', async () => {
    const query = vi.fn().mockResolvedValueOnce({ rows: [] })
    const db: Queryable = { query }

    await backfillCellEvalStyleJson(db)

    expect(query).toHaveBeenCalledOnce()
    expect(storeFns.persistCellEvalRows).not.toHaveBeenCalled()
  })
})
