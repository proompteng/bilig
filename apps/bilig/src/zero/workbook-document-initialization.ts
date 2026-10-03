import { buildWorkbookSourceProjection, type WorkbookSourceProjection } from './projection.js'
import { createEmptyWorkbookSnapshot, nowIso } from './store-support.js'
import {
  applyAxisMetadataDiff,
  applyCalculationSettings,
  applyCellDiff,
  applyDefinedNameDiff,
  applyNumberFormatDiff,
  applySheetDiff,
  applyStyleDiff,
  applyWorkbookMetadataDiff,
  insertWorkbookHeaderIfMissing,
  type Queryable,
} from './store.js'
import { runQueryableTransaction } from './transaction-support.js'

async function initializeWorkbookSourceProjection(db: Queryable, projection: WorkbookSourceProjection): Promise<void> {
  const workbookId = projection.workbook.id
  await db.query(`DELETE FROM sheets WHERE workbook_id = $1`, [workbookId])
  await db.query(`DELETE FROM cells WHERE workbook_id = $1`, [workbookId])
  await db.query(`DELETE FROM row_metadata WHERE workbook_id = $1`, [workbookId])
  await db.query(`DELETE FROM column_metadata WHERE workbook_id = $1`, [workbookId])
  await db.query(`DELETE FROM defined_names WHERE workbook_id = $1`, [workbookId])
  await db.query(`DELETE FROM workbook_metadata WHERE workbook_id = $1`, [workbookId])
  await db.query(`DELETE FROM calculation_settings WHERE workbook_id = $1`, [workbookId])
  await db.query(`DELETE FROM cell_styles WHERE workbook_id = $1`, [workbookId])
  await db.query(`DELETE FROM cell_number_formats WHERE workbook_id = $1`, [workbookId])
  await applySheetDiff(db, [], projection.sheets)
  await applyCellDiff(db, [], projection.cells)
  await applyAxisMetadataDiff(db, 'row_metadata', [], projection.rowMetadata)
  await applyAxisMetadataDiff(db, 'column_metadata', [], projection.columnMetadata)
  await applyDefinedNameDiff(db, [], projection.definedNames)
  await applyWorkbookMetadataDiff(db, [], projection.workbookMetadataEntries)
  await applyCalculationSettings(db, projection.calculationSettings)
  await applyStyleDiff(db, [], projection.styles)
  await applyNumberFormatDiff(db, [], projection.numberFormats)
}

export async function ensureWorkbookDocumentExists(db: Queryable, documentId: string, ownerUserId = 'system'): Promise<void> {
  await runQueryableTransaction(db, async (transactionDb) => {
    const snapshot = createEmptyWorkbookSnapshot(documentId)
    const updatedAt = nowIso()
    const projection = buildWorkbookSourceProjection(documentId, snapshot, {
      revision: 0,
      calculatedRevision: 0,
      ownerUserId,
      updatedBy: ownerUserId,
      updatedAt,
    })
    const inserted = await insertWorkbookHeaderIfMissing(transactionDb, documentId, projection.workbook, snapshot, null)
    if (!inserted) {
      return
    }
    await initializeWorkbookSourceProjection(transactionDb, projection)
  })
}
