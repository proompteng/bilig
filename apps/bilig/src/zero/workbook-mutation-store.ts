import { appendWorkbookChange } from './workbook-change-store.js'
import {
  buildCalculationSettingsRowFromEngine,
  buildSheetCellSourceRowsFromEngine,
  buildSheetColumnMetadataRowsFromEngine,
  buildSheetRowMetadataRowsFromEngine,
  buildSingleCellSourceRowFromEngine,
  buildWorkbookHeaderRowFromEngine,
  buildWorkbookNumberFormatRowsFromEngine,
  buildWorkbookSourceProjectionFromEngine,
  buildWorkbookStyleRowsFromEngine,
  materializeCellEvalProjection,
} from './projection.js'
import { deriveDirtyRegions, type DirtyRegion } from '@bilig/zero-sync'
import {
  eventRequiresRecalc,
  isColumnMetadataEventPayload,
  isFocusedCellEventPayload,
  isNumberFormatRangeEventPayload,
  isRowMetadataEventPayload,
  isStyleRangeEventPayload,
  nowIso,
} from './store-support.js'
import {
  applyAxisMetadataDiff,
  applyCalculationSettings,
  applyCellDiff,
  applyNumberFormatDiff,
  applySourceProjectionDiff,
  applyStyleDiff,
  persistCellSourceRange,
  upsertWorkbookHeader,
  type PersistWorkbookMutationOptions,
  type PersistWorkbookMutationResult,
  type Queryable,
  type WorkbookProjectionCommit,
} from './store.js'
import { persistCellEvalRangeDiff } from './workbook-calculation-store.js'

function buildFocusedProjectionCellRows(
  projection: PersistWorkbookMutationOptions['previousState']['projection'],
  payload: Extract<PersistWorkbookMutationOptions['eventPayload'], { kind: 'setCellValue' | 'setCellFormula' | 'clearCell' }>,
): readonly import('./projection.js').CellSourceRow[] {
  const row = projection.cells.find((entry) => entry.sheetName === payload.sheetName && entry.address === payload.address)
  return row ? [row] : []
}

function buildSheetColumnMetadataRowsFromProjection(
  projection: PersistWorkbookMutationOptions['previousState']['projection'],
  sheetName: string,
): readonly import('./projection.js').AxisMetadataSourceRow[] {
  return projection.columnMetadata.filter((entry) => entry.sheetName === sheetName)
}

function buildSheetRowMetadataRowsFromProjection(
  projection: PersistWorkbookMutationOptions['previousState']['projection'],
  sheetName: string,
): readonly import('./projection.js').AxisMetadataSourceRow[] {
  return projection.rowMetadata.filter((entry) => entry.sheetName === sheetName)
}

async function appendWorkbookEvent(db: Queryable, event: import('@bilig/zero-sync').WorkbookEventRecord): Promise<void> {
  await db.query(
    `
      INSERT INTO workbook_event (
        workbook_id,
        revision,
        actor_user_id,
        client_mutation_id,
        txn_json,
        created_at
      )
      VALUES ($1, $2, $3, $4, $5::jsonb, $6::timestamptz)
    `,
    [event.workbookId, event.revision, event.actorUserId, event.clientMutationId, JSON.stringify(event.payload), event.createdAt],
  )
}

async function supersedePendingRecalcJobs(db: Queryable, documentId: string, toRevision: number): Promise<void> {
  await db.query(
    `
      UPDATE recalc_job
      SET status = 'superseded',
          updated_at = NOW(),
          lease_until = NULL,
          lease_owner = NULL
      WHERE workbook_id = $1
        AND status = 'pending'
        AND to_revision < $2
    `,
    [documentId, toRevision],
  )
}

async function enqueueRecalcJob(
  db: Queryable,
  documentId: string,
  fromRevision: number,
  toRevision: number,
  dirtyRegions: DirtyRegion[] | null,
  updatedAt: string,
): Promise<string> {
  const jobId = `${documentId}:recalc:${toRevision}`
  await db.query(
    `
      INSERT INTO recalc_job (
        id,
        workbook_id,
        from_revision,
        to_revision,
        dirty_regions_json,
        status,
        attempts,
        last_error,
        created_at,
        updated_at
      )
      VALUES ($1, $2, $3, $4, $5::jsonb, 'pending', 0, NULL, $6::timestamptz, $6::timestamptz)
      ON CONFLICT (id)
      DO UPDATE SET
        dirty_regions_json = EXCLUDED.dirty_regions_json,
        status = 'pending',
        lease_until = NULL,
        lease_owner = NULL,
        last_error = NULL,
        updated_at = EXCLUDED.updated_at
    `,
    [jobId, documentId, fromRevision, toRevision, JSON.stringify(dirtyRegions), updatedAt],
  )
  return jobId
}

export async function persistWorkbookMutation(
  db: Queryable,
  documentId: string,
  options: PersistWorkbookMutationOptions,
): Promise<PersistWorkbookMutationResult> {
  const updatedAt = nowIso()
  const revision = options.previousState.headRevision + 1
  const needsRecalc =
    options.previousState.calculatedRevision < options.previousState.headRevision || eventRequiresRecalc(options.eventPayload)
  const nextProjectionOptions = {
    revision,
    calculatedRevision: needsRecalc ? options.previousState.calculatedRevision : revision,
    ownerUserId: options.ownerUserId,
    updatedBy: options.updatedBy,
    updatedAt,
  }
  const nextWorkbookRow = buildWorkbookHeaderRowFromEngine(documentId, options.nextEngine, nextProjectionOptions)
  const nextCalculationSettings = buildCalculationSettingsRowFromEngine(documentId, options.nextEngine)
  let projectionCommit: WorkbookProjectionCommit

  await upsertWorkbookHeader(db, documentId, nextWorkbookRow, null, null)
  if (isFocusedCellEventPayload(options.eventPayload)) {
    const previousCellRows = buildFocusedProjectionCellRows(options.previousState.projection, options.eventPayload)
    const nextCellRow = buildSingleCellSourceRowFromEngine(
      documentId,
      options.nextEngine,
      options.eventPayload.sheetName,
      options.eventPayload.address,
      nextProjectionOptions,
    )
    const nextCellRows = nextCellRow ? [nextCellRow] : []
    await applyCalculationSettings(db, nextCalculationSettings)
    await applyCellDiff(db, previousCellRows, nextCellRows)
    projectionCommit = {
      kind: 'focused-cell',
      workbook: nextWorkbookRow,
      calculationSettings: nextCalculationSettings,
      sheetName: options.eventPayload.sheetName,
      address: options.eventPayload.address,
      cell: nextCellRow,
    }
  } else if (isStyleRangeEventPayload(options.eventPayload)) {
    const nextStyleRows = buildWorkbookStyleRowsFromEngine(documentId, options.nextEngine, nextProjectionOptions)
    const nextCellRows = buildSheetCellSourceRowsFromEngine(
      documentId,
      options.nextEngine,
      options.eventPayload.range.sheetName,
      nextProjectionOptions,
      options.eventPayload.range,
    )
    await applyCalculationSettings(db, nextCalculationSettings)
    await applyStyleDiff(db, options.previousState.projection.styles, nextStyleRows)
    await persistCellSourceRange(db, documentId, options.eventPayload.range, nextCellRows)
    await persistCellEvalRangeDiff(
      db,
      documentId,
      options.eventPayload.range,
      materializeCellEvalProjection(options.nextEngine, documentId, nextProjectionOptions.calculatedRevision, updatedAt),
    )
    projectionCommit = {
      kind: 'cell-range',
      workbook: nextWorkbookRow,
      calculationSettings: nextCalculationSettings,
      range: options.eventPayload.range,
      cells: nextCellRows,
      styles: nextStyleRows,
    }
  } else if (isNumberFormatRangeEventPayload(options.eventPayload)) {
    const nextNumberFormatRows = buildWorkbookNumberFormatRowsFromEngine(documentId, options.nextEngine, nextProjectionOptions)
    const nextCellRows = buildSheetCellSourceRowsFromEngine(
      documentId,
      options.nextEngine,
      options.eventPayload.range.sheetName,
      nextProjectionOptions,
      options.eventPayload.range,
    )
    await applyCalculationSettings(db, nextCalculationSettings)
    await applyNumberFormatDiff(db, options.previousState.projection.numberFormats, nextNumberFormatRows)
    await persistCellSourceRange(db, documentId, options.eventPayload.range, nextCellRows)
    await persistCellEvalRangeDiff(
      db,
      documentId,
      options.eventPayload.range,
      materializeCellEvalProjection(options.nextEngine, documentId, nextProjectionOptions.calculatedRevision, updatedAt),
    )
    projectionCommit = {
      kind: 'cell-range',
      workbook: nextWorkbookRow,
      calculationSettings: nextCalculationSettings,
      range: options.eventPayload.range,
      cells: nextCellRows,
      numberFormats: nextNumberFormatRows,
    }
  } else if (isColumnMetadataEventPayload(options.eventPayload)) {
    const nextColumnMetadataRows = buildSheetColumnMetadataRowsFromEngine(
      documentId,
      options.nextEngine,
      options.eventPayload.sheetName,
      nextProjectionOptions,
    )
    await applyCalculationSettings(db, nextCalculationSettings)
    await applyAxisMetadataDiff(
      db,
      'column_metadata',
      buildSheetColumnMetadataRowsFromProjection(options.previousState.projection, options.eventPayload.sheetName),
      nextColumnMetadataRows,
    )
    projectionCommit = {
      kind: 'column-metadata',
      workbook: nextWorkbookRow,
      calculationSettings: nextCalculationSettings,
      sheetName: options.eventPayload.sheetName,
      columnMetadata: nextColumnMetadataRows,
    }
  } else if (isRowMetadataEventPayload(options.eventPayload)) {
    const nextRowMetadataRows = buildSheetRowMetadataRowsFromEngine(
      documentId,
      options.nextEngine,
      options.eventPayload.sheetName,
      nextProjectionOptions,
    )
    await applyCalculationSettings(db, nextCalculationSettings)
    await applyAxisMetadataDiff(
      db,
      'row_metadata',
      buildSheetRowMetadataRowsFromProjection(options.previousState.projection, options.eventPayload.sheetName),
      nextRowMetadataRows,
    )
    projectionCommit = {
      kind: 'row-metadata',
      workbook: nextWorkbookRow,
      calculationSettings: nextCalculationSettings,
      sheetName: options.eventPayload.sheetName,
      rowMetadata: nextRowMetadataRows,
    }
  } else {
    const nextProjection = buildWorkbookSourceProjectionFromEngine(documentId, options.nextEngine, nextProjectionOptions)
    await applySourceProjectionDiff(db, options.previousState.projection, nextProjection)
    projectionCommit = {
      kind: 'replace',
      projection: nextProjection,
    }
  }

  await appendWorkbookEvent(db, {
    workbookId: documentId,
    revision,
    actorUserId: options.updatedBy,
    clientMutationId: options.clientMutationId ?? null,
    payload: options.eventPayload,
    createdAt: updatedAt,
  })
  await appendWorkbookChange(db, {
    documentId,
    revision,
    actorUserId: options.updatedBy,
    clientMutationId: options.clientMutationId ?? null,
    payload: options.eventPayload,
    undoBundle: options.undoBundle,
    createdAtUnixMs: Date.parse(updatedAt),
  })

  await supersedePendingRecalcJobs(db, documentId, revision)
  const recalcJobId = needsRecalc
    ? await enqueueRecalcJob(
        db,
        documentId,
        options.previousState.calculatedRevision,
        revision,
        eventRequiresRecalc(options.eventPayload) ? deriveDirtyRegions(options.eventPayload) : null,
        updatedAt,
      )
    : null

  return {
    revision,
    calculatedRevision: nextProjectionOptions.calculatedRevision,
    updatedAt,
    recalcJobId,
    projectionCommit,
  }
}
