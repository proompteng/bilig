import { SpreadsheetEngine } from '@bilig/core'
import {
  buildWorkbookSourceProjection,
  materializeCellEvalProjection,
  type CellEvalRow,
  type WorkbookSourceProjection,
} from './projection.js'
import {
  createEmptyWorkbookSnapshot,
  nowIso,
  parseCheckpointPayload,
  parseCheckpointReplicaState,
  parseNonNegativeInteger,
  parsePositiveInteger,
} from './store-support.js'
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
  upsertWorkbookHeader,
  type QueryResultRow,
  type Queryable,
} from './store.js'
import { persistCellEvalRows } from './workbook-calculation-store.js'
import { repairWorkbookSheetIds } from './sheet-id-repair.js'

const AUTHORITATIVE_SOURCE_PROJECTION_VERSION = 2

interface WorkbookMigrationRow extends QueryResultRow {
  readonly id?: unknown
  readonly snapshot?: unknown
  readonly replica_snapshot?: unknown
  readonly calculated_revision?: unknown
  readonly head_revision?: unknown
  readonly owner_user_id?: unknown
  readonly updated_at?: unknown
}

interface NormalizedWorkbookMigrationRow {
  readonly id: string
  readonly snapshot: unknown
  readonly replicaSnapshot: unknown
  readonly calculatedRevision: number
  readonly headRevision: number
  readonly ownerUserId: string
  readonly updatedAt: string
}

interface DuplicateWorkbookClientMutationIdRow extends QueryResultRow {
  readonly workbook_id?: unknown
  readonly client_mutation_id?: unknown
  readonly duplicate_count?: unknown
  readonly first_revision?: unknown
  readonly last_revision?: unknown
}

async function tableExists(db: Queryable, name: string): Promise<boolean> {
  const result = await db.query<{ relation: string | null }>(`SELECT to_regclass($1) AS relation`, [`public.${name}`])
  return typeof result.rows[0]?.relation === 'string'
}

async function replaceWorkbookSourceProjectionForMigration(db: Queryable, projection: WorkbookSourceProjection): Promise<void> {
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

async function replaceCellEvalForMigration(db: Queryable, documentId: string, rows: readonly CellEvalRow[]): Promise<void> {
  await db.query(`DELETE FROM cell_eval WHERE workbook_id = $1`, [documentId])
  if (rows.length > 0) {
    await persistCellEvalRows(db, documentId, [], rows)
  }
}

async function collectWorkbookIds(db: Queryable, query: string, values?: readonly unknown[]): Promise<Set<string>> {
  const result = await db.query<{ workbook_id?: string | null; id?: string | null }>(query, values ? [...values] : undefined)
  return new Set(
    result.rows.flatMap((row) => {
      if (typeof row.workbook_id === 'string' && row.workbook_id.length > 0) {
        return [row.workbook_id]
      }
      if (typeof row.id === 'string' && row.id.length > 0) {
        return [row.id]
      }
      return []
    }),
  )
}

async function loadWorkbookMigrationRows(db: Queryable, workbookIds: readonly string[]): Promise<readonly WorkbookMigrationRow[]> {
  if (workbookIds.length === 0) {
    return []
  }
  const result = await db.query<WorkbookMigrationRow>(
    `
      SELECT
        id,
        snapshot,
        replica_snapshot,
        calculated_revision,
        head_revision,
        owner_user_id,
        updated_at
      FROM workbooks
      WHERE id = ANY($1::text[])
    `,
    [workbookIds],
  )
  return result.rows
}

function normalizeWorkbookMigrationRow(row: WorkbookMigrationRow): NormalizedWorkbookMigrationRow {
  const headRevision = parseNonNegativeInteger(row.head_revision)
  const calculatedRevision = parseNonNegativeInteger(row.calculated_revision)
  if (
    typeof row.id !== 'string' ||
    row.id.length === 0 ||
    headRevision === null ||
    calculatedRevision === null ||
    calculatedRevision > headRevision
  ) {
    throw new Error(`Invalid workbook migration row for ${typeof row.id === 'string' && row.id.length > 0 ? row.id : '<unknown>'}`)
  }
  const ownerUserId = typeof row.owner_user_id === 'string' && row.owner_user_id.length > 0 ? row.owner_user_id : 'system'
  return {
    id: row.id,
    snapshot: row.snapshot,
    replicaSnapshot: row.replica_snapshot,
    headRevision,
    calculatedRevision,
    ownerUserId,
    updatedAt: typeof row.updated_at === 'string' && row.updated_at.length > 0 ? row.updated_at : nowIso(),
  }
}

async function rebuildWorkbookStateForMigration(
  db: Queryable,
  row: WorkbookMigrationRow,
  options: {
    readonly replaceSourceProjection: boolean
    readonly replaceCellEval: boolean
  },
): Promise<void> {
  const normalized = normalizeWorkbookMigrationRow(row)
  const checkpointPayload = parseCheckpointPayload(normalized.snapshot, normalized.id)
  const replicaState = parseCheckpointReplicaState(normalized.replicaSnapshot)

  const engine = new SpreadsheetEngine({
    workbookName: normalized.id,
    replicaId: `migration:${normalized.id}:${normalized.headRevision}`,
  })
  await engine.ready()
  engine.importSnapshot(checkpointPayload)
  if (replicaState) {
    engine.importReplicaSnapshot(replicaState)
  }

  if (options.replaceSourceProjection) {
    const projection = buildWorkbookSourceProjection(normalized.id, checkpointPayload, {
      revision: normalized.headRevision,
      calculatedRevision: normalized.calculatedRevision,
      ownerUserId: normalized.ownerUserId,
      updatedBy: normalized.ownerUserId,
      updatedAt: normalized.updatedAt,
    })
    await replaceWorkbookSourceProjectionForMigration(db, projection)
    await upsertWorkbookHeader(db, normalized.id, projection.workbook, checkpointPayload, replicaState)
  }

  if (options.replaceCellEval) {
    await replaceCellEvalForMigration(
      db,
      normalized.id,
      materializeCellEvalProjection(engine, normalized.id, normalized.calculatedRevision, normalized.updatedAt),
    )
  }
}

async function rebuildWorkbooksForMigration(
  db: Queryable,
  workbookIds: ReadonlySet<string>,
  options: {
    readonly replaceSourceProjection: boolean
    readonly replaceCellEval: boolean
  },
): Promise<void> {
  if (workbookIds.size === 0) {
    return
  }
  const rows = await loadWorkbookMigrationRows(db, [...workbookIds])
  await Promise.all(
    rows.map(async (row) => {
      await rebuildWorkbookStateForMigration(db, row, options)
    }),
  )
}

async function loadLegacyProjectionWorkbookIds(db: Queryable): Promise<Set<string>> {
  const workbookIds = await collectWorkbookIds(
    db,
    `
      SELECT id
      FROM workbooks
      WHERE source_projection_version < $1
    `,
    [AUTHORITATIVE_SOURCE_PROJECTION_VERSION],
  )
  if (await tableExists(db, 'sheet_style_ranges')) {
    const styleWorkbookIds = await collectWorkbookIds(db, `SELECT DISTINCT workbook_id FROM sheet_style_ranges`)
    styleWorkbookIds.forEach((workbookId) => workbookIds.add(workbookId))
  }
  if (await tableExists(db, 'sheet_format_ranges')) {
    const formatWorkbookIds = await collectWorkbookIds(db, `SELECT DISTINCT workbook_id FROM sheet_format_ranges`)
    formatWorkbookIds.forEach((workbookId) => workbookIds.add(workbookId))
  }
  return workbookIds
}

export async function ensureWorkbookDocumentExists(db: Queryable, documentId: string, ownerUserId = 'system'): Promise<void> {
  const snapshot = createEmptyWorkbookSnapshot(documentId)
  const updatedAt = nowIso()
  const projection = buildWorkbookSourceProjection(documentId, snapshot, {
    revision: 0,
    calculatedRevision: 0,
    ownerUserId,
    updatedBy: ownerUserId,
    updatedAt,
  })
  const inserted = await insertWorkbookHeaderIfMissing(db, documentId, projection.workbook, snapshot, null)
  if (!inserted) {
    return
  }
  await replaceWorkbookSourceProjectionForMigration(db, projection)
}

export async function repairWorkbookSheetIdsForMigration(db: Queryable): Promise<void> {
  await db.query(`UPDATE sheets SET sheet_id = sort_order + 1 WHERE sheet_id IS NULL`)
  await repairWorkbookSheetIds(db)
}

export async function backfillWorkbookSourceProjectionVersion(db: Queryable): Promise<void> {
  await rebuildWorkbooksForMigration(db, await loadLegacyProjectionWorkbookIds(db), {
    replaceSourceProjection: true,
    replaceCellEval: false,
  })
}

export async function backfillCellEvalStyleJson(db: Queryable): Promise<void> {
  const workbookIds = await collectWorkbookIds(
    db,
    `
      SELECT DISTINCT workbook_id
      FROM cell_eval
      WHERE style_id IS NOT NULL
        AND style_json IS NULL
    `,
  )
  await rebuildWorkbooksForMigration(db, workbookIds, {
    replaceSourceProjection: false,
    replaceCellEval: true,
  })
}

export async function enforceWorkbookEventClientMutationIdUniqueness(db: Queryable): Promise<void> {
  const duplicateResult = await db.query<DuplicateWorkbookClientMutationIdRow>(`
    SELECT workbook_id,
           client_mutation_id,
           COUNT(*)::int AS duplicate_count,
           MIN(revision) AS first_revision,
           MAX(revision) AS last_revision
      FROM workbook_event
     WHERE client_mutation_id IS NOT NULL
     GROUP BY workbook_id, client_mutation_id
    HAVING COUNT(*) > 1
     ORDER BY workbook_id, client_mutation_id
     LIMIT 5
  `)
  if (duplicateResult.rows.length > 0) {
    const duplicateSummary = duplicateResult.rows.map(formatDuplicateClientMutationId).join('; ')
    throw new Error(`Cannot enforce workbook_event client mutation id uniqueness while duplicate ids exist: ${duplicateSummary}`)
  }
  await db.query(`
    CREATE UNIQUE INDEX IF NOT EXISTS workbook_event_workbook_client_mutation_idx
      ON workbook_event(workbook_id, client_mutation_id)
      WHERE client_mutation_id IS NOT NULL;
  `)
}

export async function dropLegacyZeroSyncSchemaObjects(db: Queryable): Promise<void> {
  await db.query(`DROP INDEX IF EXISTS sheet_style_ranges_workbook_sheet_idx`)
  await db.query(`DROP INDEX IF EXISTS sheet_format_ranges_workbook_sheet_idx`)
  await db.query(`DROP TABLE IF EXISTS sheet_style_ranges`)
  await db.query(`DROP TABLE IF EXISTS sheet_format_ranges`)
}

function formatDuplicateClientMutationId(row: DuplicateWorkbookClientMutationIdRow): string {
  const workbookId = typeof row.workbook_id === 'string' ? row.workbook_id : '<unknown workbook>'
  const clientMutationId = typeof row.client_mutation_id === 'string' ? row.client_mutation_id : '<unknown mutation>'
  const duplicateCount = parsePositiveInteger(row.duplicate_count) ?? '<invalid>'
  const firstRevision = parsePositiveInteger(row.first_revision) ?? '<invalid>'
  const lastRevision = parsePositiveInteger(row.last_revision) ?? '<invalid>'
  return `${workbookId}/${clientMutationId} count=${duplicateCount} revisions=${firstRevision}-${lastRevision}`
}
