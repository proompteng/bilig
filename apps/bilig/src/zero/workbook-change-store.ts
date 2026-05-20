import {
  isWorkbookEventPayload,
  normalizeWorkbookChangeRowModel,
  queries,
  type WorkbookChangeUndoBundle,
  type WorkbookChangeRange,
  type WorkbookEventPayload,
} from '@bilig/zero-sync'
import type { Row } from '@rocicorp/zero'
import { buildWorkbookChangeDescriptor, type WorkbookChangeDescriptor } from './workbook-change-descriptor.js'
import type { QueryResultRow, Queryable, ZeroQueryRunner } from './store.js'
import { parseNonNegativeInteger, parsePositiveInteger } from './store-support.js'
import { resolveWorkbookSheetRef } from './workbook-sheet-ref.js'
import { selectLatestRedoableWorkbookChangeRevision, selectLatestUndoableWorkbookChangeRevision } from './workbook-history-selector.js'
import { runQueryableTransaction, runSequentially } from './transaction-support.js'
import { addColumnIfMissing } from './schema-upgrade.js'
import { ensureZeroSchemaTable } from './zero-schema-ddl.js'

export type { WorkbookChangeRange } from '@bilig/zero-sync'
export { buildWorkbookChangeDescriptor, type WorkbookChangeDescriptor } from './workbook-change-descriptor.js'

export interface AppendWorkbookChangeInput {
  readonly documentId: string
  readonly revision: number
  readonly actorUserId: string
  readonly clientMutationId: string | null
  readonly payload: WorkbookEventPayload
  readonly undoBundle: WorkbookChangeUndoBundle | null
  readonly createdAtUnixMs: number
}

interface WorkbookChangeInsertRow {
  readonly documentId: string
  readonly revision: number
  readonly actorUserId: string
  readonly clientMutationId: string | null
  readonly descriptor: WorkbookChangeDescriptor
  readonly undoBundle: WorkbookChangeUndoBundle | null
  readonly revertsRevision: number | null
  readonly createdAtUnixMs: number
}

export interface WorkbookChangeRecord {
  readonly revision: number
  readonly actorUserId: string
  readonly clientMutationId: string | null
  readonly eventKind: WorkbookEventPayload['kind']
  readonly summary: string
  readonly sheetId: number | null
  readonly sheetName: string | null
  readonly anchorAddress: string | null
  readonly range: WorkbookChangeRange | null
  readonly rangeInvalid: boolean
  readonly undoBundle: WorkbookChangeUndoBundle | null
  readonly revertedByRevision: number | null
  readonly revertsRevision: number | null
  readonly createdAtUnixMs: number
}

interface WorkbookEventBackfillRow extends QueryResultRow {
  readonly workbookId?: unknown
  readonly revision?: unknown
  readonly actorUserId?: unknown
  readonly clientMutationId?: unknown
  readonly payload?: unknown
  readonly createdAtUnixMs?: unknown
}

interface WorkbookChangeSelectRow extends QueryResultRow {
  readonly revision?: unknown
  readonly actorUserId?: unknown
  readonly clientMutationId?: unknown
  readonly eventKind?: unknown
  readonly summary?: unknown
  readonly sheetId?: unknown
  readonly sheetName?: unknown
  readonly anchorAddress?: unknown
  readonly rangeJson?: unknown
  readonly undoBundleJson?: unknown
  readonly revertedByRevision?: unknown
  readonly revertsRevision?: unknown
  readonly createdAt?: unknown
  readonly createdAtUnixMs?: unknown
}

type ZeroWorkbookChangeRow = Row['workbook_change']

export interface WorkbookChangeStoreConnection extends Queryable {
  loadWorkbookChangeRow(input: { readonly documentId: string; readonly revision: number }): Promise<ZeroWorkbookChangeRow | null>
  listWorkbookChangesAfterRevisionRows(input: {
    readonly documentId: string
    readonly revision: number
  }): Promise<readonly ZeroWorkbookChangeRow[]>
  listWorkbookHistoryRows(input: { readonly documentId: string }): Promise<readonly ZeroWorkbookChangeRow[]>
  listRecentWorkbookChangeRows(input: { readonly documentId: string; readonly limit: number }): Promise<readonly ZeroWorkbookChangeRow[]>
}

export function createWorkbookChangeStoreConnection(db: Queryable & ZeroQueryRunner): WorkbookChangeStoreConnection {
  return {
    query: (text, values) => db.query(text, values),
    loadWorkbookChangeRow: async ({ documentId, revision }) =>
      (await db.run(queries.workbookChange.one.fn({ args: { documentId, revision }, ctx: { userID: 'system' } }))) ?? null,
    listWorkbookChangesAfterRevisionRows: async ({ documentId, revision }) =>
      await db.run(queries.workbookChange.afterRevision.fn({ args: { documentId, revision }, ctx: { userID: 'system' } })),
    listWorkbookHistoryRows: async ({ documentId }) =>
      await db.run(queries.workbookChange.historyByWorkbook.fn({ args: { documentId }, ctx: { userID: 'system' } })),
    listRecentWorkbookChangeRows: async ({ documentId, limit }) =>
      await db.run(queries.workbookChange.byWorkbook.fn({ args: { documentId, limit }, ctx: { userID: 'system' } })),
  }
}

function toWorkbookChangeSelectRow(row: ZeroWorkbookChangeRow): WorkbookChangeSelectRow {
  return {
    revision: row.revision,
    actorUserId: row.actorUserId,
    clientMutationId: row.clientMutationId,
    eventKind: row.eventKind,
    summary: row.summary,
    sheetId: row.sheetId,
    sheetName: row.sheetName,
    anchorAddress: row.anchorAddress,
    rangeJson: row.rangeJson,
    undoBundleJson: row.undoBundleJson,
    revertedByRevision: row.revertedByRevision,
    revertsRevision: row.revertsRevision,
    createdAtUnixMs: row.createdAt,
  }
}

function normalizeWorkbookChangeRecord(row: WorkbookChangeSelectRow): WorkbookChangeRecord | null {
  const model = normalizeWorkbookChangeRowModel({
    ...row,
    createdAt: row.createdAt ?? row.createdAtUnixMs,
  })
  if (!model) {
    return null
  }
  return {
    revision: model.revision,
    actorUserId: model.actorUserId,
    clientMutationId: model.clientMutationId,
    eventKind: model.eventKind,
    summary: model.summary,
    sheetId: model.sheetId,
    sheetName: model.sheetName,
    anchorAddress: model.anchorAddress,
    range: model.rangeJson,
    rangeInvalid: model.rangeJsonInvalid,
    undoBundle: model.undoBundleJson,
    revertedByRevision: model.revertedByRevision,
    revertsRevision: model.revertsRevision,
    createdAtUnixMs: model.createdAt,
  }
}

function isSafePositiveInteger(value: number): boolean {
  return Number.isSafeInteger(value) && value > 0
}

function isSafeNonNegativeInteger(value: number): boolean {
  return Number.isSafeInteger(value) && value >= 0
}

async function insertWorkbookChange(db: Queryable, row: WorkbookChangeInsertRow): Promise<void> {
  const sheetRef = await resolveWorkbookSheetRef(db, {
    documentId: row.documentId,
    sheetName: row.descriptor.sheetName,
  })
  await db.query(
    `
      INSERT INTO workbook_change (
        workbook_id,
        revision,
        actor_user_id,
        client_mutation_id,
        event_kind,
        summary,
        sheet_id,
        sheet_name,
        anchor_address,
        range_json,
        undo_bundle_json,
        reverted_by_revision,
        reverts_revision,
        created_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10::jsonb, $11::jsonb, $12, $13, $14)
      ON CONFLICT (workbook_id, revision)
      DO UPDATE SET
        actor_user_id = EXCLUDED.actor_user_id,
        client_mutation_id = EXCLUDED.client_mutation_id,
        event_kind = EXCLUDED.event_kind,
        summary = EXCLUDED.summary,
        sheet_id = EXCLUDED.sheet_id,
        sheet_name = EXCLUDED.sheet_name,
        anchor_address = EXCLUDED.anchor_address,
        range_json = EXCLUDED.range_json,
        undo_bundle_json = EXCLUDED.undo_bundle_json,
        reverted_by_revision = COALESCE(EXCLUDED.reverted_by_revision, workbook_change.reverted_by_revision),
        reverts_revision = EXCLUDED.reverts_revision,
        created_at = EXCLUDED.created_at
    `,
    [
      row.documentId,
      row.revision,
      row.actorUserId,
      row.clientMutationId,
      row.descriptor.eventKind,
      row.descriptor.summary,
      sheetRef.sheetId,
      sheetRef.sheetName,
      row.descriptor.anchorAddress,
      JSON.stringify(row.descriptor.range),
      row.undoBundle === null ? null : JSON.stringify(row.undoBundle),
      null,
      row.revertsRevision,
      row.createdAtUnixMs,
    ],
  )
}

async function markWorkbookChangeReverted(
  db: Queryable,
  input: {
    readonly documentId: string
    readonly revision: number
    readonly revertedByRevision: number
  },
): Promise<void> {
  await db.query(
    `
      UPDATE workbook_change
         SET reverted_by_revision = $3
       WHERE workbook_id = $1
         AND revision = $2
    `,
    [input.documentId, input.revision, input.revertedByRevision],
  )
}

export async function ensureWorkbookChangeSchema(db: Queryable): Promise<void> {
  await ensureZeroSchemaTable(db, 'workbook_change', {
    columnOverrides: {
      workbookId: { constraintSql: 'REFERENCES workbooks(id) ON DELETE CASCADE' },
      sheetId: { dataType: 'INTEGER' },
    },
  })
  await addColumnIfMissing(db, { tableName: 'workbook_change', columnName: 'client_mutation_id', dataType: 'TEXT' })
  await addColumnIfMissing(db, { tableName: 'workbook_change', columnName: 'sheet_id', dataType: 'INTEGER' })
  await addColumnIfMissing(db, { tableName: 'workbook_change', columnName: 'sheet_name', dataType: 'TEXT' })
  await addColumnIfMissing(db, { tableName: 'workbook_change', columnName: 'anchor_address', dataType: 'TEXT' })
  await addColumnIfMissing(db, { tableName: 'workbook_change', columnName: 'range_json', dataType: 'JSONB' })
  await db.query(`ALTER TABLE workbook_change ADD COLUMN IF NOT EXISTS undo_bundle_json JSONB;`)
  await db.query(`ALTER TABLE workbook_change ADD COLUMN IF NOT EXISTS reverted_by_revision BIGINT;`)
  await db.query(`ALTER TABLE workbook_change ADD COLUMN IF NOT EXISTS reverts_revision BIGINT;`)
  await db.query(
    `CREATE INDEX IF NOT EXISTS workbook_change_workbook_created_idx ON workbook_change(workbook_id, created_at DESC, revision DESC);`,
  )
}

export async function appendWorkbookChange(db: Queryable, input: AppendWorkbookChangeInput): Promise<void> {
  await runQueryableTransaction(db, async (transactionDb) => {
    await insertWorkbookChange(transactionDb, {
      documentId: input.documentId,
      revision: input.revision,
      actorUserId: input.actorUserId,
      clientMutationId: input.clientMutationId,
      descriptor: buildWorkbookChangeDescriptor(input.payload),
      undoBundle: input.undoBundle,
      revertsRevision: input.payload.kind === 'revertChange' || input.payload.kind === 'redoChange' ? input.payload.targetRevision : null,
      createdAtUnixMs: input.createdAtUnixMs,
    })
    if (input.payload.kind === 'revertChange' || input.payload.kind === 'redoChange') {
      await markWorkbookChangeReverted(transactionDb, {
        documentId: input.documentId,
        revision: input.payload.targetRevision,
        revertedByRevision: input.revision,
      })
    }
  })
}

export async function loadWorkbookChange(
  db: WorkbookChangeStoreConnection,
  documentId: string,
  revision: number,
): Promise<WorkbookChangeRecord | null> {
  if (!isSafePositiveInteger(revision)) {
    return null
  }
  const row = await db.loadWorkbookChangeRow({ documentId, revision })
  return row ? normalizeWorkbookChangeRecord(toWorkbookChangeSelectRow(row)) : null
}

export async function listWorkbookChangesAfterRevision(
  db: WorkbookChangeStoreConnection,
  input: {
    readonly documentId: string
    readonly revision: number
  },
): Promise<WorkbookChangeRecord[]> {
  if (!isSafeNonNegativeInteger(input.revision)) {
    return []
  }
  const rows = await db.listWorkbookChangesAfterRevisionRows(input)
  return rows.flatMap((row) => {
    const record = normalizeWorkbookChangeRecord(toWorkbookChangeSelectRow(row))
    return record ? [record] : []
  })
}

async function listWorkbookHistoryChanges(
  db: WorkbookChangeStoreConnection,
  input: {
    readonly documentId: string
  },
): Promise<WorkbookChangeRecord[]> {
  const rows = await db.listWorkbookHistoryRows(input)
  return rows.flatMap((row) => {
    const record = normalizeWorkbookChangeRecord(toWorkbookChangeSelectRow(row))
    return record ? [record] : []
  })
}

export async function loadLatestUndoableWorkbookChange(
  db: WorkbookChangeStoreConnection,
  input: {
    readonly documentId: string
    readonly actorUserId: string
  },
): Promise<WorkbookChangeRecord | null> {
  const rows = await listWorkbookHistoryChanges(db, input)
  const revision = selectLatestUndoableWorkbookChangeRevision({
    actorUserId: input.actorUserId,
    rows,
  })
  return revision === null ? null : (rows.find((row) => row.revision === revision) ?? null)
}

export async function loadLatestRedoableWorkbookChange(
  db: WorkbookChangeStoreConnection,
  input: {
    readonly documentId: string
    readonly actorUserId: string
  },
): Promise<WorkbookChangeRecord | null> {
  const rows = await listWorkbookHistoryChanges(db, input)
  const revision = selectLatestRedoableWorkbookChangeRevision({
    actorUserId: input.actorUserId,
    rows,
  })
  return revision === null ? null : (rows.find((row) => row.revision === revision) ?? null)
}

export async function listWorkbookChanges(
  db: WorkbookChangeStoreConnection,
  input: {
    readonly documentId: string
    readonly limit?: number
  },
): Promise<WorkbookChangeRecord[]> {
  if (input.limit !== undefined && !isSafePositiveInteger(input.limit)) {
    return []
  }
  const rows = await db.listRecentWorkbookChangeRows({
    documentId: input.documentId,
    limit: input.limit ?? 10,
  })
  return rows.flatMap((row) => {
    const record = normalizeWorkbookChangeRecord(toWorkbookChangeSelectRow(row))
    return record ? [record] : []
  })
}

export async function backfillWorkbookChanges(db: Queryable): Promise<void> {
  const result = await db.query<WorkbookEventBackfillRow>(
    `
      SELECT event.workbook_id AS "workbookId",
             event.revision AS "revision",
             event.actor_user_id AS "actorUserId",
             event.client_mutation_id AS "clientMutationId",
             event.txn_json AS "payload",
             CASE
               WHEN event.created_at IS NULL THEN 0
               ELSE FLOOR(EXTRACT(EPOCH FROM event.created_at) * 1000)
             END AS "createdAtUnixMs"
        FROM workbook_event AS event
        LEFT JOIN workbook_change AS change
          ON change.workbook_id = event.workbook_id
         AND change.revision = event.revision
       WHERE change.workbook_id IS NULL
       ORDER BY event.workbook_id ASC, event.revision ASC
    `,
  )

  const inputs = result.rows.flatMap((row) => {
    const revision = parsePositiveInteger(row.revision)
    const createdAtUnixMs = parseNonNegativeInteger(row.createdAtUnixMs)
    if (
      typeof row.workbookId !== 'string' ||
      typeof row.actorUserId !== 'string' ||
      revision === null ||
      createdAtUnixMs === null ||
      !isWorkbookEventPayload(row.payload)
    ) {
      return []
    }
    return [
      {
        documentId: row.workbookId,
        revision,
        actorUserId: row.actorUserId,
        clientMutationId: typeof row.clientMutationId === 'string' ? row.clientMutationId : null,
        payload: row.payload,
        undoBundle: null,
        createdAtUnixMs,
      } satisfies AppendWorkbookChangeInput,
    ]
  })
  await runQueryableTransaction(db, async (transactionDb) => {
    await runSequentially(inputs, async (input) => {
      await appendWorkbookChange(transactionDb, input)
    })
  })
}
