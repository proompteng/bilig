import type { EngineReplicaSnapshot } from '@bilig/core'
import type { CellRangeRef, WorkbookSnapshot } from '@bilig/protocol'
import { diffProjectionRows, type CellEvalRow, sourceProjectionKeys } from './projection.js'
import {
  cellEvalRowInRange,
  cellEvalSignature,
  normalizeRangeBounds,
  nowIso,
  parseCellEvalValue,
  parseCellStyleRecord,
  parseJsonKey,
  parseNonNegativeInteger,
} from './store-support.js'
import type { Queryable, QueryResultRow } from './store.js'

const WORKBOOK_CHECKPOINT_FORMAT = 'json-v1'
const WORKBOOK_CHECKPOINT_RETENTION = 5

interface CellEvalSelectRow extends QueryResultRow {
  readonly workbook_id?: unknown
  readonly sheet_name?: unknown
  readonly address?: unknown
  readonly row_num?: unknown
  readonly col_num?: unknown
  readonly value?: unknown
  readonly flags?: unknown
  readonly version?: unknown
  readonly style_id?: unknown
  readonly style_json?: unknown
  readonly format_id?: unknown
  readonly format_code?: unknown
  readonly calc_revision?: unknown
  readonly updated_at?: unknown
}

function normalizeCellEvalRow(row: CellEvalSelectRow, documentId: string): CellEvalRow {
  const rowNum = parseNonNegativeInteger(row.row_num)
  const colNum = parseNonNegativeInteger(row.col_num)
  const flags = parseNonNegativeInteger(row.flags)
  const version = parseNonNegativeInteger(row.version)
  const calcRevision = parseNonNegativeInteger(row.calc_revision)
  if (
    typeof row.workbook_id !== 'string' ||
    row.workbook_id !== documentId ||
    typeof row.sheet_name !== 'string' ||
    row.sheet_name.length === 0 ||
    typeof row.address !== 'string' ||
    row.address.length === 0 ||
    rowNum === null ||
    colNum === null ||
    flags === null ||
    version === null ||
    calcRevision === null
  ) {
    throw new Error(`Invalid cell_eval projection row for workbook ${documentId}`)
  }

  return {
    workbookId: row.workbook_id,
    sheetName: row.sheet_name,
    address: row.address,
    rowNum,
    colNum,
    value: parseCellEvalValue(row.value),
    flags,
    version,
    styleId: typeof row.style_id === 'string' ? row.style_id : null,
    styleJson: parseCellStyleRecord(row.style_json),
    formatId: typeof row.format_id === 'string' ? row.format_id : null,
    formatCode: typeof row.format_code === 'string' ? row.format_code : null,
    calcRevision,
    updatedAt: typeof row.updated_at === 'string' ? row.updated_at : nowIso(),
  }
}

async function loadCellEvalRows(db: Queryable, documentId: string): Promise<CellEvalRow[]> {
  const result = await db.query<CellEvalSelectRow>(
    `
      SELECT
        workbook_id,
        sheet_name,
        address,
        row_num,
        col_num,
        value,
        flags,
        version,
        style_id,
        style_json,
        format_id,
        format_code,
        calc_revision,
        updated_at
      FROM cell_eval
      WHERE workbook_id = $1
    `,
    [documentId],
  )
  return result.rows.map((row) => normalizeCellEvalRow(row, documentId))
}

export async function persistCellEvalRows(
  db: Queryable,
  documentId: string,
  previousRows: readonly CellEvalRow[],
  nextRows: readonly CellEvalRow[],
): Promise<void> {
  const diff = diffProjectionRows(previousRows, nextRows, sourceProjectionKeys.cellEval, cellEvalSignature)
  const tasks: Promise<unknown>[] = []
  for (const key of diff.deletes) {
    const [, sheetName, address] = parseJsonKey(key)
    tasks.push(
      db.query(`DELETE FROM cell_eval WHERE workbook_id = $1 AND sheet_name = $2 AND address = $3`, [documentId, sheetName, address]),
    )
  }
  for (const row of diff.upserts) {
    tasks.push(
      db.query(
        `
        INSERT INTO cell_eval (
          workbook_id,
          sheet_name,
          address,
          row_num,
          col_num,
          value,
          flags,
          version,
          style_id,
          style_json,
          format_id,
          format_code,
          calc_revision,
          updated_at
        )
        VALUES ($1, $2, $3, $4, $5, $6::jsonb, $7, $8, $9, $10::jsonb, $11, $12, $13, $14::timestamptz)
        ON CONFLICT (workbook_id, sheet_name, address)
        DO UPDATE SET
          row_num = EXCLUDED.row_num,
          col_num = EXCLUDED.col_num,
          value = EXCLUDED.value,
          flags = EXCLUDED.flags,
          version = EXCLUDED.version,
          style_id = EXCLUDED.style_id,
          style_json = EXCLUDED.style_json,
          format_id = EXCLUDED.format_id,
          format_code = EXCLUDED.format_code,
          calc_revision = EXCLUDED.calc_revision,
          updated_at = EXCLUDED.updated_at
      `,
        [
          row.workbookId,
          row.sheetName,
          row.address,
          row.rowNum,
          row.colNum,
          JSON.stringify(row.value),
          row.flags,
          row.version,
          row.styleId,
          JSON.stringify(row.styleJson),
          row.formatId,
          row.formatCode,
          row.calcRevision,
          row.updatedAt,
        ],
      ),
    )
  }
  await Promise.all(tasks)
}

export async function persistCellEvalIncremental(db: Queryable, _documentId: string, rows: readonly CellEvalRow[]): Promise<void> {
  const tasks: Promise<unknown>[] = []
  for (const row of rows) {
    tasks.push(
      db.query(
        `
        INSERT INTO cell_eval (
          workbook_id,
          sheet_name,
          address,
          row_num,
          col_num,
          value,
          flags,
          version,
          style_id,
          style_json,
          format_id,
          format_code,
          calc_revision,
          updated_at
        )
        VALUES ($1, $2, $3, $4, $5, $6::jsonb, $7, $8, $9, $10::jsonb, $11, $12, $13, $14::timestamptz)
        ON CONFLICT (workbook_id, sheet_name, address)
        DO UPDATE SET
          row_num = EXCLUDED.row_num,
          col_num = EXCLUDED.col_num,
          value = EXCLUDED.value,
          flags = EXCLUDED.flags,
          version = EXCLUDED.version,
          style_id = EXCLUDED.style_id,
          style_json = EXCLUDED.style_json,
          format_id = EXCLUDED.format_id,
          format_code = EXCLUDED.format_code,
          calc_revision = EXCLUDED.calc_revision,
          updated_at = EXCLUDED.updated_at
      `,
        [
          row.workbookId,
          row.sheetName,
          row.address,
          row.rowNum,
          row.colNum,
          JSON.stringify(row.value),
          row.flags,
          row.version,
          row.styleId,
          JSON.stringify(row.styleJson),
          row.formatId,
          row.formatCode,
          row.calcRevision,
          row.updatedAt,
        ],
      ),
    )
  }
  await Promise.all(tasks)
}

export async function persistCellEvalDiff(db: Queryable, documentId: string, nextRows: readonly CellEvalRow[]): Promise<void> {
  const previousRows = await loadCellEvalRows(db, documentId)
  await persistCellEvalRows(db, documentId, previousRows, nextRows)
}

export async function persistCellEvalRangeDiff(
  db: Queryable,
  documentId: string,
  range: CellRangeRef,
  nextRows: readonly CellEvalRow[],
): Promise<void> {
  const nextRowsInRange = nextRows.filter((row) => cellEvalRowInRange(row, range))
  const bounds = normalizeRangeBounds(range)
  await db.query(
    `
      DELETE FROM cell_eval
      WHERE workbook_id = $1
        AND sheet_name = $2
        AND row_num BETWEEN $3 AND $4
        AND col_num BETWEEN $5 AND $6
    `,
    [documentId, bounds.sheetName, bounds.rowStart, bounds.rowEnd, bounds.colStart, bounds.colEnd],
  )
  if (nextRowsInRange.length === 0) {
    return
  }
  await persistCellEvalRows(db, documentId, [], nextRowsInRange)
}

export async function persistWorkbookCheckpoint(
  db: Queryable,
  documentId: string,
  revision: number,
  checkpointPayload: WorkbookSnapshot,
  replicaState: EngineReplicaSnapshot | null,
): Promise<void> {
  await db.query(
    `
      INSERT INTO workbook_snapshot (
        workbook_id,
        revision,
        format,
        payload,
        replica_snapshot,
        created_at
      )
      VALUES ($1, $2, $3, $4::jsonb, $5::jsonb, NOW())
      ON CONFLICT (workbook_id, revision)
      DO NOTHING
    `,
    [documentId, revision, WORKBOOK_CHECKPOINT_FORMAT, JSON.stringify(checkpointPayload), JSON.stringify(replicaState)],
  )
  await db.query(
    `
      DELETE FROM workbook_snapshot
      WHERE workbook_id = $1
        AND revision NOT IN (
          SELECT revision
          FROM workbook_snapshot
          WHERE workbook_id = $1
          ORDER BY revision DESC
          LIMIT $2
        )
    `,
    [documentId, WORKBOOK_CHECKPOINT_RETENTION],
  )
}

export async function backfillWorkbookSnapshotsFromInlineState(db: Queryable): Promise<void> {
  await db.query(
    `
      INSERT INTO workbook_snapshot (
        workbook_id,
        revision,
        format,
        payload,
        replica_snapshot,
        created_at
      )
      SELECT
        id,
        head_revision,
        $1,
        snapshot,
        replica_snapshot,
        updated_at
      FROM workbooks
      WHERE snapshot IS NOT NULL
      ON CONFLICT (workbook_id, revision)
      DO NOTHING
    `,
    [WORKBOOK_CHECKPOINT_FORMAT],
  )
}
