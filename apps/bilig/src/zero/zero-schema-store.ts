import type { Queryable } from './store.js'
import { ensureZeroSchemaTable } from './zero-schema-ddl.js'

const workbookReferenceConstraint = 'REFERENCES workbooks(id) ON DELETE CASCADE'

export async function ensureZeroSyncSchema(db: Queryable): Promise<void> {
  await ensureZeroSchemaTable(db, 'workbooks', {
    extraColumnSql: [
      "snapshot JSONB NOT NULL DEFAULT 'null'::jsonb",
      'replica_snapshot JSONB',
      'source_projection_version BIGINT NOT NULL DEFAULT 2',
    ],
    columnOverrides: {
      ownerUserId: { defaultSql: "'system'" },
      headRevision: { defaultSql: '0' },
      calculatedRevision: { defaultSql: '0' },
      calcMode: { defaultSql: "'automatic'" },
      compatibilityMode: { defaultSql: "'excel-modern'" },
      recalcEpoch: { defaultSql: '0' },
      createdAt: { dataType: 'TIMESTAMPTZ', defaultSql: 'NOW()' },
      updatedAt: { dataType: 'TIMESTAMPTZ', defaultSql: 'NOW()' },
    },
  })

  await ensureZeroSchemaTable(db, 'sheets', {
    columnOverrides: {
      workbookId: { constraintSql: workbookReferenceConstraint },
      sheetId: { dataType: 'INTEGER', constraintSql: 'CONSTRAINT sheets_sheet_id_positive_chk CHECK (sheet_id > 0)' },
      sortOrder: { dataType: 'INTEGER' },
      freezeRows: { dataType: 'INTEGER', defaultSql: '0' },
      freezeCols: { dataType: 'INTEGER', defaultSql: '0' },
      createdAt: { dataType: 'TIMESTAMPTZ', defaultSql: 'NOW()' },
      updatedAt: { dataType: 'TIMESTAMPTZ', defaultSql: 'NOW()' },
    },
  })

  await ensureZeroSchemaTable(db, 'cells', {
    columnOverrides: {
      workbookId: { constraintSql: workbookReferenceConstraint },
      rowNum: { dataType: 'INTEGER' },
      colNum: { dataType: 'INTEGER' },
      sourceRevision: { defaultSql: '0' },
      updatedBy: { defaultSql: "'system'" },
      updatedAt: { dataType: 'TIMESTAMPTZ', defaultSql: 'NOW()' },
    },
  })

  await ensureZeroSchemaTable(db, 'cell_eval', {
    columnOverrides: {
      workbookId: { constraintSql: workbookReferenceConstraint },
      rowNum: { dataType: 'INTEGER' },
      colNum: { dataType: 'INTEGER' },
      flags: { dataType: 'INTEGER' },
      version: { dataType: 'INTEGER' },
      calcRevision: { defaultSql: '0' },
      updatedAt: { dataType: 'TIMESTAMPTZ', defaultSql: 'NOW()' },
    },
  })

  await ensureZeroSchemaTable(db, 'row_metadata', {
    columnOverrides: {
      workbookId: { constraintSql: workbookReferenceConstraint },
      startIndex: { dataType: 'INTEGER' },
      count: { dataType: 'INTEGER' },
      size: { dataType: 'INTEGER' },
      sourceRevision: { defaultSql: '0' },
      updatedAt: { dataType: 'TIMESTAMPTZ', defaultSql: 'NOW()' },
    },
  })

  await ensureZeroSchemaTable(db, 'column_metadata', {
    columnOverrides: {
      workbookId: { constraintSql: workbookReferenceConstraint },
      startIndex: { dataType: 'INTEGER' },
      count: { dataType: 'INTEGER' },
      size: { dataType: 'INTEGER' },
      sourceRevision: { defaultSql: '0' },
      updatedAt: { dataType: 'TIMESTAMPTZ', defaultSql: 'NOW()' },
    },
  })

  await ensureZeroSchemaTable(db, 'defined_names', {
    columnOverrides: {
      workbookId: { constraintSql: workbookReferenceConstraint },
    },
  })
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_metadata (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      key TEXT NOT NULL,
      value JSONB NOT NULL,
      PRIMARY KEY (workbook_id, key)
    );
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS calculation_settings (
      workbook_id TEXT PRIMARY KEY REFERENCES workbooks(id) ON DELETE CASCADE,
      mode TEXT NOT NULL,
      recalc_epoch BIGINT NOT NULL DEFAULT 0
    );
  `)

  await ensureZeroSchemaTable(db, 'cell_styles', {
    columnOverrides: {
      workbookId: { constraintSql: workbookReferenceConstraint },
      createdAt: { dataType: 'TIMESTAMPTZ', defaultSql: 'NOW()' },
    },
  })
  await ensureZeroSchemaTable(db, 'cell_number_formats', {
    columnOverrides: {
      workbookId: { constraintSql: workbookReferenceConstraint },
      createdAt: { dataType: 'TIMESTAMPTZ', defaultSql: 'NOW()' },
    },
  })
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_event (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      revision BIGINT NOT NULL,
      actor_user_id TEXT NOT NULL,
      client_mutation_id TEXT,
      txn_json JSONB NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      PRIMARY KEY (workbook_id, revision)
    );
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS recalc_job (
      id TEXT PRIMARY KEY,
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      from_revision BIGINT NOT NULL,
      to_revision BIGINT NOT NULL,
      dirty_regions_json JSONB,
      status TEXT NOT NULL,
      attempts INTEGER NOT NULL DEFAULT 0,
      lease_until TIMESTAMPTZ,
      lease_owner TEXT,
      last_error TEXT,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `)

  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_snapshot (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      revision BIGINT NOT NULL,
      format TEXT NOT NULL,
      payload JSONB NOT NULL,
      replica_snapshot JSONB,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      PRIMARY KEY (workbook_id, revision)
    );
  `)

  await db.query(`CREATE INDEX IF NOT EXISTS sheets_workbook_sort_order_idx ON sheets(workbook_id, sort_order);`)
  await db.query(`CREATE UNIQUE INDEX IF NOT EXISTS sheets_workbook_sheet_id_idx ON sheets(workbook_id, sheet_id);`)
  await db.query(`CREATE INDEX IF NOT EXISTS cells_workbook_sheet_idx ON cells(workbook_id, sheet_name);`)
  await db.query(`CREATE INDEX IF NOT EXISTS cells_workbook_sheet_row_col_idx ON cells(workbook_id, sheet_name, row_num, col_num);`)
  await db.query(`CREATE INDEX IF NOT EXISTS cell_eval_workbook_sheet_idx ON cell_eval(workbook_id, sheet_name);`)
  await db.query(`CREATE INDEX IF NOT EXISTS cell_eval_workbook_sheet_row_col_idx ON cell_eval(workbook_id, sheet_name, row_num, col_num);`)
  await db.query(`CREATE INDEX IF NOT EXISTS row_metadata_workbook_sheet_idx ON row_metadata(workbook_id, sheet_name, start_index);`)
  await db.query(`CREATE INDEX IF NOT EXISTS column_metadata_workbook_sheet_idx ON column_metadata(workbook_id, sheet_name, start_index);`)
  await db.query(`CREATE INDEX IF NOT EXISTS recalc_job_status_lease_created_idx ON recalc_job(status, lease_until, created_at);`)
  await db.query(
    `CREATE UNIQUE INDEX IF NOT EXISTS workbook_event_workbook_client_mutation_idx ON workbook_event(workbook_id, client_mutation_id) WHERE client_mutation_id IS NOT NULL;`,
  )
  await db.query(`CREATE INDEX IF NOT EXISTS workbook_event_workbook_created_idx ON workbook_event(workbook_id, created_at);`)
  await db.query(`CREATE INDEX IF NOT EXISTS workbook_snapshot_workbook_revision_idx ON workbook_snapshot(workbook_id, revision DESC);`)
}
