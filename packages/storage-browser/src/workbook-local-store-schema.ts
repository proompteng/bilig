import type { Database, SqlValue } from "@sqlite.org/sqlite-wasm";

function readSingleObjectRow(
  db: Database,
  sql: string,
  bind?: readonly SqlValue[],
): Record<string, SqlValue> | null {
  const statement = db.prepare(sql);
  try {
    if (bind) {
      statement.bind([...bind]);
    }
    if (!statement.step()) {
      return null;
    }
    return statement.get({});
  } finally {
    statement.finalize();
  }
}

export function initializeWorkbookLocalStoreSchema(db: Database): void {
  db.exec(`
    PRAGMA journal_mode = WAL;
    PRAGMA synchronous = NORMAL;
    PRAGMA temp_store = MEMORY;
    PRAGMA foreign_keys = ON;

    CREATE TABLE IF NOT EXISTS runtime_state (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      snapshot_json TEXT NOT NULL,
      replica_json TEXT NOT NULL,
      authoritative_revision INTEGER NOT NULL,
      applied_pending_local_seq INTEGER NOT NULL DEFAULT 0,
      updated_at_ms INTEGER NOT NULL
    );

    CREATE TABLE IF NOT EXISTS pending_op (
      op_id TEXT PRIMARY KEY,
      local_seq INTEGER NOT NULL UNIQUE,
      base_revision INTEGER NOT NULL,
      method TEXT NOT NULL,
      args_json TEXT NOT NULL,
      enqueued_at_ms INTEGER NOT NULL,
      submitted_at_ms INTEGER,
      status TEXT NOT NULL CHECK (status IN ('pending', 'submitted'))
    );

    CREATE INDEX IF NOT EXISTS pending_op_local_seq_idx
      ON pending_op(local_seq);

    CREATE TABLE IF NOT EXISTS authoritative_sheet (
      name TEXT PRIMARY KEY,
      sort_order INTEGER NOT NULL,
      freeze_rows INTEGER NOT NULL DEFAULT 0,
      freeze_cols INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS authoritative_cell_input (
      sheet_name TEXT NOT NULL REFERENCES authoritative_sheet(name) ON DELETE CASCADE,
      address TEXT NOT NULL,
      row_num INTEGER NOT NULL,
      col_num INTEGER NOT NULL,
      input_json TEXT,
      formula TEXT,
      format TEXT,
      PRIMARY KEY (sheet_name, address)
    );

    CREATE TABLE IF NOT EXISTS authoritative_cell_render (
      sheet_name TEXT NOT NULL REFERENCES authoritative_sheet(name) ON DELETE CASCADE,
      address TEXT NOT NULL,
      row_num INTEGER NOT NULL,
      col_num INTEGER NOT NULL,
      value_json TEXT NOT NULL,
      flags INTEGER NOT NULL,
      version INTEGER NOT NULL,
      style_id TEXT,
      number_format_id TEXT,
      PRIMARY KEY (sheet_name, address)
    );

    CREATE TABLE IF NOT EXISTS authoritative_row_axis (
      sheet_name TEXT NOT NULL REFERENCES authoritative_sheet(name) ON DELETE CASCADE,
      axis_index INTEGER NOT NULL,
      axis_id TEXT NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (sheet_name, axis_index)
    );

    CREATE TABLE IF NOT EXISTS authoritative_column_axis (
      sheet_name TEXT NOT NULL REFERENCES authoritative_sheet(name) ON DELETE CASCADE,
      axis_index INTEGER NOT NULL,
      axis_id TEXT NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (sheet_name, axis_index)
    );

    CREATE TABLE IF NOT EXISTS authoritative_style (
      style_id TEXT PRIMARY KEY,
      record_json TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS projection_overlay_cell (
      sheet_name TEXT NOT NULL,
      address TEXT NOT NULL,
      row_num INTEGER NOT NULL,
      col_num INTEGER NOT NULL,
      value_json TEXT NOT NULL,
      flags INTEGER NOT NULL,
      version INTEGER NOT NULL,
      input_json TEXT,
      formula TEXT,
      format TEXT,
      style_id TEXT,
      number_format_id TEXT,
      PRIMARY KEY (sheet_name, address)
    );

    CREATE TABLE IF NOT EXISTS projection_overlay_row_axis (
      sheet_name TEXT NOT NULL,
      axis_index INTEGER NOT NULL,
      axis_id TEXT NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (sheet_name, axis_index)
    );

    CREATE TABLE IF NOT EXISTS projection_overlay_column_axis (
      sheet_name TEXT NOT NULL,
      axis_index INTEGER NOT NULL,
      axis_id TEXT NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (sheet_name, axis_index)
    );

    CREATE TABLE IF NOT EXISTS projection_overlay_style (
      style_id TEXT PRIMARY KEY,
      record_json TEXT NOT NULL
    );

    CREATE INDEX IF NOT EXISTS authoritative_cell_render_viewport_idx
      ON authoritative_cell_render(sheet_name, row_num, col_num);

    CREATE INDEX IF NOT EXISTS authoritative_row_axis_viewport_idx
      ON authoritative_row_axis(sheet_name, axis_index);

    CREATE INDEX IF NOT EXISTS authoritative_column_axis_viewport_idx
      ON authoritative_column_axis(sheet_name, axis_index);

    CREATE INDEX IF NOT EXISTS projection_overlay_cell_viewport_idx
      ON projection_overlay_cell(sheet_name, row_num, col_num);

    CREATE INDEX IF NOT EXISTS projection_overlay_row_axis_viewport_idx
      ON projection_overlay_row_axis(sheet_name, axis_index);

    CREATE INDEX IF NOT EXISTS projection_overlay_column_axis_viewport_idx
      ON projection_overlay_column_axis(sheet_name, axis_index);
  `);
  const appliedPendingColumn = readSingleObjectRow(
    db,
    `
      SELECT 1 AS present
        FROM pragma_table_info('runtime_state')
       WHERE name = 'applied_pending_local_seq'
    `,
  );
  if (!appliedPendingColumn) {
    db.exec(`
      ALTER TABLE runtime_state
      ADD COLUMN applied_pending_local_seq INTEGER NOT NULL DEFAULT 0
    `);
  }
  const submittedAtColumn = readSingleObjectRow(
    db,
    `
      SELECT 1 AS present
        FROM pragma_table_info('pending_op')
       WHERE name = 'submitted_at_ms'
    `,
  );
  if (!submittedAtColumn) {
    db.exec(`
      ALTER TABLE pending_op
      ADD COLUMN submitted_at_ms INTEGER
    `);
  }
}
