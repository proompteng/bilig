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

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function ensureColumn(
  db: Database,
  tableName: string,
  columnName: string,
  definition: string,
): void {
  const column = readSingleObjectRow(
    db,
    `
      SELECT 1 AS present
        FROM pragma_table_info(?)
       WHERE name = ?
    `,
    [tableName, columnName],
  );
  if (!column) {
    db.exec(`ALTER TABLE ${tableName} ADD COLUMN ${columnName} ${definition}`);
  }
}

function parseSheetIdsFromSnapshotJson(snapshotJson: string | null): Map<string, number> {
  if (typeof snapshotJson !== "string") {
    return new Map();
  }
  try {
    const parsed = JSON.parse(snapshotJson) as unknown;
    if (!isRecord(parsed) || !Array.isArray(parsed["sheets"])) {
      return new Map();
    }
    const sheetIds = new Map<string, number>();
    parsed["sheets"].forEach((sheet, index) => {
      if (!isRecord(sheet) || typeof sheet["name"] !== "string") {
        return;
      }
      const id = typeof sheet["id"] === "number" ? sheet["id"] : index + 1;
      if (!sheetIds.has(sheet["name"])) {
        sheetIds.set(sheet["name"], id);
      }
    });
    return sheetIds;
  } catch {
    return new Map();
  }
}

function backfillSheetIds(db: Database): void {
  const snapshotRow = readSingleObjectRow(
    db,
    `
      SELECT snapshot_json AS snapshotJson
        FROM runtime_state
       WHERE id = 1
    `,
  );
  const legacySheetIds = parseSheetIdsFromSnapshotJson(
    typeof snapshotRow?.["snapshotJson"] === "string" ? snapshotRow["snapshotJson"] : null,
  );

  const rows: Array<{ name: string; sheetId: number | null }> = [];
  const readSheets = db.prepare(
    `
      SELECT name,
             sheet_id AS sheetId,
             sort_order AS sortOrder
        FROM authoritative_sheet
       ORDER BY sort_order ASC, name ASC
    `,
  );
  try {
    while (readSheets.step()) {
      const row = readSheets.get({});
      if (typeof row["name"] !== "string") {
        continue;
      }
      rows.push({
        name: row["name"],
        sheetId: typeof row["sheetId"] === "number" ? row["sheetId"] : null,
      });
    }
  } finally {
    readSheets.finalize();
  }
  if (rows.length === 0) {
    return;
  }

  const usedIds = new Set(rows.flatMap((row) => (row.sheetId === null ? [] : [row.sheetId])));
  let nextAutoId = usedIds.size > 0 ? Math.max(...usedIds) + 1 : 1;
  const assignments = new Map<string, number>();
  rows.forEach((row) => {
    if (row.sheetId !== null) {
      assignments.set(row.name, row.sheetId);
      return;
    }
    const legacyId = legacySheetIds.get(row.name);
    if (legacyId !== undefined && !usedIds.has(legacyId)) {
      assignments.set(row.name, legacyId);
      usedIds.add(legacyId);
      return;
    }
    while (usedIds.has(nextAutoId)) {
      nextAutoId += 1;
    }
    assignments.set(row.name, nextAutoId);
    usedIds.add(nextAutoId);
    nextAutoId += 1;
  });

  const updateSheet = db.prepare(
    `
      UPDATE authoritative_sheet
         SET sheet_id = ?
       WHERE name = ?
         AND sheet_id IS NULL
    `,
  );
  const updateChildTable = (tableName: string) =>
    db.prepare(
      `
        UPDATE ${tableName}
           SET sheet_id = ?
         WHERE sheet_name = ?
           AND sheet_id IS NULL
      `,
    );
  const childTables = [
    updateChildTable("authoritative_cell_input"),
    updateChildTable("authoritative_cell_render"),
    updateChildTable("authoritative_row_axis"),
    updateChildTable("authoritative_column_axis"),
    updateChildTable("projection_overlay_cell"),
    updateChildTable("projection_overlay_row_axis"),
    updateChildTable("projection_overlay_column_axis"),
  ];
  try {
    assignments.forEach((sheetId, sheetName) => {
      updateSheet.bind([sheetId, sheetName]);
      updateSheet.step();
      updateSheet.reset();
      childTables.forEach((statement) => {
        statement.bind([sheetId, sheetName]);
        statement.step();
        statement.reset();
      });
    });
  } finally {
    updateSheet.finalize();
    childTables.forEach((statement) => statement.finalize());
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
      sheet_id INTEGER NOT NULL UNIQUE,
      sort_order INTEGER NOT NULL,
      freeze_rows INTEGER NOT NULL DEFAULT 0,
      freeze_cols INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS authoritative_cell_input (
      sheet_id INTEGER NOT NULL,
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
      sheet_id INTEGER NOT NULL,
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
      sheet_id INTEGER NOT NULL,
      sheet_name TEXT NOT NULL REFERENCES authoritative_sheet(name) ON DELETE CASCADE,
      axis_index INTEGER NOT NULL,
      axis_id TEXT NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (sheet_name, axis_index)
    );

    CREATE TABLE IF NOT EXISTS authoritative_column_axis (
      sheet_id INTEGER NOT NULL,
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
      sheet_id INTEGER NOT NULL,
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
      sheet_id INTEGER NOT NULL,
      sheet_name TEXT NOT NULL,
      axis_index INTEGER NOT NULL,
      axis_id TEXT NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (sheet_name, axis_index)
    );

    CREATE TABLE IF NOT EXISTS projection_overlay_column_axis (
      sheet_id INTEGER NOT NULL,
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
  `);

  ensureColumn(db, "runtime_state", "applied_pending_local_seq", "INTEGER NOT NULL DEFAULT 0");
  ensureColumn(db, "pending_op", "submitted_at_ms", "INTEGER");
  ensureColumn(db, "authoritative_sheet", "sheet_id", "INTEGER");
  ensureColumn(db, "authoritative_cell_input", "sheet_id", "INTEGER");
  ensureColumn(db, "authoritative_cell_render", "sheet_id", "INTEGER");
  ensureColumn(db, "authoritative_row_axis", "sheet_id", "INTEGER");
  ensureColumn(db, "authoritative_column_axis", "sheet_id", "INTEGER");
  ensureColumn(db, "projection_overlay_cell", "sheet_id", "INTEGER");
  ensureColumn(db, "projection_overlay_row_axis", "sheet_id", "INTEGER");
  ensureColumn(db, "projection_overlay_column_axis", "sheet_id", "INTEGER");

  backfillSheetIds(db);

  db.exec(`
    CREATE UNIQUE INDEX IF NOT EXISTS authoritative_sheet_sheet_id_idx
      ON authoritative_sheet(sheet_id);

    CREATE INDEX IF NOT EXISTS authoritative_cell_render_viewport_idx
      ON authoritative_cell_render(sheet_name, row_num, col_num);
    CREATE INDEX IF NOT EXISTS authoritative_cell_render_sheet_id_viewport_idx
      ON authoritative_cell_render(sheet_id, row_num, col_num);
    CREATE INDEX IF NOT EXISTS authoritative_cell_input_sheet_id_address_idx
      ON authoritative_cell_input(sheet_id, address);

    CREATE INDEX IF NOT EXISTS authoritative_row_axis_viewport_idx
      ON authoritative_row_axis(sheet_name, axis_index);
    CREATE INDEX IF NOT EXISTS authoritative_row_axis_sheet_id_viewport_idx
      ON authoritative_row_axis(sheet_id, axis_index);

    CREATE INDEX IF NOT EXISTS authoritative_column_axis_viewport_idx
      ON authoritative_column_axis(sheet_name, axis_index);
    CREATE INDEX IF NOT EXISTS authoritative_column_axis_sheet_id_viewport_idx
      ON authoritative_column_axis(sheet_id, axis_index);

    CREATE INDEX IF NOT EXISTS projection_overlay_cell_viewport_idx
      ON projection_overlay_cell(sheet_name, row_num, col_num);
    CREATE INDEX IF NOT EXISTS projection_overlay_cell_sheet_id_viewport_idx
      ON projection_overlay_cell(sheet_id, row_num, col_num);

    CREATE INDEX IF NOT EXISTS projection_overlay_row_axis_viewport_idx
      ON projection_overlay_row_axis(sheet_name, axis_index);
    CREATE INDEX IF NOT EXISTS projection_overlay_row_axis_sheet_id_viewport_idx
      ON projection_overlay_row_axis(sheet_id, axis_index);

    CREATE INDEX IF NOT EXISTS projection_overlay_column_axis_viewport_idx
      ON projection_overlay_column_axis(sheet_name, axis_index);
    CREATE INDEX IF NOT EXISTS projection_overlay_column_axis_sheet_id_viewport_idx
      ON projection_overlay_column_axis(sheet_id, axis_index);
  `);
}
