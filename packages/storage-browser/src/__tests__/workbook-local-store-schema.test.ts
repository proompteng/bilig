import sqlite3InitModule, { type Database, type SqlValue } from "@sqlite.org/sqlite-wasm";
import { describe, expect, it } from "vitest";

import { readWorkbookViewportProjection } from "../workbook-local-store-projection.js";
import { initializeWorkbookLocalStoreSchema } from "../workbook-local-store-schema.js";

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

async function createLegacyWorkbookDb(): Promise<Database> {
  const sqlite3 = await sqlite3InitModule();
  const db = new sqlite3.oo1.DB(":memory:", "c");
  db.exec(`
    CREATE TABLE runtime_state (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      snapshot_json TEXT NOT NULL,
      replica_json TEXT NOT NULL,
      authoritative_revision INTEGER NOT NULL,
      updated_at_ms INTEGER NOT NULL
    );

    CREATE TABLE pending_op (
      op_id TEXT PRIMARY KEY,
      local_seq INTEGER NOT NULL UNIQUE,
      base_revision INTEGER NOT NULL,
      method TEXT NOT NULL,
      args_json TEXT NOT NULL,
      enqueued_at_ms INTEGER NOT NULL,
      status TEXT NOT NULL CHECK (status IN ('pending', 'submitted'))
    );

    CREATE TABLE authoritative_sheet (
      name TEXT PRIMARY KEY,
      sort_order INTEGER NOT NULL,
      freeze_rows INTEGER NOT NULL DEFAULT 0,
      freeze_cols INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE authoritative_cell_input (
      sheet_name TEXT NOT NULL,
      address TEXT NOT NULL,
      row_num INTEGER NOT NULL,
      col_num INTEGER NOT NULL,
      input_json TEXT,
      formula TEXT,
      format TEXT,
      PRIMARY KEY (sheet_name, address)
    );

    CREATE TABLE authoritative_cell_render (
      sheet_name TEXT NOT NULL,
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

    CREATE TABLE authoritative_row_axis (
      sheet_name TEXT NOT NULL,
      axis_index INTEGER NOT NULL,
      axis_id TEXT NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (sheet_name, axis_index)
    );

    CREATE TABLE authoritative_column_axis (
      sheet_name TEXT NOT NULL,
      axis_index INTEGER NOT NULL,
      axis_id TEXT NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (sheet_name, axis_index)
    );

    CREATE TABLE authoritative_style (
      style_id TEXT PRIMARY KEY,
      record_json TEXT NOT NULL
    );

    CREATE TABLE projection_overlay_cell (
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

    CREATE TABLE projection_overlay_row_axis (
      sheet_name TEXT NOT NULL,
      axis_index INTEGER NOT NULL,
      axis_id TEXT NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (sheet_name, axis_index)
    );

    CREATE TABLE projection_overlay_column_axis (
      sheet_name TEXT NOT NULL,
      axis_index INTEGER NOT NULL,
      axis_id TEXT NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (sheet_name, axis_index)
    );

    CREATE TABLE projection_overlay_style (
      style_id TEXT PRIMARY KEY,
      record_json TEXT NOT NULL
    );
  `);
  return db;
}

describe("workbook-local-store schema migration", () => {
  it("backfills sheet ids from the legacy snapshot into authoritative and overlay tables", async () => {
    const db = await createLegacyWorkbookDb();
    try {
      db.exec(
        `
          INSERT INTO runtime_state (
            id,
            snapshot_json,
            replica_json,
            authoritative_revision,
            updated_at_ms
          )
          VALUES (1, ?, '{}', 4, 1);

          INSERT INTO pending_op (
            op_id,
            local_seq,
            base_revision,
            method,
            args_json,
            enqueued_at_ms,
            status
          )
          VALUES ('op-1', 1, 4, 'setCellValue', '[]', 1, 'pending');

          INSERT INTO authoritative_sheet (name, sort_order, freeze_rows, freeze_cols)
          VALUES ('Revenue', 0, 0, 0);

          INSERT INTO authoritative_cell_input (
            sheet_name,
            address,
            row_num,
            col_num,
            input_json
          )
          VALUES ('Revenue', 'A1', 0, 0, '42');

          INSERT INTO authoritative_cell_render (
            sheet_name,
            address,
            row_num,
            col_num,
            value_json,
            flags,
            version
          )
          VALUES ('Revenue', 'A1', 0, 0, '{"tag":1,"value":42}', 0, 1);

          INSERT INTO projection_overlay_cell (
            sheet_name,
            address,
            row_num,
            col_num,
            value_json,
            flags,
            version,
            input_json
          )
          VALUES ('Revenue', 'A1', 0, 0, '{"tag":1,"value":43}', 0, 2, '43');
        `,
        {
          bind: [
            JSON.stringify({
              version: 1,
              workbook: { name: "legacy-doc" },
              sheets: [{ id: 7, name: "Revenue", order: 0, cells: [] }],
            }),
          ],
        },
      );

      initializeWorkbookLocalStoreSchema(db);

      expect(
        readSingleObjectRow(
          db,
          `
            SELECT applied_pending_local_seq AS appliedPendingLocalSeq
              FROM runtime_state
             WHERE id = 1
          `,
        ),
      ).toEqual({ appliedPendingLocalSeq: 0 });
      expect(
        readSingleObjectRow(
          db,
          `
            SELECT submitted_at_ms AS submittedAtMs
              FROM pending_op
             WHERE op_id = 'op-1'
          `,
        ),
      ).toEqual({ submittedAtMs: null });
      expect(
        readSingleObjectRow(
          db,
          `
            SELECT sheet_id AS sheetId
              FROM authoritative_sheet
             WHERE name = 'Revenue'
          `,
        ),
      ).toEqual({ sheetId: 7 });
      expect(
        readSingleObjectRow(
          db,
          `
            SELECT sheet_id AS sheetId
              FROM authoritative_cell_input
             WHERE sheet_name = 'Revenue' AND address = 'A1'
          `,
        ),
      ).toEqual({ sheetId: 7 });
      expect(
        readSingleObjectRow(
          db,
          `
            SELECT sheet_id AS sheetId
              FROM projection_overlay_cell
             WHERE sheet_name = 'Revenue' AND address = 'A1'
          `,
        ),
      ).toEqual({ sheetId: 7 });
      expect(
        readWorkbookViewportProjection(db, "Revenue", {
          rowStart: 0,
          rowEnd: 0,
          colStart: 0,
          colEnd: 0,
        }),
      ).toMatchObject({
        sheetId: 7,
        sheetName: "Revenue",
        cells: [
          {
            snapshot: {
              address: "A1",
              input: 43,
              value: { tag: 1, value: 43 },
            },
          },
        ],
      });
    } finally {
      db.close();
    }
  });

  it("assigns deterministic fallback sheet ids from sheet order when the legacy snapshot omits them", async () => {
    const db = await createLegacyWorkbookDb();
    try {
      db.exec(
        `
          INSERT INTO runtime_state (
            id,
            snapshot_json,
            replica_json,
            authoritative_revision,
            updated_at_ms
          )
          VALUES (1, ?, '{}', 0, 1);

          INSERT INTO authoritative_sheet (name, sort_order, freeze_rows, freeze_cols)
          VALUES ('Alpha', 0, 0, 0), ('Beta', 1, 0, 0);
        `,
        {
          bind: [
            JSON.stringify({
              version: 1,
              workbook: { name: "legacy-doc" },
              sheets: [
                { name: "Alpha", order: 0, cells: [] },
                { name: "Beta", order: 1, cells: [] },
              ],
            }),
          ],
        },
      );

      initializeWorkbookLocalStoreSchema(db);

      expect(
        readSingleObjectRow(
          db,
          `
            SELECT sheet_id AS alphaSheetId
              FROM authoritative_sheet
             WHERE name = 'Alpha'
          `,
        ),
      ).toEqual({ alphaSheetId: 1 });
      expect(
        readSingleObjectRow(
          db,
          `
            SELECT sheet_id AS betaSheetId
              FROM authoritative_sheet
             WHERE name = 'Beta'
          `,
        ),
      ).toEqual({ betaSheetId: 2 });
    } finally {
      db.close();
    }
  });
});
