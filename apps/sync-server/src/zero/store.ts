import type { CellValue, WorkbookSnapshot } from "@bilig/protocol";

export interface QueryResultRow {
  [key: string]: unknown;
}

export interface Queryable {
  query<T extends QueryResultRow = QueryResultRow>(
    text: string,
    values?: unknown[],
  ): Promise<{ rows: T[] }>;
}

export interface MaterializedComputedCell {
  sheetName: string;
  address: string;
  value: CellValue;
  flags: number;
  version: number;
}

export async function ensureZeroSyncSchema(db: Queryable): Promise<void> {
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbooks (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      snapshot JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS sheets (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      name TEXT NOT NULL,
      sort_order INTEGER NOT NULL,
      PRIMARY KEY (workbook_id, name)
    );
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS cells (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      sheet_name TEXT NOT NULL,
      address TEXT NOT NULL,
      input_value JSONB,
      formula TEXT,
      format TEXT,
      PRIMARY KEY (workbook_id, sheet_name, address)
    );
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS computed_cells (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      sheet_name TEXT NOT NULL,
      address TEXT NOT NULL,
      value JSONB NOT NULL,
      flags INTEGER NOT NULL,
      version INTEGER NOT NULL,
      PRIMARY KEY (workbook_id, sheet_name, address)
    );
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS row_metadata (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      sheet_name TEXT NOT NULL,
      start_index INTEGER NOT NULL,
      count INTEGER NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (workbook_id, sheet_name, start_index)
    );
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS column_metadata (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      sheet_name TEXT NOT NULL,
      start_index INTEGER NOT NULL,
      count INTEGER NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (workbook_id, sheet_name, start_index)
    );
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS defined_names (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      name TEXT NOT NULL,
      value JSONB NOT NULL,
      PRIMARY KEY (workbook_id, name)
    );
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_metadata (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      key TEXT NOT NULL,
      value JSONB NOT NULL,
      PRIMARY KEY (workbook_id, key)
    );
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS calculation_settings (
      workbook_id TEXT PRIMARY KEY REFERENCES workbooks(id) ON DELETE CASCADE,
      mode TEXT NOT NULL,
      recalc_epoch BIGINT NOT NULL DEFAULT 0
    );
  `);
  await db.query(
    `CREATE INDEX IF NOT EXISTS sheets_workbook_sort_order_idx ON sheets(workbook_id, sort_order);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS cells_workbook_sheet_idx ON cells(workbook_id, sheet_name);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS computed_cells_workbook_sheet_idx ON computed_cells(workbook_id, sheet_name);`,
  );
}

export function createEmptyWorkbookSnapshot(documentId: string): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: documentId,
    },
    sheets: [
      {
        name: "Sheet1",
        order: 0,
        cells: [],
      },
    ],
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isWorkbookSnapshot(value: unknown): value is WorkbookSnapshot {
  return (
    isRecord(value) &&
    value["version"] === 1 &&
    isRecord(value["workbook"]) &&
    typeof value["workbook"]["name"] === "string" &&
    Array.isArray(value["sheets"])
  );
}

export async function loadWorkbookSnapshot(
  db: Queryable,
  documentId: string,
): Promise<WorkbookSnapshot> {
  const result = await db.query<{ snapshot: unknown }>(
    `SELECT snapshot FROM workbooks WHERE id = $1 LIMIT 1`,
    [documentId],
  );
  const snapshot = result.rows[0]?.snapshot;
  return isWorkbookSnapshot(snapshot) ? snapshot : createEmptyWorkbookSnapshot(documentId);
}

async function clearWorkbookProjection(db: Queryable, documentId: string): Promise<void> {
  await db.query(`DELETE FROM sheets WHERE workbook_id = $1`, [documentId]);
  await db.query(`DELETE FROM cells WHERE workbook_id = $1`, [documentId]);
  await db.query(`DELETE FROM computed_cells WHERE workbook_id = $1`, [documentId]);
  await db.query(`DELETE FROM row_metadata WHERE workbook_id = $1`, [documentId]);
  await db.query(`DELETE FROM column_metadata WHERE workbook_id = $1`, [documentId]);
  await db.query(`DELETE FROM defined_names WHERE workbook_id = $1`, [documentId]);
  await db.query(`DELETE FROM workbook_metadata WHERE workbook_id = $1`, [documentId]);
  await db.query(`DELETE FROM calculation_settings WHERE workbook_id = $1`, [documentId]);
}

export async function persistWorkbookProjection(
  db: Queryable,
  documentId: string,
  snapshot: WorkbookSnapshot,
  computedCells: MaterializedComputedCell[],
): Promise<void> {
  const updatedAt = new Date().toISOString();
  await db.query(
    `
      INSERT INTO workbooks (id, name, snapshot, updated_at)
      VALUES ($1, $2, $3::jsonb, $4::timestamptz)
      ON CONFLICT (id)
      DO UPDATE SET
        name = EXCLUDED.name,
        snapshot = EXCLUDED.snapshot,
        updated_at = EXCLUDED.updated_at
    `,
    [documentId, snapshot.workbook.name, JSON.stringify(snapshot), updatedAt],
  );

  await clearWorkbookProjection(db, documentId);

  const sheetQueries: Promise<unknown>[] = [];
  const cellQueries: Promise<unknown>[] = [];
  const rowMetadataQueries: Promise<unknown>[] = [];
  const columnMetadataQueries: Promise<unknown>[] = [];

  for (const sheet of snapshot.sheets) {
    sheetQueries.push(
      db.query(`INSERT INTO sheets (workbook_id, name, sort_order) VALUES ($1, $2, $3)`, [
        documentId,
        sheet.name,
        sheet.order,
      ]),
    );

    for (const cell of sheet.cells) {
      cellQueries.push(
        db.query(
          `
          INSERT INTO cells (workbook_id, sheet_name, address, input_value, formula, format)
          VALUES ($1, $2, $3, $4::jsonb, $5, $6)
        `,
          [
            documentId,
            sheet.name,
            cell.address,
            cell.formula ? null : JSON.stringify(cell.value ?? null),
            cell.formula ?? null,
            cell.format ?? null,
          ],
        ),
      );
    }

    for (const rowEntry of sheet.metadata?.rowMetadata ?? []) {
      rowMetadataQueries.push(
        db.query(
          `
          INSERT INTO row_metadata (workbook_id, sheet_name, start_index, count, size, hidden)
          VALUES ($1, $2, $3, $4, $5, $6)
        `,
          [
            documentId,
            sheet.name,
            rowEntry.start,
            rowEntry.count,
            rowEntry.size ?? null,
            rowEntry.hidden ?? null,
          ],
        ),
      );
    }

    for (const columnEntry of sheet.metadata?.columnMetadata ?? []) {
      columnMetadataQueries.push(
        db.query(
          `
          INSERT INTO column_metadata (workbook_id, sheet_name, start_index, count, size, hidden)
          VALUES ($1, $2, $3, $4, $5, $6)
        `,
          [
            documentId,
            sheet.name,
            columnEntry.start,
            columnEntry.count,
            columnEntry.size ?? null,
            columnEntry.hidden ?? null,
          ],
        ),
      );
    }
  }

  await Promise.all([
    ...sheetQueries,
    ...cellQueries,
    ...rowMetadataQueries,
    ...columnMetadataQueries,
  ]);

  const computedCellQueries = computedCells.map((cell) =>
    db.query(
      `
        INSERT INTO computed_cells (workbook_id, sheet_name, address, value, flags, version)
        VALUES ($1, $2, $3, $4::jsonb, $5, $6)
      `,
      [
        documentId,
        cell.sheetName,
        cell.address,
        JSON.stringify(cell.value),
        cell.flags,
        cell.version,
      ],
    ),
  );
  await Promise.all(computedCellQueries);

  const definedNameQueries = (snapshot.workbook.metadata?.definedNames ?? []).map((entry) =>
    db.query(`INSERT INTO defined_names (workbook_id, name, value) VALUES ($1, $2, $3::jsonb)`, [
      documentId,
      entry.name,
      JSON.stringify(entry.value),
    ]),
  );
  await Promise.all(definedNameQueries);

  const metadataQueries = (snapshot.workbook.metadata?.properties ?? []).map((entry) =>
    db.query(`INSERT INTO workbook_metadata (workbook_id, key, value) VALUES ($1, $2, $3::jsonb)`, [
      documentId,
      entry.key,
      JSON.stringify(entry.value),
    ]),
  );
  await Promise.all(metadataQueries);

  await db.query(
    `
      INSERT INTO calculation_settings (workbook_id, mode, recalc_epoch)
      VALUES ($1, $2, $3)
    `,
    [
      documentId,
      snapshot.workbook.metadata?.calculationSettings?.mode ?? "automatic",
      snapshot.workbook.metadata?.volatileContext?.recalcEpoch ?? 0,
    ],
  );
}
