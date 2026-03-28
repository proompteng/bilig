import type { EngineReplicaSnapshot } from "@bilig/core";
import { parseCellAddress } from "@bilig/formula";
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
  rowNum: number;
  colNum: number;
  address: string;
  value: CellValue;
  flags: number;
  version: number;
}

export interface WorkbookRuntimeState {
  snapshot: WorkbookSnapshot;
  replicaSnapshot: EngineReplicaSnapshot | null;
  headRevision: number;
  ownerUserId: string;
}

export interface PersistWorkbookProjectionOptions {
  replicaSnapshot?: EngineReplicaSnapshot | null;
  revision?: number;
  updatedBy?: string;
  ownerUserId?: string;
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
  await db.query(
    `ALTER TABLE workbooks ADD COLUMN IF NOT EXISTS owner_user_id TEXT NOT NULL DEFAULT 'system';`,
  );
  await db.query(
    `ALTER TABLE workbooks ADD COLUMN IF NOT EXISTS head_revision BIGINT NOT NULL DEFAULT 0;`,
  );
  await db.query(
    `ALTER TABLE workbooks ADD COLUMN IF NOT EXISTS calculated_revision BIGINT NOT NULL DEFAULT 0;`,
  );
  await db.query(
    `ALTER TABLE workbooks ADD COLUMN IF NOT EXISTS calc_mode TEXT NOT NULL DEFAULT 'automatic';`,
  );
  await db.query(
    `ALTER TABLE workbooks ADD COLUMN IF NOT EXISTS compatibility_mode TEXT NOT NULL DEFAULT 'excel-modern';`,
  );
  await db.query(
    `ALTER TABLE workbooks ADD COLUMN IF NOT EXISTS recalc_epoch BIGINT NOT NULL DEFAULT 0;`,
  );
  await db.query(`ALTER TABLE workbooks ADD COLUMN IF NOT EXISTS replica_snapshot JSONB;`);
  await db.query(
    `ALTER TABLE workbooks ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ NOT NULL DEFAULT NOW();`,
  );

  await db.query(`
    CREATE TABLE IF NOT EXISTS sheets (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      name TEXT NOT NULL,
      sort_order INTEGER NOT NULL,
      PRIMARY KEY (workbook_id, name)
    );
  `);
  await db.query(
    `ALTER TABLE sheets ADD COLUMN IF NOT EXISTS freeze_rows INTEGER NOT NULL DEFAULT 0;`,
  );
  await db.query(
    `ALTER TABLE sheets ADD COLUMN IF NOT EXISTS freeze_cols INTEGER NOT NULL DEFAULT 0;`,
  );
  await db.query(
    `ALTER TABLE sheets ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ NOT NULL DEFAULT NOW();`,
  );
  await db.query(
    `ALTER TABLE sheets ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW();`,
  );

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
  await db.query(`ALTER TABLE cells ADD COLUMN IF NOT EXISTS row_num INTEGER;`);
  await db.query(`ALTER TABLE cells ADD COLUMN IF NOT EXISTS col_num INTEGER;`);
  await db.query(`ALTER TABLE cells ADD COLUMN IF NOT EXISTS explicit_format_id TEXT;`);
  await db.query(
    `ALTER TABLE cells ADD COLUMN IF NOT EXISTS source_revision BIGINT NOT NULL DEFAULT 0;`,
  );
  await db.query(
    `ALTER TABLE cells ADD COLUMN IF NOT EXISTS updated_by TEXT NOT NULL DEFAULT 'system';`,
  );
  await db.query(
    `ALTER TABLE cells ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW();`,
  );

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
  await db.query(`ALTER TABLE computed_cells ADD COLUMN IF NOT EXISTS row_num INTEGER;`);
  await db.query(`ALTER TABLE computed_cells ADD COLUMN IF NOT EXISTS col_num INTEGER;`);
  await db.query(
    `ALTER TABLE computed_cells ADD COLUMN IF NOT EXISTS calc_revision BIGINT NOT NULL DEFAULT 0;`,
  );
  await db.query(
    `ALTER TABLE computed_cells ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW();`,
  );

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
  await db.query(
    `ALTER TABLE row_metadata ADD COLUMN IF NOT EXISTS source_revision BIGINT NOT NULL DEFAULT 0;`,
  );
  await db.query(
    `ALTER TABLE row_metadata ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW();`,
  );

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
  await db.query(
    `ALTER TABLE column_metadata ADD COLUMN IF NOT EXISTS source_revision BIGINT NOT NULL DEFAULT 0;`,
  );
  await db.query(
    `ALTER TABLE column_metadata ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW();`,
  );

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

  await db.query(`
    CREATE TABLE IF NOT EXISTS cell_styles (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      style_id TEXT NOT NULL,
      record_json JSONB NOT NULL,
      hash TEXT NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      PRIMARY KEY (workbook_id, style_id)
    );
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS cell_number_formats (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      format_id TEXT NOT NULL,
      code TEXT NOT NULL,
      kind TEXT NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      PRIMARY KEY (workbook_id, format_id)
    );
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS sheet_style_ranges (
      id TEXT PRIMARY KEY,
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      sheet_name TEXT NOT NULL,
      start_row INTEGER NOT NULL,
      end_row INTEGER NOT NULL,
      start_col INTEGER NOT NULL,
      end_col INTEGER NOT NULL,
      style_id TEXT NOT NULL,
      source_revision BIGINT NOT NULL DEFAULT 0,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS sheet_format_ranges (
      id TEXT PRIMARY KEY,
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      sheet_name TEXT NOT NULL,
      start_row INTEGER NOT NULL,
      end_row INTEGER NOT NULL,
      start_col INTEGER NOT NULL,
      end_col INTEGER NOT NULL,
      format_id TEXT NOT NULL,
      source_revision BIGINT NOT NULL DEFAULT 0,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `);

  await db.query(
    `CREATE INDEX IF NOT EXISTS sheets_workbook_sort_order_idx ON sheets(workbook_id, sort_order);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS cells_workbook_sheet_idx ON cells(workbook_id, sheet_name);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS cells_workbook_sheet_row_col_idx ON cells(workbook_id, sheet_name, row_num, col_num);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS computed_cells_workbook_sheet_idx ON computed_cells(workbook_id, sheet_name);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS computed_cells_workbook_sheet_row_col_idx ON computed_cells(workbook_id, sheet_name, row_num, col_num);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS row_metadata_workbook_sheet_idx ON row_metadata(workbook_id, sheet_name, start_index);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS column_metadata_workbook_sheet_idx ON column_metadata(workbook_id, sheet_name, start_index);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS sheet_style_ranges_workbook_sheet_idx ON sheet_style_ranges(workbook_id, sheet_name, start_row, end_row, start_col, end_col);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS sheet_format_ranges_workbook_sheet_idx ON sheet_format_ranges(workbook_id, sheet_name, start_row, end_row, start_col, end_col);`,
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

function isEngineReplicaSnapshot(value: unknown): value is EngineReplicaSnapshot {
  return (
    isRecord(value) &&
    isRecord(value["replica"]) &&
    Array.isArray(value["entityVersions"]) &&
    Array.isArray(value["sheetDeleteVersions"])
  );
}

export async function loadWorkbookState(
  db: Queryable,
  documentId: string,
): Promise<WorkbookRuntimeState> {
  const result = await db.query<{
    snapshot: unknown;
    replica_snapshot: unknown;
    head_revision: number | string | null;
    owner_user_id: string | null;
  }>(
    `SELECT snapshot, replica_snapshot, head_revision, owner_user_id FROM workbooks WHERE id = $1 LIMIT 1`,
    [documentId],
  );
  const row = result.rows[0];
  return {
    snapshot: isWorkbookSnapshot(row?.snapshot)
      ? row.snapshot
      : createEmptyWorkbookSnapshot(documentId),
    replicaSnapshot: isEngineReplicaSnapshot(row?.replica_snapshot) ? row.replica_snapshot : null,
    headRevision:
      typeof row?.head_revision === "number"
        ? row.head_revision
        : typeof row?.head_revision === "string"
          ? Number.parseInt(row.head_revision, 10) || 0
          : 0,
    ownerUserId: row?.owner_user_id ?? "system",
  };
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
  await db.query(`DELETE FROM cell_styles WHERE workbook_id = $1`, [documentId]);
  await db.query(`DELETE FROM cell_number_formats WHERE workbook_id = $1`, [documentId]);
  await db.query(`DELETE FROM sheet_style_ranges WHERE workbook_id = $1`, [documentId]);
  await db.query(`DELETE FROM sheet_format_ranges WHERE workbook_id = $1`, [documentId]);
}

function rangeId(
  prefix: "style-range" | "format-range",
  sheetName: string,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): string {
  return `${prefix}:${sheetName}:${startRow}:${startCol}:${endRow}:${endCol}`;
}

export async function persistWorkbookProjection(
  db: Queryable,
  documentId: string,
  snapshot: WorkbookSnapshot,
  computedCells: MaterializedComputedCell[],
  options: PersistWorkbookProjectionOptions = {},
): Promise<void> {
  const updatedAt = new Date().toISOString();
  const revision = options.revision ?? 0;
  const updatedBy = options.updatedBy ?? "system";
  const ownerUserId = options.ownerUserId ?? updatedBy;
  const calcSettings = snapshot.workbook.metadata?.calculationSettings;
  const recalcEpoch = snapshot.workbook.metadata?.volatileContext?.recalcEpoch ?? 0;

  await db.query(
    `
      INSERT INTO workbooks (
        id,
        name,
        owner_user_id,
        head_revision,
        calculated_revision,
        calc_mode,
        compatibility_mode,
        recalc_epoch,
        snapshot,
        replica_snapshot,
        updated_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9::jsonb, $10::jsonb, $11::timestamptz)
      ON CONFLICT (id)
      DO UPDATE SET
        name = EXCLUDED.name,
        owner_user_id = EXCLUDED.owner_user_id,
        head_revision = EXCLUDED.head_revision,
        calculated_revision = EXCLUDED.calculated_revision,
        calc_mode = EXCLUDED.calc_mode,
        compatibility_mode = EXCLUDED.compatibility_mode,
        recalc_epoch = EXCLUDED.recalc_epoch,
        snapshot = EXCLUDED.snapshot,
        replica_snapshot = EXCLUDED.replica_snapshot,
        updated_at = EXCLUDED.updated_at
    `,
    [
      documentId,
      snapshot.workbook.name,
      ownerUserId,
      revision,
      revision,
      calcSettings?.mode ?? "automatic",
      calcSettings?.compatibilityMode ?? "excel-modern",
      recalcEpoch,
      JSON.stringify(snapshot),
      JSON.stringify(options.replicaSnapshot ?? null),
      updatedAt,
    ],
  );

  await clearWorkbookProjection(db, documentId);

  const sheetQueries: Promise<unknown>[] = [];
  const cellQueries: Promise<unknown>[] = [];
  const rowMetadataQueries: Promise<unknown>[] = [];
  const columnMetadataQueries: Promise<unknown>[] = [];
  const styleQueries: Promise<unknown>[] = [];
  const formatQueries: Promise<unknown>[] = [];
  const styleRangeQueries: Promise<unknown>[] = [];
  const formatRangeQueries: Promise<unknown>[] = [];

  for (const sheet of snapshot.sheets) {
    sheetQueries.push(
      db.query(
        `
          INSERT INTO sheets (
            workbook_id,
            name,
            sort_order,
            freeze_rows,
            freeze_cols,
            updated_at
          )
          VALUES ($1, $2, $3, $4, $5, $6::timestamptz)
        `,
        [
          documentId,
          sheet.name,
          sheet.order,
          sheet.metadata?.freezePane?.rows ?? 0,
          sheet.metadata?.freezePane?.cols ?? 0,
          updatedAt,
        ],
      ),
    );

    for (const cell of sheet.cells) {
      const parsed = parseCellAddress(cell.address, sheet.name);
      cellQueries.push(
        db.query(
          `
            INSERT INTO cells (
              workbook_id,
              sheet_name,
              address,
              row_num,
              col_num,
              input_value,
              formula,
              format,
              explicit_format_id,
              source_revision,
              updated_by,
              updated_at
            )
            VALUES ($1, $2, $3, $4, $5, $6::jsonb, $7, $8, $9, $10, $11, $12::timestamptz)
          `,
          [
            documentId,
            sheet.name,
            cell.address,
            parsed.row,
            parsed.col,
            cell.formula ? null : JSON.stringify(cell.value ?? null),
            cell.formula ?? null,
            cell.format ?? null,
            null,
            revision,
            updatedBy,
            updatedAt,
          ],
        ),
      );
    }

    for (const rowEntry of sheet.metadata?.rowMetadata ?? []) {
      rowMetadataQueries.push(
        db.query(
          `
            INSERT INTO row_metadata (
              workbook_id,
              sheet_name,
              start_index,
              count,
              size,
              hidden,
              source_revision,
              updated_at
            )
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8::timestamptz)
          `,
          [
            documentId,
            sheet.name,
            rowEntry.start,
            rowEntry.count,
            rowEntry.size ?? null,
            rowEntry.hidden ?? null,
            revision,
            updatedAt,
          ],
        ),
      );
    }

    for (const columnEntry of sheet.metadata?.columnMetadata ?? []) {
      columnMetadataQueries.push(
        db.query(
          `
            INSERT INTO column_metadata (
              workbook_id,
              sheet_name,
              start_index,
              count,
              size,
              hidden,
              source_revision,
              updated_at
            )
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8::timestamptz)
          `,
          [
            documentId,
            sheet.name,
            columnEntry.start,
            columnEntry.count,
            columnEntry.size ?? null,
            columnEntry.hidden ?? null,
            revision,
            updatedAt,
          ],
        ),
      );
    }

    for (const styleRange of sheet.metadata?.styleRanges ?? []) {
      const start = parseCellAddress(styleRange.range.startAddress, sheet.name);
      const end = parseCellAddress(styleRange.range.endAddress, sheet.name);
      styleRangeQueries.push(
        db.query(
          `
            INSERT INTO sheet_style_ranges (
              id,
              workbook_id,
              sheet_name,
              start_row,
              end_row,
              start_col,
              end_col,
              style_id,
              source_revision,
              updated_at
            )
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10::timestamptz)
          `,
          [
            rangeId("style-range", sheet.name, start.row, start.col, end.row, end.col),
            documentId,
            sheet.name,
            start.row,
            end.row,
            start.col,
            end.col,
            styleRange.styleId,
            revision,
            updatedAt,
          ],
        ),
      );
    }

    for (const formatRange of sheet.metadata?.formatRanges ?? []) {
      const start = parseCellAddress(formatRange.range.startAddress, sheet.name);
      const end = parseCellAddress(formatRange.range.endAddress, sheet.name);
      formatRangeQueries.push(
        db.query(
          `
            INSERT INTO sheet_format_ranges (
              id,
              workbook_id,
              sheet_name,
              start_row,
              end_row,
              start_col,
              end_col,
              format_id,
              source_revision,
              updated_at
            )
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10::timestamptz)
          `,
          [
            rangeId("format-range", sheet.name, start.row, start.col, end.row, end.col),
            documentId,
            sheet.name,
            start.row,
            end.row,
            start.col,
            end.col,
            formatRange.formatId,
            revision,
            updatedAt,
          ],
        ),
      );
    }
  }

  for (const style of snapshot.workbook.metadata?.styles ?? []) {
    styleQueries.push(
      db.query(
        `
          INSERT INTO cell_styles (workbook_id, style_id, record_json, hash, created_at)
          VALUES ($1, $2, $3::jsonb, $4, $5::timestamptz)
        `,
        [documentId, style.id, JSON.stringify(style), JSON.stringify(style), updatedAt],
      ),
    );
  }

  for (const format of snapshot.workbook.metadata?.formats ?? []) {
    formatQueries.push(
      db.query(
        `
          INSERT INTO cell_number_formats (workbook_id, format_id, code, kind, created_at)
          VALUES ($1, $2, $3, $4, $5::timestamptz)
        `,
        [documentId, format.id, format.code, format.kind, updatedAt],
      ),
    );
  }

  await Promise.all([
    ...sheetQueries,
    ...cellQueries,
    ...rowMetadataQueries,
    ...columnMetadataQueries,
    ...styleQueries,
    ...formatQueries,
    ...styleRangeQueries,
    ...formatRangeQueries,
  ]);

  const computedCellQueries = computedCells.map((cell) =>
    db.query(
      `
        INSERT INTO computed_cells (
          workbook_id,
          sheet_name,
          address,
          row_num,
          col_num,
          value,
          flags,
          version,
          calc_revision,
          updated_at
        )
        VALUES ($1, $2, $3, $4, $5, $6::jsonb, $7, $8, $9, $10::timestamptz)
      `,
      [
        documentId,
        cell.sheetName,
        cell.address,
        cell.rowNum,
        cell.colNum,
        JSON.stringify(cell.value),
        cell.flags,
        cell.version,
        revision,
        updatedAt,
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
    [documentId, calcSettings?.mode ?? "automatic", recalcEpoch],
  );
}
