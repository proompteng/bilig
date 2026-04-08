import { SpreadsheetEngine } from "@bilig/core";
import {
  buildWorkbookSourceProjection,
  materializeCellEvalProjection,
  type CellEvalRow,
  type WorkbookSourceProjection,
} from "./projection.js";
import {
  createEmptyWorkbookSnapshot,
  nowIso,
  parseCheckpointPayload,
  parseCheckpointReplicaState,
  parseInteger,
} from "./store-support.js";
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
  persistCellEvalRows,
  upsertWorkbookHeader,
  type Queryable,
  type ReplaceWorkbookDocumentInput,
} from "./store.js";

const AUTHORITATIVE_SOURCE_PROJECTION_VERSION = 2;

async function tableExists(db: Queryable, name: string): Promise<boolean> {
  const result = await db.query<{ relation: string | null }>(`SELECT to_regclass($1) AS relation`, [
    `public.${name}`,
  ]);
  return typeof result.rows[0]?.relation === "string";
}

async function replaceWorkbookSourceProjectionForMigration(
  db: Queryable,
  projection: WorkbookSourceProjection,
): Promise<void> {
  const workbookId = projection.workbook.id;
  await db.query(`DELETE FROM sheets WHERE workbook_id = $1`, [workbookId]);
  await db.query(`DELETE FROM cells WHERE workbook_id = $1`, [workbookId]);
  await db.query(`DELETE FROM row_metadata WHERE workbook_id = $1`, [workbookId]);
  await db.query(`DELETE FROM column_metadata WHERE workbook_id = $1`, [workbookId]);
  await db.query(`DELETE FROM defined_names WHERE workbook_id = $1`, [workbookId]);
  await db.query(`DELETE FROM workbook_metadata WHERE workbook_id = $1`, [workbookId]);
  await db.query(`DELETE FROM calculation_settings WHERE workbook_id = $1`, [workbookId]);
  await db.query(`DELETE FROM cell_styles WHERE workbook_id = $1`, [workbookId]);
  await db.query(`DELETE FROM cell_number_formats WHERE workbook_id = $1`, [workbookId]);
  await applySheetDiff(db, [], projection.sheets);
  await applyCellDiff(db, [], projection.cells);
  await applyAxisMetadataDiff(db, "row_metadata", [], projection.rowMetadata);
  await applyAxisMetadataDiff(db, "column_metadata", [], projection.columnMetadata);
  await applyDefinedNameDiff(db, [], projection.definedNames);
  await applyWorkbookMetadataDiff(db, [], projection.workbookMetadataEntries);
  await applyCalculationSettings(db, projection.calculationSettings);
  await applyStyleDiff(db, [], projection.styles);
  await applyNumberFormatDiff(db, [], projection.numberFormats);
}

async function replaceCellEvalForMigration(
  db: Queryable,
  documentId: string,
  rows: readonly CellEvalRow[],
): Promise<void> {
  await db.query(`DELETE FROM cell_eval WHERE workbook_id = $1`, [documentId]);
  if (rows.length > 0) {
    await persistCellEvalRows(db, documentId, [], rows);
  }
}

export async function replaceWorkbookDocument(
  db: Queryable,
  input: ReplaceWorkbookDocumentInput,
): Promise<void> {
  const updatedAt = input.updatedAt ?? nowIso();
  const projection = buildWorkbookSourceProjection(input.documentId, input.snapshot, {
    revision: input.revision,
    calculatedRevision: input.calculatedRevision,
    ownerUserId: input.ownerUserId,
    updatedBy: input.updatedBy,
    updatedAt,
  });
  await db.query(`DELETE FROM workbooks WHERE id = $1`, [input.documentId]);
  await upsertWorkbookHeader(db, input.documentId, projection.workbook, input.snapshot, null);
  await replaceWorkbookSourceProjectionForMigration(db, projection);
  const engine = new SpreadsheetEngine({
    workbookName: input.documentId,
    replicaId: `replace-document:${input.documentId}:${input.revision}`,
  });
  await engine.ready();
  engine.importSnapshot(input.snapshot);
  await replaceCellEvalForMigration(
    db,
    input.documentId,
    materializeCellEvalProjection(engine, input.documentId, input.calculatedRevision, updatedAt),
  );
}

export async function ensureWorkbookDocumentExists(
  db: Queryable,
  documentId: string,
  ownerUserId = "system",
): Promise<void> {
  const snapshot = createEmptyWorkbookSnapshot(documentId);
  const updatedAt = nowIso();
  const projection = buildWorkbookSourceProjection(documentId, snapshot, {
    revision: 0,
    calculatedRevision: 0,
    ownerUserId,
    updatedBy: ownerUserId,
    updatedAt,
  });
  const inserted = await insertWorkbookHeaderIfMissing(
    db,
    documentId,
    projection.workbook,
    snapshot,
    null,
  );
  if (!inserted) {
    return;
  }
  await replaceWorkbookSourceProjectionForMigration(db, projection);
}

export async function backfillAuthoritativeCellEval(db: Queryable): Promise<void> {
  const styleRangesExist = await tableExists(db, "sheet_style_ranges");
  const formatRangesExist = await tableExists(db, "sheet_format_ranges");
  const legacyWorkbookIds = new Set<string>();

  const staleProjectionRows = await db.query<{ id: string }>(
    `
      SELECT id
      FROM workbooks
      WHERE source_projection_version < $1
    `,
    [AUTHORITATIVE_SOURCE_PROJECTION_VERSION],
  );
  for (const row of staleProjectionRows.rows) {
    legacyWorkbookIds.add(row.id);
  }

  const staleRenderRows = await db.query<{ workbook_id: string }>(
    `
      SELECT DISTINCT workbook_id
      FROM cell_eval
      WHERE style_id IS NOT NULL
        AND style_json IS NULL
    `,
  );
  for (const row of staleRenderRows.rows) {
    legacyWorkbookIds.add(row.workbook_id);
  }

  if (styleRangesExist) {
    const legacyStyleRows = await db.query<{ workbook_id: string }>(
      `SELECT DISTINCT workbook_id FROM sheet_style_ranges`,
    );
    for (const row of legacyStyleRows.rows) {
      legacyWorkbookIds.add(row.workbook_id);
    }
  }

  if (formatRangesExist) {
    const legacyFormatRows = await db.query<{ workbook_id: string }>(
      `SELECT DISTINCT workbook_id FROM sheet_format_ranges`,
    );
    for (const row of legacyFormatRows.rows) {
      legacyWorkbookIds.add(row.workbook_id);
    }
  }

  if (legacyWorkbookIds.size === 0) {
    return;
  }

  const result = await db.query<{
    id: string;
    snapshot: unknown;
    replica_snapshot: unknown;
    calculated_revision: number | string | null;
    head_revision: number | string | null;
    owner_user_id: string | null;
    updated_at: string | null;
  }>(
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
    [[...legacyWorkbookIds]],
  );

  await Promise.all(
    result.rows.map(async (row) => {
      const checkpointPayload = parseCheckpointPayload(row.snapshot, row.id);
      const replicaState = parseCheckpointReplicaState(row.replica_snapshot);
      const updatedAt = row.updated_at ?? nowIso();
      const engine = new SpreadsheetEngine({
        workbookName: row.id,
        replicaId: `cell-eval-backfill:${row.id}`,
      });
      await engine.ready();
      engine.importSnapshot(checkpointPayload);
      if (replicaState) {
        engine.importReplicaSnapshot(replicaState);
      }
      const projection = buildWorkbookSourceProjection(row.id, checkpointPayload, {
        revision: parseInteger(row.head_revision),
        calculatedRevision: parseInteger(row.calculated_revision),
        ownerUserId: row.owner_user_id ?? "system",
        updatedBy: row.owner_user_id ?? "system",
        updatedAt,
      });
      await replaceWorkbookSourceProjectionForMigration(db, projection);
      await replaceCellEvalForMigration(
        db,
        row.id,
        materializeCellEvalProjection(
          engine,
          row.id,
          parseInteger(row.calculated_revision),
          updatedAt,
        ),
      );
      await upsertWorkbookHeader(db, row.id, projection.workbook, checkpointPayload, replicaState);
    }),
  );
}

export async function dropLegacyZeroSyncSchemaObjects(db: Queryable): Promise<void> {
  await db.query(`DROP INDEX IF EXISTS sheet_style_ranges_workbook_sheet_idx`);
  await db.query(`DROP INDEX IF EXISTS sheet_format_ranges_workbook_sheet_idx`);
  await db.query(`DROP TABLE IF EXISTS sheet_style_ranges`);
  await db.query(`DROP TABLE IF EXISTS sheet_format_ranges`);
}
