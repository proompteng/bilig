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
  upsertWorkbookHeader,
  type QueryResultRow,
  type Queryable,
  type ReplaceWorkbookDocumentInput,
} from "./store.js";
import { persistCellEvalRows } from "./workbook-calculation-store.js";
import { repairWorkbookSheetIds } from "./sheet-id-repair.js";

const AUTHORITATIVE_SOURCE_PROJECTION_VERSION = 2;

interface WorkbookMigrationRow extends QueryResultRow {
  readonly id: string;
  readonly snapshot: unknown;
  readonly replica_snapshot: unknown;
  readonly calculated_revision: number | string | null;
  readonly head_revision: number | string | null;
  readonly owner_user_id: string | null;
  readonly updated_at: string | null;
}

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

async function collectWorkbookIds(
  db: Queryable,
  query: string,
  values?: readonly unknown[],
): Promise<Set<string>> {
  const result = await db.query<{ workbook_id?: string | null; id?: string | null }>(
    query,
    values ? [...values] : undefined,
  );
  return new Set(
    result.rows.flatMap((row) => {
      if (typeof row.workbook_id === "string" && row.workbook_id.length > 0) {
        return [row.workbook_id];
      }
      if (typeof row.id === "string" && row.id.length > 0) {
        return [row.id];
      }
      return [];
    }),
  );
}

async function loadWorkbookMigrationRows(
  db: Queryable,
  workbookIds: readonly string[],
): Promise<readonly WorkbookMigrationRow[]> {
  if (workbookIds.length === 0) {
    return [];
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
  );
  return result.rows;
}

async function rebuildWorkbookStateForMigration(
  db: Queryable,
  row: WorkbookMigrationRow,
  options: {
    readonly replaceSourceProjection: boolean;
    readonly replaceCellEval: boolean;
  },
): Promise<void> {
  const checkpointPayload = parseCheckpointPayload(row.snapshot, row.id);
  const replicaState = parseCheckpointReplicaState(row.replica_snapshot);
  const updatedAt = row.updated_at ?? nowIso();
  const headRevision = parseInteger(row.head_revision);
  const calculatedRevision = parseInteger(row.calculated_revision);
  const ownerUserId = row.owner_user_id ?? "system";

  const engine = new SpreadsheetEngine({
    workbookName: row.id,
    replicaId: `migration:${row.id}:${headRevision}`,
  });
  await engine.ready();
  engine.importSnapshot(checkpointPayload);
  if (replicaState) {
    engine.importReplicaSnapshot(replicaState);
  }

  if (options.replaceSourceProjection) {
    const projection = buildWorkbookSourceProjection(row.id, checkpointPayload, {
      revision: headRevision,
      calculatedRevision,
      ownerUserId,
      updatedBy: ownerUserId,
      updatedAt,
    });
    await replaceWorkbookSourceProjectionForMigration(db, projection);
    await upsertWorkbookHeader(db, row.id, projection.workbook, checkpointPayload, replicaState);
  }

  if (options.replaceCellEval) {
    await replaceCellEvalForMigration(
      db,
      row.id,
      materializeCellEvalProjection(engine, row.id, calculatedRevision, updatedAt),
    );
  }
}

async function rebuildWorkbooksForMigration(
  db: Queryable,
  workbookIds: ReadonlySet<string>,
  options: {
    readonly replaceSourceProjection: boolean;
    readonly replaceCellEval: boolean;
  },
): Promise<void> {
  if (workbookIds.size === 0) {
    return;
  }
  const rows = await loadWorkbookMigrationRows(db, [...workbookIds]);
  await Promise.all(
    rows.map(async (row) => {
      await rebuildWorkbookStateForMigration(db, row, options);
    }),
  );
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
  );
  if (await tableExists(db, "sheet_style_ranges")) {
    const styleWorkbookIds = await collectWorkbookIds(
      db,
      `SELECT DISTINCT workbook_id FROM sheet_style_ranges`,
    );
    styleWorkbookIds.forEach((workbookId) => workbookIds.add(workbookId));
  }
  if (await tableExists(db, "sheet_format_ranges")) {
    const formatWorkbookIds = await collectWorkbookIds(
      db,
      `SELECT DISTINCT workbook_id FROM sheet_format_ranges`,
    );
    formatWorkbookIds.forEach((workbookId) => workbookIds.add(workbookId));
  }
  return workbookIds;
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

export async function repairWorkbookSheetIdsForMigration(db: Queryable): Promise<void> {
  await db.query(`UPDATE sheets SET sheet_id = sort_order + 1 WHERE sheet_id IS NULL`);
  await repairWorkbookSheetIds(db);
}

export async function backfillWorkbookSourceProjectionVersion(db: Queryable): Promise<void> {
  await rebuildWorkbooksForMigration(db, await loadLegacyProjectionWorkbookIds(db), {
    replaceSourceProjection: true,
    replaceCellEval: false,
  });
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
  );
  await rebuildWorkbooksForMigration(db, workbookIds, {
    replaceSourceProjection: false,
    replaceCellEval: true,
  });
}

export async function dropLegacyZeroSyncSchemaObjects(db: Queryable): Promise<void> {
  await db.query(`DROP INDEX IF EXISTS sheet_style_ranges_workbook_sheet_idx`);
  await db.query(`DROP INDEX IF EXISTS sheet_format_ranges_workbook_sheet_idx`);
  await db.query(`DROP TABLE IF EXISTS sheet_style_ranges`);
  await db.query(`DROP TABLE IF EXISTS sheet_format_ranges`);
}
