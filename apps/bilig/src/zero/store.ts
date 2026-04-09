import type { EngineReplicaSnapshot, SpreadsheetEngine } from "@bilig/core";
import type { CellRangeRef, WorkbookSnapshot } from "@bilig/protocol";
import {
  type AuthoritativeWorkbookEventRecord,
  isWorkbookEventPayload,
  type WorkbookChangeUndoBundle,
  type WorkbookEventPayload,
} from "@bilig/zero-sync";
import {
  diffProjectionRows,
  type AxisMetadataSourceRow,
  type CellSourceRow,
  type CalculationSettingsSourceRow,
  type DefinedNameSourceRow,
  type NumberFormatSourceRow,
  type SheetSourceRow,
  type StyleSourceRow,
  type WorkbookMetadataSourceRow,
  type WorkbookSourceProjection,
  sourceProjectionKeys,
} from "./projection.js";
import {
  axisSignature,
  cellSignature,
  cellSourceRowInRange,
  definedNameSignature,
  normalizeRangeBounds,
  numberFormatSignature,
  parseInteger,
  parseJsonKey,
  sheetSignature,
  styleSignature,
  workbookMetadataSignature,
} from "./store-support.js";

export interface QueryResultRow {
  [key: string]: unknown;
}

export interface Queryable {
  query<T extends QueryResultRow = QueryResultRow>(
    text: string,
    values?: unknown[],
  ): Promise<{ rows: T[] }>;
}

export interface WorkbookRuntimeState {
  snapshot: WorkbookSnapshot;
  replicaSnapshot: EngineReplicaSnapshot | null;
  headRevision: number;
  calculatedRevision: number;
  ownerUserId: string;
}

export interface WorkbookRuntimeMetadata {
  headRevision: number;
  calculatedRevision: number;
  ownerUserId: string;
}

export interface ReplaceWorkbookDocumentInput {
  readonly documentId: string;
  readonly snapshot: WorkbookSnapshot;
  readonly ownerUserId: string;
  readonly revision: number;
  readonly calculatedRevision: number;
  readonly updatedBy: string;
  readonly updatedAt?: string;
}

export interface WorkbookProjectionState {
  projection: WorkbookSourceProjection;
  headRevision: number;
  calculatedRevision: number;
  ownerUserId: string;
}

export type WorkbookProjectionCommit =
  | {
      kind: "replace";
      projection: WorkbookSourceProjection;
    }
  | {
      kind: "focused-cell";
      workbook: WorkbookSourceProjection["workbook"];
      calculationSettings: CalculationSettingsSourceRow;
      sheetName: string;
      address: string;
      cell: CellSourceRow | null;
    }
  | {
      kind: "cell-range";
      workbook: WorkbookSourceProjection["workbook"];
      calculationSettings: CalculationSettingsSourceRow;
      range: CellRangeRef;
      cells: readonly CellSourceRow[];
      styles?: readonly StyleSourceRow[];
      numberFormats?: readonly NumberFormatSourceRow[];
    }
  | {
      kind: "column-metadata";
      workbook: WorkbookSourceProjection["workbook"];
      calculationSettings: CalculationSettingsSourceRow;
      sheetName: string;
      columnMetadata: readonly AxisMetadataSourceRow[];
    }
  | {
      kind: "row-metadata";
      workbook: WorkbookSourceProjection["workbook"];
      calculationSettings: CalculationSettingsSourceRow;
      sheetName: string;
      rowMetadata: readonly AxisMetadataSourceRow[];
    };

export interface PersistWorkbookMutationOptions {
  previousState: WorkbookProjectionState;
  nextEngine: SpreadsheetEngine;
  updatedBy: string;
  ownerUserId: string;
  eventPayload: WorkbookEventPayload;
  undoBundle: WorkbookChangeUndoBundle | null;
  clientMutationId?: string | null;
}

export interface PersistWorkbookMutationResult {
  revision: number;
  calculatedRevision: number;
  updatedAt: string;
  recalcJobId: string | null;
  projectionCommit: WorkbookProjectionCommit;
}

const WORKBOOK_CHECKPOINT_INTERVAL = 64;
const AUTHORITATIVE_SOURCE_PROJECTION_VERSION = 2;

export function shouldPersistWorkbookCheckpointRevision(revision: number): boolean {
  return revision === 1 || revision % WORKBOOK_CHECKPOINT_INTERVAL === 0;
}

export async function loadWorkbookEventRecordsAfter(
  db: Queryable,
  documentId: string,
  revision: number,
): Promise<readonly AuthoritativeWorkbookEventRecord[]> {
  const result = await db.query<{
    revision: number | string | null;
    client_mutation_id: string | null;
    txn_json: unknown;
  }>(
    `
      SELECT revision, client_mutation_id, txn_json
      FROM workbook_event
      WHERE workbook_id = $1
        AND revision > $2
      ORDER BY revision ASC
    `,
    [documentId, revision],
  );
  return result.rows.flatMap((row) =>
    parseInteger(row.revision) > revision && isWorkbookEventPayload(row.txn_json)
      ? [
          {
            revision: parseInteger(row.revision),
            clientMutationId: row.client_mutation_id,
            payload: row.txn_json,
          } satisfies AuthoritativeWorkbookEventRecord,
        ]
      : [],
  );
}

export async function upsertWorkbookHeader(
  db: Queryable,
  documentId: string,
  projection: WorkbookSourceProjection["workbook"],
  checkpointPayload: WorkbookSnapshot | null,
  replicaState: EngineReplicaSnapshot | null,
): Promise<void> {
  await db.query(
    `
      INSERT INTO workbooks (
        id,
        name,
        owner_user_id,
        head_revision,
        calculated_revision,
        source_projection_version,
        calc_mode,
        compatibility_mode,
        recalc_epoch,
        snapshot,
        replica_snapshot,
        updated_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10::jsonb, $11::jsonb, $12::timestamptz)
      ON CONFLICT (id)
      DO UPDATE SET
        name = EXCLUDED.name,
        owner_user_id = EXCLUDED.owner_user_id,
        head_revision = EXCLUDED.head_revision,
        calculated_revision = EXCLUDED.calculated_revision,
        source_projection_version = EXCLUDED.source_projection_version,
        calc_mode = EXCLUDED.calc_mode,
        compatibility_mode = EXCLUDED.compatibility_mode,
        recalc_epoch = EXCLUDED.recalc_epoch,
        snapshot = EXCLUDED.snapshot,
        replica_snapshot = EXCLUDED.replica_snapshot,
        updated_at = EXCLUDED.updated_at
    `,
    [
      documentId,
      projection.name,
      projection.ownerUserId,
      projection.headRevision,
      projection.calculatedRevision,
      AUTHORITATIVE_SOURCE_PROJECTION_VERSION,
      projection.calcMode,
      projection.compatibilityMode,
      projection.recalcEpoch,
      JSON.stringify(checkpointPayload),
      JSON.stringify(replicaState),
      projection.updatedAt,
    ],
  );
}

export async function insertWorkbookHeaderIfMissing(
  db: Queryable,
  documentId: string,
  projection: WorkbookSourceProjection["workbook"],
  checkpointPayload: WorkbookSnapshot | null,
  replicaState: EngineReplicaSnapshot | null,
): Promise<boolean> {
  const result = await db.query<{ id: string }>(
    `
      INSERT INTO workbooks (
        id,
        name,
        owner_user_id,
        head_revision,
        calculated_revision,
        source_projection_version,
        calc_mode,
        compatibility_mode,
        recalc_epoch,
        snapshot,
        replica_snapshot,
        updated_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10::jsonb, $11::jsonb, $12::timestamptz)
      ON CONFLICT (id)
      DO NOTHING
      RETURNING id
    `,
    [
      documentId,
      projection.name,
      projection.ownerUserId,
      projection.headRevision,
      projection.calculatedRevision,
      AUTHORITATIVE_SOURCE_PROJECTION_VERSION,
      projection.calcMode,
      projection.compatibilityMode,
      projection.recalcEpoch,
      JSON.stringify(checkpointPayload),
      JSON.stringify(replicaState),
      projection.updatedAt,
    ],
  );
  return result.rows.length > 0;
}

export async function applySheetDiff(
  db: Queryable,
  previousRows: readonly SheetSourceRow[],
  nextRows: readonly SheetSourceRow[],
): Promise<void> {
  const diff = diffProjectionRows(
    previousRows,
    nextRows,
    sourceProjectionKeys.sheet,
    sheetSignature,
  );
  const workbookId = nextRows[0]?.workbookId ?? previousRows[0]?.workbookId;
  const tasks: Promise<unknown>[] = [];
  for (const key of diff.deletes) {
    const [, sheetId] = parseJsonKey(key);
    if (typeof sheetId !== "number") {
      continue;
    }
    tasks.push(
      db.query(`DELETE FROM sheets WHERE workbook_id = $1 AND sheet_id = $2`, [
        workbookId,
        sheetId,
      ]),
    );
  }
  for (const row of diff.upserts) {
    tasks.push(
      db.query(
        `
        INSERT INTO sheets (
          workbook_id,
          sheet_id,
          name,
          sort_order,
          freeze_rows,
          freeze_cols,
          updated_at
        )
        VALUES ($1, $2, $3, $4, $5, $6, $7::timestamptz)
        ON CONFLICT (workbook_id, sheet_id)
        DO UPDATE SET
          name = EXCLUDED.name,
          sort_order = EXCLUDED.sort_order,
          freeze_rows = EXCLUDED.freeze_rows,
          freeze_cols = EXCLUDED.freeze_cols,
          updated_at = EXCLUDED.updated_at
      `,
        [
          row.workbookId,
          row.sheetId,
          row.name,
          row.sortOrder,
          row.freezeRows,
          row.freezeCols,
          row.updatedAt,
        ],
      ),
    );
  }
  await Promise.all(tasks);
}

export async function applyCellDiff(
  db: Queryable,
  previousRows: readonly CellSourceRow[],
  nextRows: readonly CellSourceRow[],
): Promise<void> {
  const diff = diffProjectionRows(previousRows, nextRows, sourceProjectionKeys.cell, cellSignature);
  const workbookId = nextRows[0]?.workbookId ?? previousRows[0]?.workbookId;
  const tasks: Promise<unknown>[] = [];
  for (const key of diff.deletes) {
    const [, sheetName, address] = parseJsonKey(key);
    tasks.push(
      db.query(`DELETE FROM cells WHERE workbook_id = $1 AND sheet_name = $2 AND address = $3`, [
        workbookId,
        sheetName,
        address,
      ]),
    );
  }
  for (const row of diff.upserts) {
    tasks.push(
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
          style_id,
          explicit_format_id,
          source_revision,
          updated_by,
          updated_at
        )
        VALUES ($1, $2, $3, $4, $5, $6::jsonb, $7, $8, $9, $10, $11, $12, $13::timestamptz)
        ON CONFLICT (workbook_id, sheet_name, address)
        DO UPDATE SET
          row_num = EXCLUDED.row_num,
          col_num = EXCLUDED.col_num,
          input_value = EXCLUDED.input_value,
          formula = EXCLUDED.formula,
          format = EXCLUDED.format,
          style_id = EXCLUDED.style_id,
          explicit_format_id = EXCLUDED.explicit_format_id,
          source_revision = EXCLUDED.source_revision,
          updated_by = EXCLUDED.updated_by,
          updated_at = EXCLUDED.updated_at
      `,
        [
          row.workbookId,
          row.sheetName,
          row.address,
          row.rowNum,
          row.colNum,
          JSON.stringify(row.inputValue ?? null),
          row.formula,
          row.format,
          row.styleId,
          row.explicitFormatId,
          row.sourceRevision,
          row.updatedBy,
          row.updatedAt,
        ],
      ),
    );
  }
  await Promise.all(tasks);
}

export async function persistCellSourceRange(
  db: Queryable,
  documentId: string,
  range: CellRangeRef,
  nextRows: readonly CellSourceRow[],
): Promise<void> {
  const bounds = normalizeRangeBounds(range);
  const nextRowsInRange = nextRows.filter((row) => cellSourceRowInRange(row, range));
  await db.query(
    `
      DELETE FROM cells
      WHERE workbook_id = $1
        AND sheet_name = $2
        AND row_num BETWEEN $3 AND $4
        AND col_num BETWEEN $5 AND $6
    `,
    [documentId, bounds.sheetName, bounds.rowStart, bounds.rowEnd, bounds.colStart, bounds.colEnd],
  );
  if (nextRowsInRange.length === 0) {
    return;
  }
  await applyCellDiff(db, [], nextRowsInRange);
}

export async function applyAxisMetadataDiff(
  db: Queryable,
  tableName: "row_metadata" | "column_metadata",
  previousRows: readonly AxisMetadataSourceRow[],
  nextRows: readonly AxisMetadataSourceRow[],
): Promise<void> {
  const diff = diffProjectionRows(
    previousRows,
    nextRows,
    sourceProjectionKeys.axisMetadata,
    axisSignature,
  );
  const workbookId = nextRows[0]?.workbookId ?? previousRows[0]?.workbookId;
  const tasks: Promise<unknown>[] = [];
  for (const key of diff.deletes) {
    const [, sheetName, startIndex] = parseJsonKey(key);
    tasks.push(
      db.query(
        `DELETE FROM ${tableName} WHERE workbook_id = $1 AND sheet_name = $2 AND start_index = $3`,
        [workbookId, sheetName, startIndex],
      ),
    );
  }
  for (const row of diff.upserts) {
    tasks.push(
      db.query(
        `
        INSERT INTO ${tableName} (
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
        ON CONFLICT (workbook_id, sheet_name, start_index)
        DO UPDATE SET
          count = EXCLUDED.count,
          size = EXCLUDED.size,
          hidden = EXCLUDED.hidden,
          source_revision = EXCLUDED.source_revision,
          updated_at = EXCLUDED.updated_at
      `,
        [
          row.workbookId,
          row.sheetName,
          row.startIndex,
          row.count,
          row.size,
          row.hidden,
          row.sourceRevision,
          row.updatedAt,
        ],
      ),
    );
  }
  await Promise.all(tasks);
}

export async function applyDefinedNameDiff(
  db: Queryable,
  previousRows: readonly DefinedNameSourceRow[],
  nextRows: readonly DefinedNameSourceRow[],
): Promise<void> {
  const diff = diffProjectionRows(
    previousRows,
    nextRows,
    sourceProjectionKeys.definedName,
    definedNameSignature,
  );
  const workbookId = nextRows[0]?.workbookId ?? previousRows[0]?.workbookId;
  const tasks: Promise<unknown>[] = [];
  for (const key of diff.deletes) {
    const [, name] = parseJsonKey(key);
    tasks.push(
      db.query(`DELETE FROM defined_names WHERE workbook_id = $1 AND name = $2`, [
        workbookId,
        name,
      ]),
    );
  }
  for (const row of diff.upserts) {
    tasks.push(
      db.query(
        `
        INSERT INTO defined_names (workbook_id, name, value)
        VALUES ($1, $2, $3::jsonb)
        ON CONFLICT (workbook_id, name)
        DO UPDATE SET value = EXCLUDED.value
      `,
        [row.workbookId, row.name, JSON.stringify(row.value)],
      ),
    );
  }
  await Promise.all(tasks);
}

export async function applyWorkbookMetadataDiff(
  db: Queryable,
  previousRows: readonly WorkbookMetadataSourceRow[],
  nextRows: readonly WorkbookMetadataSourceRow[],
): Promise<void> {
  const diff = diffProjectionRows(
    previousRows,
    nextRows,
    sourceProjectionKeys.workbookMetadata,
    workbookMetadataSignature,
  );
  const workbookId = nextRows[0]?.workbookId ?? previousRows[0]?.workbookId;
  const tasks: Promise<unknown>[] = [];
  for (const key of diff.deletes) {
    const [, metadataKey] = parseJsonKey(key);
    tasks.push(
      db.query(`DELETE FROM workbook_metadata WHERE workbook_id = $1 AND key = $2`, [
        workbookId,
        metadataKey,
      ]),
    );
  }
  for (const row of diff.upserts) {
    tasks.push(
      db.query(
        `
        INSERT INTO workbook_metadata (workbook_id, key, value)
        VALUES ($1, $2, $3::jsonb)
        ON CONFLICT (workbook_id, key)
        DO UPDATE SET value = EXCLUDED.value
      `,
        [row.workbookId, row.key, JSON.stringify(row.value)],
      ),
    );
  }
  await Promise.all(tasks);
}

export async function applyCalculationSettings(
  db: Queryable,
  projection: WorkbookSourceProjection["calculationSettings"],
): Promise<void> {
  await db.query(
    `
      INSERT INTO calculation_settings (workbook_id, mode, recalc_epoch)
      VALUES ($1, $2, $3)
      ON CONFLICT (workbook_id)
      DO UPDATE SET
        mode = EXCLUDED.mode,
        recalc_epoch = EXCLUDED.recalc_epoch
    `,
    [projection.workbookId, projection.mode, projection.recalcEpoch],
  );
}

export async function applyStyleDiff(
  db: Queryable,
  previousRows: readonly StyleSourceRow[],
  nextRows: readonly StyleSourceRow[],
): Promise<void> {
  const diff = diffProjectionRows(
    previousRows,
    nextRows,
    sourceProjectionKeys.style,
    styleSignature,
  );
  const workbookId = nextRows[0]?.workbookId ?? previousRows[0]?.workbookId;
  const tasks: Promise<unknown>[] = [];
  for (const key of diff.deletes) {
    const [, styleId] = parseJsonKey(key);
    tasks.push(
      db.query(`DELETE FROM cell_styles WHERE workbook_id = $1 AND style_id = $2`, [
        workbookId,
        styleId,
      ]),
    );
  }
  for (const row of diff.upserts) {
    tasks.push(
      db.query(
        `
        INSERT INTO cell_styles (workbook_id, style_id, record_json, hash, created_at)
        VALUES ($1, $2, $3::jsonb, $4, $5::timestamptz)
        ON CONFLICT (workbook_id, style_id)
        DO UPDATE SET
          record_json = EXCLUDED.record_json,
          hash = EXCLUDED.hash
      `,
        [row.workbookId, row.id, JSON.stringify(row.recordJSON), row.hash, row.createdAt],
      ),
    );
  }
  await Promise.all(tasks);
}

export async function applyNumberFormatDiff(
  db: Queryable,
  previousRows: readonly NumberFormatSourceRow[],
  nextRows: readonly NumberFormatSourceRow[],
): Promise<void> {
  const diff = diffProjectionRows(
    previousRows,
    nextRows,
    sourceProjectionKeys.numberFormat,
    numberFormatSignature,
  );
  const workbookId = nextRows[0]?.workbookId ?? previousRows[0]?.workbookId;
  const tasks: Promise<unknown>[] = [];
  for (const key of diff.deletes) {
    const [, formatId] = parseJsonKey(key);
    tasks.push(
      db.query(`DELETE FROM cell_number_formats WHERE workbook_id = $1 AND format_id = $2`, [
        workbookId,
        formatId,
      ]),
    );
  }
  for (const row of diff.upserts) {
    tasks.push(
      db.query(
        `
        INSERT INTO cell_number_formats (workbook_id, format_id, code, kind, created_at)
        VALUES ($1, $2, $3, $4, $5::timestamptz)
        ON CONFLICT (workbook_id, format_id)
        DO UPDATE SET
          code = EXCLUDED.code,
          kind = EXCLUDED.kind
      `,
        [row.workbookId, row.id, row.code, row.kind, row.createdAt],
      ),
    );
  }
  await Promise.all(tasks);
}

export async function applySourceProjectionDiff(
  db: Queryable,
  previousProjection: WorkbookSourceProjection,
  nextProjection: WorkbookSourceProjection,
): Promise<void> {
  await applySheetDiff(db, previousProjection.sheets, nextProjection.sheets);
  await applyCellDiff(db, previousProjection.cells, nextProjection.cells);
  await applyAxisMetadataDiff(
    db,
    "row_metadata",
    previousProjection.rowMetadata,
    nextProjection.rowMetadata,
  );
  await applyAxisMetadataDiff(
    db,
    "column_metadata",
    previousProjection.columnMetadata,
    nextProjection.columnMetadata,
  );
  await applyDefinedNameDiff(db, previousProjection.definedNames, nextProjection.definedNames);
  await applyWorkbookMetadataDiff(
    db,
    previousProjection.workbookMetadataEntries,
    nextProjection.workbookMetadataEntries,
  );
  await applyCalculationSettings(db, nextProjection.calculationSettings);
  await applyStyleDiff(db, previousProjection.styles, nextProjection.styles);
  await applyNumberFormatDiff(db, previousProjection.numberFormats, nextProjection.numberFormats);
}
