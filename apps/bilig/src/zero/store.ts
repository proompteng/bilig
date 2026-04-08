import type { EngineReplicaSnapshot, SpreadsheetEngine } from "@bilig/core";
import type { CellRangeRef, WorkbookSnapshot } from "@bilig/protocol";
import {
  type AuthoritativeWorkbookEventRecord,
  isWorkbookEventPayload,
  type WorkbookChangeUndoBundle,
  type WorkbookEventPayload,
} from "@bilig/zero-sync";
import { repairWorkbookSheetIds } from "./sheet-id-repair.js";
import {
  diffProjectionRows,
  type AxisMetadataSourceRow,
  type CellEvalRow,
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
  cellEvalRowInRange,
  cellEvalSignature,
  cellSignature,
  cellSourceRowInRange,
  definedNameSignature,
  normalizeRangeBounds,
  nowIso,
  numberFormatSignature,
  parseCellEvalValue,
  parseCellStyleRecord,
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

const WORKBOOK_CHECKPOINT_FORMAT = "json-v1";
const WORKBOOK_CHECKPOINT_RETENTION = 5;
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

async function loadCellEvalRows(db: Queryable, documentId: string): Promise<CellEvalRow[]> {
  const result = await db.query<{
    workbook_id: string;
    sheet_name: string;
    address: string;
    row_num: number | null;
    col_num: number | null;
    value: unknown;
    flags: number | string | null;
    version: number | string | null;
    style_id: string | null;
    style_json: unknown;
    format_id: string | null;
    format_code: string | null;
    calc_revision: number | string | null;
    updated_at: string | null;
  }>(
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
  );
  return result.rows.map((row) => ({
    workbookId: row.workbook_id,
    sheetName: row.sheet_name,
    address: row.address,
    rowNum: parseInteger(row.row_num),
    colNum: parseInteger(row.col_num),
    value: parseCellEvalValue(row.value),
    flags: parseInteger(row.flags),
    version: parseInteger(row.version),
    styleId: row.style_id,
    styleJson: parseCellStyleRecord(row.style_json),
    formatId: row.format_id,
    formatCode: row.format_code,
    calcRevision: parseInteger(row.calc_revision),
    updatedAt: row.updated_at ?? nowIso(),
  }));
}

export async function persistCellEvalRows(
  db: Queryable,
  documentId: string,
  previousRows: readonly CellEvalRow[],
  nextRows: readonly CellEvalRow[],
): Promise<void> {
  const diff = diffProjectionRows(
    previousRows,
    nextRows,
    sourceProjectionKeys.cellEval,
    cellEvalSignature,
  );
  const tasks: Promise<unknown>[] = [];
  for (const key of diff.deletes) {
    const [, sheetName, address] = parseJsonKey(key);
    tasks.push(
      db.query(
        `DELETE FROM cell_eval WHERE workbook_id = $1 AND sheet_name = $2 AND address = $3`,
        [documentId, sheetName, address],
      ),
    );
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
    );
  }
  await Promise.all(tasks);
}

export async function persistCellEvalIncremental(
  db: Queryable,
  _documentId: string,
  rows: readonly CellEvalRow[],
): Promise<void> {
  const tasks: Promise<unknown>[] = [];
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
    );
  }
  await Promise.all(tasks);
}

export async function persistCellEvalDiff(
  db: Queryable,
  documentId: string,
  nextRows: readonly CellEvalRow[],
): Promise<void> {
  const previousRows = await loadCellEvalRows(db, documentId);
  await persistCellEvalRows(db, documentId, previousRows, nextRows);
}

export async function persistCellEvalRangeDiff(
  db: Queryable,
  documentId: string,
  range: CellRangeRef,
  nextRows: readonly CellEvalRow[],
): Promise<void> {
  const nextRowsInRange = nextRows.filter((row) => cellEvalRowInRange(row, range));
  const bounds = normalizeRangeBounds(range);
  await db.query(
    `
      DELETE FROM cell_eval
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
  await persistCellEvalRows(db, documentId, [], nextRowsInRange);
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
    [
      documentId,
      revision,
      WORKBOOK_CHECKPOINT_FORMAT,
      JSON.stringify(checkpointPayload),
      JSON.stringify(replicaState),
    ],
  );
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
  );
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
    `ALTER TABLE workbooks ADD COLUMN IF NOT EXISTS source_projection_version BIGINT NOT NULL DEFAULT 0;`,
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
      sheet_id INTEGER,
      name TEXT NOT NULL,
      sort_order INTEGER NOT NULL,
      PRIMARY KEY (workbook_id, name)
    );
  `);
  await db.query(`ALTER TABLE sheets ADD COLUMN IF NOT EXISTS sheet_id INTEGER;`);
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
  await db.query(`ALTER TABLE cells ADD COLUMN IF NOT EXISTS style_id TEXT;`);
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
    CREATE TABLE IF NOT EXISTS cell_eval (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      sheet_name TEXT NOT NULL,
      address TEXT NOT NULL,
      value JSONB NOT NULL,
      flags INTEGER NOT NULL,
      version INTEGER NOT NULL,
      PRIMARY KEY (workbook_id, sheet_name, address)
    );
  `);
  await db.query(`ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS row_num INTEGER;`);
  await db.query(`ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS col_num INTEGER;`);
  await db.query(`ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS style_id TEXT;`);
  await db.query(`ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS style_json JSONB;`);
  await db.query(`ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS format_id TEXT;`);
  await db.query(`ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS format_code TEXT;`);
  await db.query(
    `ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS calc_revision BIGINT NOT NULL DEFAULT 0;`,
  );
  await db.query(
    `ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW();`,
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
    CREATE TABLE IF NOT EXISTS workbook_event (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      revision BIGINT NOT NULL,
      actor_user_id TEXT NOT NULL,
      client_mutation_id TEXT,
      txn_json JSONB NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      PRIMARY KEY (workbook_id, revision)
    );
  `);

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
  `);

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
  `);

  await db.query(
    `CREATE INDEX IF NOT EXISTS sheets_workbook_sort_order_idx ON sheets(workbook_id, sort_order);`,
  );
  await db.query(`UPDATE sheets SET sheet_id = sort_order + 1 WHERE sheet_id IS NULL;`);
  await repairWorkbookSheetIds(db);
  await db.query(
    `CREATE UNIQUE INDEX IF NOT EXISTS sheets_workbook_sheet_id_idx ON sheets(workbook_id, sheet_id);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS cells_workbook_sheet_idx ON cells(workbook_id, sheet_name);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS cells_workbook_sheet_row_col_idx ON cells(workbook_id, sheet_name, row_num, col_num);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS cell_eval_workbook_sheet_idx ON cell_eval(workbook_id, sheet_name);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS cell_eval_workbook_sheet_row_col_idx ON cell_eval(workbook_id, sheet_name, row_num, col_num);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS row_metadata_workbook_sheet_idx ON row_metadata(workbook_id, sheet_name, start_index);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS column_metadata_workbook_sheet_idx ON column_metadata(workbook_id, sheet_name, start_index);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS recalc_job_status_lease_created_idx ON recalc_job(status, lease_until, created_at);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS workbook_event_workbook_created_idx ON workbook_event(workbook_id, created_at);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS workbook_snapshot_workbook_revision_idx ON workbook_snapshot(workbook_id, revision DESC);`,
  );

  await db.query(`
    DO $$
    BEGIN
      IF to_regclass('public.computed_cells') IS NOT NULL THEN
        INSERT INTO cell_eval (
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
        SELECT
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
        FROM computed_cells
        ON CONFLICT (workbook_id, sheet_name, address)
        DO NOTHING;
      END IF;
    END $$;
  `);

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
  );
}
