import type { EngineReplicaSnapshot, SpreadsheetEngine } from "@bilig/core";
import { parseCellAddress } from "@bilig/formula";
import {
  ValueTag,
  type CellRangeRef,
  type CellValue,
  type WorkbookSnapshot,
} from "@bilig/protocol";
import {
  deriveDirtyRegions,
  type DirtyRegion,
  type WorkbookEventPayload,
  type WorkbookEventRecord,
} from "./events.js";
import {
  buildCalculationSettingsRow,
  buildSheetColumnMetadataRows,
  buildSheetFormatRangeRows,
  buildSheetStyleRangeRows,
  buildSingleCellSourceRow,
  buildWorkbookHeaderRow,
  buildWorkbookNumberFormatRows,
  buildWorkbookSourceProjection,
  buildWorkbookStyleRows,
  diffProjectionRows,
  materializeCellEvalProjection,
  type AxisMetadataSourceRow,
  type CellEvalRow,
  type CellSourceRow,
  type DefinedNameSourceRow,
  type FormatRangeSourceRow,
  type NumberFormatSourceRow,
  type SheetSourceRow,
  type StyleRangeSourceRow,
  type StyleSourceRow,
  type WorkbookMetadataSourceRow,
  type WorkbookSourceProjection,
  sourceProjectionKeys,
} from "./projection.js";

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

export interface PersistWorkbookMutationOptions {
  previousState: WorkbookRuntimeState;
  nextSnapshot: WorkbookSnapshot;
  nextReplicaSnapshot: EngineReplicaSnapshot | null;
  nextEngine?: SpreadsheetEngine | null;
  updatedBy: string;
  ownerUserId: string;
  eventPayload: WorkbookEventPayload;
  clientMutationId?: string | null;
}

export interface PersistWorkbookMutationResult {
  revision: number;
  calculatedRevision: number;
  updatedAt: string;
  recalcJobId: string | null;
}

export interface RecalcJobLease {
  id: string;
  workbookId: string;
  fromRevision: number;
  toRevision: number;
  dirtyRegions: DirtyRegion[] | null;
  attempts: number;
}

interface WorkbookCheckpoint {
  revision: number;
  snapshot: WorkbookSnapshot;
  replicaSnapshot: EngineReplicaSnapshot | null;
}

const WORKBOOK_SNAPSHOT_FORMAT = "json-v1";
const WORKBOOK_SNAPSHOT_RETENTION = 5;
const WORKBOOK_SNAPSHOT_INTERVAL = 64;
const RECALC_LEASE_MS = 30_000;
const MAX_RECALC_ATTEMPTS = 3;

type FocusedCellEventPayload = Extract<
  WorkbookEventPayload,
  { kind: "setCellValue" | "setCellFormula" | "clearCell" }
>;

type StyleRangeEventPayload = Extract<
  WorkbookEventPayload,
  { kind: "setRangeStyle" | "clearRangeStyle" }
>;

type NumberFormatRangeEventPayload = Extract<
  WorkbookEventPayload,
  { kind: "setRangeNumberFormat" | "clearRangeNumberFormat" }
>;

type ColumnMetadataEventPayload = Extract<WorkbookEventPayload, { kind: "updateColumnWidth" }>;

function nowIso(): string {
  return new Date().toISOString();
}

function parseInteger(value: unknown): number {
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.trunc(value);
  }
  if (typeof value === "string" && value.length > 0) {
    const parsed = Number.parseInt(value, 10);
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }
  return 0;
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

function isDirtyRegion(value: unknown): value is DirtyRegion {
  return (
    isRecord(value) &&
    typeof value["sheetName"] === "string" &&
    typeof value["rowStart"] === "number" &&
    typeof value["rowEnd"] === "number" &&
    typeof value["colStart"] === "number" &&
    typeof value["colEnd"] === "number"
  );
}

function workbookSnapshot(value: unknown, documentId: string): WorkbookSnapshot {
  return isWorkbookSnapshot(value) ? value : createEmptyWorkbookSnapshot(documentId);
}

function workbookReplicaSnapshot(value: unknown): EngineReplicaSnapshot | null {
  return isEngineReplicaSnapshot(value) ? value : null;
}

function isFocusedCellEventPayload(
  payload: WorkbookEventPayload,
): payload is FocusedCellEventPayload {
  return (
    payload.kind === "setCellValue" ||
    payload.kind === "setCellFormula" ||
    payload.kind === "clearCell"
  );
}

function isStyleRangeEventPayload(
  payload: WorkbookEventPayload,
): payload is StyleRangeEventPayload {
  return payload.kind === "setRangeStyle" || payload.kind === "clearRangeStyle";
}

function isNumberFormatRangeEventPayload(
  payload: WorkbookEventPayload,
): payload is NumberFormatRangeEventPayload {
  return payload.kind === "setRangeNumberFormat" || payload.kind === "clearRangeNumberFormat";
}

function isColumnMetadataEventPayload(
  payload: WorkbookEventPayload,
): payload is ColumnMetadataEventPayload {
  return payload.kind === "updateColumnWidth";
}

function eventRequiresRecalc(payload: WorkbookEventPayload): boolean {
  return !(
    payload.kind === "setRangeStyle" ||
    payload.kind === "clearRangeStyle" ||
    payload.kind === "setRangeNumberFormat" ||
    payload.kind === "clearRangeNumberFormat" ||
    payload.kind === "updateColumnWidth"
  );
}

function semanticSignature(value: unknown): string {
  return JSON.stringify(value);
}

function sheetSignature(row: SheetSourceRow): string {
  return semanticSignature([row.name, row.sortOrder, row.freezeRows, row.freezeCols]);
}

function cellSignature(row: CellSourceRow): string {
  return semanticSignature([
    row.sheetName,
    row.address,
    row.rowNum,
    row.colNum,
    row.inputValue ?? null,
    row.formula ?? null,
    row.format ?? null,
    row.explicitFormatId ?? null,
  ]);
}

function axisSignature(row: AxisMetadataSourceRow): string {
  return semanticSignature([
    row.sheetName,
    row.startIndex,
    row.count,
    row.size ?? null,
    row.hidden ?? null,
  ]);
}

function definedNameSignature(row: DefinedNameSourceRow): string {
  return semanticSignature([row.name, row.value]);
}

function workbookMetadataSignature(row: WorkbookMetadataSourceRow): string {
  return semanticSignature([row.key, row.value]);
}

function styleSignature(row: StyleSourceRow): string {
  return semanticSignature([row.id, row.recordJSON, row.hash]);
}

function numberFormatSignature(row: NumberFormatSourceRow): string {
  return semanticSignature([row.id, row.code, row.kind]);
}

function styleRangeSignature(row: StyleRangeSourceRow): string {
  return semanticSignature([
    row.sheetName,
    row.startRow,
    row.endRow,
    row.startCol,
    row.endCol,
    row.styleId,
  ]);
}

function formatRangeSignature(row: FormatRangeSourceRow): string {
  return semanticSignature([
    row.sheetName,
    row.startRow,
    row.endRow,
    row.startCol,
    row.endCol,
    row.formatId,
  ]);
}

function cellEvalSignature(row: CellEvalRow): string {
  return semanticSignature([
    row.sheetName,
    row.address,
    row.rowNum,
    row.colNum,
    row.value,
    row.flags,
    row.version,
    row.styleId,
    row.formatId,
    row.formatCode,
  ]);
}

function parseJsonKey(key: string): unknown[] {
  const parsed = JSON.parse(key) as unknown;
  if (!Array.isArray(parsed)) {
    throw new Error(`Invalid projection key: ${key}`);
  }
  return parsed;
}

function isCellValue(value: unknown): value is CellValue {
  if (!isRecord(value) || typeof value["tag"] !== "number") {
    return false;
  }
  const tag = value["tag"];
  if (tag === 0) {
    return true;
  }
  if (tag === 1) {
    return typeof value["value"] === "number";
  }
  if (tag === 2) {
    return typeof value["value"] === "boolean";
  }
  if (tag === 3) {
    return typeof value["value"] === "string";
  }
  if (tag === 4) {
    return typeof value["code"] === "number";
  }
  return false;
}

function parseCellEvalValue(value: unknown): CellValue {
  return isCellValue(value) ? value : { tag: ValueTag.Empty };
}

function normalizeRangeBounds(range: CellRangeRef): {
  sheetName: string;
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
} {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return {
    sheetName: range.sheetName,
    rowStart: Math.min(start.row, end.row),
    rowEnd: Math.max(start.row, end.row),
    colStart: Math.min(start.col, end.col),
    colEnd: Math.max(start.col, end.col),
  };
}

function cellEvalRowInRange(
  row: Pick<CellEvalRow, "sheetName" | "rowNum" | "colNum">,
  range: CellRangeRef,
): boolean {
  const bounds = normalizeRangeBounds(range);
  return (
    row.sheetName === bounds.sheetName &&
    row.rowNum >= bounds.rowStart &&
    row.rowNum <= bounds.rowEnd &&
    row.colNum >= bounds.colStart &&
    row.colNum <= bounds.colEnd
  );
}

async function loadLatestWorkbookCheckpoint(
  db: Queryable,
  documentId: string,
): Promise<WorkbookCheckpoint | null> {
  const result = await db.query<{
    revision: number | string | null;
    payload: unknown;
    replica_snapshot: unknown;
  }>(
    `
      SELECT revision, payload, replica_snapshot
      FROM workbook_snapshot
      WHERE workbook_id = $1
      ORDER BY revision DESC
      LIMIT 1
    `,
    [documentId],
  );
  const row = result.rows[0];
  if (!row || !isWorkbookSnapshot(row.payload)) {
    return null;
  }
  return {
    revision: parseInteger(row.revision),
    snapshot: row.payload,
    replicaSnapshot: workbookReplicaSnapshot(row.replica_snapshot),
  };
}

async function upsertWorkbookHeader(
  db: Queryable,
  documentId: string,
  projection: WorkbookSourceProjection["workbook"],
  snapshot: WorkbookSnapshot,
  replicaSnapshot: EngineReplicaSnapshot | null,
): Promise<void> {
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
      projection.name,
      projection.ownerUserId,
      projection.headRevision,
      projection.calculatedRevision,
      projection.calcMode,
      projection.compatibilityMode,
      projection.recalcEpoch,
      JSON.stringify(snapshot),
      JSON.stringify(replicaSnapshot),
      projection.updatedAt,
    ],
  );
}

async function applySheetDiff(
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
    const [, name] = parseJsonKey(key);
    tasks.push(
      db.query(`DELETE FROM sheets WHERE workbook_id = $1 AND name = $2`, [workbookId, name]),
    );
  }
  for (const row of diff.upserts) {
    tasks.push(
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
        ON CONFLICT (workbook_id, name)
        DO UPDATE SET
          sort_order = EXCLUDED.sort_order,
          freeze_rows = EXCLUDED.freeze_rows,
          freeze_cols = EXCLUDED.freeze_cols,
          updated_at = EXCLUDED.updated_at
      `,
        [row.workbookId, row.name, row.sortOrder, row.freezeRows, row.freezeCols, row.updatedAt],
      ),
    );
  }
  await Promise.all(tasks);
}

async function applyCellDiff(
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
          explicit_format_id,
          source_revision,
          updated_by,
          updated_at
        )
        VALUES ($1, $2, $3, $4, $5, $6::jsonb, $7, $8, $9, $10, $11, $12::timestamptz)
        ON CONFLICT (workbook_id, sheet_name, address)
        DO UPDATE SET
          row_num = EXCLUDED.row_num,
          col_num = EXCLUDED.col_num,
          input_value = EXCLUDED.input_value,
          formula = EXCLUDED.formula,
          format = EXCLUDED.format,
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

async function applyAxisMetadataDiff(
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

async function applyDefinedNameDiff(
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

async function applyWorkbookMetadataDiff(
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

async function applyCalculationSettings(
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

async function applyStyleDiff(
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

async function applyNumberFormatDiff(
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

async function applyStyleRangeDiff(
  db: Queryable,
  previousRows: readonly StyleRangeSourceRow[],
  nextRows: readonly StyleRangeSourceRow[],
): Promise<void> {
  const diff = diffProjectionRows(
    previousRows,
    nextRows,
    sourceProjectionKeys.styleRange,
    styleRangeSignature,
  );
  const tasks: Promise<unknown>[] = [];
  for (const id of diff.deletes) {
    tasks.push(db.query(`DELETE FROM sheet_style_ranges WHERE id = $1`, [id]));
  }
  for (const row of diff.upserts) {
    tasks.push(
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
        ON CONFLICT (id)
        DO UPDATE SET
          sheet_name = EXCLUDED.sheet_name,
          start_row = EXCLUDED.start_row,
          end_row = EXCLUDED.end_row,
          start_col = EXCLUDED.start_col,
          end_col = EXCLUDED.end_col,
          style_id = EXCLUDED.style_id,
          source_revision = EXCLUDED.source_revision,
          updated_at = EXCLUDED.updated_at
      `,
        [
          row.id,
          row.workbookId,
          row.sheetName,
          row.startRow,
          row.endRow,
          row.startCol,
          row.endCol,
          row.styleId,
          row.sourceRevision,
          row.updatedAt,
        ],
      ),
    );
  }
  await Promise.all(tasks);
}

async function applyFormatRangeDiff(
  db: Queryable,
  previousRows: readonly FormatRangeSourceRow[],
  nextRows: readonly FormatRangeSourceRow[],
): Promise<void> {
  const diff = diffProjectionRows(
    previousRows,
    nextRows,
    sourceProjectionKeys.formatRange,
    formatRangeSignature,
  );
  const tasks: Promise<unknown>[] = [];
  for (const id of diff.deletes) {
    tasks.push(db.query(`DELETE FROM sheet_format_ranges WHERE id = $1`, [id]));
  }
  for (const row of diff.upserts) {
    tasks.push(
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
        ON CONFLICT (id)
        DO UPDATE SET
          sheet_name = EXCLUDED.sheet_name,
          start_row = EXCLUDED.start_row,
          end_row = EXCLUDED.end_row,
          start_col = EXCLUDED.start_col,
          end_col = EXCLUDED.end_col,
          format_id = EXCLUDED.format_id,
          source_revision = EXCLUDED.source_revision,
          updated_at = EXCLUDED.updated_at
      `,
        [
          row.id,
          row.workbookId,
          row.sheetName,
          row.startRow,
          row.endRow,
          row.startCol,
          row.endCol,
          row.formatId,
          row.sourceRevision,
          row.updatedAt,
        ],
      ),
    );
  }
  await Promise.all(tasks);
}

async function applySourceProjectionDiff(
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
  await applyStyleRangeDiff(db, previousProjection.styleRanges, nextProjection.styleRanges);
  await applyFormatRangeDiff(db, previousProjection.formatRanges, nextProjection.formatRanges);
}

function buildFocusedCellRows(
  documentId: string,
  snapshot: WorkbookSnapshot,
  payload: FocusedCellEventPayload,
  options: {
    revision: number;
    calculatedRevision: number;
    ownerUserId: string;
    updatedBy: string;
    updatedAt: string;
  },
): readonly CellSourceRow[] {
  const row = buildSingleCellSourceRow(
    documentId,
    snapshot,
    payload.sheetName,
    payload.address,
    options,
  );
  return row ? [row] : [];
}

async function appendWorkbookEvent(db: Queryable, event: WorkbookEventRecord): Promise<void> {
  await db.query(
    `
      INSERT INTO workbook_event (
        workbook_id,
        revision,
        actor_user_id,
        client_mutation_id,
        txn_json,
        created_at
      )
      VALUES ($1, $2, $3, $4, $5::jsonb, $6::timestamptz)
    `,
    [
      event.workbookId,
      event.revision,
      event.actorUserId,
      event.clientMutationId,
      JSON.stringify(event.payload),
      event.createdAt,
    ],
  );
}

async function supersedePendingRecalcJobs(
  db: Queryable,
  documentId: string,
  toRevision: number,
): Promise<void> {
  await db.query(
    `
      UPDATE recalc_job
      SET status = 'superseded',
          updated_at = NOW(),
          lease_until = NULL,
          lease_owner = NULL
      WHERE workbook_id = $1
        AND status = 'pending'
        AND to_revision < $2
    `,
    [documentId, toRevision],
  );
}

async function enqueueRecalcJob(
  db: Queryable,
  documentId: string,
  fromRevision: number,
  toRevision: number,
  dirtyRegions: DirtyRegion[] | null,
  updatedAt: string,
): Promise<string> {
  const jobId = `${documentId}:recalc:${toRevision}`;
  await db.query(
    `
      INSERT INTO recalc_job (
        id,
        workbook_id,
        from_revision,
        to_revision,
        dirty_regions_json,
        status,
        attempts,
        last_error,
        created_at,
        updated_at
      )
      VALUES ($1, $2, $3, $4, $5::jsonb, 'pending', 0, NULL, $6::timestamptz, $6::timestamptz)
      ON CONFLICT (id)
      DO UPDATE SET
        dirty_regions_json = EXCLUDED.dirty_regions_json,
        status = 'pending',
        lease_until = NULL,
        lease_owner = NULL,
        last_error = NULL,
        updated_at = EXCLUDED.updated_at
    `,
    [jobId, documentId, fromRevision, toRevision, JSON.stringify(dirtyRegions), updatedAt],
  );
  return jobId;
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
    formatId: row.format_id,
    formatCode: row.format_code,
    calcRevision: parseInteger(row.calc_revision),
    updatedAt: row.updated_at ?? nowIso(),
  }));
}

async function loadCellEvalRowsForRange(
  db: Queryable,
  documentId: string,
  range: CellRangeRef,
): Promise<CellEvalRow[]> {
  const bounds = normalizeRangeBounds(range);
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
        format_id,
        format_code,
        calc_revision,
        updated_at
      FROM cell_eval
      WHERE workbook_id = $1
        AND sheet_name = $2
        AND row_num BETWEEN $3 AND $4
        AND col_num BETWEEN $5 AND $6
    `,
    [documentId, bounds.sheetName, bounds.rowStart, bounds.rowEnd, bounds.colStart, bounds.colEnd],
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
    formatId: row.format_id,
    formatCode: row.format_code,
    calcRevision: parseInteger(row.calc_revision),
    updatedAt: row.updated_at ?? nowIso(),
  }));
}

async function persistCellEvalRows(
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
          format_id,
          format_code,
          calc_revision,
          updated_at
        )
        VALUES ($1, $2, $3, $4, $5, $6::jsonb, $7, $8, $9, $10, $11, $12, $13::timestamptz)
        ON CONFLICT (workbook_id, sheet_name, address)
        DO UPDATE SET
          row_num = EXCLUDED.row_num,
          col_num = EXCLUDED.col_num,
          value = EXCLUDED.value,
          flags = EXCLUDED.flags,
          version = EXCLUDED.version,
          style_id = EXCLUDED.style_id,
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

async function persistCellEvalDiff(
  db: Queryable,
  documentId: string,
  nextRows: readonly CellEvalRow[],
): Promise<void> {
  const previousRows = await loadCellEvalRows(db, documentId);
  await persistCellEvalRows(db, documentId, previousRows, nextRows);
}

async function persistCellEvalRangeDiff(
  db: Queryable,
  documentId: string,
  range: CellRangeRef,
  nextRows: readonly CellEvalRow[],
): Promise<void> {
  const previousRows = await loadCellEvalRowsForRange(db, documentId, range);
  const nextRowsInRange = nextRows.filter((row) => cellEvalRowInRange(row, range));
  await persistCellEvalRows(db, documentId, previousRows, nextRowsInRange);
}

async function persistWorkbookCheckpoint(
  db: Queryable,
  documentId: string,
  revision: number,
  snapshot: WorkbookSnapshot,
  replicaSnapshot: EngineReplicaSnapshot | null,
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
      WORKBOOK_SNAPSHOT_FORMAT,
      JSON.stringify(snapshot),
      JSON.stringify(replicaSnapshot),
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
    [documentId, WORKBOOK_SNAPSHOT_RETENTION],
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
    `CREATE INDEX IF NOT EXISTS sheet_style_ranges_workbook_sheet_idx ON sheet_style_ranges(workbook_id, sheet_name, start_row, end_row, start_col, end_col);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS sheet_format_ranges_workbook_sheet_idx ON sheet_format_ranges(workbook_id, sheet_name, start_row, end_row, start_col, end_col);`,
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
    [WORKBOOK_SNAPSHOT_FORMAT],
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

export async function loadWorkbookState(
  db: Queryable,
  documentId: string,
): Promise<WorkbookRuntimeState> {
  const result = await db.query<{
    snapshot: unknown;
    replica_snapshot: unknown;
    head_revision: number | string | null;
    calculated_revision: number | string | null;
    owner_user_id: string | null;
  }>(
    `SELECT snapshot, replica_snapshot, head_revision, calculated_revision, owner_user_id FROM workbooks WHERE id = $1 LIMIT 1`,
    [documentId],
  );
  const row = result.rows[0];
  const checkpoint = isWorkbookSnapshot(row?.snapshot)
    ? null
    : await loadLatestWorkbookCheckpoint(db, documentId);
  return {
    snapshot: workbookSnapshot(row?.snapshot ?? checkpoint?.snapshot, documentId),
    replicaSnapshot: workbookReplicaSnapshot(row?.replica_snapshot ?? checkpoint?.replicaSnapshot),
    headRevision: parseInteger(row?.head_revision ?? checkpoint?.revision ?? 0),
    calculatedRevision: parseInteger(row?.calculated_revision),
    ownerUserId: row?.owner_user_id ?? "system",
  };
}

export async function loadWorkbookRuntimeMetadata(
  db: Queryable,
  documentId: string,
): Promise<WorkbookRuntimeMetadata> {
  const result = await db.query<{
    head_revision: number | string | null;
    calculated_revision: number | string | null;
    owner_user_id: string | null;
  }>(
    `SELECT head_revision, calculated_revision, owner_user_id FROM workbooks WHERE id = $1 LIMIT 1`,
    [documentId],
  );
  const row = result.rows[0];
  return {
    headRevision: parseInteger(row?.head_revision),
    calculatedRevision: parseInteger(row?.calculated_revision),
    ownerUserId: row?.owner_user_id ?? "system",
  };
}

export async function acquireWorkbookMutationLock(
  db: Queryable,
  documentId: string,
): Promise<void> {
  await db.query(`SELECT pg_advisory_xact_lock(hashtext($1))`, [documentId]);
}

export async function persistWorkbookMutation(
  db: Queryable,
  documentId: string,
  options: PersistWorkbookMutationOptions,
): Promise<PersistWorkbookMutationResult> {
  const updatedAt = nowIso();
  const revision = options.previousState.headRevision + 1;
  const needsRecalc =
    options.previousState.calculatedRevision < options.previousState.headRevision ||
    eventRequiresRecalc(options.eventPayload);
  const nextProjectionOptions = {
    revision,
    calculatedRevision: needsRecalc ? options.previousState.calculatedRevision : revision,
    ownerUserId: options.ownerUserId,
    updatedBy: options.updatedBy,
    updatedAt,
  };
  const previousProjectionOptions = {
    revision: options.previousState.headRevision,
    calculatedRevision: options.previousState.calculatedRevision,
    ownerUserId: options.previousState.ownerUserId,
    updatedBy: options.updatedBy,
    updatedAt,
  };
  const nextWorkbookRow = buildWorkbookHeaderRow(
    documentId,
    options.nextSnapshot,
    nextProjectionOptions,
  );

  await upsertWorkbookHeader(
    db,
    documentId,
    nextWorkbookRow,
    options.nextSnapshot,
    options.nextReplicaSnapshot,
  );
  if (isFocusedCellEventPayload(options.eventPayload)) {
    const previousCellRows = buildFocusedCellRows(
      documentId,
      options.previousState.snapshot,
      options.eventPayload,
      previousProjectionOptions,
    );
    const nextCellRows = buildFocusedCellRows(
      documentId,
      options.nextSnapshot,
      options.eventPayload,
      nextProjectionOptions,
    );
    await applyCalculationSettings(
      db,
      buildCalculationSettingsRow(documentId, options.nextSnapshot),
    );
    await applyCellDiff(db, previousCellRows, nextCellRows);
  } else if (isStyleRangeEventPayload(options.eventPayload)) {
    await applyCalculationSettings(
      db,
      buildCalculationSettingsRow(documentId, options.nextSnapshot),
    );
    await applyStyleDiff(
      db,
      buildWorkbookStyleRows(documentId, options.previousState.snapshot, previousProjectionOptions),
      buildWorkbookStyleRows(documentId, options.nextSnapshot, nextProjectionOptions),
    );
    await applyStyleRangeDiff(
      db,
      buildSheetStyleRangeRows(
        documentId,
        options.previousState.snapshot,
        options.eventPayload.range.sheetName,
        previousProjectionOptions,
      ),
      buildSheetStyleRangeRows(
        documentId,
        options.nextSnapshot,
        options.eventPayload.range.sheetName,
        nextProjectionOptions,
      ),
    );
    if (options.nextEngine) {
      await persistCellEvalRangeDiff(
        db,
        documentId,
        options.eventPayload.range,
        materializeCellEvalProjection(
          options.nextEngine,
          documentId,
          nextProjectionOptions.calculatedRevision,
          updatedAt,
        ),
      );
    }
  } else if (isNumberFormatRangeEventPayload(options.eventPayload)) {
    await applyCalculationSettings(
      db,
      buildCalculationSettingsRow(documentId, options.nextSnapshot),
    );
    await applyNumberFormatDiff(
      db,
      buildWorkbookNumberFormatRows(
        documentId,
        options.previousState.snapshot,
        previousProjectionOptions,
      ),
      buildWorkbookNumberFormatRows(documentId, options.nextSnapshot, nextProjectionOptions),
    );
    await applyFormatRangeDiff(
      db,
      buildSheetFormatRangeRows(
        documentId,
        options.previousState.snapshot,
        options.eventPayload.range.sheetName,
        previousProjectionOptions,
      ),
      buildSheetFormatRangeRows(
        documentId,
        options.nextSnapshot,
        options.eventPayload.range.sheetName,
        nextProjectionOptions,
      ),
    );
    if (options.nextEngine) {
      await persistCellEvalRangeDiff(
        db,
        documentId,
        options.eventPayload.range,
        materializeCellEvalProjection(
          options.nextEngine,
          documentId,
          nextProjectionOptions.calculatedRevision,
          updatedAt,
        ),
      );
    }
  } else if (isColumnMetadataEventPayload(options.eventPayload)) {
    await applyCalculationSettings(
      db,
      buildCalculationSettingsRow(documentId, options.nextSnapshot),
    );
    await applyAxisMetadataDiff(
      db,
      "column_metadata",
      buildSheetColumnMetadataRows(
        documentId,
        options.previousState.snapshot,
        options.eventPayload.sheetName,
        previousProjectionOptions,
      ),
      buildSheetColumnMetadataRows(
        documentId,
        options.nextSnapshot,
        options.eventPayload.sheetName,
        nextProjectionOptions,
      ),
    );
  } else {
    const previousProjection = buildWorkbookSourceProjection(
      documentId,
      options.previousState.snapshot,
      previousProjectionOptions,
    );
    const nextProjection = buildWorkbookSourceProjection(
      documentId,
      options.nextSnapshot,
      nextProjectionOptions,
    );
    await applySourceProjectionDiff(db, previousProjection, nextProjection);
  }

  await appendWorkbookEvent(db, {
    workbookId: documentId,
    revision,
    actorUserId: options.updatedBy,
    clientMutationId: options.clientMutationId ?? null,
    payload: options.eventPayload,
    createdAt: updatedAt,
  });

  await supersedePendingRecalcJobs(db, documentId, revision);
  const recalcJobId = needsRecalc
    ? await enqueueRecalcJob(
        db,
        documentId,
        options.previousState.calculatedRevision,
        revision,
        eventRequiresRecalc(options.eventPayload) ? deriveDirtyRegions(options.eventPayload) : null,
        updatedAt,
      )
    : null;

  return {
    revision,
    calculatedRevision: nextProjectionOptions.calculatedRevision,
    updatedAt,
    recalcJobId,
  };
}

export async function leaseNextRecalcJob(
  db: Queryable,
  workerId: string,
): Promise<RecalcJobLease | null> {
  const result = await db.query<{
    id: string;
    workbook_id: string;
    from_revision: number | string | null;
    to_revision: number | string | null;
    dirty_regions_json: unknown;
    attempts: number | string | null;
  }>(
    `
      WITH candidate AS (
        SELECT id
        FROM recalc_job
        WHERE status = 'pending'
           OR (status = 'running' AND lease_until IS NOT NULL AND lease_until < NOW())
        ORDER BY created_at ASC
        LIMIT 1
        FOR UPDATE SKIP LOCKED
      )
      UPDATE recalc_job
      SET status = 'running',
          attempts = attempts + 1,
          lease_owner = $1,
          lease_until = NOW() + ($2 * INTERVAL '1 millisecond'),
          updated_at = NOW()
      WHERE id IN (SELECT id FROM candidate)
      RETURNING id, workbook_id, from_revision, to_revision, dirty_regions_json, attempts
    `,
    [workerId, RECALC_LEASE_MS],
  );
  const row = result.rows[0];
  if (!row) {
    return null;
  }
  const dirtyRegions = Array.isArray(row.dirty_regions_json)
    ? row.dirty_regions_json.filter(isDirtyRegion)
    : null;
  return {
    id: row.id,
    workbookId: row.workbook_id,
    fromRevision: parseInteger(row.from_revision),
    toRevision: parseInteger(row.to_revision),
    dirtyRegions: dirtyRegions && dirtyRegions.length > 0 ? dirtyRegions : null,
    attempts: parseInteger(row.attempts),
  };
}

export async function markRecalcJobCompleted(
  db: Queryable,
  lease: RecalcJobLease,
  nextRows: readonly CellEvalRow[],
  snapshot: WorkbookSnapshot,
  replicaSnapshot: EngineReplicaSnapshot | null,
): Promise<boolean> {
  const revisionResult = await db.query<{ head_revision: number | string | null }>(
    `SELECT head_revision FROM workbooks WHERE id = $1 LIMIT 1`,
    [lease.workbookId],
  );
  if (parseInteger(revisionResult.rows[0]?.head_revision) !== lease.toRevision) {
    await markRecalcJobSuperseded(db, lease);
    return false;
  }

  await persistCellEvalDiff(db, lease.workbookId, nextRows);
  await db.query(
    `
      UPDATE workbooks
      SET calculated_revision = $2
      WHERE id = $1 AND head_revision = $2
    `,
    [lease.workbookId, lease.toRevision],
  );
  await db.query(
    `
      UPDATE recalc_job
      SET status = 'completed',
          lease_until = NULL,
          lease_owner = NULL,
          updated_at = NOW()
      WHERE id = $1
    `,
    [lease.id],
  );

  if (lease.toRevision === 1 || lease.toRevision % WORKBOOK_SNAPSHOT_INTERVAL === 0) {
    await persistWorkbookCheckpoint(
      db,
      lease.workbookId,
      lease.toRevision,
      snapshot,
      replicaSnapshot,
    );
  }
  return true;
}

export async function markRecalcJobSuperseded(db: Queryable, lease: RecalcJobLease): Promise<void> {
  await db.query(
    `
      UPDATE recalc_job
      SET status = 'superseded',
          lease_until = NULL,
          lease_owner = NULL,
          updated_at = NOW()
      WHERE id = $1
    `,
    [lease.id],
  );
}

export async function markRecalcJobFailed(
  db: Queryable,
  lease: RecalcJobLease,
  error: unknown,
): Promise<void> {
  const message = error instanceof Error ? (error.stack ?? error.message) : String(error);
  const exhausted = lease.attempts >= MAX_RECALC_ATTEMPTS;
  await db.query(
    `
      UPDATE recalc_job
      SET status = $2,
          lease_until = NULL,
          lease_owner = NULL,
          last_error = $3,
          updated_at = NOW()
      WHERE id = $1
    `,
    [lease.id, exhausted ? "failed" : "pending", message],
  );
}
