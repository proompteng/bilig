import { formatAddress, parseCellAddress } from "@bilig/formula";
import {
  isWorkbookChangeUndoBundle,
  isWorkbookEventPayload,
  type WorkbookChangeUndoBundle,
  type WorkbookEventPayload,
} from "@bilig/zero-sync";
import type { CellRangeRef } from "@bilig/protocol";
import type { QueryResultRow, Queryable } from "./store.js";
import { resolveWorkbookSheetRef } from "./workbook-sheet-ref.js";

export interface WorkbookChangeRange {
  readonly sheetName: string;
  readonly startAddress: string;
  readonly endAddress: string;
}

export interface WorkbookChangeDescriptor {
  readonly eventKind: WorkbookEventPayload["kind"];
  readonly summary: string;
  readonly sheetName: string | null;
  readonly anchorAddress: string | null;
  readonly range: WorkbookChangeRange | null;
}

export interface AppendWorkbookChangeInput {
  readonly documentId: string;
  readonly revision: number;
  readonly actorUserId: string;
  readonly clientMutationId: string | null;
  readonly payload: WorkbookEventPayload;
  readonly undoBundle: WorkbookChangeUndoBundle | null;
  readonly createdAtUnixMs: number;
}

interface WorkbookChangeInsertRow {
  readonly documentId: string;
  readonly revision: number;
  readonly actorUserId: string;
  readonly clientMutationId: string | null;
  readonly descriptor: WorkbookChangeDescriptor;
  readonly undoBundle: WorkbookChangeUndoBundle | null;
  readonly revertsRevision: number | null;
  readonly createdAtUnixMs: number;
}

export interface WorkbookChangeRecord {
  readonly revision: number;
  readonly actorUserId: string;
  readonly clientMutationId: string | null;
  readonly eventKind: WorkbookEventPayload["kind"];
  readonly summary: string;
  readonly sheetId: number | null;
  readonly sheetName: string | null;
  readonly anchorAddress: string | null;
  readonly range: WorkbookChangeRange | null;
  readonly undoBundle: WorkbookChangeUndoBundle | null;
  readonly revertedByRevision: number | null;
  readonly revertsRevision: number | null;
  readonly createdAtUnixMs: number;
}

interface WorkbookEventBackfillRow extends QueryResultRow {
  readonly workbookId?: unknown;
  readonly revision?: unknown;
  readonly actorUserId?: unknown;
  readonly clientMutationId?: unknown;
  readonly payload?: unknown;
  readonly createdAtUnixMs?: unknown;
}

interface WorkbookChangeSelectRow extends QueryResultRow {
  readonly revision?: unknown;
  readonly actorUserId?: unknown;
  readonly clientMutationId?: unknown;
  readonly eventKind?: unknown;
  readonly summary?: unknown;
  readonly sheetId?: unknown;
  readonly sheetName?: unknown;
  readonly anchorAddress?: unknown;
  readonly rangeJson?: unknown;
  readonly undoBundleJson?: unknown;
  readonly revertedByRevision?: unknown;
  readonly revertsRevision?: unknown;
  readonly createdAtUnixMs?: unknown;
}

interface CommitCellOpDescriptor {
  readonly sheetName: string;
  readonly address: string;
  readonly kind: "upsertCell" | "deleteCell";
  readonly formula?: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function parseNumericValue(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.trunc(value);
  }
  if (typeof value === "string" && value.length > 0) {
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

function normalizeRange(range: CellRangeRef): WorkbookChangeRange {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return {
    sheetName: range.sheetName,
    startAddress: formatAddress(Math.min(start.row, end.row), Math.min(start.col, end.col)),
    endAddress: formatAddress(Math.max(start.row, end.row), Math.max(start.col, end.col)),
  };
}

function normalizeWorkbookChangeRange(value: unknown): WorkbookChangeRange | null {
  if (!isRecord(value)) {
    return null;
  }
  const sheetName = value["sheetName"];
  const startAddress = value["startAddress"];
  const endAddress = value["endAddress"];
  if (
    typeof sheetName !== "string" ||
    typeof startAddress !== "string" ||
    typeof endAddress !== "string"
  ) {
    return null;
  }
  return {
    sheetName,
    startAddress,
    endAddress,
  };
}

function normalizeWorkbookChangeRecord(row: WorkbookChangeSelectRow): WorkbookChangeRecord | null {
  const revision = parseNumericValue(row.revision);
  const sheetId = parseNumericValue(row.sheetId);
  const revertedByRevision = parseNumericValue(row.revertedByRevision);
  const revertsRevision = parseNumericValue(row.revertsRevision);
  const createdAtUnixMs = parseNumericValue(row.createdAtUnixMs);
  if (
    revision === null ||
    createdAtUnixMs === null ||
    typeof row.actorUserId !== "string" ||
    typeof row.eventKind !== "string" ||
    typeof row.summary !== "string"
  ) {
    return null;
  }
  const range = normalizeWorkbookChangeRange(row.rangeJson);
  const eventKind = row.eventKind;
  if (
    eventKind !== "applyBatch" &&
    eventKind !== "setCellValue" &&
    eventKind !== "setCellFormula" &&
    eventKind !== "clearCell" &&
    eventKind !== "clearRange" &&
    eventKind !== "renderCommit" &&
    eventKind !== "fillRange" &&
    eventKind !== "copyRange" &&
    eventKind !== "moveRange" &&
    eventKind !== "updateColumnWidth" &&
    eventKind !== "setRangeStyle" &&
    eventKind !== "clearRangeStyle" &&
    eventKind !== "setRangeNumberFormat" &&
    eventKind !== "clearRangeNumberFormat" &&
    eventKind !== "restoreVersion" &&
    eventKind !== "revertChange"
  ) {
    return null;
  }
  return {
    revision,
    actorUserId: row.actorUserId,
    clientMutationId: typeof row.clientMutationId === "string" ? row.clientMutationId : null,
    eventKind,
    summary: row.summary,
    sheetId,
    sheetName: typeof row.sheetName === "string" ? row.sheetName : null,
    anchorAddress: typeof row.anchorAddress === "string" ? row.anchorAddress : null,
    range,
    undoBundle: isWorkbookChangeUndoBundle(row.undoBundleJson) ? row.undoBundleJson : null,
    revertedByRevision,
    revertsRevision,
    createdAtUnixMs,
  };
}

function rangeLabel(range: WorkbookChangeRange): string {
  return range.startAddress === range.endAddress
    ? `${range.sheetName}!${range.startAddress}`
    : `${range.sheetName}!${range.startAddress}:${range.endAddress}`;
}

function columnLabel(columnIndex: number): string {
  return formatAddress(0, columnIndex).replace(/[0-9]+$/u, "");
}

function rangeFromAddresses(
  sheetName: string,
  addresses: readonly string[],
): WorkbookChangeRange | null {
  if (addresses.length === 0) {
    return null;
  }
  let rowStart = Number.POSITIVE_INFINITY;
  let rowEnd = Number.NEGATIVE_INFINITY;
  let colStart = Number.POSITIVE_INFINITY;
  let colEnd = Number.NEGATIVE_INFINITY;
  for (const address of addresses) {
    const parsed = parseCellAddress(address, sheetName);
    rowStart = Math.min(rowStart, parsed.row);
    rowEnd = Math.max(rowEnd, parsed.row);
    colStart = Math.min(colStart, parsed.col);
    colEnd = Math.max(colEnd, parsed.col);
  }
  return {
    sheetName,
    startAddress: formatAddress(rowStart, colStart),
    endAddress: formatAddress(rowEnd, colEnd),
  };
}

function summarizeCommitCellOps(ops: readonly CommitCellOpDescriptor[]): WorkbookChangeDescriptor {
  const range = rangeFromAddresses(
    ops[0]!.sheetName,
    ops.map((op) => op.address),
  );
  const allUpserts = ops.every((op) => op.kind === "upsertCell");
  const allDeletes = ops.every((op) => op.kind === "deleteCell");
  const allFormulas = ops.every((op) => op.kind === "upsertCell" && typeof op.formula === "string");

  if (ops.length === 1) {
    const op = ops[0]!;
    if (op.kind === "deleteCell") {
      return {
        eventKind: "renderCommit",
        summary: `Cleared ${op.sheetName}!${op.address}`,
        sheetName: op.sheetName,
        anchorAddress: op.address,
        range,
      };
    }
    return {
      eventKind: "renderCommit",
      summary:
        typeof op.formula === "string"
          ? `Set formula in ${op.sheetName}!${op.address}`
          : `Updated ${op.sheetName}!${op.address}`,
      sheetName: op.sheetName,
      anchorAddress: op.address,
      range,
    };
  }

  return {
    eventKind: "renderCommit",
    summary: allFormulas
      ? `Filled ${ops.length} formulas in ${rangeLabel(range!)}`
      : allUpserts
        ? `Updated ${ops.length} cells in ${rangeLabel(range!)}`
        : allDeletes
          ? `Cleared ${ops.length} cells in ${rangeLabel(range!)}`
          : `Changed ${ops.length} cells in ${rangeLabel(range!)}`,
    sheetName: range?.sheetName ?? ops[0]!.sheetName,
    anchorAddress: range?.startAddress ?? ops[0]!.address,
    range,
  };
}

function summarizeRenderCommit(
  payload: Extract<WorkbookEventPayload, { kind: "renderCommit" }>,
): WorkbookChangeDescriptor {
  const cellOps = payload.ops.flatMap((op): CommitCellOpDescriptor[] => {
    if (!isRecord(op) || typeof op["kind"] !== "string") {
      return [];
    }
    if (
      (op["kind"] === "upsertCell" || op["kind"] === "deleteCell") &&
      typeof op["sheetName"] === "string" &&
      typeof op["addr"] === "string"
    ) {
      const base = {
        sheetName: op["sheetName"],
        address: op["addr"],
        kind: op["kind"],
      } satisfies Omit<CommitCellOpDescriptor, "formula">;
      return [typeof op["formula"] === "string" ? { ...base, formula: op["formula"] } : base];
    }
    return [];
  });

  if (cellOps.length === payload.ops.length && cellOps.length > 0) {
    return summarizeCommitCellOps(cellOps);
  }

  if (payload.ops.length === 1) {
    const op = payload.ops[0]!;
    if (op.kind === "upsertSheet" && op.name) {
      return {
        eventKind: "renderCommit",
        summary: `Created sheet ${op.name}`,
        sheetName: op.name,
        anchorAddress: "A1",
        range: {
          sheetName: op.name,
          startAddress: "A1",
          endAddress: "A1",
        },
      };
    }
    if (op.kind === "renameSheet" && op.oldName && op.newName) {
      return {
        eventKind: "renderCommit",
        summary: `Renamed sheet ${op.oldName} to ${op.newName}`,
        sheetName: op.newName,
        anchorAddress: "A1",
        range: {
          sheetName: op.newName,
          startAddress: "A1",
          endAddress: "A1",
        },
      };
    }
    if (op.kind === "deleteSheet" && op.name) {
      return {
        eventKind: "renderCommit",
        summary: `Deleted sheet ${op.name}`,
        sheetName: null,
        anchorAddress: null,
        range: null,
      };
    }
  }

  return {
    eventKind: "renderCommit",
    summary: `Applied ${payload.ops.length} workbook changes`,
    sheetName: null,
    anchorAddress: null,
    range: null,
  };
}

export function buildWorkbookChangeDescriptor(
  payload: WorkbookEventPayload,
): WorkbookChangeDescriptor {
  switch (payload.kind) {
    case "setCellValue":
      return {
        eventKind: payload.kind,
        summary: `Updated ${payload.sheetName}!${payload.address}`,
        sheetName: payload.sheetName,
        anchorAddress: payload.address,
        range: {
          sheetName: payload.sheetName,
          startAddress: payload.address,
          endAddress: payload.address,
        },
      };
    case "setCellFormula":
      return {
        eventKind: payload.kind,
        summary: `Set formula in ${payload.sheetName}!${payload.address}`,
        sheetName: payload.sheetName,
        anchorAddress: payload.address,
        range: {
          sheetName: payload.sheetName,
          startAddress: payload.address,
          endAddress: payload.address,
        },
      };
    case "clearCell":
      return {
        eventKind: payload.kind,
        summary: `Cleared ${payload.sheetName}!${payload.address}`,
        sheetName: payload.sheetName,
        anchorAddress: payload.address,
        range: {
          sheetName: payload.sheetName,
          startAddress: payload.address,
          endAddress: payload.address,
        },
      };
    case "clearRange": {
      const range = normalizeRange(payload.range);
      return {
        eventKind: payload.kind,
        summary: `Cleared ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      };
    }
    case "fillRange": {
      const range = normalizeRange(payload.target);
      return {
        eventKind: payload.kind,
        summary: `Filled ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      };
    }
    case "copyRange": {
      const range = normalizeRange(payload.target);
      return {
        eventKind: payload.kind,
        summary: `Copied into ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      };
    }
    case "moveRange": {
      const range = normalizeRange(payload.target);
      return {
        eventKind: payload.kind,
        summary: `Moved cells to ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      };
    }
    case "updateColumnWidth": {
      const anchorAddress = formatAddress(0, payload.columnIndex);
      return {
        eventKind: payload.kind,
        summary: `Resized column ${columnLabel(payload.columnIndex)} on ${payload.sheetName}`,
        sheetName: payload.sheetName,
        anchorAddress,
        range: {
          sheetName: payload.sheetName,
          startAddress: anchorAddress,
          endAddress: anchorAddress,
        },
      };
    }
    case "setRangeStyle": {
      const range = normalizeRange(payload.range);
      return {
        eventKind: payload.kind,
        summary: `Formatted ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      };
    }
    case "clearRangeStyle": {
      const range = normalizeRange(payload.range);
      return {
        eventKind: payload.kind,
        summary: `Cleared formatting in ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      };
    }
    case "setRangeNumberFormat": {
      const range = normalizeRange(payload.range);
      return {
        eventKind: payload.kind,
        summary: `Changed number format in ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      };
    }
    case "clearRangeNumberFormat": {
      const range = normalizeRange(payload.range);
      return {
        eventKind: payload.kind,
        summary: `Cleared number format in ${rangeLabel(range)}`,
        sheetName: range.sheetName,
        anchorAddress: range.startAddress,
        range,
      };
    }
    case "renderCommit":
      return summarizeRenderCommit(payload);
    case "restoreVersion":
      return {
        eventKind: payload.kind,
        summary: `Restored version ${payload.versionName}`,
        sheetName: payload.sheetName ?? null,
        anchorAddress: payload.address ?? null,
        range:
          payload.sheetName && payload.address
            ? {
                sheetName: payload.sheetName,
                startAddress: payload.address,
                endAddress: payload.address,
              }
            : null,
      };
    case "revertChange":
      return {
        eventKind: payload.kind,
        summary: `Reverted r${payload.targetRevision}: ${payload.targetSummary}`,
        sheetName: payload.sheetName ?? payload.range?.sheetName ?? null,
        anchorAddress: payload.address ?? payload.range?.startAddress ?? null,
        range: payload.range ? normalizeRange(payload.range) : null,
      };
    case "applyBatch":
      return {
        eventKind: payload.kind,
        summary: `Applied ${payload.batch.ops.length} synced operations`,
        sheetName: null,
        anchorAddress: null,
        range: null,
      };
    default: {
      const exhaustive: never = payload;
      return exhaustive;
    }
  }
}

async function insertWorkbookChange(db: Queryable, row: WorkbookChangeInsertRow): Promise<void> {
  const sheetRef = await resolveWorkbookSheetRef(db, {
    documentId: row.documentId,
    sheetName: row.descriptor.sheetName,
  });
  await db.query(
    `
      INSERT INTO workbook_change (
        workbook_id,
        revision,
        actor_user_id,
        client_mutation_id,
        event_kind,
        summary,
        sheet_id,
        sheet_name,
        anchor_address,
        range_json,
        undo_bundle_json,
        reverted_by_revision,
        reverts_revision,
        created_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10::jsonb, $11::jsonb, $12, $13, $14)
      ON CONFLICT (workbook_id, revision)
      DO UPDATE SET
        actor_user_id = EXCLUDED.actor_user_id,
        client_mutation_id = EXCLUDED.client_mutation_id,
        event_kind = EXCLUDED.event_kind,
        summary = EXCLUDED.summary,
        sheet_id = EXCLUDED.sheet_id,
        sheet_name = EXCLUDED.sheet_name,
        anchor_address = EXCLUDED.anchor_address,
        range_json = EXCLUDED.range_json,
        undo_bundle_json = EXCLUDED.undo_bundle_json,
        reverted_by_revision = EXCLUDED.reverted_by_revision,
        reverts_revision = EXCLUDED.reverts_revision,
        created_at = EXCLUDED.created_at
    `,
    [
      row.documentId,
      row.revision,
      row.actorUserId,
      row.clientMutationId,
      row.descriptor.eventKind,
      row.descriptor.summary,
      sheetRef.sheetId,
      sheetRef.sheetName,
      row.descriptor.anchorAddress,
      JSON.stringify(row.descriptor.range),
      row.undoBundle === null ? null : JSON.stringify(row.undoBundle),
      null,
      row.revertsRevision,
      row.createdAtUnixMs,
    ],
  );
}

async function markWorkbookChangeReverted(
  db: Queryable,
  input: {
    readonly documentId: string;
    readonly revision: number;
    readonly revertedByRevision: number;
  },
): Promise<void> {
  await db.query(
    `
      UPDATE workbook_change
         SET reverted_by_revision = $3
       WHERE workbook_id = $1
         AND revision = $2
    `,
    [input.documentId, input.revision, input.revertedByRevision],
  );
}

export async function ensureWorkbookChangeSchema(db: Queryable): Promise<void> {
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_change (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      revision BIGINT NOT NULL,
      actor_user_id TEXT NOT NULL,
      client_mutation_id TEXT,
      event_kind TEXT NOT NULL,
      summary TEXT NOT NULL,
      sheet_id INTEGER,
      sheet_name TEXT,
      anchor_address TEXT,
      range_json JSONB,
      undo_bundle_json JSONB,
      reverted_by_revision BIGINT,
      reverts_revision BIGINT,
      created_at BIGINT NOT NULL,
      PRIMARY KEY (workbook_id, revision)
    );
  `);
  await db.query(`ALTER TABLE workbook_change ADD COLUMN IF NOT EXISTS undo_bundle_json JSONB;`);
  await db.query(
    `ALTER TABLE workbook_change ADD COLUMN IF NOT EXISTS reverted_by_revision BIGINT;`,
  );
  await db.query(`ALTER TABLE workbook_change ADD COLUMN IF NOT EXISTS reverts_revision BIGINT;`);
  await db.query(
    `CREATE INDEX IF NOT EXISTS workbook_change_workbook_created_idx ON workbook_change(workbook_id, created_at DESC, revision DESC);`,
  );
}

export async function appendWorkbookChange(
  db: Queryable,
  input: AppendWorkbookChangeInput,
): Promise<void> {
  await insertWorkbookChange(db, {
    documentId: input.documentId,
    revision: input.revision,
    actorUserId: input.actorUserId,
    clientMutationId: input.clientMutationId,
    descriptor: buildWorkbookChangeDescriptor(input.payload),
    undoBundle: input.undoBundle,
    revertsRevision: input.payload.kind === "revertChange" ? input.payload.targetRevision : null,
    createdAtUnixMs: input.createdAtUnixMs,
  });
  if (input.payload.kind === "revertChange") {
    await markWorkbookChangeReverted(db, {
      documentId: input.documentId,
      revision: input.payload.targetRevision,
      revertedByRevision: input.revision,
    });
  }
}

export async function loadWorkbookChange(
  db: Queryable,
  documentId: string,
  revision: number,
): Promise<WorkbookChangeRecord | null> {
  const result = await db.query<WorkbookChangeSelectRow>(
    `
      SELECT revision AS "revision",
             actor_user_id AS "actorUserId",
             client_mutation_id AS "clientMutationId",
             event_kind AS "eventKind",
             summary AS "summary",
             sheet_id AS "sheetId",
             sheet_name AS "sheetName",
             anchor_address AS "anchorAddress",
             range_json AS "rangeJson",
             undo_bundle_json AS "undoBundleJson",
             reverted_by_revision AS "revertedByRevision",
             reverts_revision AS "revertsRevision",
             created_at AS "createdAtUnixMs"
        FROM workbook_change
       WHERE workbook_id = $1
         AND revision = $2
       LIMIT 1
    `,
    [documentId, revision],
  );
  const row = result.rows[0];
  return row ? normalizeWorkbookChangeRecord(row) : null;
}

export async function backfillWorkbookChanges(db: Queryable): Promise<void> {
  const result = await db.query<WorkbookEventBackfillRow>(
    `
      SELECT event.workbook_id AS "workbookId",
             event.revision AS "revision",
             event.actor_user_id AS "actorUserId",
             event.client_mutation_id AS "clientMutationId",
             event.txn_json AS "payload",
             CASE
               WHEN event.created_at IS NULL THEN 0
               ELSE FLOOR(EXTRACT(EPOCH FROM event.created_at) * 1000)
             END AS "createdAtUnixMs"
        FROM workbook_event AS event
        LEFT JOIN workbook_change AS change
          ON change.workbook_id = event.workbook_id
         AND change.revision = event.revision
       WHERE change.workbook_id IS NULL
       ORDER BY event.workbook_id ASC, event.revision ASC
    `,
  );

  const inserts = result.rows.flatMap((row) => {
    const revision = parseNumericValue(row.revision);
    const createdAtUnixMs = parseNumericValue(row.createdAtUnixMs);
    if (
      typeof row.workbookId !== "string" ||
      typeof row.actorUserId !== "string" ||
      revision === null ||
      createdAtUnixMs === null ||
      !isWorkbookEventPayload(row.payload)
    ) {
      return [];
    }
    return [
      appendWorkbookChange(db, {
        documentId: row.workbookId,
        revision,
        actorUserId: row.actorUserId,
        clientMutationId: typeof row.clientMutationId === "string" ? row.clientMutationId : null,
        payload: row.payload,
        undoBundle: null,
        createdAtUnixMs,
      }),
    ];
  });
  await Promise.all(inserts);
}
