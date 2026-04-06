import type { WorkbookSnapshot } from "@bilig/protocol";
import type { Queryable } from "./store.js";
import { resolveWorkbookSheetRef } from "./workbook-sheet-ref.js";

export interface WorkbookVersionViewport {
  readonly rowStart: number;
  readonly rowEnd: number;
  readonly colStart: number;
  readonly colEnd: number;
}

export interface CreateWorkbookVersionInput {
  readonly documentId: string;
  readonly id: string;
  readonly ownerUserId: string;
  readonly name: string;
  readonly revision: number;
  readonly snapshot: WorkbookSnapshot;
  readonly sheetId?: number | null;
  readonly sheetName?: string | null;
  readonly address?: string | null;
  readonly viewport?: WorkbookVersionViewport | null;
}

export interface DeleteWorkbookVersionInput {
  readonly documentId: string;
  readonly id: string;
  readonly ownerUserId: string;
}

export interface WorkbookVersionRecord {
  readonly id: string;
  readonly ownerUserId: string;
  readonly name: string;
  readonly revision: number;
  readonly snapshot: WorkbookSnapshot;
  readonly sheetId: number | null;
  readonly sheetName: string | null;
  readonly address: string | null;
  readonly viewport: WorkbookVersionViewport | null;
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

function isWorkbookVersionViewport(value: unknown): value is WorkbookVersionViewport {
  return (
    isRecord(value) &&
    typeof value["rowStart"] === "number" &&
    typeof value["rowEnd"] === "number" &&
    typeof value["colStart"] === "number" &&
    typeof value["colEnd"] === "number" &&
    value["rowEnd"] >= value["rowStart"] &&
    value["colEnd"] >= value["colStart"]
  );
}

function parseInteger(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.trunc(value);
  }
  if (typeof value === "string" && value.length > 0) {
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

function readExistingOwnerUserId(rows: readonly { ownerUserId?: unknown }[]): string | null {
  const ownerUserId = rows[0]?.ownerUserId;
  return typeof ownerUserId === "string" ? ownerUserId : null;
}

async function ensureWorkbookVersionOwnership(
  db: Queryable,
  documentId: string,
  id: string,
  ownerUserId: string,
): Promise<void> {
  const existing = await db.query<{ ownerUserId?: unknown }>(
    `
      SELECT owner_user_id AS "ownerUserId"
      FROM workbook_version
      WHERE workbook_id = $1 AND id = $2
      LIMIT 1
    `,
    [documentId, id],
  );
  const existingOwnerUserId = readExistingOwnerUserId(existing.rows);
  if (existingOwnerUserId && existingOwnerUserId !== ownerUserId) {
    throw new Error("Workbook version is owned by another user");
  }
}

export async function ensureWorkbookVersionSchema(db: Queryable): Promise<void> {
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_version (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      id TEXT NOT NULL,
      owner_user_id TEXT NOT NULL,
      name TEXT NOT NULL,
      revision BIGINT NOT NULL,
      snapshot_json JSONB NOT NULL,
      sheet_id INTEGER,
      sheet_name TEXT,
      address TEXT,
      viewport_json JSONB,
      created_at BIGINT NOT NULL,
      updated_at BIGINT NOT NULL,
      PRIMARY KEY (workbook_id, id)
    );
  `);
  await db.query(
    `CREATE INDEX IF NOT EXISTS workbook_version_workbook_updated_idx ON workbook_version(workbook_id, updated_at DESC, created_at DESC);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS workbook_version_workbook_owner_idx ON workbook_version(workbook_id, owner_user_id);`,
  );
}

export async function createWorkbookVersion(
  db: Queryable,
  input: CreateWorkbookVersionInput,
): Promise<void> {
  await ensureWorkbookVersionOwnership(db, input.documentId, input.id, input.ownerUserId);
  const timestamp = Date.now();
  const sheetRef =
    input.sheetId == null && input.sheetName == null
      ? { sheetId: null, sheetName: null }
      : await resolveWorkbookSheetRef(db, input);
  await db.query(
    `
      INSERT INTO workbook_version (
        workbook_id,
        id,
        owner_user_id,
        name,
        revision,
        snapshot_json,
        sheet_id,
        sheet_name,
        address,
        viewport_json,
        created_at,
        updated_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12)
      ON CONFLICT (workbook_id, id)
      DO UPDATE SET
        owner_user_id = EXCLUDED.owner_user_id,
        name = EXCLUDED.name,
        revision = EXCLUDED.revision,
        snapshot_json = EXCLUDED.snapshot_json,
        sheet_id = EXCLUDED.sheet_id,
        sheet_name = EXCLUDED.sheet_name,
        address = EXCLUDED.address,
        viewport_json = EXCLUDED.viewport_json,
        updated_at = EXCLUDED.updated_at
    `,
    [
      input.documentId,
      input.id,
      input.ownerUserId,
      input.name,
      input.revision,
      JSON.stringify(input.snapshot),
      sheetRef.sheetId,
      sheetRef.sheetName,
      input.address ?? null,
      input.viewport ? JSON.stringify(input.viewport) : null,
      timestamp,
      timestamp,
    ],
  );
}

export async function deleteWorkbookVersion(
  db: Queryable,
  input: DeleteWorkbookVersionInput,
): Promise<void> {
  await ensureWorkbookVersionOwnership(db, input.documentId, input.id, input.ownerUserId);
  await db.query(
    `
      DELETE FROM workbook_version
      WHERE workbook_id = $1 AND id = $2 AND owner_user_id = $3
    `,
    [input.documentId, input.id, input.ownerUserId],
  );
}

export async function loadWorkbookVersion(
  db: Queryable,
  documentId: string,
  id: string,
): Promise<WorkbookVersionRecord | null> {
  const result = await db.query<{
    id?: unknown;
    owner_user_id?: unknown;
    name?: unknown;
    revision?: unknown;
    snapshot_json?: unknown;
    sheet_id?: unknown;
    sheet_name?: unknown;
    address?: unknown;
    viewport_json?: unknown;
  }>(
    `
      SELECT
        id,
        owner_user_id,
        name,
        revision,
        snapshot_json,
        sheet_id,
        sheet_name,
        address,
        viewport_json
      FROM workbook_version
      WHERE workbook_id = $1 AND id = $2
      LIMIT 1
    `,
    [documentId, id],
  );
  const row = result.rows[0];
  if (
    !row ||
    typeof row.id !== "string" ||
    typeof row.owner_user_id !== "string" ||
    typeof row.name !== "string" ||
    parseInteger(row.revision) === null ||
    !isWorkbookSnapshot(row.snapshot_json)
  ) {
    return null;
  }
  return {
    id: row.id,
    ownerUserId: row.owner_user_id,
    name: row.name,
    revision: parseInteger(row.revision) ?? 0,
    snapshot: row.snapshot_json,
    sheetId: typeof row.sheet_id === "number" ? row.sheet_id : null,
    sheetName: typeof row.sheet_name === "string" ? row.sheet_name : null,
    address: typeof row.address === "string" ? row.address : null,
    viewport: isWorkbookVersionViewport(row.viewport_json) ? row.viewport_json : null,
  };
}
