import type { WorkbookScenarioResponse } from "@bilig/zero-sync";
import type { Queryable } from "./store.js";
import { resolveWorkbookSheetRef } from "./workbook-sheet-ref.js";

export interface WorkbookScenarioViewport {
  readonly rowStart: number;
  readonly rowEnd: number;
  readonly colStart: number;
  readonly colEnd: number;
}

export interface CreateWorkbookScenarioInput {
  readonly workbookId: string;
  readonly documentId: string;
  readonly ownerUserId: string;
  readonly name: string;
  readonly baseRevision: number;
  readonly sheetId?: number | null;
  readonly sheetName?: string | null;
  readonly address?: string | null;
  readonly viewport?: WorkbookScenarioViewport | null;
}

export interface DeleteWorkbookScenarioInput {
  readonly workbookId: string;
  readonly documentId: string;
  readonly ownerUserId: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isWorkbookScenarioViewport(value: unknown): value is WorkbookScenarioViewport {
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

async function ensureWorkbookScenarioOwnership(
  db: Queryable,
  workbookId: string,
  documentId: string,
  ownerUserId: string,
): Promise<void> {
  const existing = await db.query<{ ownerUserId?: unknown }>(
    `
      SELECT owner_user_id AS "ownerUserId"
      FROM workbook_scenario
      WHERE workbook_id = $1 AND document_id = $2
      LIMIT 1
    `,
    [workbookId, documentId],
  );
  const existingOwnerUserId = readExistingOwnerUserId(existing.rows);
  if (existingOwnerUserId && existingOwnerUserId !== ownerUserId) {
    throw new Error("Workbook scenario is owned by another user");
  }
}

export async function ensureWorkbookScenarioSchema(db: Queryable): Promise<void> {
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_scenario (
      document_id TEXT PRIMARY KEY REFERENCES workbooks(id) ON DELETE CASCADE,
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      owner_user_id TEXT NOT NULL,
      name TEXT NOT NULL,
      base_revision BIGINT NOT NULL,
      sheet_id INTEGER,
      sheet_name TEXT,
      address TEXT,
      viewport_json JSONB,
      created_at BIGINT NOT NULL,
      updated_at BIGINT NOT NULL
    );
  `);
  await db.query(
    `CREATE INDEX IF NOT EXISTS workbook_scenario_workbook_updated_idx ON workbook_scenario(workbook_id, owner_user_id, updated_at DESC, created_at DESC);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS workbook_scenario_workbook_owner_idx ON workbook_scenario(workbook_id, owner_user_id);`,
  );
}

export async function createWorkbookScenario(
  db: Queryable,
  input: CreateWorkbookScenarioInput,
): Promise<void> {
  await ensureWorkbookScenarioOwnership(db, input.workbookId, input.documentId, input.ownerUserId);
  const timestamp = Date.now();
  const sheetRef =
    input.sheetId == null && input.sheetName == null
      ? { sheetId: null, sheetName: null }
      : await resolveWorkbookSheetRef(db, {
          documentId: input.workbookId,
          ...(input.sheetId != null ? { sheetId: input.sheetId } : {}),
          ...(input.sheetName != null ? { sheetName: input.sheetName } : {}),
        });
  await db.query(
    `
      INSERT INTO workbook_scenario (
        document_id,
        workbook_id,
        owner_user_id,
        name,
        base_revision,
        sheet_id,
        sheet_name,
        address,
        viewport_json,
        created_at,
        updated_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)
      ON CONFLICT (document_id)
      DO UPDATE SET
        workbook_id = EXCLUDED.workbook_id,
        owner_user_id = EXCLUDED.owner_user_id,
        name = EXCLUDED.name,
        base_revision = EXCLUDED.base_revision,
        sheet_id = EXCLUDED.sheet_id,
        sheet_name = EXCLUDED.sheet_name,
        address = EXCLUDED.address,
        viewport_json = EXCLUDED.viewport_json,
        updated_at = EXCLUDED.updated_at
    `,
    [
      input.documentId,
      input.workbookId,
      input.ownerUserId,
      input.name,
      input.baseRevision,
      sheetRef.sheetId,
      sheetRef.sheetName,
      input.address ?? null,
      input.viewport ? JSON.stringify(input.viewport) : null,
      timestamp,
      timestamp,
    ],
  );
}

export async function deleteWorkbookScenario(
  db: Queryable,
  input: DeleteWorkbookScenarioInput,
): Promise<void> {
  await ensureWorkbookScenarioOwnership(db, input.workbookId, input.documentId, input.ownerUserId);
  await db.query(
    `
      DELETE FROM workbook_scenario
      WHERE workbook_id = $1 AND document_id = $2 AND owner_user_id = $3
    `,
    [input.workbookId, input.documentId, input.ownerUserId],
  );
}

export async function loadWorkbookScenarioByDocument(
  db: Queryable,
  documentId: string,
): Promise<WorkbookScenarioResponse | null> {
  const result = await db.query<{
    document_id?: unknown;
    workbook_id?: unknown;
    owner_user_id?: unknown;
    name?: unknown;
    base_revision?: unknown;
    sheet_id?: unknown;
    sheet_name?: unknown;
    address?: unknown;
    viewport_json?: unknown;
    created_at?: unknown;
    updated_at?: unknown;
  }>(
    `
      SELECT
        document_id,
        workbook_id,
        owner_user_id,
        name,
        base_revision,
        sheet_id,
        sheet_name,
        address,
        viewport_json,
        created_at,
        updated_at
      FROM workbook_scenario
      WHERE document_id = $1
      LIMIT 1
    `,
    [documentId],
  );
  const row = result.rows[0];
  const baseRevision = parseInteger(row?.base_revision);
  const createdAt = parseInteger(row?.created_at);
  const updatedAt = parseInteger(row?.updated_at);
  if (
    !row ||
    typeof row.document_id !== "string" ||
    typeof row.workbook_id !== "string" ||
    typeof row.owner_user_id !== "string" ||
    typeof row.name !== "string" ||
    baseRevision === null ||
    createdAt === null ||
    updatedAt === null
  ) {
    return null;
  }
  return {
    documentId: row.document_id,
    workbookId: row.workbook_id,
    ownerUserId: row.owner_user_id,
    name: row.name,
    baseRevision,
    sheetId: typeof row.sheet_id === "number" ? row.sheet_id : null,
    sheetName: typeof row.sheet_name === "string" ? row.sheet_name : null,
    address: typeof row.address === "string" ? row.address : null,
    viewport: isWorkbookScenarioViewport(row.viewport_json) ? row.viewport_json : null,
    createdAt,
    updatedAt,
  };
}
