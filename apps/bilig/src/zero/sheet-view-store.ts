import type { WorkbookEventPayload } from "@bilig/zero-sync";
import type { Queryable } from "./store.js";
import { resolveWorkbookSheetRef } from "./workbook-sheet-ref.js";

export type WorkbookSheetViewVisibility = "private" | "shared";

export interface WorkbookSheetViewViewport {
  readonly rowStart: number;
  readonly rowEnd: number;
  readonly colStart: number;
  readonly colEnd: number;
}

export interface UpsertWorkbookSheetViewInput {
  readonly documentId: string;
  readonly id: string;
  readonly ownerUserId: string;
  readonly name: string;
  readonly visibility: WorkbookSheetViewVisibility;
  readonly sheetId?: number | null;
  readonly sheetName?: string | null;
  readonly address: string;
  readonly viewport: WorkbookSheetViewViewport;
}

export interface DeleteWorkbookSheetViewInput {
  readonly documentId: string;
  readonly id: string;
  readonly ownerUserId: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function readExistingOwnerUserId(rows: readonly { ownerUserId?: unknown }[]): string | null {
  const ownerUserId = rows[0]?.ownerUserId;
  return typeof ownerUserId === "string" ? ownerUserId : null;
}

async function ensureSheetViewOwnership(
  db: Queryable,
  documentId: string,
  id: string,
  ownerUserId: string,
): Promise<void> {
  const existing = await db.query<{ ownerUserId?: unknown }>(
    `
      SELECT owner_user_id AS "ownerUserId"
      FROM sheet_view
      WHERE workbook_id = $1 AND id = $2
      LIMIT 1
    `,
    [documentId, id],
  );
  const existingOwnerUserId = readExistingOwnerUserId(existing.rows);
  if (existingOwnerUserId && existingOwnerUserId !== ownerUserId) {
    throw new Error("Sheet view is owned by another user");
  }
}

export async function ensureWorkbookSheetViewSchema(db: Queryable): Promise<void> {
  await db.query(`
    CREATE TABLE IF NOT EXISTS sheet_view (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      id TEXT NOT NULL,
      owner_user_id TEXT NOT NULL,
      name TEXT NOT NULL,
      visibility TEXT NOT NULL CHECK (visibility IN ('private', 'shared')),
      sheet_id INTEGER NOT NULL,
      sheet_name TEXT NOT NULL,
      address TEXT NOT NULL,
      viewport_json JSONB NOT NULL,
      updated_at BIGINT NOT NULL,
      PRIMARY KEY (workbook_id, id)
    );
  `);
  await db.query(
    `CREATE INDEX IF NOT EXISTS sheet_view_workbook_updated_idx ON sheet_view(workbook_id, updated_at DESC);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS sheet_view_workbook_owner_idx ON sheet_view(workbook_id, owner_user_id);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS sheet_view_workbook_sheet_idx ON sheet_view(workbook_id, sheet_id);`,
  );
}

export async function upsertWorkbookSheetView(
  db: Queryable,
  input: UpsertWorkbookSheetViewInput,
): Promise<void> {
  await ensureSheetViewOwnership(db, input.documentId, input.id, input.ownerUserId);
  const sheetRef = await resolveWorkbookSheetRef(db, input);
  if (sheetRef.sheetId == null || sheetRef.sheetName == null) {
    throw new Error("Cannot save a workbook view for an unknown sheet");
  }
  await db.query(
    `
      INSERT INTO sheet_view (
        workbook_id,
        id,
        owner_user_id,
        name,
        visibility,
        sheet_id,
        sheet_name,
        address,
        viewport_json,
        updated_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)
      ON CONFLICT (workbook_id, id)
      DO UPDATE SET
        owner_user_id = EXCLUDED.owner_user_id,
        name = EXCLUDED.name,
        visibility = EXCLUDED.visibility,
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
      input.visibility,
      sheetRef.sheetId,
      sheetRef.sheetName,
      input.address,
      JSON.stringify(input.viewport),
      Date.now(),
    ],
  );
}

export async function deleteWorkbookSheetView(
  db: Queryable,
  input: DeleteWorkbookSheetViewInput,
): Promise<void> {
  await ensureSheetViewOwnership(db, input.documentId, input.id, input.ownerUserId);
  await db.query(
    `
      DELETE FROM sheet_view
      WHERE workbook_id = $1 AND id = $2 AND owner_user_id = $3
    `,
    [input.documentId, input.id, input.ownerUserId],
  );
}

async function renameWorkbookSheetViews(
  db: Queryable,
  documentId: string,
  nextName: string,
): Promise<void> {
  const sheetRef = await resolveWorkbookSheetRef(db, {
    documentId,
    sheetName: nextName,
  });
  if (sheetRef.sheetId == null || sheetRef.sheetName == null) {
    return;
  }
  await db.query(
    `
      UPDATE sheet_view
      SET sheet_name = $3
      WHERE workbook_id = $1 AND sheet_id = $2
    `,
    [documentId, sheetRef.sheetId, sheetRef.sheetName],
  );
}

async function deleteWorkbookSheetViews(
  db: Queryable,
  documentId: string,
  sheetName: string,
): Promise<void> {
  await db.query(
    `
      DELETE FROM sheet_view
      WHERE workbook_id = $1 AND sheet_name = $2
    `,
    [documentId, sheetName],
  );
}

export async function reconcileWorkbookSheetViews(input: {
  readonly db: Queryable;
  readonly documentId: string;
  readonly payload: WorkbookEventPayload;
}): Promise<void> {
  if (input.payload.kind !== "renderCommit") {
    return;
  }
  const work: Promise<void>[] = [];
  for (const op of input.payload.ops) {
    if (!isRecord(op) || typeof op["kind"] !== "string") {
      continue;
    }
    if (
      op["kind"] === "renameSheet" &&
      typeof op["newName"] === "string" &&
      op["newName"].length > 0
    ) {
      work.push(renameWorkbookSheetViews(input.db, input.documentId, op["newName"]));
      continue;
    }
    if (op["kind"] === "deleteSheet" && typeof op["name"] === "string" && op["name"].length > 0) {
      work.push(deleteWorkbookSheetViews(input.db, input.documentId, op["name"]));
    }
  }
  await Promise.all(work);
}
