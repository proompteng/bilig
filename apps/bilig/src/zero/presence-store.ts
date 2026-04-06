import type { Queryable } from "./store.js";

export interface WorkbookPresenceSheetRef {
  readonly sheetId: number | null;
  readonly sheetName: string | null;
}

export interface UpsertWorkbookPresenceInput {
  readonly documentId: string;
  readonly sessionId: string;
  readonly userId: string;
  readonly sheetId?: number | null;
  readonly sheetName?: string | null;
  readonly address?: string | null;
  readonly selection?: unknown;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function parseSheetRefRow(value: unknown): WorkbookPresenceSheetRef | null {
  if (!isRecord(value)) {
    return null;
  }
  const sheetId = value["sheetId"];
  const sheetName = value["sheetName"];
  return {
    sheetId: typeof sheetId === "number" ? sheetId : null,
    sheetName: typeof sheetName === "string" ? sheetName : null,
  };
}

export async function ensureWorkbookPresenceSchema(db: Queryable): Promise<void> {
  await db.query(`
    CREATE TABLE IF NOT EXISTS presence_coarse (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      session_id TEXT NOT NULL,
      user_id TEXT NOT NULL,
      sheet_id INTEGER,
      sheet_name TEXT,
      address TEXT,
      selection_json JSONB,
      updated_at BIGINT NOT NULL,
      PRIMARY KEY (workbook_id, session_id)
    );
  `);
  await db.query(
    `CREATE INDEX IF NOT EXISTS presence_coarse_workbook_updated_idx ON presence_coarse(workbook_id, updated_at DESC);`,
  );
  await db.query(
    `CREATE INDEX IF NOT EXISTS presence_coarse_updated_idx ON presence_coarse(updated_at);`,
  );
}

export async function resolveWorkbookPresenceSheetRef(
  db: Queryable,
  input: Pick<UpsertWorkbookPresenceInput, "documentId" | "sheetId" | "sheetName">,
): Promise<WorkbookPresenceSheetRef> {
  const fallback = {
    sheetId: input.sheetId ?? null,
    sheetName: input.sheetName ?? null,
  } satisfies WorkbookPresenceSheetRef;

  if (input.sheetId == null && !input.sheetName) {
    return fallback;
  }

  const rows = await db.query<{ sheetId?: unknown; sheetName?: unknown }>(
    `
      SELECT sheet_id AS "sheetId",
             name AS "sheetName"
        FROM sheets
       WHERE workbook_id = $1
         AND (
           ($2::INTEGER IS NOT NULL AND sheet_id = $2)
           OR ($3::TEXT IS NOT NULL AND name = $3)
         )
       ORDER BY sort_order ASC
       LIMIT 1
    `,
    [input.documentId, input.sheetId ?? null, input.sheetName ?? null],
  );
  return parseSheetRefRow(rows.rows[0]) ?? fallback;
}

export async function upsertWorkbookPresence(
  db: Queryable,
  input: UpsertWorkbookPresenceInput,
): Promise<void> {
  const sheetRef = await resolveWorkbookPresenceSheetRef(db, input);
  await db.query(
    `
      INSERT INTO presence_coarse (
        workbook_id,
        session_id,
        user_id,
        sheet_id,
        sheet_name,
        address,
        selection_json,
        updated_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
      ON CONFLICT (workbook_id, session_id)
      DO UPDATE SET
        user_id = EXCLUDED.user_id,
        sheet_id = EXCLUDED.sheet_id,
        sheet_name = EXCLUDED.sheet_name,
        address = EXCLUDED.address,
        selection_json = EXCLUDED.selection_json,
        updated_at = EXCLUDED.updated_at
    `,
    [
      input.documentId,
      input.sessionId,
      input.userId,
      sheetRef.sheetId,
      sheetRef.sheetName,
      input.address ?? null,
      JSON.stringify(input.selection ?? null),
      Date.now(),
    ],
  );
}
