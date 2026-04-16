import type { Queryable } from './store.js'

export interface WorkbookSheetRef {
  readonly sheetId: number | null
  readonly sheetName: string | null
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function parseSheetRefRow(value: unknown): WorkbookSheetRef | null {
  if (!isRecord(value)) {
    return null
  }
  const sheetId = value['sheetId']
  const sheetName = value['sheetName']
  return {
    sheetId: typeof sheetId === 'number' ? sheetId : null,
    sheetName: typeof sheetName === 'string' ? sheetName : null,
  }
}

export async function resolveWorkbookSheetRef(
  db: Queryable,
  input: {
    readonly documentId: string
    readonly sheetId?: number | null
    readonly sheetName?: string | null
  },
): Promise<WorkbookSheetRef> {
  const fallback = {
    sheetId: input.sheetId ?? null,
    sheetName: input.sheetName ?? null,
  } satisfies WorkbookSheetRef

  if (input.sheetId == null && !input.sheetName) {
    return fallback
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
  )
  return parseSheetRefRow(rows.rows[0]) ?? fallback
}
