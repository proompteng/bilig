import type { QueryResultRow, Queryable } from "./store.js";

const SHEET_ID_DEPENDENT_TABLES = [
  "presence_coarse",
  "sheet_view",
  "workbook_change",
  "workbook_version",
  "workbook_scenario",
] as const;

export interface WorkbookSheetIdRow {
  readonly workbookId: string;
  readonly name: string;
  readonly sortOrder: number;
  readonly sheetId: number | null;
}

export interface WorkbookSheetIdAssignment {
  readonly workbookId: string;
  readonly name: string;
  readonly previousSheetId: number | null;
  readonly nextSheetId: number;
}

function isValidSheetId(value: number | null): value is number {
  return value !== null && Number.isInteger(value) && value > 0;
}

function readSheetRows(rows: readonly QueryResultRow[]): WorkbookSheetIdRow[] {
  return rows.flatMap((row) => {
    const workbookId = row["workbookId"];
    const name = row["name"];
    const sortOrder = row["sortOrder"];
    const sheetId = row["sheetId"];
    if (
      typeof workbookId !== "string" ||
      workbookId.length === 0 ||
      typeof name !== "string" ||
      name.length === 0 ||
      typeof sortOrder !== "number"
    ) {
      return [];
    }
    return [
      {
        workbookId,
        name,
        sortOrder,
        sheetId: typeof sheetId === "number" ? sheetId : null,
      } satisfies WorkbookSheetIdRow,
    ];
  });
}

export function normalizeWorkbookSheetIdAssignments(
  rows: readonly WorkbookSheetIdRow[],
): WorkbookSheetIdAssignment[] {
  const orderedRows = rows.toSorted(
    (left, right) =>
      left.workbookId.localeCompare(right.workbookId) ||
      left.sortOrder - right.sortOrder ||
      left.name.localeCompare(right.name),
  );
  const assignments: WorkbookSheetIdAssignment[] = [];
  let index = 0;
  while (index < orderedRows.length) {
    const workbookId = orderedRows[index]!.workbookId;
    const workbookRows: WorkbookSheetIdRow[] = [];
    while (index < orderedRows.length && orderedRows[index]!.workbookId === workbookId) {
      workbookRows.push(orderedRows[index]!);
      index += 1;
    }

    const counts = new Map<number, number>();
    workbookRows.forEach((row) => {
      if (!isValidSheetId(row.sheetId)) {
        return;
      }
      counts.set(row.sheetId, (counts.get(row.sheetId) ?? 0) + 1);
    });

    const usedSheetIds = new Set<number>();
    const pendingRows: WorkbookSheetIdRow[] = [];
    workbookRows.forEach((row) => {
      if (
        !isValidSheetId(row.sheetId) ||
        counts.get(row.sheetId) !== 1 ||
        usedSheetIds.has(row.sheetId)
      ) {
        pendingRows.push(row);
        return;
      }
      usedSheetIds.add(row.sheetId);
      assignments.push({
        workbookId: row.workbookId,
        name: row.name,
        previousSheetId: row.sheetId,
        nextSheetId: row.sheetId,
      });
    });

    let nextSheetId = 1;
    pendingRows.forEach((row) => {
      while (usedSheetIds.has(nextSheetId)) {
        nextSheetId += 1;
      }
      usedSheetIds.add(nextSheetId);
      assignments.push({
        workbookId: row.workbookId,
        name: row.name,
        previousSheetId: row.sheetId,
        nextSheetId,
      });
      nextSheetId += 1;
    });
  }

  return assignments.toSorted(
    (left, right) =>
      left.workbookId.localeCompare(right.workbookId) || left.nextSheetId - right.nextSheetId,
  );
}

async function loadExistingTables(
  db: Queryable,
  tableNames: readonly string[],
): Promise<ReadonlySet<string>> {
  const result = await db.query<{ tableName?: unknown }>(
    `
      SELECT tablename AS "tableName"
      FROM pg_tables
      WHERE schemaname = 'public'
        AND tablename = ANY($1::text[])
    `,
    [[...tableNames]],
  );
  return new Set(
    result.rows.flatMap((row) =>
      typeof row.tableName === "string" && row.tableName.length > 0 ? [row.tableName] : [],
    ),
  );
}

export async function repairWorkbookSheetIds(db: Queryable): Promise<void> {
  const result = await db.query(
    `
      SELECT workbook_id AS "workbookId",
             name,
             sort_order AS "sortOrder",
             sheet_id AS "sheetId"
      FROM sheets
      ORDER BY workbook_id ASC, sort_order ASC, name ASC
    `,
  );
  const assignments = normalizeWorkbookSheetIdAssignments(readSheetRows(result.rows));
  const changedAssignments = assignments.filter(
    (assignment) => assignment.previousSheetId !== assignment.nextSheetId,
  );
  if (changedAssignments.length === 0) {
    return;
  }

  await Promise.all(
    changedAssignments.map(async (assignment) => {
      await db.query(
        `
          UPDATE sheets
          SET sheet_id = $3
          WHERE workbook_id = $1
            AND name = $2
        `,
        [assignment.workbookId, assignment.name, assignment.nextSheetId],
      );
    }),
  );

  const existingTables = await loadExistingTables(db, SHEET_ID_DEPENDENT_TABLES);
  await Promise.all(
    SHEET_ID_DEPENDENT_TABLES.filter((tableName) => existingTables.has(tableName)).flatMap(
      (tableName) =>
        changedAssignments.map(async (assignment) => {
          await db.query(
            `
              UPDATE ${tableName}
              SET sheet_id = $3
              WHERE workbook_id = $1
                AND sheet_name = $2
                AND sheet_id IS DISTINCT FROM $3
            `,
            [assignment.workbookId, assignment.name, assignment.nextSheetId],
          );
        }),
    ),
  );
}
