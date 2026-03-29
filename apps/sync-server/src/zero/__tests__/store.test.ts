import { describe, expect, it } from "vitest";
import type { WorkbookSnapshot } from "@bilig/protocol";
import {
  createEmptyWorkbookSnapshot,
  persistWorkbookMutation,
  type Queryable,
  type WorkbookRuntimeState,
} from "../store.js";

function cloneSnapshot(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  return structuredClone(snapshot);
}

function normalizeSql(text: string): string {
  return text.replaceAll(/\s+/g, " ").trim();
}

function createRecordingDb(): Queryable & { statements: string[] } {
  const statements: string[] = [];
  return {
    statements,
    async query(text) {
      statements.push(normalizeSql(text));
      return { rows: [] };
    },
  };
}

function createRuntimeState(snapshot: WorkbookSnapshot): WorkbookRuntimeState {
  return {
    snapshot,
    replicaSnapshot: null,
    headRevision: 0,
    calculatedRevision: 0,
    ownerUserId: "owner-1",
  };
}

describe("persistWorkbookMutation", () => {
  it("persists single-cell edits without diffing unrelated source tables", async () => {
    const previousSnapshot = createEmptyWorkbookSnapshot("doc-1");
    const nextSnapshot = cloneSnapshot(previousSnapshot);
    nextSnapshot.sheets[0]?.cells.push({
      address: "A1",
      value: 7,
    });
    const db = createRecordingDb();

    await persistWorkbookMutation(db, "doc-1", {
      previousState: createRuntimeState(previousSnapshot),
      nextSnapshot,
      nextReplicaSnapshot: null,
      updatedBy: "user-1",
      ownerUserId: "owner-1",
      eventPayload: {
        kind: "setCellValue",
        sheetName: "Sheet1",
        address: "A1",
        value: 7,
      },
    });

    expect(db.statements.some((statement) => statement.includes("INSERT INTO cells"))).toBe(true);
    expect(
      db.statements.some(
        (statement) =>
          statement.includes("INSERT INTO sheets") ||
          statement.includes("DELETE FROM sheets") ||
          statement.includes("INSERT INTO row_metadata") ||
          statement.includes("INSERT INTO column_metadata") ||
          statement.includes("INSERT INTO defined_names") ||
          statement.includes("INSERT INTO workbook_metadata") ||
          statement.includes("INSERT INTO cell_styles") ||
          statement.includes("INSERT INTO cell_number_formats") ||
          statement.includes("INSERT INTO sheet_style_ranges") ||
          statement.includes("INSERT INTO sheet_format_ranges"),
      ),
    ).toBe(false);
  });

  it("deletes cleared cells through the focused hot path", async () => {
    const previousSnapshot = createEmptyWorkbookSnapshot("doc-1");
    previousSnapshot.sheets[0]?.cells.push({
      address: "A1",
      value: 7,
    });
    const nextSnapshot = cloneSnapshot(previousSnapshot);
    const [sheet] = nextSnapshot.sheets;
    if (!sheet) {
      throw new Error("Expected a default worksheet");
    }
    sheet.cells = [];
    const db = createRecordingDb();

    await persistWorkbookMutation(db, "doc-1", {
      previousState: createRuntimeState(previousSnapshot),
      nextSnapshot,
      nextReplicaSnapshot: null,
      updatedBy: "user-1",
      ownerUserId: "owner-1",
      eventPayload: {
        kind: "clearCell",
        sheetName: "Sheet1",
        address: "A1",
      },
    });

    expect(
      db.statements.some((statement) =>
        statement.includes(
          "DELETE FROM cells WHERE workbook_id = $1 AND sheet_name = $2 AND address = $3",
        ),
      ),
    ).toBe(true);
  });
});
