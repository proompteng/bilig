import { describe, expect, it } from "vitest";
import type { WorkbookSnapshot } from "@bilig/protocol";
import { SpreadsheetEngine } from "@bilig/core";
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

async function createSnapshot(
  workbookName: string,
  apply: (engine: SpreadsheetEngine) => void | Promise<void>,
): Promise<WorkbookSnapshot> {
  const engine = new SpreadsheetEngine({ workbookName });
  await engine.ready();
  engine.createSheet("Sheet1");
  await apply(engine);
  return engine.exportSnapshot();
}

async function createSnapshotState(
  workbookName: string,
  apply: (engine: SpreadsheetEngine) => void | Promise<void>,
): Promise<{ engine: SpreadsheetEngine; snapshot: WorkbookSnapshot }> {
  const engine = new SpreadsheetEngine({ workbookName });
  await engine.ready();
  engine.createSheet("Sheet1");
  await apply(engine);
  return {
    engine,
    snapshot: engine.exportSnapshot(),
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
          statement.includes("INSERT INTO cell_number_formats"),
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

  it("skips recalc jobs for style-only edits when the workbook is already calculated", async () => {
    const previousSnapshot = createEmptyWorkbookSnapshot("doc-1");
    const nextSnapshot = await createSnapshot("doc-1", (engine) => {
      engine.setRangeStyle(
        { sheetName: "Sheet1", startAddress: "B2", endAddress: "C3" },
        { fill: { backgroundColor: "#abcdef" } },
      );
    });
    const db = createRecordingDb();

    const result = await persistWorkbookMutation(db, "doc-1", {
      previousState: createRuntimeState(previousSnapshot),
      nextSnapshot,
      nextReplicaSnapshot: null,
      updatedBy: "user-1",
      ownerUserId: "owner-1",
      eventPayload: {
        kind: "setRangeStyle",
        range: { sheetName: "Sheet1", startAddress: "B2", endAddress: "C3" },
        patch: { fill: { backgroundColor: "#abcdef" } },
      },
    });

    expect(result.calculatedRevision).toBe(result.revision);
    expect(result.recalcJobId).toBeNull();
    expect(db.statements.some((statement) => statement.includes("INSERT INTO cell_styles"))).toBe(
      true,
    );
    expect(db.statements.some((statement) => statement.includes("DELETE FROM cells"))).toBe(true);
    expect(db.statements.some((statement) => statement.includes("INSERT INTO cells"))).toBe(true);
    expect(db.statements.some((statement) => statement.includes("INSERT INTO recalc_job"))).toBe(
      false,
    );
  });

  it("keeps recalc queued for style-only edits when an older value edit is still pending", async () => {
    const previousSnapshot = createEmptyWorkbookSnapshot("doc-1");
    const nextSnapshot = await createSnapshot("doc-1", (engine) => {
      engine.setRangeStyle(
        { sheetName: "Sheet1", startAddress: "B2", endAddress: "B2" },
        { fill: { backgroundColor: "#abcdef" } },
      );
    });
    const db = createRecordingDb();

    const result = await persistWorkbookMutation(db, "doc-1", {
      previousState: {
        ...createRuntimeState(previousSnapshot),
        headRevision: 5,
        calculatedRevision: 3,
      },
      nextSnapshot,
      nextReplicaSnapshot: null,
      updatedBy: "user-1",
      ownerUserId: "owner-1",
      eventPayload: {
        kind: "setRangeStyle",
        range: { sheetName: "Sheet1", startAddress: "B2", endAddress: "B2" },
        patch: { fill: { backgroundColor: "#abcdef" } },
      },
    });

    expect(result.calculatedRevision).toBe(3);
    expect(result.recalcJobId).toBe("doc-1:recalc:6");
    expect(db.statements.some((statement) => statement.includes("INSERT INTO recalc_job"))).toBe(
      true,
    );
  });

  it("refreshes cell_eval rows when clearing a style range over populated cells", async () => {
    const previousState = await createSnapshotState("doc-1", (engine) => {
      engine.setCellValue("Sheet1", "D5", "asdf");
      engine.setRangeStyle(
        { sheetName: "Sheet1", startAddress: "C4", endAddress: "E8" },
        { fill: { backgroundColor: "#ff0000" } },
      );
    });
    const nextState = await createSnapshotState("doc-1", (engine) => {
      engine.setCellValue("Sheet1", "D5", "asdf");
    });
    const db = createRecordingDb();

    await persistWorkbookMutation(db, "doc-1", {
      previousState: createRuntimeState(previousState.snapshot),
      nextSnapshot: nextState.snapshot,
      nextReplicaSnapshot: null,
      nextEngine: nextState.engine,
      updatedBy: "user-1",
      ownerUserId: "owner-1",
      eventPayload: {
        kind: "clearRangeStyle",
        range: { sheetName: "Sheet1", startAddress: "C4", endAddress: "E8" },
      },
    });

    expect(db.statements.some((statement) => statement.includes("DELETE FROM cells"))).toBe(true);
    expect(db.statements.some((statement) => statement.includes("INSERT INTO cell_eval"))).toBe(
      true,
    );
  });

  it("persists column-width edits through the column metadata hot path without enqueuing recalc", async () => {
    const previousSnapshot = createEmptyWorkbookSnapshot("doc-1");
    const nextSnapshot = await createSnapshot("doc-1", (engine) => {
      engine.updateColumnMetadata("Sheet1", 2, 1, 180, null);
    });
    const db = createRecordingDb();

    const result = await persistWorkbookMutation(db, "doc-1", {
      previousState: createRuntimeState(previousSnapshot),
      nextSnapshot,
      nextReplicaSnapshot: null,
      updatedBy: "user-1",
      ownerUserId: "owner-1",
      eventPayload: {
        kind: "updateColumnWidth",
        sheetName: "Sheet1",
        columnIndex: 2,
        width: 180,
      },
    });

    expect(result.recalcJobId).toBeNull();
    expect(
      db.statements.some((statement) => statement.includes("INSERT INTO column_metadata")),
    ).toBe(true);
    expect(db.statements.some((statement) => statement.includes("INSERT INTO recalc_job"))).toBe(
      false,
    );
  });
});
