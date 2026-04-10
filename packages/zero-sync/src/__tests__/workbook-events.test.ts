import { describe, expect, it } from "vitest";
import { isWorkbookChangeUndoBundle, isWorkbookEventPayload } from "../workbook-events.js";

describe("workbook event guards", () => {
  it("accepts applyBatch payloads with a valid engine op batch", () => {
    expect(
      isWorkbookEventPayload({
        kind: "applyBatch",
        batch: {
          id: "batch-1",
          replicaId: "replica-1",
          clock: { counter: 2 },
          ops: [{ kind: "upsertWorkbook", name: "Book" }],
        },
      }),
    ).toBe(true);
  });

  it("rejects renderCommit payloads with malformed commit ops", () => {
    expect(
      isWorkbookEventPayload({
        kind: "renderCommit",
        ops: [{ kind: "renameSheet", oldName: "Sheet1" }],
      }),
    ).toBe(false);
  });

  it("accepts structural metadata payloads", () => {
    expect(
      isWorkbookEventPayload({
        kind: "updateRowMetadata",
        sheetName: "Sheet1",
        startRow: 1,
        count: 2,
        height: 32,
        hidden: false,
      }),
    ).toBe(true);

    expect(
      isWorkbookEventPayload({
        kind: "setFreezePane",
        sheetName: "Sheet1",
        rows: 1,
        cols: 2,
      }),
    ).toBe(true);

    expect(
      isWorkbookEventPayload({
        kind: "insertRows",
        sheetName: "Sheet1",
        start: 1,
        count: 2,
      }),
    ).toBe(true);

    expect(
      isWorkbookEventPayload({
        kind: "deleteColumns",
        sheetName: "Sheet1",
        start: 3,
        count: 1,
      }),
    ).toBe(true);

    expect(
      isWorkbookEventPayload({
        kind: "redoChange",
        targetRevision: 12,
        targetSummary: "Updated Sheet1!A1",
        sheetName: "Sheet1",
        address: "A1",
        appliedBundle: {
          kind: "engineOps",
          ops: [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 1 }],
        },
      }),
    ).toBe(true);
  });

  it("rejects engine undo bundles with malformed engine ops", () => {
    expect(
      isWorkbookChangeUndoBundle({
        kind: "engineOps",
        ops: [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1" }],
      }),
    ).toBe(false);
  });
});
