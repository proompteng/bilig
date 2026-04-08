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

  it("rejects engine undo bundles with malformed engine ops", () => {
    expect(
      isWorkbookChangeUndoBundle({
        kind: "engineOps",
        ops: [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1" }],
      }),
    ).toBe(false);
  });
});
