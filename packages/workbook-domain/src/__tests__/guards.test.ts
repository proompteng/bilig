import { describe, expect, it } from "vitest";
import { isEngineOp, isEngineOpBatch } from "../index.js";

describe("workbook domain guards", () => {
  it("accepts engine op batches with valid workbook ops", () => {
    expect(
      isEngineOpBatch({
        id: "batch-1",
        replicaId: "replica-1",
        clock: { counter: 4 },
        ops: [
          { kind: "upsertWorkbook", name: "Book" },
          {
            kind: "setDataValidation",
            validation: {
              range: {
                sheetName: "Sheet1",
                startAddress: "D2",
                endAddress: "D10",
              },
              rule: {
                kind: "list",
                values: ["Draft", "Final"],
              },
              allowBlank: false,
            },
          },
          {
            kind: "upsertCommentThread",
            thread: {
              threadId: "thread-1",
              sheetName: "Sheet1",
              address: "E2",
              comments: [{ id: "comment-1", body: "Check this total." }],
            },
          },
          {
            kind: "upsertNote",
            note: {
              sheetName: "Sheet1",
              address: "F3",
              text: "Manual override",
            },
          },
          {
            kind: "upsertConditionalFormat",
            format: {
              id: "cf-1",
              range: {
                sheetName: "Sheet1",
                startAddress: "A1",
                endAddress: "A5",
              },
              rule: {
                kind: "cellIs",
                operator: "greaterThan",
                values: [10],
              },
              style: {
                fill: { backgroundColor: "#ff0000" },
              },
            },
          },
          {
            kind: "setSheetProtection",
            protection: {
              sheetName: "Sheet1",
              hideFormulas: true,
            },
          },
          {
            kind: "upsertRangeProtection",
            protection: {
              id: "protect-a1",
              range: {
                sheetName: "Sheet1",
                startAddress: "A1",
                endAddress: "B2",
              },
              hideFormulas: true,
            },
          },
          {
            kind: "upsertPivotTable",
            name: "Pivot1",
            sheetName: "Sheet1",
            address: "F1",
            source: {
              sheetName: "Sheet1",
              startAddress: "A1",
              endAddress: "C10",
            },
            groupBy: ["Region"],
            values: [{ sourceColumn: "Sales", summarizeBy: "sum", outputLabel: "Total Sales" }],
            rows: 10,
            cols: 3,
          },
        ],
      }),
    ).toBe(true);
  });

  it("rejects engine ops with malformed nested payloads", () => {
    expect(
      isEngineOp({
        kind: "upsertCellStyle",
        style: {
          id: "style-1",
          fill: {
            backgroundColor: 42,
          },
        },
      }),
    ).toBe(false);

    expect(
      isEngineOpBatch({
        id: "batch-1",
        replicaId: "replica-1",
        clock: { counter: 4 },
        ops: [
          { kind: "setSort", sheetName: "Sheet1", range: { sheetName: "Sheet1" }, keys: [] },
          {
            kind: "setDataValidation",
            validation: {
              range: {
                sheetName: "Sheet1",
                startAddress: "A1",
                endAddress: "A5",
              },
              rule: {
                kind: "list",
                values: [undefined],
              },
            },
          },
          {
            kind: "upsertCommentThread",
            thread: {
              threadId: "thread-1",
              sheetName: "Sheet1",
              address: "E2",
              comments: [{ id: "comment-1" }],
            },
          },
          {
            kind: "upsertConditionalFormat",
            format: {
              id: "cf-1",
              range: {
                sheetName: "Sheet1",
                startAddress: "A1",
                endAddress: "A5",
              },
              rule: {
                kind: "cellIs",
                operator: "greaterThan",
                values: [],
              },
              style: "bad",
            },
          },
          {
            kind: "upsertRangeProtection",
            protection: {
              id: "protect-a1",
              range: {
                sheetName: "Sheet1",
                startAddress: "A1",
                endAddress: "B2",
              },
              hideFormulas: "yes",
            },
          },
        ],
      }),
    ).toBe(false);
  });
});
