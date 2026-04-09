import { describe, expect, it } from "vitest";
import { buildZeroWorkbookMutation } from "../workbook-sync.js";

describe("buildZeroWorkbookMutation", () => {
  it("builds structural metadata mutations", () => {
    expect(
      buildZeroWorkbookMutation("doc-1", {
        method: "updateRowMetadata",
        args: ["Sheet1", 2, 3, 48, true],
      }),
    ).toMatchObject({
      args: {
        clientMutationId: undefined,
        count: 3,
        documentId: "doc-1",
        height: 48,
        hidden: true,
        sheetName: "Sheet1",
        startRow: 2,
      },
      "~": "MutateRequest",
    });

    expect(
      buildZeroWorkbookMutation("doc-1", {
        method: "setFreezePane",
        args: ["Sheet1", 1, 2],
      }),
    ).toMatchObject({
      args: {
        clientMutationId: undefined,
        cols: 2,
        documentId: "doc-1",
        rows: 1,
        sheetName: "Sheet1",
      },
      "~": "MutateRequest",
    });
  });

  it("accepts renderCommit mutations with valid commit ops", () => {
    expect(() =>
      buildZeroWorkbookMutation("doc-1", {
        method: "renderCommit",
        args: [[{ kind: "upsertCell", sheetName: "Sheet1", addr: "A1", value: 7 }]],
      }),
    ).not.toThrow();
  });

  it("rejects renderCommit mutations with malformed commit ops", () => {
    expect(() =>
      buildZeroWorkbookMutation("doc-1", {
        method: "renderCommit",
        args: [[{ kind: "upsertCell", sheetName: "Sheet1", addr: "A1" }]],
      }),
    ).toThrow("Invalid renderCommit args");
  });
});
