import { describe, expect, it } from "vitest";
import {
  buildZeroWorkbookMutation,
  isPendingWorkbookMutationInput,
  isWorkbookMutationMethod,
} from "../workbook-sync.js";

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
        method: "updateColumnMetadata",
        args: ["Sheet1", 3, 1, 144, null],
      }),
    ).toMatchObject({
      args: {
        clientMutationId: undefined,
        count: 1,
        documentId: "doc-1",
        hidden: null,
        sheetName: "Sheet1",
        startCol: 3,
        width: 144,
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

    expect(
      buildZeroWorkbookMutation("doc-1", {
        method: "insertRows",
        args: ["Sheet1", 4, 2],
      }),
    ).toMatchObject({
      args: {
        clientMutationId: undefined,
        count: 2,
        documentId: "doc-1",
        sheetName: "Sheet1",
        start: 4,
      },
      "~": "MutateRequest",
    });

    expect(
      buildZeroWorkbookMutation("doc-1", {
        method: "deleteColumns",
        args: ["Sheet1", 1, 3],
      }),
    ).toMatchObject({
      args: {
        clientMutationId: undefined,
        count: 3,
        documentId: "doc-1",
        sheetName: "Sheet1",
        start: 1,
      },
      "~": "MutateRequest",
    });
  });

  it("normalizes legacy updateColumnWidth journal entries onto column metadata mutations", () => {
    const legacyMutation = {
      method: "updateColumnWidth",
      args: ["Sheet1", 5, 168],
    };

    expect(isPendingWorkbookMutationInput(legacyMutation)).toBe(true);
    if (!isPendingWorkbookMutationInput(legacyMutation)) {
      throw new Error("expected legacy mutation to remain readable");
    }

    expect(buildZeroWorkbookMutation("doc-1", legacyMutation)).toMatchObject({
      args: {
        clientMutationId: undefined,
        count: 1,
        documentId: "doc-1",
        hidden: null,
        sheetName: "Sheet1",
        startCol: 5,
        width: 168,
      },
      "~": "MutateRequest",
    });
  });

  it("does not advertise updateColumnWidth as an active workbook mutation method", () => {
    expect(isWorkbookMutationMethod("updateColumnWidth")).toBe(false);
    expect(isWorkbookMutationMethod("updateColumnMetadata")).toBe(true);
    expect(isWorkbookMutationMethod("insertRows")).toBe(true);
    expect(isWorkbookMutationMethod("deleteColumns")).toBe(true);
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
