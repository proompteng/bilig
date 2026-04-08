import { describe, expect, it } from "vitest";
import { buildZeroWorkbookMutation } from "../workbook-sync.js";

describe("buildZeroWorkbookMutation", () => {
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
