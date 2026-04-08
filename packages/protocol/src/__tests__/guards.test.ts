import { describe, expect, it } from "vitest";
import { ValueTag } from "../enums.js";
import { isCellRangeRef, isCellSnapshot, isLiteralInput, isWorkbookSnapshot } from "../guards.js";

describe("protocol guards", () => {
  it("accepts workbook snapshots with the shipped shape", () => {
    expect(
      isWorkbookSnapshot({
        version: 1,
        workbook: { name: "guarded" },
        sheets: [],
      }),
    ).toBe(true);
  });

  it("rejects workbook snapshots without a workbook name", () => {
    expect(
      isWorkbookSnapshot({
        version: 1,
        workbook: {},
        sheets: [],
      }),
    ).toBe(false);
  });

  it("accepts cell snapshots with a valid value tag", () => {
    expect(
      isCellSnapshot({
        sheetName: "Sheet1",
        address: "A1",
        value: { tag: ValueTag.Number, value: 7 },
        flags: 0,
        version: 1,
      }),
    ).toBe(true);
  });

  it("rejects cell snapshots with an invalid value tag", () => {
    expect(
      isCellSnapshot({
        sheetName: "Sheet1",
        address: "A1",
        value: { tag: 99, value: 7 },
        flags: 0,
        version: 1,
      }),
    ).toBe(false);
  });

  it("accepts literal inputs and cell range refs with the shipped shapes", () => {
    expect(isLiteralInput(null)).toBe(true);
    expect(isLiteralInput("text")).toBe(true);
    expect(
      isCellRangeRef({
        sheetName: "Sheet1",
        startAddress: "A1",
        endAddress: "B2",
      }),
    ).toBe(true);
  });

  it("rejects malformed cell range refs", () => {
    expect(
      isCellRangeRef({
        sheetName: "Sheet1",
        startAddress: "A1",
      }),
    ).toBe(false);
  });
});
