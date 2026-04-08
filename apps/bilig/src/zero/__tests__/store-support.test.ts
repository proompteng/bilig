import { describe, expect, it } from "vitest";
import {
  createEmptyWorkbookSnapshot,
  eventRequiresRecalc,
  normalizeRangeBounds,
  parseCellStyleRecord,
  parseCheckpointPayload,
  parseInteger,
} from "../store-support.js";

describe("store support helpers", () => {
  it("falls back to an empty workbook snapshot when checkpoint payloads are invalid", () => {
    expect(parseCheckpointPayload(null, "book-1")).toEqual(createEmptyWorkbookSnapshot("book-1"));
  });

  it("normalizes reversed range bounds", () => {
    expect(
      normalizeRangeBounds({
        sheetName: "Sheet1",
        startAddress: "D5",
        endAddress: "B2",
      }),
    ).toEqual({
      sheetName: "Sheet1",
      rowStart: 1,
      rowEnd: 4,
      colStart: 1,
      colEnd: 3,
    });
  });

  it("keeps style records but drops invalid nested fields", () => {
    expect(
      parseCellStyleRecord({
        id: "style-1",
        font: { family: "Aptos", size: 12, bold: true, color: 17 },
        alignment: { horizontal: "center", wrap: true, indent: "x" },
        borders: {
          top: { style: "solid", weight: "thin", color: "#111" },
          left: { style: "invalid", weight: "thin", color: "#222" },
        },
      }),
    ).toEqual({
      id: "style-1",
      font: { family: "Aptos", size: 12, bold: true },
      alignment: { horizontal: "center", wrap: true },
      borders: {
        top: { style: "solid", weight: "thin", color: "#111" },
      },
    });
  });

  it("treats formatting-only mutations as no-recalc events", () => {
    expect(
      eventRequiresRecalc({
        kind: "setRangeStyle",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
        patch: { font: { bold: true } },
      }),
    ).toBe(false);

    expect(
      eventRequiresRecalc({
        kind: "setCellValue",
        sheetName: "Sheet1",
        address: "A1",
        value: 123,
      }),
    ).toBe(true);
  });

  it("parses numeric strings into integers", () => {
    expect(parseInteger("42")).toBe(42);
    expect(parseInteger("")).toBe(0);
  });
});
