import { describe, expect, it } from "vitest";

import { ErrorCode, ValueTag, type Viewport } from "@bilig/protocol";

import {
  decodeViewportPatch,
  encodeViewportPatch,
  encodeViewportPatchJson,
  type ViewportPatch,
} from "../index.js";

function createViewport(): Viewport {
  return {
    rowStart: 0,
    rowEnd: 12,
    colStart: 0,
    colEnd: 8,
  };
}

function createPatch(): ViewportPatch {
  return {
    version: 2,
    full: false,
    freezeRows: 1,
    freezeCols: 2,
    viewport: {
      sheetName: "Sheet1",
      ...createViewport(),
    },
    metrics: {
      batchId: 9,
      changedInputCount: 3,
      dirtyFormulaCount: 5,
      wasmFormulaCount: 4,
      jsFormulaCount: 1,
      rangeNodeVisits: 7,
      recalcMs: 1.25,
      compileMs: 0.125,
    },
    styles: [
      {
        id: "style-1",
        fill: { backgroundColor: "#d9ead3" },
        font: {
          family: "JetBrains Mono",
          size: 11,
          bold: true,
          color: "#202124",
        },
        alignment: {
          horizontal: "right",
          vertical: "middle",
          wrap: true,
          indent: 2,
        },
        borders: {
          top: { style: "solid", weight: "thin", color: "#dadce0" },
          bottom: { style: "double", weight: "medium", color: "#1f1f1f" },
        },
      },
    ],
    cells: [
      {
        row: 0,
        col: 0,
        snapshot: {
          sheetName: "Sheet1",
          address: "A1",
          input: 42,
          value: { tag: ValueTag.Number, value: 42 },
          styleId: "style-1",
          format: "0.00",
          flags: 3,
          version: 7,
        },
        displayText: "42.00",
        copyText: "42.00",
        editorText: "42",
        formatId: 4,
        styleId: "style-1",
      },
      {
        row: 1,
        col: 1,
        snapshot: {
          sheetName: "Sheet1",
          address: "B2",
          formula: "=A1+1",
          value: { tag: ValueTag.Error, code: ErrorCode.Value },
          flags: 9,
          version: 11,
        },
        displayText: "#VALUE!",
        copyText: "=A1+1",
        editorText: "=A1+1",
        formatId: 0,
        styleId: "style-0",
      },
      {
        row: 2,
        col: 2,
        snapshot: {
          sheetName: "Sheet1",
          address: "C3",
          input: "hello",
          value: { tag: ValueTag.String, value: "hello", stringId: 17 },
          numberFormatId: "text",
          flags: 1,
          version: 12,
        },
        displayText: "hello",
        copyText: "hello",
        editorText: "hello",
        formatId: 0,
        styleId: "style-0",
      },
    ],
    columns: [{ index: 1, size: 140, hidden: false }],
    rows: [{ index: 2, size: 28, hidden: true }],
  };
}

describe("viewport patch codec", () => {
  it("round-trips viewport patches through the binary codec", () => {
    const patch = createPatch();

    const bytes = encodeViewportPatch(patch);
    const decoded = decodeViewportPatch(bytes);

    expect(bytes[0]).not.toBe("{".charCodeAt(0));
    expect(decoded).toEqual(patch);
  });

  it("decodes legacy JSON viewport patch payloads", () => {
    const patch = createPatch();

    const decoded = decodeViewportPatch(encodeViewportPatchJson(patch));

    expect(decoded).toEqual(patch);
  });
});
