import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import {
  createViewportProjectionState,
  projectViewportPatch,
  type CellEvalRow,
  type CellSourceRow,
  type FormatRangeRow,
  type StyleRangeRow,
} from "./viewport-projector.js";

describe("projectViewportPatch", () => {
  it("prefers fresh style and format ranges over stale computed formatting", () => {
    const state = createViewportProjectionState();
    const sourceCells = new Map<string, CellSourceRow>([
      [
        "A1",
        {
          workbookId: "doc-1",
          sheetName: "Sheet1",
          address: "A1",
          inputValue: 7,
        },
      ],
    ]);
    const cellEval = new Map<string, CellEvalRow>([
      [
        "A1",
        {
          workbookId: "doc-1",
          sheetName: "Sheet1",
          address: "A1",
          value: { tag: ValueTag.Number, value: 7 },
          flags: 0,
          version: 4,
          styleId: "style-stale",
          formatId: "format-stale",
          formatCode: "0.00",
        },
      ],
    ]);
    const styleRanges = new Map<string, StyleRangeRow>([
      [
        "style-range:Sheet1:0:0:0:0",
        {
          id: "style-range:Sheet1:0:0:0:0",
          workbookId: "doc-1",
          sheetName: "Sheet1",
          startRow: 0,
          endRow: 0,
          startCol: 0,
          endCol: 0,
          styleId: "style-fresh",
        },
      ],
    ]);
    const formatRanges = new Map<string, FormatRangeRow>([
      [
        "format-range:Sheet1:0:0:0:0",
        {
          id: "format-range:Sheet1:0:0:0:0",
          workbookId: "doc-1",
          sheetName: "Sheet1",
          startRow: 0,
          endRow: 0,
          startCol: 0,
          endCol: 0,
          formatId: "format-fresh",
        },
      ],
    ]);

    const patch = projectViewportPatch(
      state,
      {
        viewport: { sheetName: "Sheet1", rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 },
        sourceCells,
        cellEval,
        rowMetadata: new Map(),
        columnMetadata: new Map(),
        styleRanges,
        formatRanges,
        stylesById: new Map([
          ["style-0", { id: "style-0" }],
          ["style-stale", { id: "style-stale", fill: { backgroundColor: "#cccccc" } }],
          ["style-fresh", { id: "style-fresh", fill: { backgroundColor: "#00ff00" } }],
        ]),
        numberFormatCodeById: new Map([
          ["format-stale", "0.00"],
          ["format-fresh", "$#,##0.00"],
        ]),
      },
      true,
    );

    expect(patch.cells).toEqual([
      expect.objectContaining({
        row: 0,
        col: 0,
        styleId: "style-fresh",
        snapshot: expect.objectContaining({
          styleId: "style-fresh",
          numberFormatId: "format-fresh",
          format: "$#,##0.00",
        }),
      }),
    ]);
    expect(patch.styles).toEqual(
      expect.arrayContaining([expect.objectContaining({ id: "style-fresh" })]),
    );
  });
});
