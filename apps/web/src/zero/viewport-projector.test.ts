import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import {
  createViewportProjectionState,
  projectViewportPatch,
  type CellEvalRow,
  type CellSourceRow,
} from "./viewport-projector.js";

describe("projectViewportPatch", () => {
  it("renders from authoritative computed formatting without overlay queries", () => {
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
          styleId: "style-fresh",
          styleJson: {
            id: "style-fresh",
            fill: { backgroundColor: "#00ff00" },
          },
          formatId: "format-fresh",
          formatCode: "$#,##0.00",
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
