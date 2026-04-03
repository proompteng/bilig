import { describe, expect, it } from "vitest";
import { ValueTag, type RecalcMetrics } from "@bilig/protocol";
import type { ViewportPatch } from "@bilig/worker-transport";
import { WorkerViewportCache } from "../viewport-cache.js";

const TEST_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
};

function createPatch(styleId?: string): ViewportPatch {
  return {
    version: 1,
    full: false,
    viewport: {
      sheetName: "Sheet1",
      rowStart: 3,
      rowEnd: 7,
      colStart: 2,
      colEnd: 4,
    },
    metrics: TEST_METRICS,
    styles: [],
    cells: [
      {
        row: 4,
        col: 3,
        snapshot: {
          sheetName: "Sheet1",
          address: "D5",
          value: { tag: ValueTag.Empty },
          flags: 0,
          version: 1,
          ...(styleId ? { styleId } : {}),
        },
        displayText: "",
        copyText: "",
        editorText: "",
        formatId: 0,
        styleId: styleId ?? "style-0",
      },
    ],
    columns: [],
    rows: [],
  };
}

describe("WorkerViewportCache", () => {
  it("accepts equal-version empty snapshots that clear stale styling", () => {
    const cache = new WorkerViewportCache();

    cache.applyViewportPatch(createPatch("style-red"));
    expect(cache.getCell("Sheet1", "D5").styleId).toBe("style-red");

    cache.applyViewportPatch(createPatch());

    expect(cache.getCell("Sheet1", "D5").styleId).toBeUndefined();
  });

  it("reports damage when a style record changes without a newer cell snapshot", () => {
    const cache = new WorkerViewportCache();

    cache.applyViewportPatch({
      ...createPatch("style-fill"),
      styles: [{ id: "style-fill", fill: { backgroundColor: "#c9daf8" } }],
    });

    const damage = cache.applyViewportPatch({
      ...createPatch("style-fill"),
      styles: [{ id: "style-fill", fill: { backgroundColor: "#a4c2f4" } }],
    });

    expect(damage).toEqual([{ cell: [3, 4] }]);
    expect(cache.getCellStyle("style-fill")).toEqual({
      id: "style-fill",
      fill: { backgroundColor: "#a4c2f4" },
    });
  });

  it("drops stale sheet cache entries when sheets disappear", () => {
    const cache = new WorkerViewportCache();

    cache.applyViewportPatch(createPatch());
    expect(cache.peekCell("Sheet1", "D5")).toBeDefined();

    cache.setKnownSheets(["Sheet2"]);

    expect(cache.peekCell("Sheet1", "D5")).toBeUndefined();
  });
});
