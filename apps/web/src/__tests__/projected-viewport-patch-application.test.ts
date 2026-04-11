import { describe, expect, it } from "vitest";
import { ValueTag, type RecalcMetrics } from "@bilig/protocol";
import type { ViewportPatch } from "@bilig/worker-transport";
import {
  applyProjectedViewportPatch,
  type ProjectedViewportPatchState,
} from "../projected-viewport-patch-application.js";

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
    freezeRows: 0,
    freezeCols: 0,
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

function createPatchState(): ProjectedViewportPatchState {
  return {
    cellSnapshots: new Map(),
    cellKeysBySheet: new Map(),
    cellStyles: new Map([["style-0", { id: "style-0" }]]),
    columnSizesBySheet: new Map(),
    columnWidthsBySheet: new Map(),
    pendingColumnWidthsBySheet: new Map(),
    rowSizesBySheet: new Map(),
    rowHeightsBySheet: new Map(),
    pendingRowHeightsBySheet: new Map(),
    hiddenColumnsBySheet: new Map(),
    hiddenRowsBySheet: new Map(),
    freezeRowsBySheet: new Map(),
    freezeColsBySheet: new Map(),
    knownSheets: new Set(),
    cellAccessTicks: new Map(),
    nextCellAccessTick: 1,
  };
}

describe("applyProjectedViewportPatch", () => {
  it("reports damage when a style record changes without a newer cell snapshot", () => {
    const state = createPatchState();

    applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch("style-fill"),
        styles: [{ id: "style-fill", fill: { backgroundColor: "#c9daf8" } }],
      },
    });

    const result = applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch("style-fill"),
        styles: [{ id: "style-fill", fill: { backgroundColor: "#a4c2f4" } }],
      },
    });

    expect(result.damage).toEqual([{ cell: [3, 4] }]);
    expect(state.cellStyles.get("style-fill")).toEqual({
      id: "style-fill",
      fill: { backgroundColor: "#a4c2f4" },
    });
  });

  it("clears stale viewport cells on full patches without dropping cells outside the viewport", () => {
    const state = createPatchState();
    state.cellSnapshots.set("Sheet1!A1", {
      sheetName: "Sheet1",
      address: "A1",
      value: { tag: ValueTag.String, value: "pinned", stringId: 1 },
      flags: 0,
      version: 1,
    });
    state.cellKeysBySheet.set("Sheet1", new Set(["Sheet1!A1"]));

    applyProjectedViewportPatch({
      state,
      patch: { ...createPatch(), full: true },
    });

    const result = applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch(),
        full: true,
        cells: [],
      },
    });

    expect(result.damage).toEqual([{ cell: [3, 4] }]);
    expect(state.cellSnapshots.get("Sheet1!D5")).toBeUndefined();
    expect(state.cellSnapshots.get("Sheet1!A1")?.value).toEqual({
      tag: ValueTag.String,
      value: "pinned",
      stringId: 1,
    });
  });

  it("tracks freeze metadata and axis changes from a viewport patch", () => {
    const state = createPatchState();

    const result = applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch(),
        freezeRows: 2,
        freezeCols: 1,
        cells: [],
        columns: [{ index: 0, size: 93, hidden: true }],
        rows: [{ index: 0, size: 44, hidden: false }],
      },
    });

    expect(result.freezeChanged).toBe(true);
    expect(result.axisChanged).toBe(true);
    expect(state.freezeRowsBySheet.get("Sheet1")).toBe(2);
    expect(state.freezeColsBySheet.get("Sheet1")).toBe(1);
    expect(state.columnWidthsBySheet.get("Sheet1")?.[0]).toBe(0);
    expect(state.hiddenColumnsBySheet.get("Sheet1")?.[0]).toBe(true);
    expect(state.rowHeightsBySheet.get("Sheet1")?.[0]).toBe(44);
  });
});
