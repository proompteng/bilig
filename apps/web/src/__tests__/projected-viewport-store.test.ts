import { describe, expect, it } from "vitest";
import { ValueTag, type RecalcMetrics } from "@bilig/protocol";
import type { ViewportPatch } from "@bilig/worker-transport";
import { ProjectedViewportStore } from "../projected-viewport-store.js";

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

function createColumnPatch(size: number): ViewportPatch {
  return {
    ...createPatch(),
    cells: [],
    columns: [{ index: 0, size, hidden: false }],
  };
}

describe("ProjectedViewportStore", () => {
  it("accepts equal-version empty snapshots that clear stale styling", () => {
    const cache = new ProjectedViewportStore();

    cache.applyViewportPatch(createPatch("style-red"));
    expect(cache.getCell("Sheet1", "D5").styleId).toBe("style-red");

    cache.applyViewportPatch(createPatch());

    expect(cache.getCell("Sheet1", "D5").styleId).toBeUndefined();
  });

  it("keeps an equal-version local formula snapshot when a later patch drops the formula", () => {
    const cache = new ProjectedViewportStore();

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: "Sheet1",
            address: "D5",
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: "TRUE",
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: "style-0",
        },
      ],
    });

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: "Sheet1",
            address: "D5",
            value: { tag: ValueTag.Boolean, value: false },
            flags: 0,
            version: 3,
          },
          displayText: "FALSE",
          copyText: "FALSE",
          editorText: "FALSE",
          formatId: 0,
          styleId: "style-0",
        },
      ],
    });

    expect(cache.getCell("Sheet1", "D5")).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    });
  });

  it("keeps a local formula snapshot when a newer eval-only patch drops source metadata", () => {
    const cache = new ProjectedViewportStore();

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: "Sheet1",
            address: "D5",
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: "TRUE",
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: "style-0",
        },
      ],
    });

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: "Sheet1",
            address: "D5",
            value: { tag: ValueTag.Boolean, value: false },
            flags: 0,
            version: 4,
          },
          displayText: "FALSE",
          copyText: "FALSE",
          editorText: "FALSE",
          formatId: 0,
          styleId: "style-0",
        },
      ],
    });

    expect(cache.getCell("Sheet1", "D5")).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    });
  });

  it("accepts a newer literal snapshot when the source input is present", () => {
    const cache = new ProjectedViewportStore();

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: "Sheet1",
            address: "D5",
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: "TRUE",
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: "style-0",
        },
      ],
    });

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: "Sheet1",
            address: "D5",
            value: { tag: ValueTag.Boolean, value: false },
            input: false,
            flags: 0,
            version: 4,
          },
          displayText: "FALSE",
          copyText: "FALSE",
          editorText: "FALSE",
          formatId: 0,
          styleId: "style-0",
        },
      ],
    });

    const snapshot = cache.getCell("Sheet1", "D5");
    expect(snapshot).toMatchObject({
      value: { tag: ValueTag.Boolean, value: false },
      input: false,
      version: 4,
    });
    expect("formula" in snapshot).toBe(false);
  });

  it("reports damage when a style record changes without a newer cell snapshot", () => {
    const cache = new ProjectedViewportStore();

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
    const cache = new ProjectedViewportStore();

    cache.applyViewportPatch(createPatch());
    expect(cache.peekCell("Sheet1", "D5")).toBeDefined();

    cache.setKnownSheets(["Sheet2"]);

    expect(cache.peekCell("Sheet1", "D5")).toBeUndefined();
  });

  it("keeps a pending local column width across matching patches until the mutation is acked", () => {
    const cache = new ProjectedViewportStore();

    cache.setColumnWidth("Sheet1", 0, 68);
    cache.applyViewportPatch(createColumnPatch(68));
    cache.applyViewportPatch(createColumnPatch(93));

    expect(cache.getColumnWidths("Sheet1")[0]).toBe(68);

    cache.ackColumnWidth("Sheet1", 0, 68);
    cache.applyViewportPatch(createColumnPatch(93));

    expect(cache.getColumnWidths("Sheet1")[0]).toBe(93);
  });

  it("rolls back a failed local column width mutation without leaving a pending width behind", () => {
    const cache = new ProjectedViewportStore();

    cache.setColumnWidth("Sheet1", 0, 68);
    cache.rollbackColumnWidth("Sheet1", 0, undefined);
    cache.applyViewportPatch(createColumnPatch(104));

    expect(cache.getColumnWidths("Sheet1")[0]).toBe(104);
  });
});
