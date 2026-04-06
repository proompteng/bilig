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

  it("keeps an equal-version local formula snapshot when a later patch drops the formula", () => {
    const cache = new WorkerViewportCache();

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
    const cache = new WorkerViewportCache();

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
    const cache = new WorkerViewportCache();

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

  it("applies optimistic style patches to cached cells in range", () => {
    const cache = new WorkerViewportCache();

    cache.applyViewportPatch(createPatch());
    cache.applyOptimisticRangeStyle(
      { sheetName: "Sheet1", startAddress: "D5", endAddress: "D5" },
      { fill: { backgroundColor: "#c9daf8" } },
    );

    const snapshot = cache.getCell("Sheet1", "D5");
    const style = cache.getCellStyle(snapshot.styleId);

    expect(style).toMatchObject({
      fill: { backgroundColor: "#c9daf8" },
    });
  });

  it("clears optimistic style fields without dropping unrelated style state", () => {
    const cache = new WorkerViewportCache();

    cache.applyViewportPatch({
      ...createPatch("style-filled"),
      styles: [
        {
          id: "style-filled",
          fill: { backgroundColor: "#c9daf8" },
          font: { bold: true, color: "#111827" },
        },
      ],
    });

    cache.clearOptimisticRangeStyle({ sheetName: "Sheet1", startAddress: "D5", endAddress: "D5" }, [
      "backgroundColor",
    ]);

    const snapshot = cache.getCell("Sheet1", "D5");
    const style = cache.getCellStyle(snapshot.styleId);

    expect(style).toMatchObject({
      font: { bold: true, color: "#111827" },
    });
    expect(style?.fill).toBeUndefined();
  });

  it("clears cached cell contents optimistically without dropping formatting", () => {
    const cache = new WorkerViewportCache();

    cache.applyViewportPatch({
      ...createPatch("style-filled"),
      styles: [
        {
          id: "style-filled",
          fill: { backgroundColor: "#c9daf8" },
          font: { bold: true, color: "#111827" },
        },
      ],
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: "Sheet1",
            address: "D5",
            value: { tag: ValueTag.String, value: "hello", stringId: 1 },
            input: "hello",
            formula: 'UPPER("hello")',
            flags: 0,
            version: 3,
            styleId: "style-filled",
            format: "0.00",
          },
          displayText: "HELLO",
          copyText: "HELLO",
          editorText: '=UPPER("hello")',
          formatId: 0,
          styleId: "style-filled",
        },
      ],
    });

    cache.applyOptimisticClearRange({ sheetName: "Sheet1", startAddress: "D5", endAddress: "D5" });

    const snapshot = cache.getCell("Sheet1", "D5");
    const style = cache.getCellStyle(snapshot.styleId);

    expect(snapshot.value).toEqual({ tag: ValueTag.Empty });
    expect(snapshot.input).toBeUndefined();
    expect(snapshot.formula).toBeUndefined();
    expect(snapshot.styleId).toBe("style-filled");
    expect(snapshot.format).toBe("0.00");
    expect(snapshot.version).toBe(4);
    expect(style).toMatchObject({
      fill: { backgroundColor: "#c9daf8" },
      font: { bold: true, color: "#111827" },
    });
  });
});
