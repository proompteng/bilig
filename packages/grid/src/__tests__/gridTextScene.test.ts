import { describe, expect, test } from "vitest";
import { ValueTag, type CellSnapshot, type CellStyleRecord } from "@bilig/protocol";
import { buildGridTextScene } from "../gridTextScene.js";
import type { GridEngineLike } from "../grid-engine.js";
import { getGridMetrics } from "../gridMetrics.js";

function createCellSnapshot(
  value: CellSnapshot["value"],
  styleId: string | undefined = "style-1",
): CellSnapshot {
  return {
    address: "A1",
    input: "",
    formula: null,
    styleId,
    value,
    format: null,
    transient: null,
    dependencies: [],
    volatile: false,
  };
}

function makeEngine(
  styles: Record<string, CellStyleRecord>,
  snapshot: CellSnapshot = createCellSnapshot({ tag: ValueTag.String, value: "hello" }),
): GridEngineLike {
  return {
    getCell: () => snapshot,
    getCellStyle: (styleId) => (styleId ? styles[styleId] : undefined),
    subscribeCells: () => () => {},
    workbook: {
      getSheet: () => undefined,
    },
  };
}

describe("gridTextScene", () => {
  test("builds cell text items with resolved alignment and style", () => {
    const engine = makeEngine({
      "style-1": {
        alignment: { horizontal: "right" },
        font: { bold: true, color: "#ff0000", italic: true, size: 14 },
      },
    });

    const scene = buildGridTextScene({
      engine,
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [0, 0],
      sheetName: "Sheet1",
      visibleItems: [[0, 0]],
      visibleRegion: { range: { x: 0, y: 0, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 100, top: 200 },
      getCellBounds: () => ({ x: 110, y: 220, width: 90, height: 22 }),
    });

    expect(scene.items).toContainEqual({
      x: 10,
      y: 20,
      width: 90,
      height: 22,
      text: "hello",
      align: "right",
      wrap: false,
      color: "#ff0000",
      font: 'italic 700 14px "JetBrainsMono Nerd Font","JetBrains Mono",monospace',
      fontSize: 14,
      underline: false,
      strike: false,
    });
  });

  test("adds column headers and row markers with selected header emphasis", () => {
    const scene = buildGridTextScene({
      engine: makeEngine({}, createCellSnapshot({ tag: ValueTag.Empty })),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [2, 3],
      selectionRange: { x: 2, y: 3, width: 1, height: 1 },
      sheetName: "Sheet1",
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0 },
      getCellBounds: () => ({ x: 254, y: 90, width: 104, height: 22 }),
    });

    expect(scene.items).toContainEqual({
      x: 46,
      y: 0,
      width: 104,
      height: 24,
      text: "C",
      align: "center",
      wrap: false,
      color: "#1f7a43",
      font: '500 11px "JetBrainsMono Nerd Font","JetBrains Mono",monospace',
      fontSize: 11,
      underline: false,
      strike: false,
    });
    expect(scene.items).toContainEqual({
      x: 0,
      y: 24,
      width: 46,
      height: 22,
      text: "4",
      align: "right",
      wrap: false,
      color: "#1f7a43",
      font: '500 11px "JetBrainsMono Nerd Font","JetBrains Mono",monospace',
      fontSize: 11,
      underline: false,
      strike: false,
    });
  });

  test("adds hovered and active-drag header text emphasis", () => {
    const scene = buildGridTextScene({
      engine: makeEngine({}, createCellSnapshot({ tag: ValueTag.Empty })),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      activeHeaderDrag: { kind: "column", index: 2 },
      hoveredHeader: { kind: "row", index: 4 },
      selectedCell: [2, 3],
      selectionRange: { x: 2, y: 3, width: 1, height: 1 },
      sheetName: "Sheet1",
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 2 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0 },
      getCellBounds: () => ({ x: 254, y: 90, width: 104, height: 22 }),
    });

    expect(scene.items).toContainEqual({
      x: 46,
      y: 0,
      width: 104,
      height: 24,
      text: "C",
      align: "center",
      wrap: false,
      color: "#176239",
      font: '500 11px "JetBrainsMono Nerd Font","JetBrains Mono",monospace',
      fontSize: 11,
      underline: false,
      strike: false,
    });
    expect(scene.items).toContainEqual({
      x: 0,
      y: 24,
      width: 46,
      height: 22,
      text: "4",
      align: "right",
      wrap: false,
      color: "#1f7a43",
      font: '500 11px "JetBrainsMono Nerd Font","JetBrains Mono",monospace',
      fontSize: 11,
      underline: false,
      strike: false,
    });
    expect(scene.items).toContainEqual({
      x: 0,
      y: 46,
      width: 46,
      height: 22,
      text: "5",
      align: "right",
      wrap: false,
      color: "#3c4043",
      font: '500 11px "JetBrainsMono Nerd Font","JetBrains Mono",monospace',
      fontSize: 11,
      underline: false,
      strike: false,
    });
  });

  test("adds resize-guide emphasis to the active column header label", () => {
    const scene = buildGridTextScene({
      engine: makeEngine({}, createCellSnapshot({ tag: ValueTag.Empty })),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      resizeGuideColumn: 2,
      selectedCell: [0, 0],
      sheetName: "Sheet1",
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0 },
      getCellBounds: () => ({ x: 254, y: 90, width: 104, height: 22 }),
    });

    expect(scene.items).toContainEqual({
      x: 46,
      y: 0,
      width: 104,
      height: 24,
      text: "C",
      align: "center",
      wrap: false,
      color: "#176239",
      font: '500 11px "JetBrainsMono Nerd Font","JetBrains Mono",monospace',
      fontSize: 11,
      underline: false,
      strike: false,
    });
  });
});
