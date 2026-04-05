import { describe, expect, test } from "vitest";
import { ValueTag, type CellSnapshot, type CellStyleRecord } from "@bilig/protocol";
import { buildGridGpuScene, parseGpuColor } from "../gridGpuScene.js";
import type { GridEngineLike } from "../grid-engine.js";
import { getGridMetrics } from "../gridMetrics.js";
import { CompactSelection } from "../gridTypes.js";

function createCellSnapshot(
  styleId: string | undefined,
  value: CellSnapshot["value"] = { tag: ValueTag.String, value: "" },
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
  snapshot: CellSnapshot = createCellSnapshot("style-1"),
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

function createSelection() {
  return {
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty(),
    current: undefined,
  };
}

describe("gridGpuScene", () => {
  test("parses rgb and hex colors into normalized GPU channels", () => {
    expect(parseGpuColor("#336699")).toEqual({
      r: 0x33 / 255,
      g: 0x66 / 255,
      b: 0x99 / 255,
      a: 1,
    });
    expect(parseGpuColor("rgba(12, 34, 56, 0.5)")).toEqual({
      r: 12 / 255,
      g: 34 / 255,
      b: 56 / 255,
      a: 0.5,
    });
  });

  test("builds fill rects and solid border rects from visible cells", () => {
    const engine = makeEngine({
      "style-1": {
        fill: { backgroundColor: "#ff0000" },
        borders: {
          top: { color: "#111111", style: "solid", weight: "medium" },
        },
      },
    });

    const scene = buildGridGpuScene({
      engine,
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      gridSelection: createSelection(),
      selectedCell: [0, 0],
      sheetName: "Sheet1",
      visibleItems: [[0, 0]],
      visibleRegion: { range: { x: 0, y: 0, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 100, top: 200 },
      getCellBounds: () => ({ x: 110, y: 220, width: 90, height: 22 }),
    });

    expect(scene.fillRects).toContainEqual({
      x: 11,
      y: 21,
      width: 88,
      height: 20,
      color: { r: 1, g: 0, b: 0, a: 1 },
    });
    expect(scene.borderRects).toContainEqual({
      x: 10,
      y: 19,
      width: 90,
      height: 2,
      color: {
        r: 0x11 / 255,
        g: 0x11 / 255,
        b: 0x11 / 255,
        a: 1,
      },
    });
    expect(scene.borderRects).toContainEqual({
      x: 10,
      y: 20,
      width: 1,
      height: 22,
      color: { r: 0xe3 / 255, g: 0xe9 / 255, b: 0xf0 / 255, a: 1 },
    });
  });

  test("expands patterned and double borders into GPU rectangles", () => {
    const engine = makeEngine({
      "style-1": {
        borders: {
          left: { color: "#000000", style: "dashed", weight: "thin" },
          bottom: { color: "#000000", style: "double", weight: "thin" },
        },
      },
    });

    const scene = buildGridGpuScene({
      engine,
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      gridSelection: createSelection(),
      selectedCell: [0, 0],
      sheetName: "Sheet1",
      visibleItems: [[0, 0]],
      visibleRegion: { range: { x: 0, y: 0, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0 },
      getCellBounds: () => ({ x: 0, y: 0, width: 20, height: 12 }),
    });

    expect(scene.borderRects.length).toBeGreaterThanOrEqual(4);
    expect(scene.borderRects.some((rect) => rect.width === 1 && rect.height <= 6)).toBe(true);
    expect(scene.borderRects.filter((rect) => rect.height === 1).length).toBeGreaterThanOrEqual(2);
  });

  test("adds GPU selection fill and outline for the active range", () => {
    const scene = buildGridGpuScene({
      engine: makeEngine({}),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      gridSelection: createSelection(),
      selectedCell: [1, 2],
      selectionRange: { x: 1, y: 2, width: 2, height: 3 },
      sheetName: "Sheet1",
      visibleItems: [
        [1, 2],
        [2, 2],
        [1, 3],
        [2, 3],
        [1, 4],
        [2, 4],
      ],
      visibleRegion: { range: { x: 1, y: 2, width: 2, height: 3 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0 },
      getCellBounds: (col, row) => ({
        x: col * 100,
        y: row * 24,
        width: 100,
        height: 24,
      }),
    });

    expect(scene.fillRects).toContainEqual({
      x: 101,
      y: 49,
      width: 198,
      height: 70,
      color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 0.06 },
    });
    expect(scene.borderRects).toContainEqual({
      x: 100,
      y: 48,
      width: 200,
      height: 1,
      color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 1 },
    });
    expect(scene.fillRects.some((rect) => rect.width === 6 && rect.height === 6)).toBe(false);
  });

  test("adds GPU-backed header backgrounds and selection highlights", () => {
    const scene = buildGridGpuScene({
      engine: makeEngine({}),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      gridSelection: createSelection(),
      selectedCell: [2, 3],
      selectionRange: { x: 2, y: 3, width: 1, height: 1 },
      sheetName: "Sheet1",
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0 },
      getCellBounds: () => ({ x: 254, y: 90, width: 104, height: 22 }),
    });

    expect(scene.fillRects).toContainEqual({
      x: 0,
      y: 0,
      width: 46,
      height: 24,
      color: { r: 248 / 255, g: 249 / 255, b: 250 / 255, a: 1 },
    });
    expect(scene.fillRects).toContainEqual({
      x: 46,
      y: 0,
      width: 104,
      height: 24,
      color: { r: 230 / 255, g: 244 / 255, b: 234 / 255, a: 1 },
    });
    expect(scene.fillRects).toContainEqual({
      x: 0,
      y: 24,
      width: 46,
      height: 22,
      color: { r: 230 / 255, g: 244 / 255, b: 234 / 255, a: 1 },
    });
  });

  test("adds GPU-backed header and body hover affordances", () => {
    const scene = buildGridGpuScene({
      engine: makeEngine({}),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      gridSelection: createSelection(),
      hoveredCell: [2, 3],
      hoveredHeader: { kind: "column", index: 2 },
      selectedCell: [0, 0],
      sheetName: "Sheet1",
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0 },
      getCellBounds: () => ({ x: 254, y: 90, width: 104, height: 22 }),
    });

    expect(scene.fillRects).toContainEqual({
      x: 46,
      y: 0,
      width: 104,
      height: 24,
      color: { r: 241 / 255, g: 243 / 255, b: 244 / 255, a: 1 },
    });
    expect(scene.fillRects).toContainEqual({
      x: 255,
      y: 91,
      width: 102,
      height: 20,
      color: { r: 95 / 255, g: 99 / 255, b: 104 / 255, a: 0.06 },
    });
  });

  test("adds a GPU resize guide for hovered or active column resize", () => {
    const scene = buildGridGpuScene({
      engine: makeEngine({}),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      gridSelection: createSelection(),
      resizeGuideColumn: 2,
      selectedCell: [0, 0],
      sheetName: "Sheet1",
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 2 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0 },
      getCellBounds: () => ({ x: 254, y: 90, width: 104, height: 22 }),
    });

    expect(scene.fillRects).toContainEqual({
      x: 147,
      y: 0,
      width: 6,
      height: 68,
      color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 0.18 },
    });
    expect(scene.borderRects).toContainEqual({
      x: 149,
      y: 0,
      width: 2,
      height: 68,
      color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 1 },
    });
  });

  test("adds GPU drag guides for active column header drags", () => {
    const scene = buildGridGpuScene({
      engine: makeEngine({}),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      gridSelection: {
        columns: CompactSelection.fromSingleSelection([1, 3]),
        rows: CompactSelection.empty(),
        current: undefined,
      },
      activeHeaderDrag: { kind: "column", index: 1 },
      selectedCell: [1, 2],
      sheetName: "Sheet1",
      visibleItems: [
        [1, 2],
        [2, 2],
      ],
      visibleRegion: { range: { x: 1, y: 2, width: 2, height: 2 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0 },
      getCellBounds: () => ({ x: 146, y: 48, width: 100, height: 24 }),
    });

    expect(scene.borderRects).toContainEqual({
      x: 46,
      y: 0,
      width: 2,
      height: 68,
      color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 1 },
    });
    expect(scene.borderRects).toContainEqual({
      x: 252,
      y: 0,
      width: 2,
      height: 68,
      color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 1 },
    });
    expect(scene.fillRects).toContainEqual({
      x: 46,
      y: 21,
      width: 104,
      height: 3,
      color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 1 },
    });
  });

  test("adds GPU drag guides for active row header drags", () => {
    const scene = buildGridGpuScene({
      engine: makeEngine({}),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      gridSelection: {
        columns: CompactSelection.empty(),
        rows: CompactSelection.fromSingleSelection([2, 4]),
        current: undefined,
      },
      activeHeaderDrag: { kind: "row", index: 2 },
      selectedCell: [1, 2],
      sheetName: "Sheet1",
      visibleItems: [
        [1, 2],
        [2, 2],
      ],
      visibleRegion: { range: { x: 1, y: 2, width: 2, height: 2 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0 },
      getCellBounds: () => ({ x: 146, y: 48, width: 100, height: 24 }),
    });

    expect(scene.borderRects).toContainEqual({
      x: 0,
      y: 24,
      width: 254,
      height: 2,
      color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 1 },
    });
    expect(scene.borderRects).toContainEqual({
      x: 0,
      y: 66,
      width: 254,
      height: 2,
      color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 1 },
    });
    expect(scene.fillRects).toContainEqual({
      x: 43,
      y: 24,
      width: 3,
      height: 22,
      color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 1 },
    });
  });

  test("adds GPU body highlights for column and row slice selections", () => {
    const scene = buildGridGpuScene({
      engine: makeEngine({}),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      gridSelection: {
        columns: CompactSelection.fromSingleSelection([1, 3]),
        rows: CompactSelection.fromSingleSelection([2, 4]),
        current: undefined,
      },
      selectedCell: [1, 2],
      sheetName: "Sheet1",
      visibleItems: [
        [1, 2],
        [2, 2],
        [1, 3],
        [2, 3],
      ],
      visibleRegion: { range: { x: 1, y: 2, width: 2, height: 2 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0 },
      getCellBounds: () => ({ x: 146, y: 48, width: 100, height: 24 }),
    });

    expect(scene.fillRects).toContainEqual({
      x: 47,
      y: 25,
      width: 102,
      height: 42,
      color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 0.06 },
    });
    expect(scene.fillRects).toContainEqual({
      x: 47,
      y: 25,
      width: 206,
      height: 20,
      color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 0.06 },
    });
  });

  test("adds GPU-backed boolean checkbox chrome", () => {
    const scene = buildGridGpuScene({
      engine: makeEngine({}, createCellSnapshot(undefined, { tag: ValueTag.Boolean, value: true })),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      gridSelection: createSelection(),
      selectedCell: [1, 1],
      sheetName: "Sheet1",
      visibleItems: [[1, 1]],
      visibleRegion: { range: { x: 1, y: 1, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0 },
      getCellBounds: () => ({ x: 100, y: 24, width: 100, height: 24 }),
    });

    expect(scene.fillRects).toContainEqual({
      x: 143,
      y: 29,
      width: 14,
      height: 14,
      color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 1 },
    });
    expect(
      scene.fillRects.filter(
        (rect) =>
          rect.color.r === 1 && rect.color.g === 1 && rect.color.b === 1 && rect.width === 2,
      ),
    ).toHaveLength(3);
  });
});
