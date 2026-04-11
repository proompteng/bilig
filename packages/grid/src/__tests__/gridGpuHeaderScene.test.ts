import { expect, test } from "vitest";
import { CompactSelection } from "../gridTypes.js";
import { getGridMetrics } from "../gridMetrics.js";
import { parseGpuColor } from "../gridGpuScene.js";
import { buildGridGpuHeaderScene } from "../gridGpuHeaderScene.js";

const palette = {
  gridLineColor: parseGpuColor("#e3e9f0"),
  headerFillColor: parseGpuColor("#f8f9fa"),
  headerSelectedFillColor: parseGpuColor("#e6f4ea"),
  headerHoverFillColor: parseGpuColor("#f1f3f4"),
  headerDragAnchorFillColor: parseGpuColor("#d7eadf"),
  selectionFillColor: parseGpuColor("rgba(31, 122, 67, 0.06)"),
  resizeGuideColor: parseGpuColor("#1f7a43"),
  resizeGuideGlowColor: parseGpuColor("rgba(31, 122, 67, 0.18)"),
};

function createSelection() {
  return {
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty(),
    current: undefined,
  };
}

test("builds GPU-backed header backgrounds and selection highlights", () => {
  const scene = buildGridGpuHeaderScene({
    palette,
    columnWidths: {},
    gridMetrics: getGridMetrics(),
    gridSelection: createSelection(),
    selectedCell: [2, 3],
    selectionRange: { x: 2, y: 3, width: 1, height: 1 },
    visibleItems: [[2, 3]],
    visibleRegion: { range: { x: 2, y: 3, width: 1, height: 1 }, tx: 0, ty: 0 },
    getCellBounds: () => ({ x: 254, y: 90, width: 104, height: 22 }),
    rowHeights: {},
    hoveredHeader: null,
    resizeGuideColumn: null,
    resizeGuideRow: null,
    activeHeaderDrag: null,
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

test("builds GPU resize guides for hovered columns", () => {
  const scene = buildGridGpuHeaderScene({
    palette,
    columnWidths: {},
    gridMetrics: getGridMetrics(),
    gridSelection: createSelection(),
    selectedCell: [0, 0],
    selectionRange: null,
    visibleItems: [[2, 3]],
    visibleRegion: { range: { x: 2, y: 3, width: 1, height: 2 }, tx: 0, ty: 0 },
    getCellBounds: () => ({ x: 254, y: 90, width: 104, height: 22 }),
    rowHeights: {},
    hoveredHeader: null,
    resizeGuideColumn: 2,
    resizeGuideRow: null,
    activeHeaderDrag: null,
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

test("builds GPU drag guides for active column header drags", () => {
  const scene = buildGridGpuHeaderScene({
    palette,
    columnWidths: {},
    gridMetrics: getGridMetrics(),
    gridSelection: {
      columns: CompactSelection.fromSingleSelection([1, 3]),
      rows: CompactSelection.empty(),
      current: undefined,
    },
    selectedCell: [1, 2],
    selectionRange: null,
    visibleItems: [
      [1, 2],
      [2, 2],
    ],
    visibleRegion: { range: { x: 1, y: 2, width: 2, height: 2 }, tx: 0, ty: 0 },
    getCellBounds: () => ({ x: 146, y: 48, width: 100, height: 24 }),
    rowHeights: {},
    hoveredHeader: null,
    resizeGuideColumn: null,
    resizeGuideRow: null,
    activeHeaderDrag: { kind: "column", index: 1 },
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
