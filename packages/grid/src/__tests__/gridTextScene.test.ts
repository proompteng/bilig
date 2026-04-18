import { describe, expect, test } from 'vitest'
import { MAX_COLS, MAX_ROWS, ValueTag, type CellStyleRecord } from '@bilig/protocol'
import { buildGridTextScene } from '../gridTextScene.js'
import { getResolvedCellFontFamily } from '../gridCells.js'
import type { GridEngineLike } from '../grid-engine.js'
import { getGridMetrics } from '../gridMetrics.js'

type TestCellValue =
  | { tag: ValueTag.Empty }
  | { tag: ValueTag.Number; value: number }
  | { tag: ValueTag.Boolean; value: boolean }
  | { tag: ValueTag.String; value: string; stringId?: number }
  | { tag: ValueTag.Error; code: number }

function createCellSnapshot(value: TestCellValue, styleId: string | undefined = 'style-1') {
  return {
    sheetName: 'Sheet1',
    address: 'A1',
    input: '',
    value,
    flags: 0,
    version: 0,
    ...(styleId ? { styleId } : {}),
  }
}

type TestCellSnapshot = ReturnType<typeof createCellSnapshot>
const CELL_FONT_FAMILY = getResolvedCellFontFamily()
const HEADER_FONT = `500 11px ${CELL_FONT_FAMILY}`

function makeEngine(
  styles: Record<string, CellStyleRecord>,
  snapshots: TestCellSnapshot | Record<string, TestCellSnapshot> = createCellSnapshot({
    tag: ValueTag.String,
    value: 'hello',
  }),
): GridEngineLike {
  return {
    getCell: (_sheetName, address) =>
      'address' in snapshots ? snapshots : (snapshots[address] ?? createCellSnapshot({ tag: ValueTag.Empty }, undefined)),
    getCellStyle: (styleId) => (styleId ? styles[styleId] : undefined),
    subscribeCells: () => () => {},
    workbook: {
      getSheet: () => undefined,
    },
  }
}

describe('gridTextScene', () => {
  test('builds cell text items with resolved alignment and style', () => {
    const engine = makeEngine({
      'style-1': {
        alignment: { horizontal: 'right' },
        font: { bold: true, color: '#ff0000', italic: true, size: 14 },
      },
    })

    const scene = buildGridTextScene({
      engine,
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [0, 0],
      sheetName: 'Sheet1',
      visibleItems: [[0, 0]],
      visibleRegion: { range: { x: 0, y: 0, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 100, top: 200, width: 300, height: 200 },
      getCellBounds: () => ({ x: 110, y: 220, width: 90, height: 22 }),
    })

    expect(scene.items).toContainEqual({
      x: 10,
      y: 20,
      width: 90,
      height: 22,
      clipInsetTop: 4,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 36,
      text: 'hello',
      align: 'right',
      wrap: false,
      color: '#ff0000',
      font: `italic 700 14px ${CELL_FONT_FAMILY}`,
      fontSize: 14,
      underline: false,
      strike: false,
    })
  })

  test('adds column headers and row markers with selected header emphasis', () => {
    const scene = buildGridTextScene({
      engine: makeEngine({}, createCellSnapshot({ tag: ValueTag.Empty })),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [2, 3],
      selectionRange: { x: 2, y: 3, width: 1, height: 1 },
      sheetName: 'Sheet1',
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: 254, y: 90, width: 104, height: 22 }),
    })

    expect(scene.items).toContainEqual({
      x: 46,
      y: 0,
      width: 104,
      height: 24,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'C',
      align: 'center',
      wrap: false,
      color: '#1f7a43',
      font: HEADER_FONT,
      fontSize: 11,
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: 0,
      y: 24,
      width: 46,
      height: 22,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: '4',
      align: 'right',
      wrap: false,
      color: '#1f7a43',
      font: HEADER_FONT,
      fontSize: 11,
      underline: false,
      strike: false,
    })
  })

  test('adds hovered and active-drag header text emphasis', () => {
    const scene = buildGridTextScene({
      engine: makeEngine({}, createCellSnapshot({ tag: ValueTag.Empty })),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      activeHeaderDrag: { kind: 'column', index: 2 },
      hoveredHeader: { kind: 'row', index: 4 },
      selectedCell: [2, 3],
      selectionRange: { x: 2, y: 3, width: 1, height: 1 },
      sheetName: 'Sheet1',
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 2 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: 254, y: 90, width: 104, height: 22 }),
    })

    expect(scene.items).toContainEqual({
      x: 46,
      y: 0,
      width: 104,
      height: 24,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'C',
      align: 'center',
      wrap: false,
      color: '#176239',
      font: HEADER_FONT,
      fontSize: 11,
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: 0,
      y: 24,
      width: 46,
      height: 22,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: '4',
      align: 'right',
      wrap: false,
      color: '#1f7a43',
      font: HEADER_FONT,
      fontSize: 11,
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: 0,
      y: 46,
      width: 46,
      height: 22,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: '5',
      align: 'right',
      wrap: false,
      color: '#3c4043',
      font: HEADER_FONT,
      fontSize: 11,
      underline: false,
      strike: false,
    })
  })

  test('renders frozen header labels beside the scrollable pane', () => {
    const scene = buildGridTextScene({
      engine: makeEngine({}, createCellSnapshot({ tag: ValueTag.Empty })),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [0, 0],
      selectionRange: { x: 0, y: 0, width: 1, height: 1 },
      sheetName: 'Sheet1',
      visibleItems: [
        [0, 0],
        [2, 0],
        [0, 3],
        [2, 3],
      ],
      visibleRegion: {
        range: { x: 2, y: 3, width: 1, height: 1 },
        tx: 0,
        ty: 0,
        freezeRows: 1,
        freezeCols: 1,
      },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: (col, row) => ({
        x: col === 0 ? 46 : 150,
        y: row === 0 ? 24 : 46,
        width: 104,
        height: 22,
      }),
    })

    expect(scene.items).toContainEqual({
      x: 46,
      y: 0,
      width: 104,
      height: 24,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'A',
      align: 'center',
      wrap: false,
      color: '#1f7a43',
      font: HEADER_FONT,
      fontSize: 11,
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: 150,
      y: 0,
      width: 104,
      height: 24,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'C',
      align: 'center',
      wrap: false,
      color: '#5f6368',
      font: HEADER_FONT,
      fontSize: 11,
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: 0,
      y: 24,
      width: 46,
      height: 22,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: '1',
      align: 'right',
      wrap: false,
      color: '#1f7a43',
      font: HEADER_FONT,
      fontSize: 11,
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: 0,
      y: 46,
      width: 46,
      height: 22,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: '4',
      align: 'right',
      wrap: false,
      color: '#5f6368',
      font: HEADER_FONT,
      fontSize: 11,
      underline: false,
      strike: false,
    })
  })

  test('adds resize-guide emphasis to the active column header label', () => {
    const scene = buildGridTextScene({
      engine: makeEngine({}, createCellSnapshot({ tag: ValueTag.Empty })),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      resizeGuideColumn: 2,
      selectedCell: [0, 0],
      sheetName: 'Sheet1',
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: 254, y: 90, width: 104, height: 22 }),
    })

    expect(scene.items).toContainEqual({
      x: 46,
      y: 0,
      width: 104,
      height: 24,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'C',
      align: 'center',
      wrap: false,
      color: '#176239',
      font: HEADER_FONT,
      fontSize: 11,
      underline: false,
      strike: false,
    })
  })

  test('omits the active editing cell text item while preserving headers', () => {
    const scene = buildGridTextScene({
      engine: makeEngine({}, createCellSnapshot({ tag: ValueTag.String, value: 'editing' })),
      columnWidths: {},
      editingCell: [2, 3],
      gridMetrics: getGridMetrics(),
      selectedCell: [2, 3],
      sheetName: 'Sheet1',
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: 254, y: 90, width: 104, height: 22 }),
    })

    expect(scene.items.some((item) => item.text === 'editing')).toBe(false)
    expect(scene.items).toContainEqual({
      x: 46,
      y: 0,
      width: 104,
      height: 24,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'C',
      align: 'center',
      wrap: false,
      color: '#1f7a43',
      font: HEADER_FONT,
      fontSize: 11,
      underline: false,
      strike: false,
    })
  })

  test('renders the selected cell from the authoritative snapshot when the engine cache lags', () => {
    const scene = buildGridTextScene({
      engine: makeEngine(
        {},
        {
          C5: createCellSnapshot({ tag: ValueTag.Empty }, undefined),
        },
      ),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [2, 4],
      selectedCellSnapshot: {
        ...createCellSnapshot({ tag: ValueTag.String, value: 'selected' }),
        address: 'C5',
      },
      sheetName: 'Sheet1',
      visibleItems: [[2, 4]],
      visibleRegion: { range: { x: 2, y: 4, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: 254, y: 112, width: 104, height: 22 }),
    })

    expect(scene.items.find((item) => item.text === 'selected')).toEqual({
      x: 254,
      y: 112,
      width: 104,
      height: 22,
      clipInsetTop: 0,
      clipInsetRight: 38,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'selected',
      align: 'left',
      wrap: false,
      color: '#202124',
      font: `400 13px ${CELL_FONT_FAMILY}`,
      fontSize: 13,
      underline: false,
      strike: false,
    })
  })

  test('falls back to the engine cell when the selected snapshot address does not match', () => {
    const scene = buildGridTextScene({
      engine: makeEngine(
        {},
        {
          C5: createCellSnapshot({ tag: ValueTag.String, value: 'engine text' }),
        },
      ),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [2, 4],
      selectedCellSnapshot: {
        ...createCellSnapshot({ tag: ValueTag.String, value: 'stale text' }),
        address: 'B4',
      },
      sheetName: 'Sheet1',
      visibleItems: [[2, 4]],
      visibleRegion: { range: { x: 2, y: 4, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: 254, y: 112, width: 104, height: 22 }),
    })

    expect(scene.items.at(-1)?.text).toBe('engine text')
  })

  test('keeps the engine text when the selected snapshot is temporarily empty', () => {
    const scene = buildGridTextScene({
      engine: makeEngine(
        {},
        {
          C5: createCellSnapshot({ tag: ValueTag.String, value: 'engine text' }),
        },
      ),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [2, 4],
      selectedCellSnapshot: {
        ...createCellSnapshot({ tag: ValueTag.Empty }, undefined),
        address: 'C5',
      },
      sheetName: 'Sheet1',
      visibleItems: [[2, 4]],
      visibleRegion: { range: { x: 2, y: 4, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: 254, y: 112, width: 104, height: 22 }),
    })

    expect(scene.items.at(-1)?.text).toBe('engine text')
  })

  test('scales header text with the selected font size for full-sheet selections', () => {
    const scene = buildGridTextScene({
      engine: makeEngine(
        {
          'style-selected': {
            id: 'style-selected',
            font: { size: 20 },
          },
        },
        {
          A1: {
            ...createCellSnapshot({ tag: ValueTag.Number, value: 1200 }, 'style-selected'),
            address: 'A1',
          },
        },
      ),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [0, 0],
      selectedCellSnapshot: {
        ...createCellSnapshot({ tag: ValueTag.Number, value: 1200 }, 'style-selected'),
        address: 'A1',
      },
      selectionRange: { x: 0, y: 0, width: MAX_COLS, height: MAX_ROWS },
      sheetName: 'Sheet1',
      visibleItems: [[0, 0]],
      visibleRegion: { range: { x: 0, y: 0, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: 46, y: 24, width: 104, height: 22 }),
    })

    expect(scene.items).toContainEqual({
      x: 46,
      y: 0,
      width: 104,
      height: 24,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'A',
      align: 'center',
      wrap: false,
      color: '#1f7a43',
      font: `500 20px ${CELL_FONT_FAMILY}`,
      fontSize: 20,
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: 0,
      y: 24,
      width: 46,
      height: 22,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: '1',
      align: 'right',
      wrap: false,
      color: '#1f7a43',
      font: `500 20px ${CELL_FONT_FAMILY}`,
      fontSize: 20,
      underline: false,
      strike: false,
    })
  })

  test('spills left-aligned string text across contiguous empty cells', () => {
    const scene = buildGridTextScene({
      engine: makeEngine(
        {},
        {
          B12: createCellSnapshot({ tag: ValueTag.String, value: 'spill text' }),
          C12: createCellSnapshot({ tag: ValueTag.Empty }, undefined),
          D12: createCellSnapshot({ tag: ValueTag.Empty }, undefined),
        },
      ),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [0, 0],
      sheetName: 'Sheet1',
      visibleItems: [
        [1, 11],
        [2, 11],
        [3, 11],
      ],
      visibleRegion: { range: { x: 1, y: 11, width: 3, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: (col) => ({
        x: 46 + col * 104,
        y: 266,
        width: 104,
        height: 22,
      }),
    })

    expect(scene.items).toContainEqual({
      x: 150,
      y: 266,
      width: 312,
      height: 22,
      clipInsetTop: 0,
      clipInsetRight: 142,
      clipInsetBottom: 48,
      clipInsetLeft: 0,
      text: 'spill text',
      align: 'left',
      wrap: false,
      color: '#202124',
      font: `400 13px ${CELL_FONT_FAMILY}`,
      fontSize: 13,
      underline: false,
      strike: false,
    })
  })
})
