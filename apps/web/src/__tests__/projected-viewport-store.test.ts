import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type RecalcMetrics } from '@bilig/protocol'
import type { ViewportPatch } from '@bilig/worker-transport'
import {
  GRID_SCENE_PACKET_V2_MAGIC,
  GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_VERSION,
  createGridTileKeyV2,
  type GridScenePacketV2,
} from '../../../../packages/grid/src/renderer-v2/scene-packet-v2.js'
import { ProjectedViewportStore } from '../projected-viewport-store.js'

const TEST_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

function createPatch(styleId?: string): ViewportPatch {
  return {
    version: 1,
    full: false,
    freezeRows: 0,
    freezeCols: 0,
    viewport: {
      sheetName: 'Sheet1',
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
          sheetName: 'Sheet1',
          address: 'D5',
          value: { tag: ValueTag.Empty },
          flags: 0,
          version: 1,
          ...(styleId ? { styleId } : {}),
        },
        displayText: '',
        copyText: '',
        editorText: '',
        formatId: 0,
        styleId: styleId ?? 'style-0',
      },
    ],
    columns: [],
    rows: [],
  }
}

function createColumnPatch(size: number, hidden = false): ViewportPatch {
  return {
    ...createPatch(),
    cells: [],
    columns: [{ index: 0, size, hidden }],
  }
}

function createRowPatch(size: number, hidden = false): ViewportPatch {
  return {
    ...createPatch(),
    cells: [],
    rows: [{ index: 0, size, hidden }],
  }
}

function columnLabel(columnIndex: number): string {
  let index = columnIndex + 1
  let label = ''
  while (index > 0) {
    const remainder = (index - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    index = Math.floor((index - 1) / 26)
  }
  return label
}

function createLargePatch(rowCount: number, columnCount: number): ViewportPatch {
  return {
    version: 1,
    full: false,
    viewport: {
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: rowCount - 1,
      colStart: 0,
      colEnd: columnCount - 1,
    },
    metrics: TEST_METRICS,
    styles: [],
    cells: Array.from({ length: rowCount * columnCount }, (_, index) => {
      const row = Math.floor(index / columnCount)
      const col = index % columnCount
      return {
        row,
        col,
        snapshot: {
          sheetName: 'Sheet1',
          address: `${columnLabel(col)}${row + 1}`,
          value: { tag: ValueTag.Number, value: index },
          flags: 0,
          version: 1,
        },
        displayText: String(index),
        copyText: String(index),
        editorText: String(index),
        formatId: 0,
        styleId: 'style-0',
      }
    }),
    columns: [],
    rows: [],
  }
}

function countSheetCells(cache: ProjectedViewportStore, sheetName: string): number {
  let count = 0
  cache.workbook.getSheet(sheetName)?.grid.forEachCellEntry(() => {
    count += 1
  })
  return count
}

function createResidentScene(
  request: {
    readonly sheetName: string
    readonly residentViewport: {
      readonly rowStart: number
      readonly rowEnd: number
      readonly colStart: number
      readonly colEnd: number
    }
  },
  generation: number,
  requestSeq = generation,
  packetOverrides: Partial<GridScenePacketV2> = {},
) {
  const viewport = request.residentViewport
  const surfaceSize = { width: 400, height: 200 }
  return {
    generation,
    paneId: 'body' as const,
    viewport,
    surfaceSize,
    gpuScene: { fillRects: [], borderRects: [] },
    textScene: { items: [] },
    packedScene: {
      borderRectCount: 0,
      cameraSeq: requestSeq,
      fillRectCount: 0,
      generatedAt: 1_000 + requestSeq,
      generation,
      key: createGridTileKeyV2({ paneId: 'body', sheetName: request.sheetName, viewport }),
      magic: GRID_SCENE_PACKET_V2_MAGIC,
      paneId: 'body' as const,
      rectCount: 0,
      rectInstances: new Float32Array(GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT),
      rects: new Float32Array(GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT),
      requestSeq,
      sheetName: request.sheetName,
      surfaceSize,
      textMetrics: new Float32Array(GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT),
      textCount: 0,
      version: GRID_SCENE_PACKET_V2_VERSION,
      viewport,
      ...packetOverrides,
    },
  }
}

describe('ProjectedViewportStore', () => {
  it('accepts equal-version empty snapshots that clear stale styling', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createPatch('style-red'))
    expect(cache.getCell('Sheet1', 'D5').styleId).toBe('style-red')

    cache.applyViewportPatch(createPatch())

    expect(cache.getCell('Sheet1', 'D5').styleId).toBeUndefined()
  })

  it('keeps an equal-version local formula snapshot when a later patch drops the formula', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: 'TRUE',
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: false },
            flags: 0,
            version: 3,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    })
  })

  it('keeps a local formula snapshot when a newer eval-only patch drops source metadata', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: 'TRUE',
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: false },
            flags: 0,
            version: 4,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    })
  })

  it('keeps a local formula snapshot when a direct cell refresh drops source metadata', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.Boolean, value: true },
      input: '=A1="HELLO"',
      formula: 'A1="HELLO"',
      flags: 0,
      version: 3,
    })

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.Boolean, value: false },
      flags: 0,
      version: 4,
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    })
  })

  it('accepts a newer literal snapshot when the source input is present', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: 'TRUE',
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: false },
            input: false,
            flags: 0,
            version: 4,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    const snapshot = cache.getCell('Sheet1', 'D5')
    expect(snapshot).toMatchObject({
      value: { tag: ValueTag.Boolean, value: false },
      input: false,
      version: 4,
    })
    expect('formula' in snapshot).toBe(false)
  })

  it('reports damage when a style record changes without a newer cell snapshot', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch('style-fill'),
      styles: [{ id: 'style-fill', fill: { backgroundColor: '#c9daf8' } }],
    })

    const damage = cache.applyViewportPatch({
      ...createPatch('style-fill'),
      styles: [{ id: 'style-fill', fill: { backgroundColor: '#a4c2f4' } }],
    })

    expect(damage).toEqual([{ cell: [3, 4] }])
    expect(cache.getCellStyle('style-fill')).toEqual({
      id: 'style-fill',
      fill: { backgroundColor: '#a4c2f4' },
    })
  })

  it('tracks freeze pane metadata from viewport patches', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      freezeRows: 2,
      freezeCols: 1,
    })

    expect(cache.getFreezeRows('Sheet1')).toBe(2)
    expect(cache.getFreezeCols('Sheet1')).toBe(1)
  })

  it('clears stale viewport cells on full patches without dropping cells outside the viewport', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'A1',
      value: { tag: ValueTag.String, value: 'pinned', stringId: 1 },
      flags: 0,
      version: 1,
    })
    cache.applyViewportPatch({ ...createPatch(), full: true })

    const damage = cache.applyViewportPatch({
      ...createPatch(),
      full: true,
      cells: [],
    })

    expect(damage).toEqual([{ cell: [3, 4] }])
    expect(cache.peekCell('Sheet1', 'D5')).toBeUndefined()
    expect(cache.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.String,
      value: 'pinned',
      stringId: 1,
    })
  })

  it('drops stale sheet cache entries when sheets disappear', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createPatch())
    expect(cache.peekCell('Sheet1', 'D5')).toBeDefined()

    cache.setKnownSheets(['Sheet2'])

    expect(cache.peekCell('Sheet1', 'D5')).toBeUndefined()
  })

  it('clears a pending local column width once the authoritative patch matches it', () => {
    const cache = new ProjectedViewportStore()

    cache.setColumnWidth('Sheet1', 0, 68)
    cache.applyViewportPatch(createColumnPatch(68))
    cache.applyViewportPatch(createColumnPatch(93))

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(93)
  })

  it('rolls back a failed local column width mutation without leaving a pending width behind', () => {
    const cache = new ProjectedViewportStore()

    cache.setColumnWidth('Sheet1', 0, 68)
    cache.rollbackColumnWidth('Sheet1', 0, undefined)
    cache.applyViewportPatch(createColumnPatch(104))

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(104)
  })

  it('clears a pending local row height once the authoritative patch matches it', () => {
    const cache = new ProjectedViewportStore()

    cache.setRowHeight('Sheet1', 0, 30)
    cache.applyViewportPatch(createRowPatch(30))
    cache.applyViewportPatch(createRowPatch(44))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(44)
  })

  it('rolls back a failed local row height mutation without leaving a pending height behind', () => {
    const cache = new ProjectedViewportStore()

    cache.setRowHeight('Sheet1', 0, 30)
    cache.rollbackRowHeight('Sheet1', 0, undefined)
    cache.applyViewportPatch(createRowPatch(22))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(22)
  })

  it('preserves hidden column metadata and collapses hidden columns from the visible axis map', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createColumnPatch(93, true))

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(0)
    expect(cache.getColumnSizes('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBe(true)

    cache.applyViewportPatch(createColumnPatch(93, false))

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBeUndefined()
  })

  it('preserves hidden row metadata and collapses hidden rows from the visible axis map', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createRowPatch(44, true))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(0)
    expect(cache.getRowSizes('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBe(true)

    cache.applyViewportPatch(createRowPatch(44, false))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBeUndefined()
  })

  it('supports optimistic column hide and rollback using preserved raw sizes', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createColumnPatch(93))
    cache.setColumnHidden('Sheet1', 0, true, 93)

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(0)
    expect(cache.getColumnSizes('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBe(true)

    cache.rollbackColumnHidden('Sheet1', 0, { hidden: false, size: 93 })

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(93)
    expect(cache.getColumnSizes('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBeUndefined()
  })

  it('supports optimistic row hide and rollback using preserved raw sizes', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createRowPatch(44))
    cache.setRowHidden('Sheet1', 0, true, 44)

    expect(cache.getRowHeights('Sheet1')[0]).toBe(0)
    expect(cache.getRowSizes('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBe(true)

    cache.rollbackRowHidden('Sheet1', 0, { hidden: false, size: 44 })

    expect(cache.getRowHeights('Sheet1')[0]).toBe(44)
    expect(cache.getRowSizes('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBeUndefined()
  })

  it('prunes back to the cache cap after the last viewport unsubscribes', () => {
    const cache = new ProjectedViewportStore({
      invoke: async () => undefined,
      ready: async () => undefined,
      subscribe: () => () => undefined,
      subscribeBatches: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      dispose: () => undefined,
    })

    const unsubscribe = cache.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 600, colStart: 0, colEnd: 9 }, () => undefined)

    cache.applyViewportPatch(createLargePatch(601, 10))

    expect(countSheetCells(cache, 'Sheet1')).toBe(6010)

    unsubscribe()

    expect(countSheetCells(cache, 'Sheet1')).toBe(6000)
  })

  it('keeps pinned cell subscriptions while pruning after viewport teardown', () => {
    const cache = new ProjectedViewportStore({
      invoke: async () => undefined,
      ready: async () => undefined,
      subscribe: () => () => undefined,
      subscribeBatches: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      dispose: () => undefined,
    })

    const unsubscribeViewport = cache.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 600, colStart: 0, colEnd: 9 }, () => undefined)
    const unsubscribeCell = cache.subscribeCells('Sheet1', ['A1'], () => undefined)

    cache.applyViewportPatch(createLargePatch(601, 10))
    unsubscribeViewport()

    expect(cache.peekCell('Sheet1', 'A1')).toBeDefined()
    expect(countSheetCells(cache, 'Sheet1')).toBe(6000)

    unsubscribeCell()
  })

  it('publishes patched resident scenes immediately when subscribed viewport patches arrive', async () => {
    vi.useFakeTimers()
    const typedPatch: ViewportPatch = {
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.String, value: 'typed value', stringId: 1 },
            input: 'typed value',
            flags: 0,
            version: 1,
          },
          displayText: 'typed value',
          copyText: 'typed value',
          editorText: 'typed value',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    }
    const encodedPatch = new TextEncoder().encode(JSON.stringify(typedPatch))
    const subscribeViewportPatches = vi.fn((_viewport, listener: (bytes: Uint8Array) => void) => {
      setTimeout(() => listener(encodedPatch), 0)
      return () => undefined
    })
    const request = {
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 0,
      freezeCols: 0,
    } as const
    const pendingRefresh = new Promise<never>(() => undefined)
    const invoke = vi
      .fn()
      .mockResolvedValueOnce([createResidentScene(request, 1, 1)])
      .mockReturnValueOnce(pendingRefresh)
    const cache = new ProjectedViewportStore({
      invoke,
      ready: async () => undefined,
      subscribe: () => () => undefined,
      subscribeBatches: () => () => undefined,
      subscribeViewportPatches,
      dispose: () => undefined,
    })
    const sceneListener = vi.fn()

    const unsubscribeScene = cache.subscribeResidentPaneScenes(request, sceneListener)
    await vi.runAllTimersAsync()
    expect(cache.peekResidentPaneScenes(request)?.[0]?.generation).toBe(1)
    sceneListener.mockClear()

    const unsubscribeViewport = cache.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 }, () => undefined)
    await vi.runOnlyPendingTimersAsync()

    expect(cache.getCell('Sheet1', 'D5').version).toBe(1)
    const immediateScene = cache.peekResidentPaneScenes(request)
    expect(immediateScene?.[0]?.generation).toBe(2)
    expect(immediateScene?.[0]?.textScene.items.some((item) => item.text === 'typed value')).toBe(true)
    expect(invoke).toHaveBeenCalledTimes(1)
    expect(sceneListener).toHaveBeenCalledTimes(1)

    unsubscribeViewport()
    unsubscribeScene()
    vi.useRealTimers()
  })

  it('publishes immediate resident scenes for local cell snapshots', async () => {
    vi.useFakeTimers()
    const request = {
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 0,
      freezeCols: 0,
    } as const
    const pendingRefresh = new Promise<never>(() => undefined)
    const invoke = vi
      .fn()
      .mockResolvedValueOnce([createResidentScene(request, 1, 1)])
      .mockReturnValueOnce(pendingRefresh)
    const cache = new ProjectedViewportStore({
      invoke,
      ready: async () => undefined,
      subscribe: () => () => undefined,
      subscribeBatches: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      dispose: () => undefined,
    })
    const sceneListener = vi.fn()

    const unsubscribeScene = cache.subscribeResidentPaneScenes(request, sceneListener)
    await vi.runAllTimersAsync()
    sceneListener.mockClear()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.String, value: 'local typed', stringId: 1 },
      input: 'local typed',
      flags: 0,
      version: 1,
    })

    const immediateScene = cache.peekResidentPaneScenes(request)
    expect(immediateScene?.[0]?.generation).toBe(2)
    expect(immediateScene?.[0]?.textScene.items.some((item) => item.text === 'local typed')).toBe(true)
    expect(invoke).toHaveBeenCalledTimes(1)
    expect(sceneListener).toHaveBeenCalledTimes(1)

    unsubscribeScene()
    vi.useRealTimers()
  })

  it('notifies selected-cell listeners only for the addressed cell', () => {
    const cache = new ProjectedViewportStore()
    const selectedCellListener = vi.fn()
    const unrelatedCellListener = vi.fn()

    const unsubscribeSelected = cache.subscribeCell('Sheet1', 'A1', selectedCellListener)
    const unsubscribeUnrelated = cache.subscribeCell('Sheet1', 'B2', unrelatedCellListener)

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'A1',
      value: { tag: ValueTag.Number, value: 42 },
      flags: 0,
      version: 1,
    })

    expect(selectedCellListener).toHaveBeenCalledTimes(1)
    expect(unrelatedCellListener).not.toHaveBeenCalled()

    unsubscribeSelected()
    unsubscribeUnrelated()
  })

  it('notifies only the subscribed sheet-axis channels for viewport axis patches', () => {
    const cache = new ProjectedViewportStore()
    const freezeListener = vi.fn()
    const columnsListener = vi.fn()
    const rowsListener = vi.fn()

    const unsubscribeFreeze = cache.subscribeSheetChannel('Sheet1', 'freeze', freezeListener)
    const unsubscribeColumns = cache.subscribeSheetChannel('Sheet1', 'columnWidths', columnsListener)
    const unsubscribeRows = cache.subscribeSheetChannel('Sheet1', 'rowHeights', rowsListener)

    cache.applyViewportPatch({
      viewport: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 2, colStart: 0, colEnd: 2 },
      full: false,
      cells: [],
      styles: [],
      columns: [{ index: 1, size: 140, hidden: false }],
      rows: [],
      freezeRows: 1,
      freezeCols: 1,
    })

    expect(freezeListener).toHaveBeenCalledTimes(1)
    expect(columnsListener).toHaveBeenCalledTimes(1)
    expect(rowsListener).not.toHaveBeenCalled()

    unsubscribeFreeze()
    unsubscribeColumns()
    unsubscribeRows()
  })
})
