// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  GRID_SCENE_PACKET_V2_MAGIC,
  GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_VERSION,
} from '../../../../packages/grid/src/renderer-v2/scene-packet-v2.js'
import { ProjectedSceneStore } from '../projected-scene-store.js'

describe('ProjectedSceneStore', () => {
  const originalRequestAnimationFrame = window.requestAnimationFrame
  const originalCancelAnimationFrame = window.cancelAnimationFrame

  beforeEach(() => {
    window.requestAnimationFrame = ((callback: FrameRequestCallback) =>
      window.setTimeout(() => callback(performance.now()), 0)) as typeof window.requestAnimationFrame
    window.cancelAnimationFrame = ((handle: number) => window.clearTimeout(handle)) as typeof window.cancelAnimationFrame
  })

  afterEach(() => {
    window.requestAnimationFrame = originalRequestAnimationFrame
    window.cancelAnimationFrame = originalCancelAnimationFrame
  })

  it('fetches resident pane scenes on first subscription and exposes them via peek', async () => {
    const request = {
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 1,
      freezeCols: 1,
      selectedCell: { col: 0, row: 0 },
      selectedCellSnapshot: null,
      selectionRange: null,
      editingCell: null,
    } as const
    const scenes = [
      {
        generation: 1,
        paneId: 'body',
        viewport: request.residentViewport,
        surfaceSize: { width: 400, height: 200 },
        gpuScene: { fillRects: [], borderRects: [] },
        textScene: { items: [] },
        packedScene: {
          borderRectCount: 0,
          fillRectCount: 0,
          generation: 1,
          magic: GRID_SCENE_PACKET_V2_MAGIC,
          paneId: 'body',
          rectCount: 0,
          rectInstances: new Float32Array(GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT),
          rects: new Float32Array(GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT),
          sheetName: request.sheetName,
          surfaceSize: { width: 400, height: 200 },
          textMetrics: new Float32Array(GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT),
          textCount: 0,
          version: GRID_SCENE_PACKET_V2_VERSION,
          viewport: request.residentViewport,
        },
      },
    ] as const
    const client = {
      invoke: vi.fn(async (_method: string, _request: unknown) => scenes),
    }
    const store = new ProjectedSceneStore(client)
    const listener = vi.fn()

    const unsubscribe = store.subscribeResidentPaneScenes(request, listener)
    await new Promise((resolve) => window.setTimeout(resolve, 10))

    expect(client.invoke).toHaveBeenCalledWith('getResidentPaneScenes', request)
    expect(listener).toHaveBeenCalledTimes(1)
    expect(store.peekResidentPaneScenes(request)).toEqual(scenes)

    unsubscribe()
  })

  it('coalesces intersecting viewport patches into a refresh', async () => {
    const request = {
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 1,
      freezeCols: 1,
      selectedCell: { col: 0, row: 0 },
      selectedCellSnapshot: null,
      selectionRange: null,
      editingCell: null,
    } as const
    const client = {
      invoke: vi
        .fn()
        .mockResolvedValueOnce([
          {
            generation: 1,
            paneId: 'body',
            viewport: request.residentViewport,
            surfaceSize: { width: 400, height: 200 },
            gpuScene: { fillRects: [], borderRects: [] },
            textScene: { items: [] },
          },
        ])
        .mockResolvedValueOnce([
          {
            generation: 2,
            paneId: 'body',
            viewport: request.residentViewport,
            surfaceSize: { width: 400, height: 200 },
            gpuScene: { fillRects: [], borderRects: [] },
            textScene: { items: [] },
          },
        ]),
    }
    const store = new ProjectedSceneStore(client)

    const unsubscribe = store.subscribeResidentPaneScenes(request, () => undefined)
    await new Promise((resolve) => window.setTimeout(resolve, 10))

    store.noteViewportPatch({
      version: 2,
      full: true,
      viewport: { sheetName: 'Sheet1', rowStart: 5, rowEnd: 6, colStart: 5, colEnd: 6 },
      metrics: {
        batchId: 0,
        changedInputCount: 0,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      },
      styles: [],
      cells: [],
      columns: [],
      rows: [],
    })
    store.noteViewportPatch({
      version: 3,
      full: true,
      viewport: { sheetName: 'Sheet1', rowStart: 7, rowEnd: 7, colStart: 7, colEnd: 7 },
      metrics: {
        batchId: 0,
        changedInputCount: 0,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      },
      styles: [],
      cells: [],
      columns: [],
      rows: [],
    })

    await new Promise((resolve) => window.setTimeout(resolve, 10))

    expect(client.invoke).toHaveBeenCalledTimes(2)

    unsubscribe()
  })
})
