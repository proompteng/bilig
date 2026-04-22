// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest'
import {
  GRID_SCENE_PACKET_V2_MAGIC,
  GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_VERSION,
  createGridTileKeyV2,
  type GridScenePacketV2,
} from '../../../../packages/grid/src/renderer-v2/scene-packet-v2.js'
import { ProjectedSceneStore } from '../projected-scene-store.js'

interface TestSceneRequest {
  readonly sheetName: string
  readonly residentViewport: {
    readonly rowStart: number
    readonly rowEnd: number
    readonly colStart: number
    readonly colEnd: number
  }
}

function createResidentScene(
  request: TestSceneRequest,
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

describe('ProjectedSceneStore', () => {
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
    const scenes = [createResidentScene(request, 1)] as const
    const client = {
      invoke: vi.fn(async (_method: string, _request: unknown) => scenes),
    }
    const store = new ProjectedSceneStore(client)
    const listener = vi.fn()

    const unsubscribe = store.subscribeResidentPaneScenes(request, listener)
    await new Promise((resolve) => window.setTimeout(resolve, 10))

    expect(client.invoke).toHaveBeenCalledWith(
      'getResidentPaneScenes',
      expect.objectContaining({
        ...request,
        cameraSeq: 1,
        priority: 0,
        reason: 'visible',
        requestSeq: 1,
      }),
    )
    expect(listener).toHaveBeenCalledTimes(1)
    expect(store.peekResidentPaneScenes(request)).toEqual(scenes)

    unsubscribe()
  })

  it('tags prefetch scene requests with lower worker priority while keeping semantic cache lookup stable', async () => {
    const request = {
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 0,
      freezeCols: 0,
      cameraSeq: 7,
      priority: 4,
      reason: 'prefetch' as const,
      selectedCell: { col: 0, row: 0 },
      selectedCellSnapshot: null,
      selectionRange: null,
      editingCell: null,
    }
    const scenes = [createResidentScene(request, 1, 1, { cameraSeq: 7 })] as const
    const client = {
      invoke: vi.fn(async (_method: string, _request: unknown) => scenes),
    }
    const store = new ProjectedSceneStore(client)

    const unsubscribe = store.subscribeResidentPaneScenes(request, () => undefined)
    await new Promise((resolve) => window.setTimeout(resolve, 10))

    expect(client.invoke).toHaveBeenCalledWith(
      'getResidentPaneScenes',
      expect.objectContaining({
        cameraSeq: 7,
        priority: 4,
        reason: 'prefetch',
        requestSeq: 1,
      }),
    )
    expect(store.peekResidentPaneScenes({ ...request, cameraSeq: 99, priority: 0, reason: 'visible' })).toEqual(scenes)

    unsubscribe()
  })

  it('rejects stale or packetless worker scene responses before publishing', async () => {
    const request = {
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 0,
      freezeCols: 0,
      selectedCell: { col: 0, row: 0 },
      selectedCellSnapshot: null,
      selectionRange: null,
      editingCell: null,
    } as const
    const staleClient = {
      invoke: vi.fn(async () => [createResidentScene(request, 1, 0)]),
    }
    const staleStore = new ProjectedSceneStore(staleClient)
    const staleListener = vi.fn()

    const unsubscribeStale = staleStore.subscribeResidentPaneScenes(request, staleListener)
    await new Promise((resolve) => window.setTimeout(resolve, 10))

    expect(staleListener).not.toHaveBeenCalled()
    expect(staleStore.peekResidentPaneScenes(request)).toBeNull()
    expect(staleStore.getLastRejectedPacketReason()).toBe('request-sequence-stale')

    const packetlessClient = {
      invoke: vi.fn(async () => [{ ...createResidentScene(request, 1, 1), packedScene: undefined }]),
    }
    const packetlessStore = new ProjectedSceneStore(packetlessClient)
    const packetlessListener = vi.fn()

    const unsubscribePacketless = packetlessStore.subscribeResidentPaneScenes(request, packetlessListener)
    await new Promise((resolve) => window.setTimeout(resolve, 10))

    expect(packetlessListener).not.toHaveBeenCalled()
    expect(packetlessStore.peekResidentPaneScenes(request)).toBeNull()
    expect(packetlessStore.getLastRejectedPacketReason()).toBe('missing-packed-scene')

    unsubscribePacketless()
    unsubscribeStale()
  })

  it('lets a visible request supersede an older prefetch request for the same tile before dispatch', async () => {
    const prefetchRequest = {
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 0,
      freezeCols: 0,
      cameraSeq: 3,
      priority: 4,
      reason: 'prefetch' as const,
      selectedCell: { col: 0, row: 0 },
      selectedCellSnapshot: null,
      selectionRange: null,
      editingCell: null,
    }
    const visibleRequest = {
      ...prefetchRequest,
      cameraSeq: 4,
      priority: 0,
      reason: 'visible' as const,
    }
    const scenes = [createResidentScene(visibleRequest, 1, 1, { cameraSeq: 4 })] as const
    const client = {
      invoke: vi.fn(async (_method: string, _request: unknown) => scenes),
    }
    const store = new ProjectedSceneStore(client)

    const cleanupPrefetch = store.subscribeResidentPaneScenes(prefetchRequest, () => undefined)
    const cleanupVisible = store.subscribeResidentPaneScenes(visibleRequest, () => undefined)
    await new Promise((resolve) => window.setTimeout(resolve, 10))

    expect(client.invoke).toHaveBeenCalledTimes(1)
    expect(client.invoke).toHaveBeenCalledWith(
      'getResidentPaneScenes',
      expect.objectContaining({
        cameraSeq: 4,
        priority: 0,
        reason: 'visible',
        requestSeq: 1,
      }),
    )

    cleanupVisible()
    cleanupPrefetch()
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
        .mockResolvedValueOnce([createResidentScene(request, 1, 1)])
        .mockResolvedValueOnce([createResidentScene(request, 2, 2)]),
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

  it('does not publish a stale in-flight scene response after a newer visible refresh is queued', async () => {
    const request = {
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 0,
      freezeCols: 0,
      selectedCell: { col: 0, row: 0 },
      selectedCellSnapshot: null,
      selectionRange: null,
      editingCell: null,
    } as const
    let resolveFirst: ((value: readonly unknown[]) => void) | null = null
    const first = new Promise<readonly unknown[]>((resolve) => {
      resolveFirst = resolve
    })
    const second = [createResidentScene(request, 2, 2)] as const
    const client = {
      invoke: vi.fn().mockReturnValueOnce(first).mockResolvedValueOnce(second),
    }
    const listener = vi.fn()
    const store = new ProjectedSceneStore(client)

    const unsubscribe = store.subscribeResidentPaneScenes(request, listener)
    await new Promise((resolve) => window.setTimeout(resolve, 10))

    store.noteViewportPatch({
      version: 2,
      full: true,
      viewport: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
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
    resolveFirst?.([createResidentScene(request, 1, 1)])
    await new Promise((resolve) => window.setTimeout(resolve, 10))

    expect(client.invoke).toHaveBeenCalledTimes(2)
    expect(listener).toHaveBeenCalledTimes(1)
    expect(store.peekResidentPaneScenes(request)?.[0]?.generation).toBe(2)

    unsubscribe()
  })

  it('refreshes resident pane scenes for style-only patches that intersect the request', async () => {
    const request = {
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 0,
      freezeCols: 0,
      selectedCell: { col: 0, row: 0 },
      selectedCellSnapshot: null,
      selectionRange: null,
      editingCell: null,
    } as const
    const client = {
      invoke: vi
        .fn()
        .mockResolvedValueOnce([createResidentScene(request, 1, 1)])
        .mockResolvedValueOnce([createResidentScene(request, 2, 2)]),
    }
    const store = new ProjectedSceneStore(client)

    const unsubscribe = store.subscribeResidentPaneScenes(request, () => undefined)
    await new Promise((resolve) => window.setTimeout(resolve, 10))

    store.noteViewportPatch({
      version: 2,
      full: false,
      viewport: { sheetName: 'Sheet1', rowStart: 4, rowEnd: 4, colStart: 3, colEnd: 3 },
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
      styles: [{ id: 'style-fill', fill: { backgroundColor: '#a4c2f4' } }],
      cells: [],
      columns: [],
      rows: [],
    })

    await new Promise((resolve) => window.setTimeout(resolve, 10))

    expect(client.invoke).toHaveBeenCalledTimes(2)

    unsubscribe()
  })

  it('ignores style-only patches outside the resident pane request', async () => {
    const request = {
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 0,
      freezeCols: 0,
      selectedCell: { col: 0, row: 0 },
      selectedCellSnapshot: null,
      selectionRange: null,
      editingCell: null,
    } as const
    const client = {
      invoke: vi.fn().mockResolvedValue([createResidentScene(request, 1, 1)]),
    }
    const store = new ProjectedSceneStore(client)

    const unsubscribe = store.subscribeResidentPaneScenes(request, () => undefined)
    await new Promise((resolve) => window.setTimeout(resolve, 10))

    store.noteViewportPatch({
      version: 2,
      full: false,
      viewport: { sheetName: 'Sheet1', rowStart: 50, rowEnd: 50, colStart: 50, colEnd: 50 },
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
      styles: [{ id: 'style-fill', fill: { backgroundColor: '#a4c2f4' } }],
      cells: [],
      columns: [],
      rows: [],
    })

    await new Promise((resolve) => window.setTimeout(resolve, 10))

    expect(client.invoke).toHaveBeenCalledTimes(1)

    unsubscribe()
  })
})
