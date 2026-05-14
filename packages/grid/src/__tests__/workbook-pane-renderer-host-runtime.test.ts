// @vitest-environment jsdom
import { afterEach, describe, expect, test, vi } from 'vitest'
import { WorkbookPaneRendererHostRuntimeV3 } from '../renderer-v3/workbook-pane-renderer-host-runtime.js'
import {
  WorkbookPaneRendererRuntimeV3,
  type WorkbookPaneFrameDrawerV3,
  type WorkbookPaneRendererRuntimeStateV3,
} from '../renderer-v3/workbook-pane-renderer-runtime.js'
import { WorkbookPaneSurfaceRuntimeV3 } from '../renderer-v3/workbook-pane-surface-runtime.js'
import { DirtyMaskV3 } from '../renderer-v3/tile-damage-index.js'

function createHost(width: number, height: number): HTMLDivElement {
  const host = document.createElement('div')
  Object.defineProperty(host, 'clientWidth', { configurable: true, value: width })
  Object.defineProperty(host, 'clientHeight', { configurable: true, value: height })
  return host
}

function installManualAnimationFrames(): { flushNextFrame: () => void; restore: () => void } {
  const callbacks = new Map<number, FrameRequestCallback>()
  let nextHandle = 1
  const requestFrame = vi.spyOn(window, 'requestAnimationFrame').mockImplementation((callback) => {
    const handle = nextHandle
    nextHandle += 1
    callbacks.set(handle, callback)
    return handle
  })
  const cancelFrame = vi.spyOn(window, 'cancelAnimationFrame').mockImplementation((handle) => {
    callbacks.delete(handle)
  })
  return {
    flushNextFrame: () => {
      const next = callbacks.entries().next()
      if (next.done) {
        throw new Error('no animation frame is scheduled')
      }
      const [handle, callback] = next.value
      callbacks.delete(handle)
      callback(performance.now())
    },
    restore: () => {
      requestFrame.mockRestore()
      cancelFrame.mockRestore()
    },
  }
}

function createDirtyTilePane(): WorkbookPaneRendererRuntimeStateV3['tilePanes'][number] {
  return {
    contentOffset: { x: 0, y: 0 },
    drawVisible: true,
    frame: { height: 360, width: 640, x: 0, y: 0 },
    generation: 1,
    paneId: 'body',
    scrollAxes: { x: true, y: true },
    surfaceSize: { height: 360, width: 640 },
    tile: {
      bounds: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      coord: {
        colTile: 0,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 0,
        sheetId: 1,
      },
      dirtyMasks: new Uint32Array([DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
      lastBatchId: 1,
      lastCameraSeq: 1,
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: 0,
      textMetrics: new Float32Array(),
      textRuns: [],
      tileId: 1,
      version: {
        axisX: 1,
        axisY: 2,
        freeze: 0,
        styles: 1,
        text: 1,
        values: 1,
      },
    },
    viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
  }
}

describe('WorkbookPaneRendererHostRuntimeV3', () => {
  afterEach(() => {
    vi.restoreAllMocks()
  })

  test('owns the surface-to-renderer handoff outside React state', async () => {
    const animationFrames = installManualAnimationFrames()
    const backend = {}
    const createBackend = vi.fn(async () => backend)
    const destroyBackend = vi.fn()
    const syncSurface = vi.fn()
    const drawFrame = vi.fn<WorkbookPaneFrameDrawerV3>()
    const runtime = new WorkbookPaneRendererHostRuntimeV3({
      rendererRuntime: new WorkbookPaneRendererRuntimeV3(drawFrame),
      surfaceRuntime: new WorkbookPaneSurfaceRuntimeV3({
        createBackend,
        createResizeObserver: () => null,
        destroyBackend,
        getDevicePixelRatio: () => 2,
        syncSurface,
      }),
    })
    const host = createHost(640, 360)
    const canvas = document.createElement('canvas')

    runtime.updateProps({
      active: true,
      cameraStore: null,
      geometry: null,
      headerPanes: [],
      host,
      overlay: null,
      overlayBuilder: null,
      preloadTilePanes: [],
      scrollTransformStore: null,
      tilePanes: [],
    })
    runtime.setCanvas(canvas)
    await Promise.resolve()
    animationFrames.flushNextFrame()

    expect(createBackend).toHaveBeenCalledWith(canvas)
    expect(syncSurface).toHaveBeenCalledWith({
      backend,
      canvas,
      size: {
        dpr: 2,
        height: 360,
        pixelHeight: 720,
        pixelWidth: 1280,
        width: 640,
      },
    })
    expect(drawFrame).toHaveBeenCalled()
    expect(drawFrame.mock.calls.at(-1)?.[0]).toMatchObject({
      backend,
      surface: {
        dpr: 2,
        height: 360,
        pixelHeight: 720,
        pixelWidth: 1280,
        width: 640,
      },
      tilePanes: [],
    })

    runtime.dispose()
    animationFrames.restore()

    expect(destroyBackend).toHaveBeenCalledWith(backend)
    expect(canvas.width).toBe(0)
    expect(canvas.height).toBe(0)
  })

  test('publishes backend status changes for the React shell without callback wiring', async () => {
    const backend = {}
    const runtime = new WorkbookPaneRendererHostRuntimeV3({
      rendererRuntime: new WorkbookPaneRendererRuntimeV3(vi.fn<WorkbookPaneFrameDrawerV3>()),
      surfaceRuntime: new WorkbookPaneSurfaceRuntimeV3({
        createBackend: vi.fn(async () => backend),
        createResizeObserver: () => null,
        destroyBackend: vi.fn(),
        syncSurface: vi.fn(),
      }),
    })
    const statuses: string[] = []
    const unsubscribe = runtime.subscribeBackendStatus(() => {
      statuses.push(runtime.getBackendStatusSnapshot())
    })

    runtime.updateProps({
      active: true,
      cameraStore: null,
      geometry: null,
      headerPanes: [],
      host: createHost(640, 360),
      overlay: null,
      overlayBuilder: null,
      preloadTilePanes: [],
      scrollTransformStore: null,
      tilePanes: [],
    })
    runtime.setCanvas(document.createElement('canvas'))
    await Promise.resolve()

    expect(statuses).toContain('initializing')
    expect(statuses).toContain('ready')

    unsubscribe()
    runtime.dispose()
  })

  test('detaches the WebGPU surface when the pane becomes inactive', async () => {
    const backend = {}
    const destroyBackend = vi.fn()
    const runtime = new WorkbookPaneRendererHostRuntimeV3({
      rendererRuntime: new WorkbookPaneRendererRuntimeV3(vi.fn<WorkbookPaneFrameDrawerV3>()),
      surfaceRuntime: new WorkbookPaneSurfaceRuntimeV3({
        createBackend: vi.fn(async () => backend),
        createResizeObserver: () => null,
        destroyBackend,
        syncSurface: vi.fn(),
      }),
    })

    runtime.updateProps({
      active: true,
      cameraStore: null,
      geometry: null,
      headerPanes: [],
      host: createHost(640, 360),
      overlay: null,
      overlayBuilder: null,
      preloadTilePanes: [],
      scrollTransformStore: null,
      tilePanes: [],
    })
    runtime.setCanvas(document.createElement('canvas'))
    await Promise.resolve()

    runtime.updateProps({
      active: false,
      cameraStore: null,
      geometry: null,
      headerPanes: [],
      host: createHost(640, 360),
      overlay: null,
      overlayBuilder: null,
      preloadTilePanes: [],
      scrollTransformStore: null,
      tilePanes: [],
    })

    expect(destroyBackend).toHaveBeenCalledWith(backend)
    runtime.dispose()
  })

  test('defers preload sync for dirty structural tile updates', () => {
    const animationFrames = installManualAnimationFrames()
    const drawFrame = vi.fn<WorkbookPaneFrameDrawerV3>()
    const runtime = new WorkbookPaneRendererRuntimeV3(drawFrame)
    const dirtyPane = createDirtyTilePane()

    runtime.updateState({
      active: true,
      backend: {},
      headerPanes: [],
      overlay: null,
      overlayBuilder: null,
      preloadTilePanes: [dirtyPane],
      scrollTransformStore: null,
      surface: {
        dpr: 1,
        height: 360,
        pixelHeight: 360,
        pixelWidth: 640,
        width: 640,
      },
      tilePanes: [dirtyPane],
      webGpuReady: true,
    })
    runtime.requestDraw()
    animationFrames.flushNextFrame()

    expect(drawFrame).toHaveBeenCalledWith(expect.objectContaining({ syncPreloadPanes: false }))
    runtime.dispose()
    animationFrames.restore()
  })
})
