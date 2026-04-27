// @vitest-environment jsdom
import { describe, expect, test, vi } from 'vitest'
import { WorkbookPaneRendererHostRuntimeV3 } from '../renderer-v3/workbook-pane-renderer-host-runtime.js'
import { WorkbookPaneRendererRuntimeV3, type WorkbookPaneFrameDrawerV3 } from '../renderer-v3/workbook-pane-renderer-runtime.js'
import { WorkbookPaneSurfaceRuntimeV3 } from '../renderer-v3/workbook-pane-surface-runtime.js'

function createHost(width: number, height: number): HTMLDivElement {
  const host = document.createElement('div')
  Object.defineProperty(host, 'clientWidth', { configurable: true, value: width })
  Object.defineProperty(host, 'clientHeight', { configurable: true, value: height })
  return host
}

describe('WorkbookPaneRendererHostRuntimeV3', () => {
  test('owns the surface-to-renderer handoff outside React state', async () => {
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

    expect(destroyBackend).toHaveBeenCalledWith(backend)
    expect(canvas.width).toBe(0)
    expect(canvas.height).toBe(0)
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
})
