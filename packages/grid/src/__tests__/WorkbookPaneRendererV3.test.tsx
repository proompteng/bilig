// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from '../gridGeometry.js'
import { getGridMetrics } from '../gridMetrics.js'
import { GridCameraStore } from '../runtime/gridCameraStore.js'
import {
  TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS,
  WorkbookPaneRendererV3,
  resolveTypeGpuV3DrawScrollSnapshot,
  shouldDeferTypeGpuV3PreloadSync,
} from '../renderer-v3/WorkbookPaneRendererV3.js'
import { GridDrawSchedulerV3 } from '../renderer-v3/draw-scheduler.js'
import { GridRenderLoop } from '../renderer-v3/gridRenderLoop.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import { WorkbookPaneRendererRuntimeV3, type WorkbookPaneFrameDrawerV3 } from '../renderer-v3/workbook-pane-renderer-runtime.js'
import { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'

function createTilePane(rowStart = 0): WorkbookRenderTilePaneState {
  const tile: GridRenderTile = {
    bounds: { colEnd: 127, colStart: 0, rowEnd: rowStart + 31, rowStart },
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: Math.floor(rowStart / 32),
      sheetId: 7,
    },
    lastBatchId: 1,
    lastCameraSeq: 1,
    rectCount: 0,
    rectInstances: new Float32Array(20),
    textCount: 0,
    textMetrics: new Float32Array(8),
    textRuns: [],
    tileId: 101 + rowStart,
    version: {
      axisX: 1,
      axisY: 1,
      freeze: 0,
      styles: 1,
      text: 1,
      values: 1,
    },
  }
  return {
    contentOffset: { x: 0, y: 0 },
    frame: { x: 46, y: 24, width: 594, height: 336 },
    generation: 1,
    paneId: 'body',
    scrollAxes: { x: true, y: true },
    surfaceSize: { width: 640, height: 360 },
    tile,
    viewport: tile.bounds,
  }
}

describe('WorkbookPaneRendererV3', () => {
  const originalResizeObserver = globalThis.ResizeObserver

  beforeEach(() => {
    class TestResizeObserver {
      constructor(private readonly listener: ResizeObserverCallback) {}

      observe() {
        Reflect.apply(this.listener, undefined, [[], undefined])
      }

      disconnect() {}

      unobserve() {}
    }

    Object.defineProperty(globalThis, 'ResizeObserver', {
      configurable: true,
      value: TestResizeObserver,
      writable: true,
    })
  })

  afterEach(() => {
    Object.defineProperty(globalThis, 'ResizeObserver', {
      configurable: true,
      value: originalResizeObserver,
      writable: true,
    })
    document.body.innerHTML = ''
  })

  test('mounts the V3 TypeGPU canvas for the workbook surface path', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const host = document.createElement('div')
    Object.defineProperty(host, 'clientWidth', { configurable: true, value: 640 })
    Object.defineProperty(host, 'clientHeight', { configurable: true, value: 360 })
    const root = createRoot(host)
    const rendererHost = document.createElement('div')
    Object.defineProperty(rendererHost, 'clientWidth', { configurable: true, value: 640 })
    Object.defineProperty(rendererHost, 'clientHeight', { configurable: true, value: 360 })
    host.appendChild(rendererHost)

    await act(async () => {
      root.render(<WorkbookPaneRendererV3 active host={rendererHost} geometry={null} tilePanes={[createTilePane()]} />)
    })

    const canvas = host.querySelector('[data-testid="grid-pane-renderer"]')
    const fallbackCanvas = host.querySelector('[data-testid="grid-pane-renderer-fallback"]')
    expect(canvas).toBeInstanceOf(HTMLCanvasElement)
    expect(fallbackCanvas).toBeInstanceOf(HTMLCanvasElement)
    expect(fallbackCanvas?.getAttribute('data-renderer-mode')).toBe('canvas2d-v3-fallback')
    expect(canvas?.getAttribute('data-pane-renderer')).toBe('workbook-pane-renderer-v3')
    expect(canvas?.getAttribute('data-renderer-mode')).toBe('typegpu-v3')
    expect(canvas?.getAttribute('data-v3-tile-pane-count')).toBe('1')
    expect(canvas?.getAttribute('data-v3-header-pane-count')).toBe('0')
    if (!(canvas instanceof HTMLCanvasElement)) {
      throw new Error('expected V3 renderer canvas to mount')
    }
    expect(canvas.style.width).toBe('100%')
    expect(canvas.style.height).toBe('100%')
    expect(canvas.style.contain).toBe('strict')

    await act(async () => {
      root.unmount()
    })
  })

  test('derives draw scroll from live scroll and the current V3 body tile viewport', () => {
    const metrics = getGridMetrics()
    const geometry = createGridGeometrySnapshotFromAxes({
      columns: createGridAxisWorldIndex({ axisLength: 1024, defaultSize: metrics.columnWidth }),
      dpr: 1,
      gridMetrics: metrics,
      hostHeight: 720,
      hostWidth: 1280,
      rows: createGridAxisWorldIndex({ axisLength: 1024, defaultSize: metrics.rowHeight }),
      scrollLeft: 0,
      scrollTop: 0,
      sheetName: 'Sheet1',
    })

    expect(
      resolveTypeGpuV3DrawScrollSnapshot({
        fallback: {
          scrollLeft: 64 * metrics.columnWidth,
          scrollTop: 32 * metrics.rowHeight,
          tx: 0,
          ty: 0,
        },
        geometry,
        panes: [createTilePane(32)],
      }),
    ).toMatchObject({
      renderTx: 64 * metrics.columnWidth,
      renderTy: 0,
    })
  })

  test('defers V3 preload resource sync only while scroll input or camera velocity is fresh', () => {
    expect(
      shouldDeferTypeGpuV3PreloadSync({
        camera: null,
        lastScrollSignalAt: 1_000,
        now: 1_000 + TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS - 1,
      }),
    ).toBe(true)

    expect(
      shouldDeferTypeGpuV3PreloadSync({
        camera: {
          updatedAt: 2_000,
          velocityX: 0,
          velocityY: 0,
        },
        lastScrollSignalAt: 1_000,
        now: 2_000 + TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS + 1,
      }),
    ).toBe(false)
  })

  test('draw runtime uses live camera and scroll stores outside React render state', () => {
    const metrics = getGridMetrics()
    const geometry = createGridGeometrySnapshotFromAxes({
      columns: createGridAxisWorldIndex({ axisLength: 1024, defaultSize: metrics.columnWidth }),
      dpr: 1,
      gridMetrics: metrics,
      hostHeight: 720,
      hostWidth: 1280,
      rows: createGridAxisWorldIndex({ axisLength: 1024, defaultSize: metrics.rowHeight }),
      scrollLeft: 0,
      scrollTop: 0,
      sheetName: 'Sheet1',
    })
    const cameraStore = new GridCameraStore()
    cameraStore.setSnapshot(geometry)
    const scrollStore = new WorkbookGridScrollStore()
    scrollStore.setSnapshot({
      scrollLeft: 64 * metrics.columnWidth,
      scrollTop: 32 * metrics.rowHeight,
      tx: 0,
      ty: 0,
    })
    const drawFrame = vi.fn<WorkbookPaneFrameDrawerV3>()
    const runtime = new WorkbookPaneRendererRuntimeV3(drawFrame)

    runtime.updateState({
      active: true,
      backend: {},
      cameraStore,
      geometry: null,
      scrollTransformStore: scrollStore,
      surface: { dpr: 1, height: 720, pixelHeight: 720, pixelWidth: 1280, width: 1280 },
      tilePanes: [createTilePane(32)],
      webGpuReady: true,
    })
    runtime.drawNow()

    expect(drawFrame).toHaveBeenCalledTimes(1)
    expect(drawFrame.mock.calls[0]?.[0].scrollSnapshot).toMatchObject({
      renderTx: 64 * metrics.columnWidth,
      renderTy: 0,
    })

    runtime.dispose()
  })

  test('draw runtime owns camera and scroll store subscriptions', () => {
    const metrics = getGridMetrics()
    const geometry = createGridGeometrySnapshotFromAxes({
      columns: createGridAxisWorldIndex({ axisLength: 1024, defaultSize: metrics.columnWidth }),
      dpr: 1,
      gridMetrics: metrics,
      hostHeight: 720,
      hostWidth: 1280,
      rows: createGridAxisWorldIndex({ axisLength: 1024, defaultSize: metrics.rowHeight }),
      scrollLeft: 0,
      scrollTop: 0,
      sheetName: 'Sheet1',
    })
    const cameraStore = new GridCameraStore()
    cameraStore.setSnapshot(geometry)
    const scrollStore = new WorkbookGridScrollStore()
    const frameCallbacks: FrameRequestCallback[] = []
    const requestFrame = vi.fn((callback: FrameRequestCallback) => {
      frameCallbacks.push(callback)
      return frameCallbacks.length
    })
    const scheduler = new GridDrawSchedulerV3(
      () => 1,
      () => undefined,
      () => 1_000,
      new GridRenderLoop(requestFrame, () => undefined),
    )
    const drawFrame = vi.fn<WorkbookPaneFrameDrawerV3>()
    const runtime = new WorkbookPaneRendererRuntimeV3(drawFrame, scheduler)

    runtime.updateState({
      active: true,
      backend: {},
      cameraStore,
      geometry: null,
      scrollTransformStore: scrollStore,
      surface: { dpr: 1, height: 720, pixelHeight: 720, pixelWidth: 1280, width: 1280 },
      tilePanes: [createTilePane(32)],
      webGpuReady: true,
    })

    scrollStore.setSnapshot({
      scrollLeft: 64 * metrics.columnWidth,
      scrollTop: 32 * metrics.rowHeight,
      tx: 0,
      ty: 0,
    })
    expect(requestFrame).toHaveBeenCalledTimes(1)
    frameCallbacks.shift()?.(1_000)
    expect(drawFrame).toHaveBeenCalledTimes(1)
    expect(drawFrame.mock.calls[0]?.[0].scrollSnapshot).toMatchObject({
      renderTx: 64 * metrics.columnWidth,
      renderTy: 0,
    })

    runtime.updateState({ active: false })
    scrollStore.setSnapshot({
      scrollLeft: 65 * metrics.columnWidth,
      scrollTop: 32 * metrics.rowHeight,
      tx: 0,
      ty: 0,
    })

    expect(requestFrame).toHaveBeenCalledTimes(1)
    runtime.dispose()
  })
})
