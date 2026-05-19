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
  resolveWorkbookPanePresentedRevisionV3,
  resolveWorkbookPaneTileSceneCameraSeqV3,
  resolveWorkbookPaneTileSceneRevisionV3,
  resolveWorkbookPaneTypeGpuCanvasOpacityV3,
  resolveTypeGpuV3DrawScrollSnapshot,
  shouldMountWorkbookCanvasProofLayerV3,
  shouldDeferTypeGpuV3PreloadSync,
} from '../renderer-v3/WorkbookPaneRendererV3.js'
import { WorkbookPaneCanvasFallbackV3 } from '../renderer-v3/WorkbookPaneCanvasFallbackV3.js'
import { GridDrawSchedulerV3 } from '../renderer-v3/draw-scheduler.js'
import type { DynamicGridOverlayBatchV3 } from '../renderer-v3/dynamic-overlay-batch.js'
import { GridRenderLoop } from '../renderer-v3/gridRenderLoop.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import { resolveWorkbookPaneFrameProofSignatureV3 } from '../renderer-v3/workbook-pane-renderer-host-runtime.js'
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
      sheetOrdinal: 7,
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

function createTextTilePane(rowStart = 0): WorkbookRenderTilePaneState {
  const pane = createTilePane(rowStart)
  return {
    ...pane,
    tile: {
      ...pane.tile,
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 18,
          clipWidth: 80,
          clipX: 4,
          clipY: 4,
          color: '#1f2933',
          font: '400 14.667px Arial, "Helvetica Neue", Helvetica, sans-serif',
          fontSize: 14.667,
          height: 22,
          strike: false,
          text: 'Expense Recognized',
          underline: false,
          width: 104,
          wrap: false,
          x: 104,
          y: 88,
        },
      ],
    },
  }
}

function createOverlayBatch(overrides: Partial<DynamicGridOverlayBatchV3> = {}): DynamicGridOverlayBatchV3 {
  return {
    borderRectCount: 1,
    cameraSeq: 9,
    fillRectCount: 1,
    generatedAt: 1_000,
    rectCount: 2,
    rectInstances: new Float32Array(8),
    rects: new Float32Array(8),
    rectSignature: 'selection-a1',
    seq: 9,
    sheetName: 'Sheet1',
    surfaceSize: { height: 360, width: 640 },
    ...overrides,
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
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null)
  })

  afterEach(() => {
    vi.restoreAllMocks()
    Object.defineProperty(globalThis, 'ResizeObserver', {
      configurable: true,
      value: originalResizeObserver,
      writable: true,
    })
    document.body.innerHTML = ''
  })

  test('falls back when the V3 TypeGPU backend is unavailable', async () => {
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
    expect(canvas).toBeNull()
    expect(fallbackCanvas).toBeInstanceOf(HTMLCanvasElement)
    expect(fallbackCanvas?.getAttribute('data-renderer-mode')).toBe('canvas2d-v3-fallback')

    await act(async () => {
      root.unmount()
    })
  })

  test('mounts the Canvas2D fallback only when explicitly enabled', async () => {
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
      root.render(<WorkbookPaneRendererV3 active enableCanvasFallback host={rendererHost} geometry={null} tilePanes={[createTilePane()]} />)
    })

    const fallbackCanvas = host.querySelector('[data-testid="grid-pane-renderer-fallback"]')
    expect(fallbackCanvas).toBeInstanceOf(HTMLCanvasElement)
    expect(fallbackCanvas?.getAttribute('data-renderer-mode')).toBe('canvas2d-v3-fallback')

    await act(async () => {
      root.unmount()
    })
  })

  test('keeps workbook text on the native DOM layer while the Canvas2D fallback is mounted', async () => {
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
      root.render(<WorkbookPaneRendererV3 active host={rendererHost} geometry={null} tilePanes={[createTextTilePane()]} />)
    })

    const fallbackCanvas = host.querySelector('[data-testid="grid-pane-renderer-fallback"]')
    const nativeTextLayer = host.querySelector('[data-testid="grid-native-text-layer"]')
    expect(fallbackCanvas).toBeInstanceOf(HTMLCanvasElement)
    expect(fallbackCanvas?.getAttribute('data-v3-draw-text')).toBe('false')
    expect(nativeTextLayer?.textContent).toContain('Expense Recognized')

    await act(async () => {
      root.unmount()
    })
  })

  test('labels the non-text Canvas2D grid floor separately from fallback rendering', async () => {
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
      root.render(
        <WorkbookPaneCanvasFallbackV3
          active
          drawText={false}
          geometry={null}
          headerPanes={[]}
          host={rendererHost}
          layer="grid-floor"
          overlay={null}
          scrollTransformStore={null}
          tilePanes={[createTilePane()]}
        />,
      )
    })

    const gridFloorCanvas = host.querySelector('[data-testid="grid-pane-renderer-floor"]')
    expect(gridFloorCanvas).toBeInstanceOf(HTMLCanvasElement)
    expect(gridFloorCanvas?.className).toContain('z-[12]')
    expect(gridFloorCanvas?.getAttribute('data-renderer-mode')).toBe('canvas2d-v3-grid-floor')
    expect(gridFloorCanvas?.getAttribute('data-v3-draw-text')).toBe('false')
    expect(host.querySelector('[data-testid="grid-pane-renderer-fallback"]')).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  test('exposes visible workbook revisions only after frame proof is presented', () => {
    expect(resolveWorkbookPanePresentedRevisionV3('presented', 14)).toBe(14)
    expect(resolveWorkbookPanePresentedRevisionV3('presented', 0)).toBe(0)
    expect(resolveWorkbookPanePresentedRevisionV3('presented', null)).toBeNull()
    expect(resolveWorkbookPanePresentedRevisionV3('pending', 14)).toBeNull()
    expect(resolveWorkbookPanePresentedRevisionV3('idle', 14)).toBeNull()
  })

  test('keeps the Canvas2D proof layer until the current TypeGPU frame presents', () => {
    expect(
      shouldMountWorkbookCanvasProofLayerV3({
        backendStatus: 'ready',
        frameProofStatus: 'pending',
        headerPaneCount: 1,
        tilePaneCount: 1,
      }),
    ).toBe(true)
    expect(
      shouldMountWorkbookCanvasProofLayerV3({
        backendStatus: 'ready',
        frameProofStatus: 'presented',
        headerPaneCount: 1,
        tilePaneCount: 1,
      }),
    ).toBe(false)
    expect(
      shouldMountWorkbookCanvasProofLayerV3({
        backendStatus: 'ready',
        frameProofStatus: 'pending',
        hasPresentedFrame: true,
        headerPaneCount: 1,
        tilePaneCount: 1,
      }),
    ).toBe(true)
    expect(
      shouldMountWorkbookCanvasProofLayerV3({
        backendStatus: 'ready',
        frameProofStatus: 'pending',
        hasPresentedVisibleFrame: true,
        headerPaneCount: 1,
        tilePaneCount: 1,
      }),
    ).toBe(true)
    expect(
      shouldMountWorkbookCanvasProofLayerV3({
        backendStatus: 'ready',
        frameProofStatus: 'idle',
        headerPaneCount: 0,
        tilePaneCount: 0,
      }),
    ).toBe(false)
    expect(
      shouldMountWorkbookCanvasProofLayerV3({
        backendStatus: 'initializing',
        hasPresentedFrame: true,
        headerPaneCount: 1,
        tilePaneCount: 1,
      }),
    ).toBe(true)
    expect(
      shouldMountWorkbookCanvasProofLayerV3({
        backendStatus: 'unavailable',
        hasPresentedFrame: true,
        headerPaneCount: 1,
        tilePaneCount: 1,
      }),
    ).toBe(true)
    expect(
      shouldMountWorkbookCanvasProofLayerV3({
        backendStatus: 'initializing',
        headerPaneCount: 0,
        tilePaneCount: 0,
      }),
    ).toBe(true)
  })

  test('keeps the last presented TypeGPU frame visible while a replacement frame is pending', () => {
    expect(
      resolveWorkbookPaneTypeGpuCanvasOpacityV3({
        frameProofStatus: 'pending',
        hasPresentedVisibleFrame: true,
        showCanvasFallback: true,
      }),
    ).toBe(1)
    expect(
      resolveWorkbookPaneTypeGpuCanvasOpacityV3({
        frameProofStatus: 'pending',
        hasPresentedVisibleFrame: false,
        showCanvasFallback: true,
      }),
    ).toBe(0)
    expect(
      resolveWorkbookPaneTypeGpuCanvasOpacityV3({
        frameProofStatus: 'presented',
        hasPresentedVisibleFrame: true,
        showCanvasFallback: false,
      }),
    ).toBe(1)
  })

  test('includes surface size and device scale in frame proof identity', () => {
    const basePane = createTilePane()
    const baseSignature = resolveWorkbookPaneFrameProofSignatureV3({
      headerPanes: [],
      overlay: null,
      surface: { dpr: 2, height: 360, pixelHeight: 720, pixelWidth: 1280, width: 640 },
      tilePanes: [basePane],
    })

    expect(baseSignature).not.toBe(
      resolveWorkbookPaneFrameProofSignatureV3({
        headerPanes: [],
        overlay: null,
        surface: { dpr: 2, height: 360, pixelHeight: 720, pixelWidth: 1440, width: 720 },
        tilePanes: [basePane],
      }),
    )
    expect(baseSignature).not.toBe(
      resolveWorkbookPaneFrameProofSignatureV3({
        headerPanes: [],
        overlay: null,
        surface: { dpr: 1, height: 360, pixelHeight: 360, pixelWidth: 640, width: 640 },
        tilePanes: [basePane],
      }),
    )
  })

  test('resolves visible tile-scene revision counters from rendered panes', () => {
    const firstPane = createTilePane()
    const secondPane = {
      ...createTilePane(32),
      tile: {
        ...createTilePane(32).tile,
        lastBatchId: 9,
        lastCameraSeq: 13,
      },
    }

    expect(resolveWorkbookPaneTileSceneRevisionV3([firstPane, secondPane])).toBe(9)
    expect(resolveWorkbookPaneTileSceneCameraSeqV3([firstPane, secondPane])).toBe(13)
    expect(resolveWorkbookPaneTileSceneRevisionV3([])).toBeNull()
    expect(resolveWorkbookPaneTileSceneCameraSeqV3([])).toBeNull()
  })

  test('includes tile payload signatures in frame proof identity', () => {
    const basePane = createTilePane()
    const changedTextPane: WorkbookRenderTilePaneState = {
      ...basePane,
      tile: {
        ...basePane.tile,
        textCount: 1,
        textRuns: [
          {
            align: 'left',
            clipHeight: 18,
            clipWidth: 80,
            clipX: 4,
            clipY: 4,
            color: '#1f2933',
            font: '400 13px Arial, "Helvetica Neue", Helvetica, "Segoe UI", sans-serif',
            fontSize: 13,
            height: 21,
            row: 52,
            col: 3,
            strike: false,
            text: 'Month 1',
            underline: false,
            width: 104,
            wrap: false,
            x: 312,
            y: 420,
          },
        ],
      },
    }
    const changedRectPane: WorkbookRenderTilePaneState = {
      ...basePane,
      tile: {
        ...basePane.tile,
        rectSignature: 'changed-grid-lines',
      },
    }

    expect(resolveWorkbookPaneFrameProofSignatureV3({ headerPanes: [], overlay: null, tilePanes: [basePane] })).not.toBe(
      resolveWorkbookPaneFrameProofSignatureV3({ headerPanes: [], overlay: null, tilePanes: [changedTextPane] }),
    )
    expect(resolveWorkbookPaneFrameProofSignatureV3({ headerPanes: [], overlay: null, tilePanes: [basePane] })).not.toBe(
      resolveWorkbookPaneFrameProofSignatureV3({ headerPanes: [], overlay: null, tilePanes: [changedRectPane] }),
    )
  })

  test('includes workbook render revisions in frame proof identity', () => {
    const basePane = createTilePane()
    const baseSignature = resolveWorkbookPaneFrameProofSignatureV3({
      headerPanes: [],
      overlay: null,
      renderRevisionSnapshot: {
        authoritativeRevision: 1,
        localRevision: 0,
        projectedRevision: 1,
        tileSceneCameraSeq: 1,
        tileSceneRevision: 1,
      },
      tilePanes: [basePane],
    })

    expect(baseSignature).not.toBe(
      resolveWorkbookPaneFrameProofSignatureV3({
        headerPanes: [],
        overlay: null,
        renderRevisionSnapshot: {
          authoritativeRevision: 1,
          localRevision: 0,
          projectedRevision: 2,
          tileSceneCameraSeq: 1,
          tileSceneRevision: 1,
        },
        tilePanes: [basePane],
      }),
    )
  })

  test('includes dynamic overlay payload signatures in frame proof identity', () => {
    const basePane = createTilePane()
    const baseOverlay = createOverlayBatch()
    const movedSelectionSameCamera = createOverlayBatch({
      rectInstances: new Float32Array([24, 48, 104, 22, 0, 0, 0, 0]),
      rectSignature: 'selection-b2',
    })
    const resizedSurfaceSameRects = createOverlayBatch({
      surfaceSize: { height: 720, width: 1280 },
    })

    const baseSignature = resolveWorkbookPaneFrameProofSignatureV3({
      headerPanes: [],
      overlay: baseOverlay,
      tilePanes: [basePane],
    })

    expect(baseSignature).not.toBe(
      resolveWorkbookPaneFrameProofSignatureV3({
        headerPanes: [],
        overlay: movedSelectionSameCamera,
        tilePanes: [basePane],
      }),
    )
    expect(baseSignature).not.toBe(
      resolveWorkbookPaneFrameProofSignatureV3({
        headerPanes: [],
        overlay: resizedSurfaceSameRects,
        tilePanes: [basePane],
      }),
    )
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

  test('draw runtime forwards native-text-layer mode to the TypeGPU pass', () => {
    const drawFrame = vi.fn<WorkbookPaneFrameDrawerV3>()
    const runtime = new WorkbookPaneRendererRuntimeV3(drawFrame)

    runtime.updateState({
      active: true,
      backend: {},
      drawText: false,
      surface: { dpr: 1, height: 720, pixelHeight: 720, pixelWidth: 1280, width: 1280 },
      tilePanes: [createTilePane()],
      webGpuReady: true,
    })
    runtime.drawNow()

    expect(drawFrame).toHaveBeenCalledTimes(1)
    expect(drawFrame.mock.calls[0]?.[0].drawText).toBe(false)

    runtime.dispose()
  })

  test('draw runtime prefers fresher geometry props over stale camera store snapshots', () => {
    const metrics = getGridMetrics()
    const axes = {
      columns: createGridAxisWorldIndex({ axisLength: 1024, defaultSize: metrics.columnWidth }),
      rows: createGridAxisWorldIndex({ axisLength: 1024, defaultSize: metrics.rowHeight }),
    }
    const staleGeometry = createGridGeometrySnapshotFromAxes({
      ...axes,
      dpr: 1,
      gridMetrics: metrics,
      hostHeight: 720,
      hostWidth: 1280,
      scrollLeft: 0,
      scrollTop: 13 * metrics.rowHeight + 10,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })
    const freshGeometry = createGridGeometrySnapshotFromAxes({
      ...axes,
      dpr: 1,
      freezeRows: 7,
      gridMetrics: metrics,
      hostHeight: 720,
      hostWidth: 1280,
      previousCamera: staleGeometry.camera,
      scrollLeft: 0,
      scrollTop: 13 * metrics.rowHeight + 10,
      sheetName: 'Sheet1',
      updatedAt: 200,
    })
    const cameraStore = new GridCameraStore()
    cameraStore.setSnapshot(staleGeometry)
    const scrollStore = new WorkbookGridScrollStore()
    scrollStore.setSnapshot({
      scrollLeft: 0,
      scrollTop: 13 * metrics.rowHeight + 10,
      tx: 0,
      ty: 10,
    })
    const drawFrame = vi.fn<WorkbookPaneFrameDrawerV3>()
    const runtime = new WorkbookPaneRendererRuntimeV3(drawFrame)

    runtime.updateState({
      active: true,
      backend: {},
      cameraStore,
      geometry: freshGeometry,
      scrollTransformStore: scrollStore,
      surface: { dpr: 1, height: 720, pixelHeight: 720, pixelWidth: 1280, width: 1280 },
      tilePanes: [createTilePane()],
      webGpuReady: true,
    })
    runtime.drawNow()

    expect(drawFrame).toHaveBeenCalledTimes(1)
    expect(drawFrame.mock.calls[0]?.[0].scrollSnapshot).toMatchObject({
      renderTy: freshGeometry.camera.bodyWorldY,
    })
    expect(drawFrame.mock.calls[0]?.[0].scrollSnapshot.renderTy).not.toBe(staleGeometry.camera.bodyWorldY)

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
