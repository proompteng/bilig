// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, test } from 'vitest'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from '../gridGeometry.js'
import { getGridMetrics } from '../gridMetrics.js'
import {
  TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS,
  WorkbookPaneRendererV3,
  resolveTypeGpuV3DrawScrollSnapshot,
  shouldDeferTypeGpuV3PreloadSync,
} from '../renderer-v3/WorkbookPaneRendererV3.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'

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
    expect(canvas).toBeInstanceOf(HTMLCanvasElement)
    expect(canvas?.getAttribute('data-pane-renderer')).toBe('workbook-pane-renderer-v3')
    expect(canvas?.getAttribute('data-renderer-mode')).toBe('typegpu-v3')

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
})
