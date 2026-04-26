// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, test } from 'vitest'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from '../gridGeometry.js'
import { getGridMetrics } from '../gridMetrics.js'
import type { WorkbookRenderPaneState } from '../renderer/pane-scene-types.js'
import {
  TYPEGPU_ACTIVE_RESOURCE_DEFER_MS,
  WorkbookPaneRendererV2,
  resolveTypeGpuV2DrawScrollSnapshot,
  shouldDeferTypeGpuPreloadSync,
} from '../renderer-v2/WorkbookPaneRendererV2.js'
import { packGridScenePacketV2 } from '../renderer-v2/scene-packet-v2.js'

describe('WorkbookPaneRendererV2', () => {
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

  test('mounts a distinct V2 canvas for the hard migration path', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const host = document.createElement('div')
    Object.defineProperty(host, 'clientWidth', { configurable: true, value: 640 })
    Object.defineProperty(host, 'clientHeight', { configurable: true, value: 360 })
    const root = createRoot(host)
    const rendererHost = document.createElement('div')
    Object.defineProperty(rendererHost, 'clientWidth', { configurable: true, value: 640 })
    Object.defineProperty(rendererHost, 'clientHeight', { configurable: true, value: 360 })
    host.appendChild(rendererHost)
    const emptyGpuScene = { borderRects: [], fillRects: [] }
    const emptyTextScene = { items: [] }
    const viewport = { colStart: 0, colEnd: 1, rowStart: 0, rowEnd: 1 }

    await act(async () => {
      root.render(
        <WorkbookPaneRendererV2
          active
          host={rendererHost}
          geometry={null}
          panes={[
            {
              contentOffset: { x: 0, y: 0 },
              frame: { x: 46, y: 24, width: 594, height: 336 },
              generation: 1,
              gpuScene: emptyGpuScene,
              packedScene: packGridScenePacketV2({
                generation: 1,
                gpuScene: emptyGpuScene,
                paneId: 'body',
                sheetName: 'Sheet1',
                surfaceSize: { width: 640, height: 360 },
                textScene: emptyTextScene,
                viewport,
              }),
              paneId: 'body',
              scrollAxes: { x: true, y: true },
              surfaceSize: { width: 640, height: 360 },
              textScene: emptyTextScene,
              viewport,
            },
          ]}
        />,
      )
    })

    const canvas = host.querySelector('[data-testid="grid-pane-renderer"]')
    expect(canvas).toBeInstanceOf(HTMLCanvasElement)
    expect(canvas?.getAttribute('data-pane-renderer')).toBe('workbook-pane-renderer-v2')

    await act(async () => {
      root.unmount()
    })
  })

  test('derives draw scroll from live scroll and the current resident body viewport', () => {
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
    const makeBodyPane = (rowStart: number): WorkbookRenderPaneState => ({
      contentOffset: { x: 0, y: 0 },
      frame: { x: metrics.rowMarkerWidth, y: metrics.headerHeight, width: 1234, height: 696 },
      generation: 1,
      gpuScene: { borderRects: [], fillRects: [] },
      paneId: 'body',
      packedScene: packGridScenePacketV2({
        generation: 1,
        gpuScene: { borderRects: [], fillRects: [] },
        paneId: 'body',
        sheetName: 'Sheet1',
        surfaceSize: { width: 1234, height: 696 },
        textScene: { items: [] },
        viewport: { colStart: 0, colEnd: 127, rowStart, rowEnd: rowStart + 31 },
      }),
      scrollAxes: { x: true, y: true },
      surfaceSize: { width: 1234, height: 696 },
      textScene: { items: [] },
      viewport: { colStart: 0, colEnd: 127, rowStart, rowEnd: rowStart + 31 },
    })

    expect(
      resolveTypeGpuV2DrawScrollSnapshot({
        fallback: {
          scrollLeft: 64 * metrics.columnWidth,
          scrollTop: 20 * metrics.rowHeight,
          tx: 0,
          ty: 0,
        },
        geometry,
        panes: [makeBodyPane(0)],
      }),
    ).toMatchObject({
      renderTx: 64 * metrics.columnWidth,
      renderTy: 20 * metrics.rowHeight,
    })

    expect(
      resolveTypeGpuV2DrawScrollSnapshot({
        fallback: {
          scrollLeft: 64 * metrics.columnWidth,
          scrollTop: 32 * metrics.rowHeight,
          tx: 0,
          ty: 0,
        },
        geometry,
        panes: [makeBodyPane(32)],
      }),
    ).toMatchObject({
      renderTx: 64 * metrics.columnWidth,
      renderTy: 0,
    })
  })

  test('defers preload resource sync only while scroll input or camera velocity is fresh', () => {
    expect(
      shouldDeferTypeGpuPreloadSync({
        camera: null,
        lastScrollSignalAt: 1_000,
        now: 1_000 + TYPEGPU_ACTIVE_RESOURCE_DEFER_MS - 1,
      }),
    ).toBe(true)

    expect(
      shouldDeferTypeGpuPreloadSync({
        camera: {
          updatedAt: 2_000,
          velocityX: 1,
          velocityY: 0,
        },
        lastScrollSignalAt: 0,
        now: 2_000 + TYPEGPU_ACTIVE_RESOURCE_DEFER_MS - 1,
      }),
    ).toBe(true)

    expect(
      shouldDeferTypeGpuPreloadSync({
        camera: {
          updatedAt: 2_000,
          velocityX: 1,
          velocityY: 0,
        },
        lastScrollSignalAt: 1_000,
        now: 2_000 + TYPEGPU_ACTIVE_RESOURCE_DEFER_MS + 1,
      }),
    ).toBe(false)
  })
})
