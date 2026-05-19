// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, test, vi } from 'vitest'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from '../gridGeometry.js'
import { getGridMetrics } from '../gridMetrics.js'
import { GridCameraStore } from '../runtime/gridCameraStore.js'
import {
  WorkbookPaneCanvasFallbackV3,
  drawTextRuns,
  resolveWorkbookPaneCanvasFallbackFrame,
  type CanvasTextRunContext,
} from '../renderer-v3/WorkbookPaneCanvasFallbackV3.js'
import type { DynamicGridOverlayBatchV3 } from '../renderer-v3/dynamic-overlay-batch.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import { WORKBOOK_FONT_SANS } from '../workbookTheme.js'

const canvasGetContextDescriptor = Object.getOwnPropertyDescriptor(HTMLCanvasElement.prototype, 'getContext')

function createCanvasContextMock(): {
  readonly context: CanvasTextRunContext
  readonly fillText: ReturnType<typeof vi.fn<(text: string, x: number, y: number, maxWidth?: number) => void>>
  readonly lineTo: ReturnType<typeof vi.fn<(x: number, y: number) => void>>
} {
  const fillText = vi.fn<(text: string, x: number, y: number, maxWidth?: number) => void>()
  const lineTo = vi.fn<(x: number, y: number) => void>()
  const context = {
    beginPath: vi.fn(),
    clip: vi.fn(),
    fillStyle: '',
    fillText,
    font: '',
    lineTo,
    lineWidth: 1,
    measureText: vi.fn(() => ({ width: 520 })),
    moveTo: vi.fn(),
    rect: vi.fn(),
    restore: vi.fn(),
    save: vi.fn(),
    stroke: vi.fn(),
    strokeStyle: '',
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'middle' as CanvasTextBaseline,
  }
  return { context, fillText, lineTo }
}

function createOverlay(cameraSeq: number): DynamicGridOverlayBatchV3 {
  return {
    borderRectCount: 0,
    cameraSeq,
    fillRectCount: 0,
    generatedAt: 0,
    rectCount: 0,
    rectInstances: new Float32Array(20),
    rectSignature: `camera:${cameraSeq}`,
    rects: new Float32Array(20),
    seq: cameraSeq,
    sheetName: 'Sheet1',
    surfaceSize: { height: 240, width: 480 },
  }
}

function createFilledOverlay(cameraSeq: number): DynamicGridOverlayBatchV3 {
  const rectInstances = new Float32Array(20)
  rectInstances[0] = 12
  rectInstances[1] = 18
  rectInstances[2] = 44
  rectInstances[3] = 24
  rectInstances[4] = 0
  rectInstances[5] = 1
  rectInstances[6] = 0
  rectInstances[7] = 1
  return {
    ...createOverlay(cameraSeq),
    fillRectCount: 1,
    rectCount: 1,
    rectInstances,
    rectSignature: `filled-camera:${cameraSeq}`,
  }
}

function createBodyPane(rowStart = 0): WorkbookRenderTilePaneState {
  const tile: GridRenderTile = {
    bounds: { colEnd: 31, colStart: 0, rowEnd: rowStart + 31, rowStart },
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: Math.floor(rowStart / 32),
      sheetId: 1,
    },
    lastBatchId: 1,
    lastCameraSeq: 1,
    rectCount: 0,
    rectInstances: new Float32Array(20),
    textCount: 0,
    textMetrics: new Float32Array(8),
    textRuns: [],
    tileId: 1,
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
    frame: { x: 46, y: 24, width: 434, height: 216 },
    generation: 1,
    paneId: 'body',
    scrollAxes: { x: true, y: true },
    surfaceSize: { height: 240, width: 480 },
    tile,
    viewport: tile.bounds,
  }
}

function createTextBodyPane(): WorkbookRenderTilePaneState {
  const pane = createBodyPane()
  return {
    ...pane,
    tile: {
      ...pane.tile,
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 120,
          clipX: 0,
          clipY: 0,
          color: '#111827',
          col: 1,
          font: '400 14px Arial, sans-serif',
          fontSize: 14,
          height: 24,
          row: 1,
          strike: false,
          text: 'stale canvas text',
          underline: false,
          width: 120,
          wrap: false,
          x: 104,
          y: 24,
        },
      ],
    },
  }
}

describe('WorkbookPaneCanvasFallbackV3', () => {
  afterEach(() => {
    vi.restoreAllMocks()
    if (canvasGetContextDescriptor) {
      Object.defineProperty(HTMLCanvasElement.prototype, 'getContext', canvasGetContextDescriptor)
    }
    document.body.innerHTML = ''
  })

  test('clips long text without using canvas maxWidth font compression', () => {
    const { context, fillText } = createCanvasContextMock()

    drawTextRuns(context, [
      {
        align: 'left',
        clipHeight: 22,
        clipWidth: 180,
        clipX: 10,
        clipY: 20,
        color: '#111827',
        font: '11px system-ui, sans-serif',
        fontSize: 11,
        height: 22,
        strike: false,
        text: 'Amortization schedule note that must clip instead of squeezing into one cell',
        underline: false,
        width: 180,
        x: 10,
        y: 20,
      },
    ])

    expect(fillText).toHaveBeenCalledWith('Amortization schedule note that must clip instead of squeezing into one cell', 16, 31)
    expect(fillText.mock.calls[0]).toHaveLength(3)
  })

  test('bounds fallback text decoration to the clip rectangle', () => {
    const { context, lineTo } = createCanvasContextMock()

    drawTextRuns(context, [
      {
        align: 'left',
        clipHeight: 22,
        clipWidth: 80,
        clipX: 10,
        clipY: 20,
        color: '#111827',
        font: '11px system-ui, sans-serif',
        fontSize: 11,
        height: 22,
        strike: false,
        text: 'Underlined text that is far wider than the visible cell',
        underline: true,
        width: 500,
        x: 10,
        y: 20,
      },
    ])

    expect(lineTo).toHaveBeenCalledWith(96, 38)
  })

  test('uses the workbook font stack when a V3 text run does not carry a font', () => {
    const { context, fillText } = createCanvasContextMock()

    drawTextRuns(context, [
      {
        clipHeight: 22,
        clipWidth: 120,
        clipX: 0,
        clipY: 0,
        font: '',
        fontSize: 13,
        height: 22,
        text: 'Expense Recognized',
        width: 120,
        x: 0,
        y: 0,
      },
    ])

    expect(fillText).toHaveBeenCalled()
    expect(context.font).toBe(`400 13px ${WORKBOOK_FONT_SANS}`)
  })

  test('builds fallback overlay and scroll offsets from the live camera store', () => {
    const metrics = getGridMetrics()
    const axes = {
      columns: createGridAxisWorldIndex({ axisLength: 256, defaultSize: metrics.columnWidth }),
      rows: createGridAxisWorldIndex({ axisLength: 256, defaultSize: metrics.rowHeight }),
    }
    const initialGeometry = createGridGeometrySnapshotFromAxes({
      ...axes,
      dpr: 1,
      gridMetrics: metrics,
      hostHeight: 240,
      hostWidth: 480,
      scrollLeft: 0,
      scrollTop: 0,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })
    const liveGeometry = createGridGeometrySnapshotFromAxes({
      ...axes,
      dpr: 1,
      gridMetrics: metrics,
      hostHeight: 240,
      hostWidth: 480,
      previousCamera: initialGeometry.camera,
      scrollLeft: 0,
      scrollTop: metrics.rowHeight * 32,
      sheetName: 'Sheet1',
      updatedAt: 200,
    })
    const cameraStore = new GridCameraStore()
    cameraStore.setSnapshot(liveGeometry)
    const scrollStore = new WorkbookGridScrollStore()
    scrollStore.setSnapshot({
      scrollLeft: 0,
      scrollTop: metrics.rowHeight * 32,
      tx: 0,
      ty: 0,
    })
    const overlayBuilder = vi.fn((geometry: typeof liveGeometry) => createOverlay(geometry.camera.seq))

    const frame = resolveWorkbookPaneCanvasFallbackFrame({
      cameraStore,
      geometry: initialGeometry,
      overlay: null,
      overlayBuilder,
      scrollTransformStore: scrollStore,
      tilePanes: [createBodyPane(32)],
    })

    expect(frame.geometry).toBe(liveGeometry)
    expect(frame.overlay?.cameraSeq).toBe(liveGeometry.camera.seq)
    expect(frame.scrollSnapshot.renderTy).toBe(0)
    expect(overlayBuilder).toHaveBeenCalledWith(liveGeometry)
  })

  test('clears stale fallback text synchronously when native text takes ownership', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const host = document.createElement('div')
    Object.defineProperty(host, 'clientWidth', { configurable: true, value: 480 })
    Object.defineProperty(host, 'clientHeight', { configurable: true, value: 240 })
    document.body.appendChild(host)

    const fillText = vi.fn()
    const clearRect = vi.fn()
    const context = {
      beginPath: vi.fn(),
      clearRect,
      clip: vi.fn(),
      fillRect: vi.fn(),
      fillStyle: '',
      fillText,
      font: '',
      lineTo: vi.fn(),
      lineWidth: 1,
      measureText: vi.fn(() => ({ width: 120 })),
      moveTo: vi.fn(),
      rect: vi.fn(),
      restore: vi.fn(),
      save: vi.fn(),
      setTransform: vi.fn(),
      stroke: vi.fn(),
      strokeStyle: '',
      textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'middle' as CanvasTextBaseline,
      translate: vi.fn(),
    }
    Object.defineProperty(HTMLCanvasElement.prototype, 'getContext', {
      configurable: true,
      value: vi.fn(() => context),
    })
    vi.spyOn(window, 'requestAnimationFrame').mockReturnValue(1)
    vi.spyOn(window, 'cancelAnimationFrame').mockImplementation(() => undefined)

    const root = createRoot(host)
    const pane = createTextBodyPane()
    await act(async () => {
      root.render(
        createElement(WorkbookPaneCanvasFallbackV3, {
          active: true,
          drawText: true,
          geometry: null,
          headerPanes: [],
          host,
          overlay: null,
          scrollTransformStore: null,
          tilePanes: [pane],
        }),
      )
    })

    expect(clearRect).toHaveBeenCalled()
    expect(fillText).toHaveBeenCalledWith('stale canvas text', 110, 36)
    fillText.mockClear()
    clearRect.mockClear()

    await act(async () => {
      root.render(
        createElement(WorkbookPaneCanvasFallbackV3, {
          active: true,
          drawText: false,
          geometry: null,
          headerPanes: [],
          host,
          overlay: null,
          scrollTransformStore: null,
          tilePanes: [pane],
        }),
      )
    })

    expect(clearRect).toHaveBeenCalled()
    expect(fillText).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  test('prefers fresher geometry props over stale live camera store snapshots', () => {
    const metrics = getGridMetrics()
    const axes = {
      columns: createGridAxisWorldIndex({ axisLength: 256, defaultSize: metrics.columnWidth }),
      rows: createGridAxisWorldIndex({ axisLength: 256, defaultSize: metrics.rowHeight }),
    }
    const staleGeometry = createGridGeometrySnapshotFromAxes({
      ...axes,
      dpr: 1,
      gridMetrics: metrics,
      hostHeight: 240,
      hostWidth: 480,
      scrollLeft: 0,
      scrollTop: metrics.rowHeight * 13 + 10,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })
    const freshGeometry = createGridGeometrySnapshotFromAxes({
      ...axes,
      dpr: 1,
      freezeRows: 7,
      gridMetrics: metrics,
      hostHeight: 240,
      hostWidth: 480,
      previousCamera: staleGeometry.camera,
      scrollLeft: 0,
      scrollTop: metrics.rowHeight * 13 + 10,
      sheetName: 'Sheet1',
      updatedAt: 200,
    })
    const cameraStore = new GridCameraStore()
    cameraStore.setSnapshot(staleGeometry)
    const scrollStore = new WorkbookGridScrollStore()
    scrollStore.setSnapshot({
      scrollLeft: 0,
      scrollTop: metrics.rowHeight * 13 + 10,
      tx: 0,
      ty: 10,
    })
    const overlayBuilder = vi.fn((geometry: typeof freshGeometry) => createOverlay(geometry.camera.seq))

    const frame = resolveWorkbookPaneCanvasFallbackFrame({
      cameraStore,
      geometry: freshGeometry,
      overlay: null,
      overlayBuilder,
      scrollTransformStore: scrollStore,
      tilePanes: [createBodyPane()],
    })

    expect(frame.geometry).toBe(freshGeometry)
    expect(frame.overlay?.cameraSeq).toBe(freshGeometry.camera.seq)
    expect(frame.scrollSnapshot.renderTy).toBe(freshGeometry.camera.bodyWorldY)
    expect(frame.scrollSnapshot.renderTy).not.toBe(staleGeometry.camera.bodyWorldY)
    expect(overlayBuilder).toHaveBeenCalledWith(freshGeometry)
  })

  test('keeps grid-floor canvas free of selection overlays', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const host = document.createElement('div')
    Object.defineProperty(host, 'clientWidth', { configurable: true, value: 480 })
    Object.defineProperty(host, 'clientHeight', { configurable: true, value: 240 })
    document.body.appendChild(host)

    const fillRect = vi.fn()
    const context = {
      beginPath: vi.fn(),
      clearRect: vi.fn(),
      clip: vi.fn(),
      fillRect,
      fillStyle: '',
      fillText: vi.fn(),
      font: '',
      lineTo: vi.fn(),
      lineWidth: 1,
      measureText: vi.fn(() => ({ width: 120 })),
      moveTo: vi.fn(),
      rect: vi.fn(),
      restore: vi.fn(),
      save: vi.fn(),
      setTransform: vi.fn(),
      stroke: vi.fn(),
      strokeStyle: '',
      textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'middle' as CanvasTextBaseline,
      translate: vi.fn(),
    }
    Object.defineProperty(HTMLCanvasElement.prototype, 'getContext', {
      configurable: true,
      value: vi.fn(() => context),
    })
    vi.spyOn(window, 'requestAnimationFrame').mockReturnValue(1)
    vi.spyOn(window, 'cancelAnimationFrame').mockImplementation(() => undefined)

    const root = createRoot(host)
    await act(async () => {
      root.render(
        createElement(WorkbookPaneCanvasFallbackV3, {
          active: true,
          drawText: false,
          geometry: null,
          headerPanes: [],
          host,
          layer: 'grid-floor',
          overlay: createFilledOverlay(1),
          scrollTransformStore: null,
          tilePanes: [createBodyPane()],
        }),
      )
    })

    expect(host.querySelector('[data-testid="grid-pane-renderer-floor"]')?.getAttribute('data-v3-overlay-enabled')).toBe('false')
    expect(fillRect).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })
})
