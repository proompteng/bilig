// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ValueTag, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT, type CellSnapshot } from '@bilig/protocol'
import { packGridScenePacketV2 } from '../renderer-v2/scene-packet-v2.js'
import { packTileKey53 } from '../renderer-v3/tile-key.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import { useWorkbookGridRenderState } from '../useWorkbookGridRenderState.js'

function createEmptySnapshot(sheetName: string, address: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0,
  }
}

const engine = {
  workbook: {
    getSheet: () => undefined,
  },
  getCell(sheetName: string, address: string): CellSnapshot {
    return createEmptySnapshot(sheetName, address)
  },
  getCellStyle(_styleId: string | undefined) {
    return undefined
  },
  subscribeCells(): () => void {
    return () => undefined
  },
}

function createRenderTile(input: { readonly sheetId: number; readonly rowTile: number; readonly colTile: number }): GridRenderTile {
  const rowStart = input.rowTile * VIEWPORT_TILE_ROW_COUNT
  const colStart = input.colTile * VIEWPORT_TILE_COLUMN_COUNT
  return {
    bounds: {
      rowStart,
      rowEnd: rowStart + VIEWPORT_TILE_ROW_COUNT - 1,
      colStart,
      colEnd: colStart + VIEWPORT_TILE_COLUMN_COUNT - 1,
    },
    coord: {
      colTile: input.colTile,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: input.rowTile,
      sheetId: input.sheetId,
    },
    lastBatchId: 3,
    lastCameraSeq: 5,
    rectCount: 0,
    rectInstances: new Float32Array(20),
    textCount: 0,
    textMetrics: new Float32Array(8),
    textRuns: [],
    tileId: packTileKey53({
      colTile: input.colTile,
      dprBucket: 1,
      rowTile: input.rowTile,
      sheetOrdinal: input.sheetId,
    }),
    version: {
      axisX: 11,
      axisY: 12,
      freeze: 13,
      styles: 14,
      text: 15,
      values: 16,
    },
  }
}

describe('useWorkbookGridRenderState viewport residency', () => {
  const originalResizeObserver = globalThis.ResizeObserver
  const originalRequestAnimationFrame = window.requestAnimationFrame
  const originalCancelAnimationFrame = window.cancelAnimationFrame

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
    window.requestAnimationFrame = ((callback: FrameRequestCallback) => {
      const handle = window.setTimeout(() => callback(performance.now()), 0)
      return handle
    }) as typeof window.requestAnimationFrame
    window.cancelAnimationFrame = ((handle: number) => {
      window.clearTimeout(handle)
    }) as typeof window.cancelAnimationFrame
  })

  afterEach(() => {
    Object.defineProperty(globalThis, 'ResizeObserver', {
      configurable: true,
      value: originalResizeObserver,
      writable: true,
    })
    window.requestAnimationFrame = originalRequestAnimationFrame
    window.cancelAnimationFrame = originalCancelAnimationFrame
    document.body.innerHTML = ''
  })

  it('keeps the viewport subscription stable while horizontal scroll stays inside one resident window', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const subscribeViewport = vi.fn(() => () => undefined)
    let latestScrollTransformStore: ReturnType<typeof useWorkbookGridRenderState>['scrollTransformStore'] | null = null
    let hostElement: HTMLDivElement | null = null
    let scrollViewport: HTMLDivElement | null = null

    function Harness() {
      const renderState = useWorkbookGridRenderState({
        engine,
        sheetName: 'Sheet1',
        selectedAddr: 'A1',
        selectedCellSnapshot: createEmptySnapshot('Sheet1', 'A1'),
        editorValue: '',
        isEditingCell: false,
        subscribeViewport,
      })
      latestScrollTransformStore = renderState.scrollTransformStore

      return (
        <div
          ref={(node) => {
            renderState.handleHostRef(node)
            hostElement = node
          }}
        >
          <div
            ref={(node) => {
              renderState.scrollViewportRef.current = node
              scrollViewport = node
            }}
          />
        </div>
      )
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    Object.defineProperty(hostElement!, 'clientWidth', { configurable: true, value: 480 })
    Object.defineProperty(hostElement!, 'clientHeight', { configurable: true, value: 180 })
    Object.defineProperty(scrollViewport!, 'clientWidth', { configurable: true, value: 480 })
    Object.defineProperty(scrollViewport!, 'clientHeight', { configurable: true, value: 180 })

    await act(async () => {
      scrollViewport!.dispatchEvent(new Event('scroll'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    const initialSubscriptionCount = subscribeViewport.mock.calls.length
    expect(initialSubscriptionCount).toBeGreaterThan(0)
    expect(subscribeViewport).toHaveBeenLastCalledWith(
      'Sheet1',
      expect.objectContaining({
        rowStart: 0,
        rowEnd: 95,
        colStart: 0,
        colEnd: 255,
      }),
      expect.any(Function),
    )

    await act(async () => {
      scrollViewport!.scrollLeft = 64 * 104
      scrollViewport!.dispatchEvent(new Event('scroll'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(subscribeViewport).toHaveBeenCalledTimes(initialSubscriptionCount)
    expect(latestScrollTransformStore?.getSnapshot()).toMatchObject({
      renderTx: 64 * 104,
      scrollLeft: 64 * 104,
      tx: 0,
    })

    await act(async () => {
      scrollViewport!.scrollTop = 8 * 22
      scrollViewport!.dispatchEvent(new Event('scroll'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(subscribeViewport).toHaveBeenCalledTimes(initialSubscriptionCount)
    expect(latestScrollTransformStore?.getSnapshot()).toMatchObject({
      renderTy: 8 * 22,
      scrollTop: 8 * 22,
      ty: 0,
    })

    await act(async () => {
      scrollViewport!.scrollLeft = 256 * 104
      scrollViewport!.dispatchEvent(new Event('scroll'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(subscribeViewport.mock.calls.length).toBeGreaterThan(initialSubscriptionCount)
    expect(latestScrollTransformStore?.getSnapshot()).toMatchObject({ tx: 0 })
    expect(subscribeViewport).toHaveBeenCalledWith(
      'Sheet1',
      expect.objectContaining({
        rowStart: 0,
        rowEnd: 95,
        colStart: 512,
        colEnd: 767,
      }),
      expect.any(Function),
    )
    expect(subscribeViewport).toHaveBeenLastCalledWith(
      'Sheet1',
      expect.objectContaining({
        rowStart: 0,
        rowEnd: 95,
        colStart: 256,
        colEnd: 511,
      }),
      expect.any(Function),
    )

    await act(async () => {
      root.unmount()
    })
  })

  it('subscribes to worker resident pane scenes when the engine exposes them', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const subscribeResidentPaneScenes = vi.fn(() => () => undefined)
    const residentPaneScenes: readonly [] = []
    const peekResidentPaneScenes = vi.fn(() => residentPaneScenes)
    let hostElement: HTMLDivElement | null = null

    function Harness() {
      const renderState = useWorkbookGridRenderState({
        engine: {
          ...engine,
          subscribeResidentPaneScenes,
          peekResidentPaneScenes,
        },
        sheetName: 'Sheet1',
        selectedAddr: 'A1',
        selectedCellSnapshot: createEmptySnapshot('Sheet1', 'A1'),
        editorValue: '',
        isEditingCell: false,
      })

      return (
        <div
          ref={(node) => {
            renderState.handleHostRef(node)
            hostElement = node
          }}
        />
      )
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    Object.defineProperty(hostElement!, 'clientWidth', { configurable: true, value: 480 })
    Object.defineProperty(hostElement!, 'clientHeight', { configurable: true, value: 180 })

    await act(async () => {
      root.render(<Harness />)
      window.dispatchEvent(new Event('resize'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(subscribeResidentPaneScenes).toHaveBeenCalledWith(
      expect.objectContaining({
        sheetName: 'Sheet1',
        residentViewport: expect.objectContaining({
          rowStart: 0,
          colStart: 0,
        }),
      }),
      expect.any(Function),
    )
    expect(subscribeResidentPaneScenes.mock.calls[0]?.[0]).not.toHaveProperty('selectedCell')
    expect(peekResidentPaneScenes).toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps viewport subscriptions mounted when worker resident scenes become usable', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    let residentPaneScenes: ReturnType<typeof useWorkbookGridRenderState>['renderPanes'] = []
    const residentListeners = new Set<() => void>()
    const subscribeResidentPaneScenes = vi.fn((_request, listener: () => void) => {
      residentListeners.add(listener)
      return () => {
        residentListeners.delete(listener)
      }
    })
    const peekResidentPaneScenes = vi.fn(() => residentPaneScenes)
    const subscribeViewport = vi.fn(() => () => undefined)
    let hostElement: HTMLDivElement | null = null

    function Harness() {
      const renderState = useWorkbookGridRenderState({
        engine: {
          ...engine,
          subscribeResidentPaneScenes,
          peekResidentPaneScenes,
        },
        sheetName: 'Sheet1',
        selectedAddr: 'A1',
        selectedCellSnapshot: createEmptySnapshot('Sheet1', 'A1'),
        editorValue: '',
        isEditingCell: false,
        subscribeViewport,
      })

      return (
        <div
          ref={(node) => {
            renderState.handleHostRef(node)
            hostElement = node
          }}
        />
      )
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    Object.defineProperty(hostElement!, 'clientWidth', { configurable: true, value: 480 })
    Object.defineProperty(hostElement!, 'clientHeight', { configurable: true, value: 180 })

    await act(async () => {
      root.render(<Harness />)
      window.dispatchEvent(new Event('resize'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    const initialSubscriptionCount = subscribeViewport.mock.calls.length
    expect(initialSubscriptionCount).toBeGreaterThan(0)
    expect(subscribeViewport).toHaveBeenCalledWith('Sheet1', expect.any(Object), expect.any(Function), { initialPatch: 'none' })

    const viewport = { rowStart: 0, rowEnd: 95, colStart: 0, colEnd: 255 }
    residentPaneScenes = [
      {
        contentOffset: { x: 0, y: 0 },
        frame: { x: 0, y: 0, width: 480, height: 180 },
        generation: 1,
        packedScene: packGridScenePacketV2({
          generation: 1,
          paneId: 'body',
          sheetName: 'Sheet1',
          surfaceSize: { width: 480, height: 180 },
          gpuScene: { fillRects: [], borderRects: [] },
          textScene: { items: [] },
          viewport,
        }),
        paneId: 'body',
        scrollAxes: { x: true, y: true },
        surfaceSize: { width: 480, height: 180 },
        viewport,
      },
    ]

    await act(async () => {
      residentListeners.forEach((listener) => listener())
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(subscribeViewport).toHaveBeenCalledTimes(initialSubscriptionCount)

    await act(async () => {
      root.unmount()
    })
  })

  it('uses fixed render tile deltas instead of resident scene and viewport subscriptions when available', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const subscribeResidentPaneScenes = vi.fn(() => () => undefined)
    const peekResidentPaneScenes = vi.fn(() => [])
    const subscribeViewport = vi.fn(() => () => undefined)
    const subscribeRenderTileDeltas = vi.fn(() => () => undefined)
    const peekRenderTile = vi.fn(() => null)
    let hostElement: HTMLDivElement | null = null

    function Harness() {
      const renderState = useWorkbookGridRenderState({
        engine: {
          ...engine,
          subscribeResidentPaneScenes,
          peekResidentPaneScenes,
        },
        sheetId: 7,
        renderTileSource: {
          subscribeRenderTileDeltas,
          peekRenderTile,
        },
        sheetName: 'Sheet1',
        selectedAddr: 'A1',
        selectedCellSnapshot: createEmptySnapshot('Sheet1', 'A1'),
        editorValue: '',
        isEditingCell: false,
        subscribeViewport,
      })

      return (
        <div
          ref={(node) => {
            renderState.handleHostRef(node)
            hostElement = node
          }}
        />
      )
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    Object.defineProperty(hostElement!, 'clientWidth', { configurable: true, value: 480 })
    Object.defineProperty(hostElement!, 'clientHeight', { configurable: true, value: 180 })

    await act(async () => {
      window.dispatchEvent(new Event('resize'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(subscribeRenderTileDeltas).toHaveBeenCalledWith(
      expect.objectContaining({
        sheetId: 7,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 95,
        colStart: 0,
        colEnd: 255,
        initialDelta: 'full',
      }),
      expect.any(Function),
    )
    expect(subscribeResidentPaneScenes).not.toHaveBeenCalled()
    expect(peekResidentPaneScenes).not.toHaveBeenCalled()
    expect(subscribeViewport).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  it('retains fixed render tile panes across a transient tile miss', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const tiles = new Map<number, GridRenderTile>()
    for (let rowTile = 0; rowTile <= 2; rowTile += 1) {
      for (let colTile = 0; colTile <= 1; colTile += 1) {
        const tile = createRenderTile({ sheetId: 7, rowTile, colTile })
        tiles.set(tile.tileId, tile)
      }
    }
    let renderTileListener: (() => void) | null = null
    let tileMiss = false
    const subscribeRenderTileDeltas = vi.fn((_subscription, listener: () => void) => {
      renderTileListener = listener
      return () => {
        renderTileListener = null
      }
    })
    const peekRenderTile = vi.fn((tileId: number) => (tileMiss ? null : (tiles.get(tileId) ?? null)))
    let hostElement: HTMLDivElement | null = null
    let latestRenderState: ReturnType<typeof useWorkbookGridRenderState> | null = null

    function Harness() {
      const renderState = useWorkbookGridRenderState({
        engine,
        sheetId: 7,
        renderTileSource: {
          subscribeRenderTileDeltas,
          peekRenderTile,
        },
        sheetName: 'Sheet1',
        selectedAddr: 'A1',
        selectedCellSnapshot: createEmptySnapshot('Sheet1', 'A1'),
        editorValue: '',
        isEditingCell: false,
      })
      latestRenderState = renderState

      return (
        <div
          ref={(node) => {
            renderState.handleHostRef(node)
            hostElement = node
          }}
        />
      )
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    Object.defineProperty(hostElement!, 'clientWidth', { configurable: true, value: 480 })
    Object.defineProperty(hostElement!, 'clientHeight', { configurable: true, value: 180 })

    await act(async () => {
      root.render(<Harness />)
      window.dispatchEvent(new Event('resize'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    const bodyPane = latestRenderState?.renderPanes.find((pane) => pane.paneId === 'body')
    expect(bodyPane?.packedScene.key.valueVersion).toBe(16)

    tileMiss = true
    await act(async () => {
      renderTileListener?.()
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(latestRenderState?.renderPanes.find((pane) => pane.paneId === 'body')?.packedScene.key.valueVersion).toBe(16)

    await act(async () => {
      root.unmount()
    })
  })
})
