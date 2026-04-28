// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ValueTag, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT, type CellSnapshot } from '@bilig/protocol'
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
    Reflect.deleteProperty(window, '__biligScrollPerf')
    document.body.innerHTML = ''
  })

  it('builds local fixed render tiles without renderer viewport subscriptions', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const subscribeViewport = vi.fn(() => () => undefined)
    const subscribeCells = vi.fn(() => () => undefined)
    let latestRenderState: ReturnType<typeof useWorkbookGridRenderState> | null = null
    let hostElement: HTMLDivElement | null = null
    let scrollViewport: HTMLDivElement | null = null

    function Harness() {
      const renderState = useWorkbookGridRenderState({
        engine: {
          ...engine,
          subscribeCells,
        },
        sheetName: 'Sheet1',
        selectedAddr: 'A1',
        selectedCellSnapshot: createEmptySnapshot('Sheet1', 'A1'),
        editorValue: '',
        isEditingCell: false,
        subscribeViewport,
      })
      latestRenderState = renderState

      return (
        <div
          ref={(node) => {
            if (node) {
              Object.defineProperty(node, 'clientWidth', { configurable: true, value: 480 })
              Object.defineProperty(node, 'clientHeight', { configurable: true, value: 180 })
            }
            renderState.handleHostRef(node)
            hostElement = node
          }}
        >
          <div
            ref={(node) => {
              if (node) {
                Object.defineProperty(node, 'clientWidth', { configurable: true, value: 480 })
                Object.defineProperty(node, 'clientHeight', { configurable: true, value: 180 })
              }
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

    expect(subscribeViewport).not.toHaveBeenCalled()
    expect(subscribeCells).toHaveBeenCalled()
    expect(latestRenderState?.renderTilePanes.some((pane) => pane.paneId === 'body')).toBe(true)

    await act(async () => {
      scrollViewport!.scrollLeft = 64 * 104
      scrollViewport!.dispatchEvent(new Event('scroll'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(subscribeViewport).not.toHaveBeenCalled()
    expect(latestRenderState?.scrollTransformStore.getSnapshot()).toMatchObject({
      renderTx: 64 * 104,
      scrollLeft: 64 * 104,
      tx: 0,
    })

    await act(async () => {
      scrollViewport!.scrollTop = 8 * 22
      scrollViewport!.dispatchEvent(new Event('scroll'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(subscribeViewport).not.toHaveBeenCalled()
    expect(latestRenderState?.scrollTransformStore.getSnapshot()).toMatchObject({
      renderTy: 8 * 22,
      scrollTop: 8 * 22,
      ty: 0,
    })

    await act(async () => {
      root.unmount()
    })
  })

  it('uses fixed render tile deltas instead of resident scene and viewport subscriptions when available', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const subscribeViewport = vi.fn(() => () => undefined)
    const subscribeRenderTileDeltas = vi.fn(() => () => undefined)
    const peekRenderTile = vi.fn(() => null)
    const scrollPerf = {
      noteRendererTileReadiness: vi.fn(),
    }
    Reflect.set(window, '__biligScrollPerf', scrollPerf)
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
        subscribeViewport,
      })
      latestRenderState = renderState

      return (
        <div
          ref={(node) => {
            if (node) {
              Object.defineProperty(node, 'clientWidth', { configurable: true, value: 480 })
              Object.defineProperty(node, 'clientHeight', { configurable: true, value: 180 })
            }
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
    expect(subscribeViewport).not.toHaveBeenCalled()
    expect(latestRenderState?.renderTilePanes.some((pane) => pane.paneId === 'body')).toBe(true)
    expect(latestRenderState?.renderTilePanes.find((pane) => pane.paneId === 'body')?.tile.coord.sheetId).toBe(7)
    expect(scrollPerf.noteRendererTileReadiness).toHaveBeenCalledWith(
      expect.objectContaining({
        exactHits: expect.any(Number),
        misses: 0,
        staleHits: 0,
      }),
    )
    expect(scrollPerf.noteRendererTileReadiness.mock.calls.at(-1)?.[0].exactHits).toBeGreaterThan(0)

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps local fallback render tiles fresh when remote tiles are stale', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    let cellText = ''
    let cellInvalidationListener: (() => void) | null = null
    let subscribedAddresses: readonly string[] = []
    const subscribeCells = vi.fn((_sheetName: string, addresses: readonly string[], listener: () => void) => {
      subscribedAddresses = addresses
      cellInvalidationListener = listener
      return () => {
        cellInvalidationListener = null
      }
    })
    const fallbackEngine = {
      workbook: {
        getSheet: () => undefined,
      },
      getCell(sheetName: string, address: string): CellSnapshot {
        if (address === 'B10' && cellText.length > 0) {
          return {
            sheetName,
            address,
            value: { tag: ValueTag.String, value: cellText, stringId: 1 },
            flags: 0,
            version: 1,
          }
        }
        return createEmptySnapshot(sheetName, address)
      },
      getCellStyle() {
        return undefined
      },
      subscribeCells,
    }
    const tiles = new Map<number, GridRenderTile>()
    for (let rowTile = 0; rowTile <= 2; rowTile += 1) {
      for (let colTile = 0; colTile <= 1; colTile += 1) {
        const tile = createRenderTile({ sheetId: 7, rowTile, colTile })
        tiles.set(tile.tileId, tile)
      }
    }
    const subscribeRenderTileDeltas = vi.fn(() => () => undefined)
    const peekRenderTile = vi.fn((tileId: number) => tiles.get(tileId) ?? null)
    let hostElement: HTMLDivElement | null = null
    let latestRenderState: ReturnType<typeof useWorkbookGridRenderState> | null = null

    function Harness() {
      const renderState = useWorkbookGridRenderState({
        engine: fallbackEngine,
        sheetId: 7,
        renderTileSource: {
          subscribeRenderTileDeltas,
          peekRenderTile,
        },
        sheetName: 'Sheet1',
        selectedAddr: 'B10',
        selectedCellSnapshot: fallbackEngine.getCell('Sheet1', 'B10'),
        editorValue: '',
        isEditingCell: false,
      })
      latestRenderState = renderState

      return (
        <div
          ref={(node) => {
            if (node) {
              Object.defineProperty(node, 'clientWidth', { configurable: true, value: 480 })
              Object.defineProperty(node, 'clientHeight', { configurable: true, value: 180 })
            }
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

    expect(subscribeCells).toHaveBeenCalled()
    expect(subscribedAddresses).toContain('B10')
    expect(latestRenderState?.renderTilePanes.some((pane) => pane.tile.textRuns.some((run) => run.text === cellText))).toBe(false)

    cellText = 'ghost content fixed'
    await act(async () => {
      cellInvalidationListener?.()
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(latestRenderState?.renderTilePanes.some((pane) => pane.tile.textRuns.some((run) => run.text === cellText))).toBe(true)

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

    const bodyPane = latestRenderState?.renderTilePanes.find((pane) => pane.paneId === 'body')
    expect(bodyPane?.tile.version.values).toBe(16)

    tileMiss = true
    await act(async () => {
      renderTileListener?.()
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(latestRenderState?.renderTilePanes.find((pane) => pane.paneId === 'body')?.tile.version.values).toBe(16)

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps resident header panes stable across single-cell selection moves', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const scrollPerf = {
      noteHeaderPaneBuild: vi.fn(),
    }
    Reflect.set(window, '__biligScrollPerf', scrollPerf)

    let scrollViewport: HTMLDivElement | null = null

    function Harness({ selectedAddr }: { readonly selectedAddr: string }) {
      const renderState = useWorkbookGridRenderState({
        engine,
        sheetName: 'Sheet1',
        selectedAddr,
        selectedCellSnapshot: createEmptySnapshot('Sheet1', selectedAddr),
        editorValue: '',
        isEditingCell: false,
      })

      return (
        <div
          ref={(node) => {
            if (node) {
              Object.defineProperty(node, 'clientWidth', { configurable: true, value: 640 })
              Object.defineProperty(node, 'clientHeight', { configurable: true, value: 360 })
            }
            renderState.handleHostRef(node)
          }}
        >
          <div
            ref={(node) => {
              if (node) {
                Object.defineProperty(node, 'clientWidth', { configurable: true, value: 640 })
                Object.defineProperty(node, 'clientHeight', { configurable: true, value: 360 })
              }
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
      root.render(<Harness selectedAddr="A1" />)
    })

    await act(async () => {
      scrollViewport!.dispatchEvent(new Event('scroll'))
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    const headerBuildsAfterMount = scrollPerf.noteHeaderPaneBuild.mock.calls.length

    await act(async () => {
      root.render(<Harness selectedAddr="C4" />)
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    await act(async () => {
      root.render(<Harness selectedAddr="F7" />)
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    expect(scrollPerf.noteHeaderPaneBuild).toHaveBeenCalledTimes(headerBuildsAfterMount)

    await act(async () => {
      root.unmount()
    })
    Reflect.deleteProperty(window, '__biligScrollPerf')
  })
})
