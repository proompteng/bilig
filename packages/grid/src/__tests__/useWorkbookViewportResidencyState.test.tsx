// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import type { VisibleRegionState } from '../gridPointer.js'
import { useWorkbookViewportResidencyState, type WorkbookViewportResidencyState } from '../useWorkbookViewportResidencyState.js'
import { GridRuntimeHost } from '../runtime/gridRuntimeHost.js'
import { GridViewportResidencyRuntime } from '../runtime/gridViewportResidencyRuntime.js'

const gridMetrics = {
  columnWidth: 104,
  headerHeight: 24,
  rowHeight: 22,
  rowMarkerWidth: 44,
}

function createEmptySnapshot(sheetName: string, address: string): CellSnapshot {
  return {
    address,
    flags: 0,
    sheetName,
    value: { tag: ValueTag.Empty },
    version: 0,
  }
}

function createEngine(subscribeCells: GridEngineLike['subscribeCells']): GridEngineLike {
  return {
    getCell(sheetName: string, address: string): CellSnapshot {
      return createEmptySnapshot(sheetName, address)
    },
    getCellStyle: () => undefined,
    subscribeCells,
    workbook: {
      getSheet: () => undefined,
    },
  }
}

function createRuntimeHost(): GridRuntimeHost {
  return new GridRuntimeHost({
    columnCount: 1000,
    defaultColumnWidth: gridMetrics.columnWidth,
    defaultRowHeight: gridMetrics.rowHeight,
    gridMetrics,
    rowCount: 1000,
    viewportHeight: 240,
    viewportWidth: 640,
  })
}

const visibleRegion: VisibleRegionState = {
  freezeCols: 2,
  freezeRows: 1,
  range: {
    height: 8,
    width: 10,
    x: 260,
    y: 110,
  },
  tx: 0,
  ty: 0,
}

describe('useWorkbookViewportResidencyState', () => {
  it('keeps resident windows and frozen-pane tile interest out of the main render hook', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    let latestState: WorkbookViewportResidencyState | null = null
    const subscribeCells = vi.fn(() => () => undefined)
    const engine = createEngine(subscribeCells)
    const gridRuntimeHost = createRuntimeHost()

    function Harness() {
      latestState = useWorkbookViewportResidencyState({
        engine,
        freezeCols: 2,
        freezeRows: 1,
        gridRuntimeHost,
        sheetName: 'Sheet1',
        shouldUseRemoteRenderTileSource: true,
        visibleRegion,
      })
      return null
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    expect(subscribeCells).not.toHaveBeenCalled()
    expect(latestState?.viewport).toEqual({
      colEnd: 269,
      colStart: 260,
      rowEnd: 117,
      rowStart: 110,
    })
    expect(latestState?.residentViewport).toEqual({
      colEnd: 511,
      colStart: 256,
      rowEnd: 191,
      rowStart: 96,
    })
    expect(latestState?.renderTileViewport).toEqual({
      colEnd: 511,
      colStart: 0,
      rowEnd: 191,
      rowStart: 0,
    })
    expect(latestState?.residentHeaderRegion.range).toEqual({
      height: 96,
      width: 256,
      x: 256,
      y: 96,
    })

    await act(async () => {
      root.unmount()
    })
  })

  it('subscribes local fallback scenes to resident cells and bumps scene revision on invalidation', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    let latestState: WorkbookViewportResidencyState | null = null
    let invalidateScene: (() => void) | null = null
    const unsubscribe = vi.fn()
    const subscribeCells = vi.fn((_sheetName: string, _addresses: readonly string[], listener: () => void) => {
      invalidateScene = listener
      return unsubscribe
    })
    const engine = createEngine(subscribeCells)
    const gridRuntimeHost = createRuntimeHost()

    function Harness() {
      latestState = useWorkbookViewportResidencyState({
        engine,
        freezeCols: 2,
        freezeRows: 1,
        gridRuntimeHost,
        sheetName: 'Sheet1',
        shouldUseRemoteRenderTileSource: false,
        visibleRegion,
      })
      return null
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    expect(subscribeCells).toHaveBeenCalledTimes(1)
    expect(subscribeCells.mock.calls[0]?.[0]).toBe('Sheet1')
    expect(subscribeCells.mock.calls[0]?.[1]).toContain('A1')
    expect(subscribeCells.mock.calls[0]?.[1]).toContain('IW97')
    expect(latestState?.sceneRevision).toBe(0)

    await act(async () => {
      invalidateScene?.()
    })

    expect(latestState?.sceneRevision).toBe(1)

    await act(async () => {
      root.unmount()
    })

    expect(unsubscribe).toHaveBeenCalledTimes(1)
  })
})

describe('GridViewportResidencyRuntime', () => {
  it('retains the resident viewport until the visible range crosses a resident boundary', () => {
    const runtime = new GridViewportResidencyRuntime()
    const first = runtime.resolve({
      freezeCols: 2,
      freezeRows: 1,
      visibleRegion,
    })
    runtime.invalidateScene()
    const sameResident = runtime.resolve({
      freezeCols: 2,
      freezeRows: 1,
      visibleRegion: {
        ...visibleRegion,
        range: {
          ...visibleRegion.range,
          x: visibleRegion.range.x + 1,
          y: visibleRegion.range.y + 1,
        },
      },
    })
    runtime.invalidateScene()
    const nextResident = runtime.resolve({
      freezeCols: 2,
      freezeRows: 1,
      visibleRegion: {
        ...visibleRegion,
        range: {
          ...visibleRegion.range,
          x: visibleRegion.range.x + 260,
          y: visibleRegion.range.y,
        },
      },
    })

    expect(sameResident.residentViewport).toBe(first.residentViewport)
    expect(sameResident.visibleAddresses).toBe(first.visibleAddresses)
    expect(sameResident.renderTileViewport).toBe(first.renderTileViewport)
    expect(sameResident.sceneRevision).toBe(1)
    expect(nextResident.residentViewport).not.toBe(first.residentViewport)
    expect(nextResident.renderTileViewport.colStart).toBe(0)
    expect(nextResident.renderTileViewport.rowStart).toBe(0)
  })

  it('owns local scene invalidation subscriptions', () => {
    const runtime = new GridViewportResidencyRuntime()
    let invalidateScene: (() => void) | null = null
    const unsubscribe = vi.fn()
    const subscribeCells = vi.fn((_sheetName: string, _addresses: readonly string[], listener: () => void) => {
      invalidateScene = listener
      return unsubscribe
    })
    const invalidations: string[] = []

    const remoteUnsubscribe = runtime.connectLocalSceneInvalidation(
      {
        engine: createEngine(subscribeCells),
        sheetName: 'Sheet1',
        shouldUseRemoteRenderTileSource: true,
        visibleAddresses: ['A1'],
      },
      () => invalidations.push('remote'),
    )
    const localUnsubscribe = runtime.connectLocalSceneInvalidation(
      {
        engine: createEngine(subscribeCells),
        sheetName: 'Sheet1',
        shouldUseRemoteRenderTileSource: false,
        visibleAddresses: ['A1', 'B2'],
      },
      () => invalidations.push('local'),
    )

    expect(remoteUnsubscribe).toBeUndefined()
    expect(subscribeCells).toHaveBeenCalledTimes(1)
    expect(subscribeCells.mock.calls[0]?.[0]).toBe('Sheet1')
    expect(subscribeCells.mock.calls[0]?.[1]).toEqual(['A1', 'B2'])

    invalidateScene?.()

    expect(invalidations).toEqual(['local'])
    expect(runtime.resolve({ freezeCols: 0, freezeRows: 0, visibleRegion }).sceneRevision).toBe(1)
    localUnsubscribe?.()
    expect(unsubscribe).toHaveBeenCalledTimes(1)
  })
})
