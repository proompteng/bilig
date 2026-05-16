// @vitest-environment jsdom
import { act, createElement, useEffect } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellSnapshot, type WorkbookMergeRangeSnapshot } from '@bilig/protocol'
import type { WorkerRuntimeSelection } from '../runtime-session.js'
import { emptyCellSnapshot } from '../worker-workbook-app-model.js'
import { useWorkerWorkbookGridState } from '../use-worker-workbook-grid-state.js'

const EMPTY_AXIS_SIZES: Readonly<Record<number, number>> = Object.freeze({})
const EMPTY_HIDDEN_AXES: Readonly<Record<number, true>> = Object.freeze({})

function GridStateHarness(props: {
  workerHandle: { viewportStore: TestViewportStore } | null
  selection: WorkerRuntimeSelection
  invokeMutation: (
    method: 'insertRows' | 'deleteRows' | 'insertColumns' | 'deleteColumns' | 'setFreezePane',
    ...args: unknown[]
  ) => Promise<void>
  capture: (state: ReturnType<typeof useWorkerWorkbookGridState>) => void
}) {
  const state = useWorkerWorkbookGridState({
    workerHandle: props.workerHandle,
    selection: props.selection,
    emptySelectedCell: emptyCellSnapshot(props.selection.sheetName, props.selection.address),
    invokeMutation: props.invokeMutation,
  })

  useEffect(() => {
    props.capture(state)
  }, [props, state])

  return createElement('div')
}

function mountHarness(): {
  host: HTMLDivElement
  root: Root
  render: (props: Parameters<typeof GridStateHarness>[0]) => Promise<void>
} {
  const host = document.createElement('div')
  document.body.appendChild(host)
  const root = createRoot(host)
  const render = async (props: Parameters<typeof GridStateHarness>[0]) => {
    await act(async () => {
      root.render(createElement(GridStateHarness, props))
    })
  }
  return { host, root, render }
}

describe('useWorkerWorkbookGridState', () => {
  afterEach(() => {
    document.body.innerHTML = ''
    vi.clearAllMocks()
  })

  it('reads the active sheet grid snapshots and switches with the selection sheet', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const viewportStore = new TestViewportStore()
    const sheet1ColumnWidths = { 1: 120 } as const
    const sheet1RowHeights = { 2: 34 } as const
    const sheet1HiddenColumns = { 4: true } as const
    const sheet1HiddenRows = { 5: true } as const
    const sheet1Cell = createStringCellSnapshot('Sheet1', 'A1', 'sheet-1')
    viewportStore.setColumnWidths('Sheet1', sheet1ColumnWidths)
    viewportStore.setRowHeights('Sheet1', sheet1RowHeights)
    viewportStore.setHiddenColumns('Sheet1', sheet1HiddenColumns)
    viewportStore.setHiddenRows('Sheet1', sheet1HiddenRows)
    viewportStore.setFreeze('Sheet1', 2, 3)
    viewportStore.setCell(sheet1Cell)
    const sheet1MergeRanges = [{ sheetName: 'Sheet1', startAddress: 'C4', endAddress: 'D4' }] as const
    viewportStore.setMergeRanges('Sheet1', sheet1MergeRanges)

    const sheet2ColumnWidths = { 7: 240 } as const
    const sheet2RowHeights = { 8: 52 } as const
    const sheet2HiddenColumns = { 1: true } as const
    const sheet2HiddenRows = { 9: true } as const
    const sheet2Cell = createStringCellSnapshot('Sheet2', 'B3', 'sheet-2')
    viewportStore.setColumnWidths('Sheet2', sheet2ColumnWidths)
    viewportStore.setRowHeights('Sheet2', sheet2RowHeights)
    viewportStore.setHiddenColumns('Sheet2', sheet2HiddenColumns)
    viewportStore.setHiddenRows('Sheet2', sheet2HiddenRows)
    viewportStore.setFreeze('Sheet2', 4, 1)
    viewportStore.setCell(sheet2Cell)

    const harness = mountHarness()
    let capturedState: ReturnType<typeof useWorkerWorkbookGridState> | null = null

    // Arrange
    await harness.render({
      workerHandle: { viewportStore },
      selection: { sheetName: 'Sheet1', address: 'A1' },
      invokeMutation: vi.fn(async () => undefined),
      capture: (state) => {
        capturedState = state
      },
    })

    // Assert
    expect(capturedState?.columnWidths).toBe(sheet1ColumnWidths)
    expect(capturedState?.rowHeights).toBe(sheet1RowHeights)
    expect(capturedState?.hiddenColumns).toBe(sheet1HiddenColumns)
    expect(capturedState?.hiddenRows).toBe(sheet1HiddenRows)
    expect(capturedState?.freezeRows).toBe(2)
    expect(capturedState?.freezeCols).toBe(3)
    expect(capturedState?.mergeRanges).toBe(sheet1MergeRanges)
    expect(capturedState?.selectedCell).toBe(sheet1Cell)

    // Act
    await harness.render({
      workerHandle: { viewportStore },
      selection: { sheetName: 'Sheet2', address: 'B3' },
      invokeMutation: vi.fn(async () => undefined),
      capture: (state) => {
        capturedState = state
      },
    })

    // Assert
    expect(capturedState?.columnWidths).toBe(sheet2ColumnWidths)
    expect(capturedState?.rowHeights).toBe(sheet2RowHeights)
    expect(capturedState?.hiddenColumns).toBe(sheet2HiddenColumns)
    expect(capturedState?.hiddenRows).toBe(sheet2HiddenRows)
    expect(capturedState?.freezeRows).toBe(4)
    expect(capturedState?.freezeCols).toBe(1)
    expect(capturedState?.mergeRanges).toEqual([])
    expect(capturedState?.selectedCell).toBe(sheet2Cell)

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('updates subscribed grid snapshots and preserves structural mutation promises', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const viewportStore = new TestViewportStore()
    const initialCell = createStringCellSnapshot('Sheet1', 'A1', 'before')
    viewportStore.setColumnWidths('Sheet1', { 0: 96 })
    viewportStore.setFreeze('Sheet1', 1, 0)
    viewportStore.setCell(initialCell)

    const insertRowsTask = Promise.resolve()
    const deleteRowsTask = Promise.resolve()
    const insertColumnsTask = Promise.resolve()
    const deleteColumnsTask = Promise.resolve()
    const freezeTask = Promise.resolve()
    const invokeMutation = vi.fn((method: 'insertRows' | 'deleteRows' | 'insertColumns' | 'deleteColumns' | 'setFreezePane') => {
      switch (method) {
        case 'insertRows':
          return insertRowsTask
        case 'deleteRows':
          return deleteRowsTask
        case 'insertColumns':
          return insertColumnsTask
        case 'deleteColumns':
          return deleteColumnsTask
        case 'setFreezePane':
          return freezeTask
      }
    })

    const harness = mountHarness()
    let capturedState: ReturnType<typeof useWorkerWorkbookGridState> | null = null
    await harness.render({
      workerHandle: { viewportStore },
      selection: { sheetName: 'Sheet1', address: 'A1' },
      invokeMutation,
      capture: (state) => {
        capturedState = state
      },
    })
    if (!capturedState) {
      throw new Error('Expected grid state to be captured')
    }

    // Act
    const nextColumnWidths = { 3: 180 } as const
    const nextMergeRanges = [{ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' }] as const
    const nextCell = createStringCellSnapshot('Sheet1', 'A1', 'after')
    await act(async () => {
      viewportStore.setColumnWidths('Sheet1', nextColumnWidths)
      viewportStore.setFreeze('Sheet1', 5, 2)
      viewportStore.setMergeRanges('Sheet1', nextMergeRanges)
      viewportStore.setCell(nextCell)
    })
    const returnedInsertRowsTask = capturedState.invokeInsertRowsMutation('Sheet1', 2, 3)
    const returnedDeleteRowsTask = capturedState.invokeDeleteRowsMutation('Sheet1', 4, 1)
    const returnedInsertColumnsTask = capturedState.invokeInsertColumnsMutation('Sheet1', 5, 2)
    const returnedDeleteColumnsTask = capturedState.invokeDeleteColumnsMutation('Sheet1', 7, 1)
    const returnedFreezeTask = capturedState.invokeSetFreezePaneMutation('Sheet1', 6, 4)

    // Assert
    expect(capturedState.columnWidths).toBe(nextColumnWidths)
    expect(capturedState.freezeRows).toBe(5)
    expect(capturedState.freezeCols).toBe(2)
    expect(capturedState.mergeRanges).toBe(nextMergeRanges)
    expect(capturedState.selectedCell).toBe(nextCell)
    expect(viewportStore.clearOptimisticCellFlagsForSheet).toHaveBeenCalledTimes(4)
    expect(viewportStore.clearOptimisticCellFlagsForSheet).toHaveBeenNthCalledWith(1, 'Sheet1')
    expect(viewportStore.clearOptimisticCellFlagsForSheet).toHaveBeenNthCalledWith(2, 'Sheet1')
    expect(viewportStore.clearOptimisticCellFlagsForSheet).toHaveBeenNthCalledWith(3, 'Sheet1')
    expect(viewportStore.clearOptimisticCellFlagsForSheet).toHaveBeenNthCalledWith(4, 'Sheet1')
    expect(invokeMutation).toHaveBeenNthCalledWith(1, 'insertRows', 'Sheet1', 2, 3)
    expect(invokeMutation).toHaveBeenNthCalledWith(2, 'deleteRows', 'Sheet1', 4, 1)
    expect(invokeMutation).toHaveBeenNthCalledWith(3, 'insertColumns', 'Sheet1', 5, 2)
    expect(invokeMutation).toHaveBeenNthCalledWith(4, 'deleteColumns', 'Sheet1', 7, 1)
    expect(invokeMutation).toHaveBeenNthCalledWith(5, 'setFreezePane', 'Sheet1', 6, 4)
    viewportStore.clearOptimisticCellFlagsForSheet.mock.invocationCallOrder.forEach((clearCallOrder, index) => {
      expect(clearCallOrder).toBeLessThan(invokeMutation.mock.invocationCallOrder[index] ?? 0)
    })
    expect(returnedInsertRowsTask).toBe(insertRowsTask)
    expect(returnedDeleteRowsTask).toBe(deleteRowsTask)
    expect(returnedInsertColumnsTask).toBe(insertColumnsTask)
    expect(returnedDeleteColumnsTask).toBe(deleteColumnsTask)
    expect(returnedFreezeTask).toBe(freezeTask)

    await act(async () => {
      await Promise.all([
        returnedInsertRowsTask,
        returnedDeleteRowsTask,
        returnedInsertColumnsTask,
        returnedDeleteColumnsTask,
        returnedFreezeTask,
      ])
      harness.root.unmount()
    })
  })
})

class TestViewportStore {
  private readonly channelListeners = new Map<string, Map<string, Set<() => void>>>()
  private readonly cellListeners = new Map<string, Set<() => void>>()
  private readonly columnWidthsBySheet = new Map<string, Readonly<Record<number, number>>>()
  private readonly rowHeightsBySheet = new Map<string, Readonly<Record<number, number>>>()
  private readonly hiddenColumnsBySheet = new Map<string, Readonly<Record<number, true>>>()
  private readonly hiddenRowsBySheet = new Map<string, Readonly<Record<number, true>>>()
  private readonly freezeRowsBySheet = new Map<string, number>()
  private readonly freezeColsBySheet = new Map<string, number>()
  private readonly mergeRangesBySheet = new Map<string, readonly WorkbookMergeRangeSnapshot[]>()
  private readonly cells = new Map<string, CellSnapshot>()
  readonly clearOptimisticCellFlagsForSheet = vi.fn()

  subscribeSheetChannel(sheetName: string, channel: string, listener: () => void): () => void {
    const channels = this.channelListeners.get(sheetName) ?? new Map<string, Set<() => void>>()
    const listeners = channels.get(channel) ?? new Set<() => void>()
    listeners.add(listener)
    channels.set(channel, listeners)
    this.channelListeners.set(sheetName, channels)
    return () => {
      listeners.delete(listener)
      if (listeners.size === 0) {
        channels.delete(channel)
      }
      if (channels.size === 0) {
        this.channelListeners.delete(sheetName)
      }
    }
  }

  subscribeCell(sheetName: string, address: string, listener: () => void): () => void {
    const key = toCellKey(sheetName, address)
    const listeners = this.cellListeners.get(key) ?? new Set<() => void>()
    listeners.add(listener)
    this.cellListeners.set(key, listeners)
    return () => {
      listeners.delete(listener)
      if (listeners.size === 0) {
        this.cellListeners.delete(key)
      }
    }
  }

  peekCell(sheetName: string, address: string): CellSnapshot | undefined {
    return this.cells.get(toCellKey(sheetName, address))
  }

  getFreezeRows(sheetName: string): number {
    return this.freezeRowsBySheet.get(sheetName) ?? 0
  }

  getFreezeCols(sheetName: string): number {
    return this.freezeColsBySheet.get(sheetName) ?? 0
  }

  getColumnWidths(sheetName: string): Readonly<Record<number, number>> {
    return this.columnWidthsBySheet.get(sheetName) ?? EMPTY_AXIS_SIZES
  }

  getRowHeights(sheetName: string): Readonly<Record<number, number>> {
    return this.rowHeightsBySheet.get(sheetName) ?? EMPTY_AXIS_SIZES
  }

  getHiddenColumns(sheetName: string): Readonly<Record<number, true>> {
    return this.hiddenColumnsBySheet.get(sheetName) ?? EMPTY_HIDDEN_AXES
  }

  getHiddenRows(sheetName: string): Readonly<Record<number, true>> {
    return this.hiddenRowsBySheet.get(sheetName) ?? EMPTY_HIDDEN_AXES
  }

  listMergeRanges(sheetName: string): readonly WorkbookMergeRangeSnapshot[] {
    return this.mergeRangesBySheet.get(sheetName) ?? []
  }

  setColumnWidths(sheetName: string, value: Readonly<Record<number, number>>): void {
    this.columnWidthsBySheet.set(sheetName, value)
    this.emitSheetChannel(sheetName, 'columnWidths')
  }

  setRowHeights(sheetName: string, value: Readonly<Record<number, number>>): void {
    this.rowHeightsBySheet.set(sheetName, value)
    this.emitSheetChannel(sheetName, 'rowHeights')
  }

  setHiddenColumns(sheetName: string, value: Readonly<Record<number, true>>): void {
    this.hiddenColumnsBySheet.set(sheetName, value)
    this.emitSheetChannel(sheetName, 'hiddenColumns')
  }

  setHiddenRows(sheetName: string, value: Readonly<Record<number, true>>): void {
    this.hiddenRowsBySheet.set(sheetName, value)
    this.emitSheetChannel(sheetName, 'hiddenRows')
  }

  setFreeze(sheetName: string, rows: number, cols: number): void {
    this.freezeRowsBySheet.set(sheetName, rows)
    this.freezeColsBySheet.set(sheetName, cols)
    this.emitSheetChannel(sheetName, 'freeze')
  }

  setMergeRanges(sheetName: string, ranges: readonly WorkbookMergeRangeSnapshot[]): void {
    this.mergeRangesBySheet.set(sheetName, ranges)
    this.emitSheetChannel(sheetName, 'merges')
  }

  setCell(snapshot: CellSnapshot): void {
    this.cells.set(toCellKey(snapshot.sheetName, snapshot.address), snapshot)
    this.emitCell(snapshot.sheetName, snapshot.address)
  }

  private emitSheetChannel(sheetName: string, channel: string): void {
    this.channelListeners
      .get(sheetName)
      ?.get(channel)
      ?.forEach((listener) => listener())
  }

  private emitCell(sheetName: string, address: string): void {
    this.cellListeners.get(toCellKey(sheetName, address))?.forEach((listener) => listener())
  }
}

function createStringCellSnapshot(sheetName: string, address: string, value: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: {
      tag: ValueTag.String,
      value,
    },
    input: value,
    flags: 0,
    version: 1,
  }
}

function toCellKey(sheetName: string, address: string): string {
  return `${sheetName}:${address}`
}
