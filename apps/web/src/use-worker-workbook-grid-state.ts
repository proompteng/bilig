import { useCallback, useMemo, useSyncExternalStore } from 'react'
import type { CellSnapshot, WorkbookMergeRangeSnapshot } from '@bilig/protocol'
import type { WorkerRuntimeSelection } from './runtime-session.js'
import {
  readViewportColumnWidths,
  readViewportHiddenColumns,
  readViewportHiddenRows,
  readViewportRowHeights,
} from './worker-workbook-view-state.js'

type SheetViewportChannel = 'columnWidths' | 'rowHeights' | 'hiddenColumns' | 'hiddenRows' | 'freeze' | 'merges'

const EMPTY_MERGE_RANGES: readonly WorkbookMergeRangeSnapshot[] = Object.freeze([])

type ViewportStoreLike = {
  subscribeSheetChannel(sheetName: string, channel: SheetViewportChannel, listener: () => void): () => void
  subscribeCell(sheetName: string, address: string, listener: () => void): () => void
  peekCell(sheetName: string, address: string): CellSnapshot | undefined
  getFreezeRows(sheetName: string): number
  getFreezeCols(sheetName: string): number
  getColumnWidths(sheetName: string): Readonly<Record<number, number>>
  getRowHeights(sheetName: string): Readonly<Record<number, number>>
  getHiddenColumns(sheetName: string): Readonly<Record<number, true>>
  getHiddenRows(sheetName: string): Readonly<Record<number, true>>
  listMergeRanges(sheetName: string): readonly WorkbookMergeRangeSnapshot[]
  clearOptimisticCellFlagsForSheet?(sheetName: string): void
}

type WorkerHandleLike = {
  readonly viewportStore: ViewportStoreLike
}

type StructuralMutationMethod = 'insertRows' | 'deleteRows' | 'insertColumns' | 'deleteColumns' | 'setFreezePane'

function readMergeRangeSignature(workerHandle: WorkerHandleLike | null | undefined, sheetName: string): string {
  const ranges = workerHandle?.viewportStore.listMergeRanges(sheetName)
  if (!ranges || ranges.length === 0) {
    return ''
  }
  return ranges
    .map((range) => `${range.sheetName}:${range.startAddress}:${range.endAddress}`)
    .toSorted()
    .join('|')
}

export function useWorkerWorkbookGridState(input: {
  workerHandle: WorkerHandleLike | null | undefined
  selection: WorkerRuntimeSelection
  emptySelectedCell: CellSnapshot
  invokeMutation: (method: StructuralMutationMethod, ...args: unknown[]) => Promise<void>
}) {
  const { workerHandle, selection, emptySelectedCell, invokeMutation } = input

  const columnWidths = useSyncExternalStore(
    useCallback(
      (listener: () => void) =>
        workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'columnWidths', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => readViewportColumnWidths(workerHandle, selection.sheetName),
    () => readViewportColumnWidths(workerHandle, selection.sheetName),
  )

  const rowHeights = useSyncExternalStore(
    useCallback(
      (listener: () => void) =>
        workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'rowHeights', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => readViewportRowHeights(workerHandle, selection.sheetName),
    () => readViewportRowHeights(workerHandle, selection.sheetName),
  )

  const hiddenColumns = useSyncExternalStore(
    useCallback(
      (listener: () => void) =>
        workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'hiddenColumns', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => readViewportHiddenColumns(workerHandle, selection.sheetName),
    () => readViewportHiddenColumns(workerHandle, selection.sheetName),
  )

  const hiddenRows = useSyncExternalStore(
    useCallback(
      (listener: () => void) =>
        workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'hiddenRows', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => readViewportHiddenRows(workerHandle, selection.sheetName),
    () => readViewportHiddenRows(workerHandle, selection.sheetName),
  )

  const freezeRows = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'freeze', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => workerHandle?.viewportStore.getFreezeRows(selection.sheetName) ?? 0,
    () => workerHandle?.viewportStore.getFreezeRows(selection.sheetName) ?? 0,
  )

  const freezeCols = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'freeze', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => workerHandle?.viewportStore.getFreezeCols(selection.sheetName) ?? 0,
    () => workerHandle?.viewportStore.getFreezeCols(selection.sheetName) ?? 0,
  )

  const mergeRangeSignature = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'merges', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => readMergeRangeSignature(workerHandle, selection.sheetName),
    () => '',
  )

  const mergeRanges = useMemo(() => {
    if (!workerHandle || mergeRangeSignature === '') {
      return EMPTY_MERGE_RANGES
    }
    const ranges = workerHandle.viewportStore.listMergeRanges(selection.sheetName)
    return ranges.length === 0 ? EMPTY_MERGE_RANGES : ranges
  }, [mergeRangeSignature, selection.sheetName, workerHandle])

  const selectedCell = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribeCell(selection.sheetName, selection.address, listener) ?? (() => {}),
      [selection.address, selection.sheetName, workerHandle],
    ),
    () => workerHandle?.viewportStore.peekCell(selection.sheetName, selection.address) ?? emptySelectedCell,
    () => emptySelectedCell,
  )

  const invokeStructuralMutation = useCallback(
    (method: Exclude<StructuralMutationMethod, 'setFreezePane'>, sheetName: string, start: number, count: number): Promise<void> => {
      workerHandle?.viewportStore.clearOptimisticCellFlagsForSheet?.(sheetName)
      return invokeMutation(method, sheetName, start, count)
    },
    [invokeMutation, workerHandle],
  )

  const invokeInsertRowsMutation = useCallback(
    (sheetName: string, startRow: number, count: number): Promise<void> =>
      invokeStructuralMutation('insertRows', sheetName, startRow, count),
    [invokeStructuralMutation],
  )

  const invokeDeleteRowsMutation = useCallback(
    (sheetName: string, startRow: number, count: number): Promise<void> =>
      invokeStructuralMutation('deleteRows', sheetName, startRow, count),
    [invokeStructuralMutation],
  )

  const invokeInsertColumnsMutation = useCallback(
    (sheetName: string, startCol: number, count: number): Promise<void> =>
      invokeStructuralMutation('insertColumns', sheetName, startCol, count),
    [invokeStructuralMutation],
  )

  const invokeDeleteColumnsMutation = useCallback(
    (sheetName: string, startCol: number, count: number): Promise<void> =>
      invokeStructuralMutation('deleteColumns', sheetName, startCol, count),
    [invokeStructuralMutation],
  )

  const invokeSetFreezePaneMutation = useCallback(
    (sheetName: string, rows: number, cols: number): Promise<void> => invokeMutation('setFreezePane', sheetName, rows, cols),
    [invokeMutation],
  )

  return {
    columnWidths,
    rowHeights,
    hiddenColumns,
    hiddenRows,
    freezeRows,
    freezeCols,
    mergeRanges,
    selectedCell,
    invokeInsertRowsMutation,
    invokeDeleteRowsMutation,
    invokeInsertColumnsMutation,
    invokeDeleteColumnsMutation,
    invokeSetFreezePaneMutation,
  }
}
