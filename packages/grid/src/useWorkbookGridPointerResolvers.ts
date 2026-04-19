import { useCallback, type MutableRefObject } from 'react'
import {
  createPointerGeometry,
  resolveColumnResizeTarget,
  resolveRowResizeTarget,
  resolveHeaderSelection as resolveHeaderSelectionFromGeometry,
  resolveHeaderSelectionForDrag as resolveHeaderSelectionForDragFromGeometry,
  resolvePointerCell as resolvePointerCellFromGeometry,
  type HeaderSelection,
  type PointerGeometry,
  type VisibleRegionState,
} from './gridPointer.js'
import type { Rectangle, GridSelection, Item } from './gridTypes.js'
import type { getGridMetrics } from './gridMetrics.js'

export function useWorkbookGridPointerResolvers(input: {
  hostRef: MutableRefObject<HTMLDivElement | null>
  getVisibleRegion: () => VisibleRegionState
  columnWidths: Readonly<Record<number, number>>
  rowHeights: Readonly<Record<number, number>>
  gridMetrics: ReturnType<typeof getGridMetrics>
  selectedCell: { col: number; row: number }
  gridSelection: GridSelection
  getCellScreenBounds: (col: number, row: number) => Rectangle | undefined
}) {
  const { hostRef, getVisibleRegion, columnWidths, rowHeights, gridMetrics, selectedCell, gridSelection, getCellScreenBounds } = input

  const resolvePointerGeometry = useCallback(
    (region: VisibleRegionState = getVisibleRegion()): PointerGeometry | null => {
      const hostBounds = hostRef.current?.getBoundingClientRect()
      if (!hostBounds) {
        return null
      }
      return createPointerGeometry(
        {
          left: hostBounds.left,
          top: hostBounds.top,
          right: hostBounds.right,
          bottom: hostBounds.bottom,
        },
        region,
        columnWidths,
        rowHeights,
        gridMetrics,
      )
    },
    [columnWidths, getVisibleRegion, gridMetrics, hostRef, rowHeights],
  )

  const resolvePointerCell = useCallback(
    (clientX: number, clientY: number, region: VisibleRegionState = getVisibleRegion(), geometry?: PointerGeometry | null): Item | null => {
      const activeGeometry = geometry ?? resolvePointerGeometry(region)
      if (!activeGeometry) {
        return null
      }
      return resolvePointerCellFromGeometry({
        clientX,
        clientY,
        region,
        geometry: activeGeometry,
        columnWidths,
        rowHeights,
        gridMetrics,
        selectedCell: [selectedCell.col, selectedCell.row],
        selectedCellBounds: getCellScreenBounds(selectedCell.col, selectedCell.row) ?? null,
        selectionRange: gridSelection.current?.range ?? null,
        hasColumnSelection: gridSelection.columns.length > 0,
        hasRowSelection: gridSelection.rows.length > 0,
      })
    },
    [
      columnWidths,
      getCellScreenBounds,
      gridMetrics,
      gridSelection,
      rowHeights,
      resolvePointerGeometry,
      selectedCell.col,
      selectedCell.row,
      getVisibleRegion,
    ],
  )

  const resolveHeaderSelectionAtPointer = useCallback(
    (
      clientX: number,
      clientY: number,
      region: VisibleRegionState = getVisibleRegion(),
      geometry?: PointerGeometry | null,
    ): HeaderSelection | null => {
      const activeGeometry = geometry ?? resolvePointerGeometry(region)
      if (!activeGeometry) {
        return null
      }
      return resolveHeaderSelectionFromGeometry(clientX, clientY, region, activeGeometry, columnWidths, rowHeights, gridMetrics)
    },
    [columnWidths, getVisibleRegion, gridMetrics, resolvePointerGeometry, rowHeights],
  )

  const resolveHeaderSelectionForPointerDrag = useCallback(
    (
      kind: HeaderSelection['kind'],
      clientX: number,
      clientY: number,
      region: VisibleRegionState = getVisibleRegion(),
      geometry?: PointerGeometry | null,
    ): HeaderSelection | null => {
      const activeGeometry = geometry ?? resolvePointerGeometry(region)
      if (!activeGeometry) {
        return null
      }
      return resolveHeaderSelectionForDragFromGeometry(
        kind,
        clientX,
        clientY,
        region,
        activeGeometry,
        columnWidths,
        rowHeights,
        gridMetrics,
      )
    },
    [columnWidths, getVisibleRegion, gridMetrics, resolvePointerGeometry, rowHeights],
  )

  return {
    resolveColumnResizeTarget,
    resolveRowResizeTarget,
    resolveHeaderSelectionAtPointer,
    resolveHeaderSelectionForPointerDrag,
    resolvePointerCell,
    resolvePointerGeometry,
  }
}
