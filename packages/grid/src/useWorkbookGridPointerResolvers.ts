import { useCallback, type MutableRefObject } from 'react'
import {
  createPointerGeometry,
  resolveColumnResizeTarget as resolveLegacyColumnResizeTarget,
  resolveRowResizeTarget as resolveLegacyRowResizeTarget,
  resolveHeaderSelection as resolveHeaderSelectionFromGeometry,
  resolveHeaderSelectionForDrag as resolveHeaderSelectionForDragFromGeometry,
  resolvePointerCell as resolvePointerCellFromGeometry,
  type HeaderSelection,
  type PointerGeometry,
  type VisibleRegionState,
} from './gridPointer.js'
import type { Rectangle, GridSelection, Item } from './gridTypes.js'
import type { getGridMetrics } from './gridMetrics.js'
import type { GridGeometrySnapshot } from './gridGeometry.js'

export function useWorkbookGridPointerResolvers(input: {
  hostRef: MutableRefObject<HTMLDivElement | null>
  getVisibleRegion: () => VisibleRegionState
  columnWidths: Readonly<Record<number, number>>
  rowHeights: Readonly<Record<number, number>>
  gridMetrics: ReturnType<typeof getGridMetrics>
  selectedCell: { col: number; row: number }
  gridSelection: GridSelection
  getCellScreenBounds: (col: number, row: number) => Rectangle | undefined
  getGeometrySnapshot?: (() => GridGeometrySnapshot | null) | undefined
}) {
  const {
    hostRef,
    getVisibleRegion,
    columnWidths,
    rowHeights,
    gridMetrics,
    selectedCell,
    gridSelection,
    getCellScreenBounds,
    getGeometrySnapshot,
  } = input

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
      const geometrySnapshot = getGeometrySnapshot?.()
      const hostBounds = hostRef.current?.getBoundingClientRect()
      const geometrySelectedCellBounds =
        geometrySnapshot && hostBounds ? geometrySnapshot.cellScreenRect(selectedCell.col, selectedCell.row) : null
      const selectedCellBounds =
        geometrySelectedCellBounds && hostBounds
          ? {
              x: hostBounds.left + geometrySelectedCellBounds.x,
              y: hostBounds.top + geometrySelectedCellBounds.y,
              width: geometrySelectedCellBounds.width,
              height: geometrySelectedCellBounds.height,
            }
          : geometrySnapshot
            ? null
            : (getCellScreenBounds(selectedCell.col, selectedCell.row) ?? null)
      const selectionRange = gridSelection.current?.range ?? null
      if (
        selectedCellBounds &&
        gridSelection.columns.length === 0 &&
        gridSelection.rows.length === 0 &&
        selectionRange?.width === 1 &&
        selectionRange?.height === 1 &&
        clientX >= selectedCellBounds.x - 1 &&
        clientX < selectedCellBounds.x + selectedCellBounds.width &&
        clientY >= selectedCellBounds.y - 1 &&
        clientY < selectedCellBounds.y + selectedCellBounds.height
      ) {
        return [selectedCell.col, selectedCell.row]
      }
      if (geometrySnapshot && hostBounds) {
        const hit = geometrySnapshot.hitTestScreenPoint({
          x: clientX - hostBounds.left,
          y: clientY - hostBounds.top,
        })
        return hit ? [hit.col, hit.row] : null
      }
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
        selectedCellBounds,
        selectionRange,
        hasColumnSelection: gridSelection.columns.length > 0,
        hasRowSelection: gridSelection.rows.length > 0,
      })
    },
    [
      columnWidths,
      getCellScreenBounds,
      getGeometrySnapshot,
      gridMetrics,
      gridSelection,
      hostRef,
      rowHeights,
      resolvePointerGeometry,
      selectedCell.col,
      selectedCell.row,
      getVisibleRegion,
    ],
  )

  const resolveColumnResizeTarget = useCallback(
    (
      clientX: number,
      clientY: number,
      region: VisibleRegionState = getVisibleRegion(),
      geometry?: PointerGeometry | null,
      inputColumnWidths: Readonly<Record<number, number>> = columnWidths,
      defaultWidth: number = gridMetrics.columnWidth,
    ): number | null => {
      const geometrySnapshot = getGeometrySnapshot?.()
      const hostBounds = hostRef.current?.getBoundingClientRect()
      if (geometrySnapshot && hostBounds) {
        const hit = geometrySnapshot.hitTestResizeHandleScreenPoint({
          x: clientX - hostBounds.left,
          y: clientY - hostBounds.top,
        })
        if (hit?.kind === 'column') {
          return hit.index
        }
      }
      const activeGeometry = geometry ?? resolvePointerGeometry(region)
      return activeGeometry
        ? resolveLegacyColumnResizeTarget(clientX, clientY, region, activeGeometry, inputColumnWidths, defaultWidth)
        : null
    },
    [columnWidths, getGeometrySnapshot, getVisibleRegion, gridMetrics.columnWidth, hostRef, resolvePointerGeometry],
  )

  const resolveRowResizeTarget = useCallback(
    (
      clientX: number,
      clientY: number,
      region: VisibleRegionState = getVisibleRegion(),
      geometry?: PointerGeometry | null,
      inputRowHeights: Readonly<Record<number, number>> = rowHeights,
      defaultHeight: number = gridMetrics.rowHeight,
    ): number | null => {
      const geometrySnapshot = getGeometrySnapshot?.()
      const hostBounds = hostRef.current?.getBoundingClientRect()
      if (geometrySnapshot && hostBounds) {
        const hit = geometrySnapshot.hitTestResizeHandleScreenPoint({
          x: clientX - hostBounds.left,
          y: clientY - hostBounds.top,
        })
        if (hit?.kind === 'row') {
          return hit.index
        }
      }
      const activeGeometry = geometry ?? resolvePointerGeometry(region)
      return activeGeometry ? resolveLegacyRowResizeTarget(clientX, clientY, region, activeGeometry, inputRowHeights, defaultHeight) : null
    },
    [getGeometrySnapshot, getVisibleRegion, gridMetrics.rowHeight, hostRef, resolvePointerGeometry, rowHeights],
  )

  const resolveHeaderSelectionAtPointer = useCallback(
    (
      clientX: number,
      clientY: number,
      region: VisibleRegionState = getVisibleRegion(),
      geometry?: PointerGeometry | null,
    ): HeaderSelection | null => {
      const geometrySnapshot = getGeometrySnapshot?.()
      const hostBounds = hostRef.current?.getBoundingClientRect()
      if (geometrySnapshot && hostBounds) {
        const hit = geometrySnapshot.hitTestHeaderScreenPoint({
          x: clientX - hostBounds.left,
          y: clientY - hostBounds.top,
        })
        if (hit) {
          return hit
        }
      }
      const activeGeometry = geometry ?? resolvePointerGeometry(region)
      if (!activeGeometry) {
        return null
      }
      return resolveHeaderSelectionFromGeometry(clientX, clientY, region, activeGeometry, columnWidths, rowHeights, gridMetrics)
    },
    [columnWidths, getGeometrySnapshot, getVisibleRegion, gridMetrics, hostRef, resolvePointerGeometry, rowHeights],
  )

  const resolveHeaderSelectionForPointerDrag = useCallback(
    (
      kind: HeaderSelection['kind'],
      clientX: number,
      clientY: number,
      region: VisibleRegionState = getVisibleRegion(),
      geometry?: PointerGeometry | null,
    ): HeaderSelection | null => {
      const geometrySnapshot = getGeometrySnapshot?.()
      const hostBounds = hostRef.current?.getBoundingClientRect()
      if (geometrySnapshot && hostBounds) {
        const hit = geometrySnapshot.hitTestHeaderDragScreenPoint(kind, {
          x: clientX - hostBounds.left,
          y: clientY - hostBounds.top,
        })
        if (hit) {
          return hit
        }
      }
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
    [columnWidths, getGeometrySnapshot, getVisibleRegion, gridMetrics, hostRef, resolvePointerGeometry, rowHeights],
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
