import type { GridMetrics } from './gridMetrics.js'
import type { Item, Rectangle } from './gridTypes.js'

export interface VisibleColumnAxisBound {
  readonly index: number
  readonly left: number
  readonly right: number
  readonly width: number
}

export interface VisibleRowAxisBound {
  readonly index: number
  readonly top: number
  readonly bottom: number
  readonly height: number
}

export function collectVisibleColumnBounds(
  visibleItems: readonly Item[],
  getCellBounds: (col: number, row: number) => Rectangle | undefined,
  gridMetrics: GridMetrics,
): VisibleColumnAxisBound[] {
  const boundsByColumn = new Map<number, VisibleColumnAxisBound>()
  for (const [col, row] of visibleItems) {
    if (boundsByColumn.has(col)) {
      continue
    }
    const bounds = getCellBounds(col, row)
    if (!bounds) {
      continue
    }
    boundsByColumn.set(col, {
      index: col,
      left: bounds.x,
      right: bounds.x + bounds.width,
      width: bounds.width,
    })
  }
  const bounds = [...boundsByColumn.values()].toSorted((left, right) => left.left - right.left || left.index - right.index)
  if (bounds.length === 0) {
    return bounds
  }
  const offsetX = bounds[0]!.left - gridMetrics.rowMarkerWidth
  if (offsetX === 0) {
    return bounds
  }
  return bounds.map((entry) => ({
    index: entry.index,
    left: entry.left - offsetX,
    right: entry.right - offsetX,
    width: entry.width,
  }))
}

export function collectVisibleRowBounds(
  visibleItems: readonly Item[],
  getCellBounds: (col: number, row: number) => Rectangle | undefined,
  gridMetrics: GridMetrics,
): VisibleRowAxisBound[] {
  const boundsByRow = new Map<number, VisibleRowAxisBound>()
  for (const [col, row] of visibleItems) {
    if (boundsByRow.has(row)) {
      continue
    }
    const bounds = getCellBounds(col, row)
    if (!bounds) {
      continue
    }
    boundsByRow.set(row, {
      index: row,
      top: bounds.y,
      bottom: bounds.y + bounds.height,
      height: bounds.height,
    })
  }
  const bounds = [...boundsByRow.values()].toSorted((left, right) => left.top - right.top || left.index - right.index)
  if (bounds.length === 0) {
    return bounds
  }
  const offsetY = bounds[0]!.top - gridMetrics.headerHeight
  if (offsetY === 0) {
    return bounds
  }
  return bounds.map((entry) => ({
    index: entry.index,
    top: entry.top - offsetY,
    bottom: entry.bottom - offsetY,
    height: entry.height,
  }))
}
