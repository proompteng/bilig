import type { GridAxisWorldIndex } from './gridAxisWorldIndex.js'
import type { GridMetrics } from './gridMetrics.js'

export function applyHiddenAxisSizes(
  sizes: Readonly<Record<number, number>>,
  hidden: Readonly<Record<number, true>> | undefined,
): Readonly<Record<number, number>> {
  const hiddenEntries = Object.keys(hidden ?? {})
  if (hiddenEntries.length === 0) {
    return sizes
  }
  const next: Record<number, number> = { ...sizes }
  for (const rawIndex of hiddenEntries) {
    next[Number(rawIndex)] = 0
  }
  return next
}

export function resolveGridScrollSpacerSize(input: {
  readonly columnAxis: GridAxisWorldIndex
  readonly rowAxis: GridAxisWorldIndex
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly hostWidth: number
  readonly hostHeight: number
  readonly gridMetrics: GridMetrics
}): { readonly width: number; readonly height: number } {
  const bodyPaneX = input.gridMetrics.rowMarkerWidth + input.frozenColumnWidth
  const bodyPaneY = input.gridMetrics.headerHeight + input.frozenRowHeight
  const bodyViewportWidth = Math.max(0, input.hostWidth - bodyPaneX)
  const bodyViewportHeight = Math.max(0, input.hostHeight - bodyPaneY)
  const scrollableBodyWidth = Math.max(0, input.columnAxis.totalSize - input.frozenColumnWidth)
  const scrollableBodyHeight = Math.max(0, input.rowAxis.totalSize - input.frozenRowHeight)
  const maxScrollX = Math.max(0, scrollableBodyWidth - bodyViewportWidth)
  const maxScrollY = Math.max(0, scrollableBodyHeight - bodyViewportHeight)
  return {
    height: Math.max(1, input.hostHeight + maxScrollY),
    width: Math.max(1, input.hostWidth + maxScrollX),
  }
}
