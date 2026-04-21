import type { Viewport } from '@bilig/protocol'

export function collectViewportSubscriptions(input: {
  readonly viewport: Viewport
  readonly freezeRows: number
  readonly freezeCols: number
  readonly warmViewports?: readonly Viewport[]
}): Viewport[] {
  const { viewport, freezeRows, freezeCols, warmViewports = [] } = input
  const viewports: Viewport[] = [...warmViewports, viewport]
  if (freezeRows > 0) {
    viewports.push({
      colEnd: viewport.colEnd,
      colStart: viewport.colStart,
      rowEnd: freezeRows - 1,
      rowStart: 0,
    })
  }
  if (freezeCols > 0) {
    viewports.push({
      colEnd: freezeCols - 1,
      colStart: 0,
      rowEnd: viewport.rowEnd,
      rowStart: viewport.rowStart,
    })
  }
  if (freezeRows > 0 && freezeCols > 0) {
    viewports.push({
      colEnd: freezeCols - 1,
      colStart: 0,
      rowEnd: freezeRows - 1,
      rowStart: 0,
    })
  }
  return [...new Map(viewports.map((entry) => [`${entry.rowStart}:${entry.rowEnd}:${entry.colStart}:${entry.colEnd}`, entry])).values()]
}
