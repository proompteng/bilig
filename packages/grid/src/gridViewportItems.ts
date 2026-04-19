import { MAX_COLS, MAX_ROWS, type Viewport } from '@bilig/protocol'
import type { Item } from './gridTypes.js'

export function collectViewportItems(
  viewport: Viewport,
  options: {
    readonly freezeRows?: number
    readonly freezeCols?: number
  } = {},
): Item[] {
  const freezeCols = Math.max(0, Math.min(MAX_COLS, options.freezeCols ?? 0))
  const freezeRows = Math.max(0, Math.min(MAX_ROWS, options.freezeRows ?? 0))
  const rowStart = Math.max(0, Math.min(MAX_ROWS - 1, viewport.rowStart))
  const rowEnd = Math.max(rowStart, Math.min(MAX_ROWS - 1, viewport.rowEnd))
  const colStart = Math.max(0, Math.min(MAX_COLS - 1, viewport.colStart))
  const colEnd = Math.max(colStart, Math.min(MAX_COLS - 1, viewport.colEnd))
  const items: Item[] = []
  const seen = new Set<string>()

  const pushItem = (col: number, row: number) => {
    if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) {
      return
    }
    const key = `${col}:${row}`
    if (seen.has(key)) {
      return
    }
    seen.add(key)
    items.push([col, row])
  }

  for (let row = 0; row < freezeRows; row += 1) {
    for (let col = 0; col < freezeCols; col += 1) {
      pushItem(col, row)
    }
    for (let col = colStart; col <= colEnd; col += 1) {
      pushItem(col, row)
    }
  }

  for (let row = rowStart; row <= rowEnd; row += 1) {
    for (let col = 0; col < freezeCols; col += 1) {
      pushItem(col, row)
    }
    for (let col = colStart; col <= colEnd; col += 1) {
      pushItem(col, row)
    }
  }

  return items
}
