import { MAX_COLS, MAX_ROWS, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT } from '@bilig/protocol'
import type { ViewportPatchSubscription } from '@bilig/worker-transport'

type ViewportBounds = Pick<ViewportPatchSubscription, 'rowStart' | 'rowEnd' | 'colStart' | 'colEnd'>

function normalizeViewportBounds(viewport: ViewportBounds): ViewportBounds {
  const rowStart = Math.max(0, Math.min(MAX_ROWS - 1, viewport.rowStart))
  const rowEnd = Math.max(rowStart, Math.min(MAX_ROWS - 1, viewport.rowEnd))
  const colStart = Math.max(0, Math.min(MAX_COLS - 1, viewport.colStart))
  const colEnd = Math.max(colStart, Math.min(MAX_COLS - 1, viewport.colEnd))
  return {
    rowStart,
    rowEnd,
    colStart,
    colEnd,
  }
}

export function listViewportTileBounds(viewport: ViewportBounds): ViewportBounds[] {
  const normalized = normalizeViewportBounds(viewport)
  const rowTileStart = Math.floor(normalized.rowStart / VIEWPORT_TILE_ROW_COUNT)
  const rowTileEnd = Math.floor(normalized.rowEnd / VIEWPORT_TILE_ROW_COUNT)
  const colTileStart = Math.floor(normalized.colStart / VIEWPORT_TILE_COLUMN_COUNT)
  const colTileEnd = Math.floor(normalized.colEnd / VIEWPORT_TILE_COLUMN_COUNT)
  const bounds: ViewportBounds[] = []

  for (let rowTile = rowTileStart; rowTile <= rowTileEnd; rowTile += 1) {
    for (let colTile = colTileStart; colTile <= colTileEnd; colTile += 1) {
      const rowStart = rowTile * VIEWPORT_TILE_ROW_COUNT
      const colStart = colTile * VIEWPORT_TILE_COLUMN_COUNT
      bounds.push({
        rowStart,
        rowEnd: Math.min(MAX_ROWS - 1, rowStart + VIEWPORT_TILE_ROW_COUNT - 1),
        colStart,
        colEnd: Math.min(MAX_COLS - 1, colStart + VIEWPORT_TILE_COLUMN_COUNT - 1),
      })
    }
  }

  return bounds
}
