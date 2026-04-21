import type { Viewport } from '@bilig/protocol'
import { MAX_COLS, MAX_ROWS, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT } from '@bilig/protocol'
import { resolveResidentViewport } from './workbookGridViewport.js'

export interface GridTileKey {
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}

export interface GridTileResidencyPlan {
  readonly visible: GridTileKey
  readonly warm: readonly GridTileKey[]
  readonly all: readonly GridTileKey[]
}

export function resolveGridTileResidency(input: {
  readonly visibleViewport: Viewport
  readonly velocityX?: number
  readonly velocityY?: number
  readonly warmNeighbors?: number
}): GridTileResidencyPlan {
  const visible = viewportToTileKey(resolveResidentViewport(input.visibleViewport))
  const warmNeighbors = Math.max(0, input.warmNeighbors ?? 1)
  const warm: GridTileKey[] = []
  const colDirection = Math.sign(input.velocityX ?? 0)
  const rowDirection = Math.sign(input.velocityY ?? 0)

  for (let step = 1; step <= warmNeighbors; step += 1) {
    if (colDirection !== 0) {
      warm.push(offsetTile(visible, 0, colDirection * step))
    }
    if (rowDirection !== 0) {
      warm.push(offsetTile(visible, rowDirection * step, 0))
    }
    if (colDirection !== 0 && rowDirection !== 0) {
      warm.push(offsetTile(visible, rowDirection * step, colDirection * step))
    }
  }

  const uniqueWarm = dedupeTiles(warm.filter(isNonEmptyTile))
  return {
    all: dedupeTiles([visible, ...uniqueWarm]),
    visible,
    warm: uniqueWarm,
  }
}

function viewportToTileKey(viewport: Viewport): GridTileKey {
  return {
    colEnd: viewport.colEnd,
    colStart: viewport.colStart,
    rowEnd: viewport.rowEnd,
    rowStart: viewport.rowStart,
  }
}

function offsetTile(tile: GridTileKey, rowSteps: number, colSteps: number): GridTileKey {
  const rowStart = clampStart(tile.rowStart + rowSteps * VIEWPORT_TILE_ROW_COUNT, MAX_ROWS)
  const colStart = clampStart(tile.colStart + colSteps * VIEWPORT_TILE_COLUMN_COUNT, MAX_COLS)
  return {
    colEnd: Math.min(MAX_COLS - 1, colStart + tile.colEnd - tile.colStart),
    colStart,
    rowEnd: Math.min(MAX_ROWS - 1, rowStart + tile.rowEnd - tile.rowStart),
    rowStart,
  }
}

function clampStart(value: number, axisMax: number): number {
  return Math.max(0, Math.min(axisMax - 1, value))
}

function isNonEmptyTile(tile: GridTileKey): boolean {
  return tile.rowEnd >= tile.rowStart && tile.colEnd >= tile.colStart
}

function dedupeTiles(tiles: readonly GridTileKey[]): readonly GridTileKey[] {
  const seen = new Set<string>()
  const unique: GridTileKey[] = []
  for (const tile of tiles) {
    const key = `${tile.rowStart}:${tile.rowEnd}:${tile.colStart}:${tile.colEnd}`
    if (seen.has(key)) {
      continue
    }
    seen.add(key)
    unique.push(tile)
  }
  return unique
}
