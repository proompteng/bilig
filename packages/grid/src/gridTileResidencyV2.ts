import type { Viewport } from '@bilig/protocol'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
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

export function resolveGridTileResidencyV2(input: {
  readonly visibleViewport: Viewport
  readonly velocityX?: number
  readonly velocityY?: number
  readonly warmNeighbors?: number
}): GridTileResidencyPlan {
  const visible = viewportToTileKey(resolveResidentViewport(input.visibleViewport))
  const warmNeighbors = Math.max(0, input.warmNeighbors ?? 1)
  const warm: GridTileKey[] = []
  const rowDirection = Math.sign(input.velocityY ?? 0)
  const colDirection = Math.sign(input.velocityX ?? 0)

  for (let radius = 1; radius <= warmNeighbors; radius += 1) {
    const offsets = prioritizeOffsets(createNeighborOffsets(radius), rowDirection, colDirection)
    for (const [rowStep, colStep] of offsets) {
      warm.push(offsetTile(visible, rowStep, colStep))
    }
  }

  const uniqueWarm = dedupeTiles(warm.filter((tile) => isNonEmptyTile(tile) && !sameTile(tile, visible)))
  return {
    all: [visible, ...uniqueWarm],
    visible,
    warm: uniqueWarm,
  }
}

function createNeighborOffsets(radius: number): readonly (readonly [number, number])[] {
  const offsets: Array<readonly [number, number]> = []
  for (let row = -radius; row <= radius; row += 1) {
    for (let col = -radius; col <= radius; col += 1) {
      if (row === 0 && col === 0) {
        continue
      }
      if (Math.max(Math.abs(row), Math.abs(col)) === radius) {
        offsets.push([row, col])
      }
    }
  }
  return offsets
}

function prioritizeOffsets(
  offsets: readonly (readonly [number, number])[],
  rowDirection: number,
  colDirection: number,
): readonly (readonly [number, number])[] {
  if (rowDirection === 0 && colDirection === 0) {
    return offsets
  }
  return [...offsets].toSorted(
    (left, right) => scoreOffset(right, rowDirection, colDirection) - scoreOffset(left, rowDirection, colDirection),
  )
}

function scoreOffset(offset: readonly [number, number], rowDirection: number, colDirection: number): number {
  return offset[0] * rowDirection + offset[1] * colDirection
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
  const rowSpan = tile.rowEnd - tile.rowStart + 1
  const colSpan = tile.colEnd - tile.colStart + 1
  const rowStart = clampStart(tile.rowStart + rowSteps * rowSpan, MAX_ROWS)
  const colStart = clampStart(tile.colStart + colSteps * colSpan, MAX_COLS)
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

function sameTile(left: GridTileKey, right: GridTileKey): boolean {
  return (
    left.rowStart === right.rowStart && left.rowEnd === right.rowEnd && left.colStart === right.colStart && left.colEnd === right.colEnd
  )
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
