import { MAX_COLS, MAX_ROWS, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT, type Viewport } from '@bilig/protocol'
import type { GridMetrics } from '../gridMetrics.js'
import type { GridEngineLike } from '../grid-engine.js'
import { materializeGridRenderTileV3 } from './grid-tile-materializer.js'
import { unpackTileKey53, tileKeysForViewport } from './tile-key.js'
import type { GridRenderTile } from './render-tile-source.js'

export function buildLocalFixedRenderTiles(input: {
  readonly engine: GridEngineLike
  readonly sheetName: string
  readonly sheetId: number
  readonly viewport: Viewport
  readonly columnWidths: Readonly<Record<number, number>>
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly gridMetrics: GridMetrics
  readonly dprBucket: number
  readonly generation: number
  readonly cameraSeq: number
}): readonly GridRenderTile[] {
  const axisVersionX = hashAxisOverrides(input.sortedColumnWidthOverrides)
  const axisVersionY = hashAxisOverrides(input.sortedRowHeightOverrides)
  return tileKeysForViewport({
    dprBucket: input.dprBucket,
    sheetOrdinal: input.sheetId,
    viewport: input.viewport,
  }).map((tileId) => {
    const key = unpackTileKey53(tileId)
    const tileViewport = viewportFromTileKey(key.rowTile, key.colTile)
    return materializeGridRenderTileV3({
      ...input,
      axisSeqX: axisVersionX,
      axisSeqY: axisVersionY,
      freezeSeq: 0,
      glyphAtlasSeq: 0,
      materializedAtSeq: input.generation,
      packetSeq: input.generation,
      rectSeq: input.generation,
      styleSeq: input.generation,
      textSeq: input.generation,
      viewport: tileViewport,
      valueSeq: input.generation,
    })
  })
}

function viewportFromTileKey(rowTile: number, colTile: number): Viewport {
  const rowStart = rowTile * VIEWPORT_TILE_ROW_COUNT
  const colStart = colTile * VIEWPORT_TILE_COLUMN_COUNT
  return {
    colEnd: Math.min(MAX_COLS - 1, colStart + VIEWPORT_TILE_COLUMN_COUNT - 1),
    colStart,
    rowEnd: Math.min(MAX_ROWS - 1, rowStart + VIEWPORT_TILE_ROW_COUNT - 1),
    rowStart,
  }
}

function hashAxisOverrides(entries: readonly (readonly [number, number])[]): number {
  if (entries.length === 0) {
    return 0
  }
  let hash = 2_166_136_261
  for (const [index, size] of entries) {
    hash = mixRevisionInteger(hash, index)
    hash = mixRevisionInteger(hash, Math.round(size * 1_000))
  }
  return hash >>> 0
}

function mixRevisionInteger(hash: number, value: number): number {
  return Math.imul((hash ^ value) >>> 0, 16_777_619) >>> 0
}
