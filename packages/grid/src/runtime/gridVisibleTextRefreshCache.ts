import { formatAddress } from '@bilig/formula'
import type { Viewport } from '@bilig/protocol'

import type { GridEngineLike } from '../grid-engine.js'
import { snapshotToRenderCell } from '../gridCells.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'

interface VisibleTextRefreshCacheInput {
  readonly engine: GridEngineLike
  readonly sceneRevision: number
  readonly sheetName: string
  readonly visibleViewport: Viewport
}

interface VisibleTextRefreshCacheEntry extends VisibleTextRefreshCacheInput {
  readonly needsLocalRefresh: boolean
  readonly tile: GridRenderTile
  readonly visibleColEnd: number
  readonly visibleColStart: number
  readonly visibleRowEnd: number
  readonly visibleRowStart: number
}

export class GridVisibleTextRefreshCache {
  private readonly entries = new Map<number, VisibleTextRefreshCacheEntry>()

  needsLocalRefresh(tileKey: number, tile: GridRenderTile | null, input: VisibleTextRefreshCacheInput): boolean {
    if (!tile) {
      this.entries.delete(tileKey)
      return false
    }
    const visibleRowStart = Math.max(tile.bounds.rowStart, input.visibleViewport.rowStart)
    const visibleRowEnd = Math.min(tile.bounds.rowEnd, input.visibleViewport.rowEnd)
    const visibleColStart = Math.max(tile.bounds.colStart, input.visibleViewport.colStart)
    const visibleColEnd = Math.min(tile.bounds.colEnd, input.visibleViewport.colEnd)
    if (visibleRowStart > visibleRowEnd || visibleColStart > visibleColEnd) {
      this.entries.delete(tileKey)
      return false
    }

    const cached = this.entries.get(tileKey)
    if (
      cached &&
      cached.tile === tile &&
      cached.engine === input.engine &&
      cached.sheetName === input.sheetName &&
      cached.sceneRevision === input.sceneRevision &&
      cached.visibleRowStart === visibleRowStart &&
      cached.visibleRowEnd === visibleRowEnd &&
      cached.visibleColStart === visibleColStart &&
      cached.visibleColEnd === visibleColEnd
    ) {
      return cached.needsLocalRefresh
    }

    const needsLocalRefresh = tileVisibleTextNeedsLocalRefresh(tile, input, {
      visibleColEnd,
      visibleColStart,
      visibleRowEnd,
      visibleRowStart,
    })
    this.entries.set(tileKey, {
      ...input,
      needsLocalRefresh,
      tile,
      visibleColEnd,
      visibleColStart,
      visibleRowEnd,
      visibleRowStart,
    })
    return needsLocalRefresh
  }
}

function tileVisibleTextNeedsLocalRefresh(
  tile: GridRenderTile,
  input: Pick<VisibleTextRefreshCacheInput, 'engine' | 'sheetName'>,
  visibleBounds: {
    readonly visibleColEnd: number
    readonly visibleColStart: number
    readonly visibleRowEnd: number
    readonly visibleRowStart: number
  },
): boolean {
  const textRunsByCell = new Map<string, string>()
  for (const run of tile.textRuns) {
    if (run.text.length > 0) {
      textRunsByCell.set(`${run.row}:${run.col}`, run.text)
    }
  }

  for (let row = visibleBounds.visibleRowStart; row <= visibleBounds.visibleRowEnd; row += 1) {
    for (let col = visibleBounds.visibleColStart; col <= visibleBounds.visibleColEnd; col += 1) {
      const snapshot = input.engine.getCell(input.sheetName, formatAddress(row, col))
      const renderCell = snapshotToRenderCell(snapshot, input.engine.getCellStyle(snapshot.styleId))
      if (renderCell.displayText.length === 0) {
        continue
      }
      if (textRunsByCell.get(`${row}:${col}`) !== renderCell.displayText) {
        return true
      }
    }
  }
  return false
}
