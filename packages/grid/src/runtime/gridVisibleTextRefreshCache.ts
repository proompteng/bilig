import { formatAddress } from '@bilig/formula'
import type { CellStyleRecord, Viewport } from '@bilig/protocol'

import type { GridEngineLike } from '../grid-engine.js'
import { snapshotToRenderCell } from '../gridCells.js'
import { parseGpuColor, type GridGpuColor } from '../gridGpuPrimitives.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../renderer-v3/rect-instance-buffer.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'

interface VisibleTextRefreshCacheInput {
  readonly engine: GridEngineLike
  readonly sceneRevision: number
  readonly sheetName: string
  readonly visibleViewport: Viewport
}

interface VisibleTextRefreshCacheEntry extends VisibleTextRefreshCacheInput {
  readonly needsLocalRefresh: boolean
  readonly renderRevisionKey: string
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

    const renderRevisionKey = resolveRenderRevisionKey(input.engine)
    const cached = this.entries.get(tileKey)
    if (
      cached &&
      cached.tile === tile &&
      cached.engine === input.engine &&
      cached.sheetName === input.sheetName &&
      cached.sceneRevision === input.sceneRevision &&
      cached.renderRevisionKey === renderRevisionKey &&
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
      renderRevisionKey,
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
  if (tile.textRuns.length !== tile.textCount) {
    return true
  }

  const textRunsByCell = new Map<string, GridRenderTile['textRuns'][number]>()
  for (const run of tile.textRuns) {
    if (run.text.length === 0) {
      continue
    }
    const row = run.row
    const col = run.col
    if (
      !Number.isInteger(row) ||
      !Number.isInteger(col) ||
      row === undefined ||
      col === undefined ||
      row < tile.bounds.rowStart ||
      row > tile.bounds.rowEnd ||
      col < tile.bounds.colStart ||
      col > tile.bounds.colEnd
    ) {
      return true
    }
    if (
      row >= visibleBounds.visibleRowStart &&
      row <= visibleBounds.visibleRowEnd &&
      col >= visibleBounds.visibleColStart &&
      col <= visibleBounds.visibleColEnd
    ) {
      const key = `${row}:${col}`
      if (textRunsByCell.has(key)) {
        return true
      }
      textRunsByCell.set(key, run)
    }
  }

  const projectedRevision = resolveProjectedRevision(input.engine)
  const tileStyleRevisionIsBehind = projectedRevision !== null && tile.version.styles < projectedRevision
  let needsCurrentStyleProof = tileStyleRevisionIsBehind && tileHasAuthoredPaintRects(tile)
  const expectedVisibleFillColorKeys = new Set<string>()

  for (let row = visibleBounds.visibleRowStart; row <= visibleBounds.visibleRowEnd; row += 1) {
    for (let col = visibleBounds.visibleColStart; col <= visibleBounds.visibleColEnd; col += 1) {
      const snapshot = input.engine.getCell(input.sheetName, formatAddress(row, col))
      const style = input.engine.getCellStyle(snapshot.styleId)
      if (style?.fill?.backgroundColor) {
        expectedVisibleFillColorKeys.add(gpuColorKey(parseGpuColor(style.fill.backgroundColor)))
      }
      const renderCell = snapshotToRenderCell(snapshot, style)
      const visibleTileRun = textRunsByCell.get(`${row}:${col}`) ?? null
      if ((visibleTileRun?.text ?? '') !== renderCell.displayText) {
        return true
      }
      if (visibleTileRun && textRunStyleDiffersFromRenderCell(visibleTileRun, renderCell)) {
        return true
      }
      needsCurrentStyleProof = needsCurrentStyleProof || (tileStyleRevisionIsBehind && styleAffectsVisibleGridPaint(style))
    }
  }
  if (!tileContainsExpectedVisibleFillColors(tile, expectedVisibleFillColorKeys)) {
    return true
  }
  return needsCurrentStyleProof
}

function resolveRenderRevisionKey(engine: GridEngineLike): string {
  const revision = engine.getRenderRevisionSnapshot?.()
  if (!revision) {
    return 'untracked'
  }
  return [
    revision.authoritativeRevision ?? 'none',
    revision.localRevision ?? 'none',
    revision.projectedRevision,
    revision.tileSceneRevision ?? 'none',
    revision.tileSceneCameraSeq ?? 'none',
  ].join(':')
}

function resolveProjectedRevision(engine: GridEngineLike): number | null {
  const revision = engine.getRenderRevisionSnapshot?.().projectedRevision
  return typeof revision === 'number' && Number.isInteger(revision) && revision >= 0 ? revision : null
}

function styleAffectsVisibleGridPaint(style: CellStyleRecord | undefined): boolean {
  return Boolean(style?.fill?.backgroundColor || style?.borders)
}

function tileHasAuthoredPaintRects(tile: GridRenderTile): boolean {
  const readableRectCount = Math.min(tile.rectCount, Math.floor(tile.rectInstances.length / GRID_RECT_INSTANCE_FLOAT_COUNT_V3))
  for (let index = 0; index < readableRectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    const instanceKind = tile.rectInstances[offset + 13] ?? -1
    const fillAlpha = tile.rectInstances[offset + 7] ?? 0
    if (instanceKind === 0 && fillAlpha > 0.01) {
      return true
    }
  }
  return tile.rectCount > expectedBaseGridLineRectCount(tile)
}

function tileContainsExpectedVisibleFillColors(tile: GridRenderTile, expectedFillColorKeys: ReadonlySet<string>): boolean {
  const tileFillColorKeys = collectTileFillColorKeys(tile)
  if (tileFillColorKeys.size !== expectedFillColorKeys.size) {
    return false
  }
  for (const expectedKey of expectedFillColorKeys) {
    if (!tileFillColorKeys.has(expectedKey)) {
      return false
    }
  }
  return true
}

function collectTileFillColorKeys(tile: GridRenderTile): Set<string> {
  const readableRectCount = Math.min(tile.rectCount, Math.floor(tile.rectInstances.length / GRID_RECT_INSTANCE_FLOAT_COUNT_V3))
  const keys = new Set<string>()
  for (let index = 0; index < readableRectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    const instanceKind = tile.rectInstances[offset + 13] ?? -1
    const fillAlpha = tile.rectInstances[offset + 7] ?? 0
    if (instanceKind !== 0 || fillAlpha <= 0.01) {
      continue
    }
    keys.add(
      gpuColorKey({
        a: fillAlpha,
        b: tile.rectInstances[offset + 6] ?? 0,
        g: tile.rectInstances[offset + 5] ?? 0,
        r: tile.rectInstances[offset + 4] ?? 0,
      }),
    )
  }
  return keys
}

function gpuColorKey(color: GridGpuColor): string {
  return [color.r, color.g, color.b, color.a].map((channel) => String(Math.round(channel * 1000))).join(':')
}

function expectedBaseGridLineRectCount(tile: GridRenderTile): number {
  const rowCount = Math.max(0, tile.bounds.rowEnd - tile.bounds.rowStart + 1)
  const colCount = Math.max(0, tile.bounds.colEnd - tile.bounds.colStart + 1)
  return rowCount + colCount
}

function textRunStyleDiffersFromRenderCell(
  run: GridRenderTile['textRuns'][number],
  renderCell: ReturnType<typeof snapshotToRenderCell>,
): boolean {
  return (
    (run.font ?? '') !== renderCell.font ||
    (run.color ?? '') !== renderCell.color ||
    (run.align ?? 'left') !== renderCell.align ||
    (run.wrap ?? false) !== renderCell.wrap ||
    run.underline !== renderCell.underline
  )
}
