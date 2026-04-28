import type { GridRenderTile, GridRenderTileDirtySpan, GridRenderTileDirtySpans } from './render-tile-source.js'
import { DirtyMaskV3 } from './tile-damage-index.js'

const RECT_DIRTY_MASK_V3 =
  DirtyMaskV3.Style | DirtyMaskV3.Rect | DirtyMaskV3.Border | DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Freeze
const TEXT_DIRTY_MASK_V3 =
  DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Freeze

const EMPTY_DIRTY_SPANS: GridRenderTileDirtySpans = Object.freeze({
  glyphSpans: Object.freeze([]),
  rectSpans: Object.freeze([]),
  textSpans: Object.freeze([]),
})

export function resolveFullGridRenderTileDirtySpansV3(tile: Pick<GridRenderTile, 'rectCount' | 'textCount'>): GridRenderTileDirtySpans {
  return {
    glyphSpans: [],
    rectSpans: tile.rectCount > 0 ? [{ offset: 0, length: tile.rectCount }] : [],
    textSpans: tile.textCount > 0 ? [{ offset: 0, length: tile.textCount }] : [],
  }
}

export function resolveGridRenderTileDirtySpansV3(tile: GridRenderTile): GridRenderTileDirtySpans {
  const dirtyMasks = tile.dirtyMasks
  const dirtyLocalRows = tile.dirtyLocalRows
  const dirtyLocalCols = tile.dirtyLocalCols
  if (!dirtyMasks || !dirtyLocalRows || !dirtyLocalCols) {
    return resolveFullGridRenderTileDirtySpansV3(tile)
  }
  if (dirtyMasks.length === 0) {
    return EMPTY_DIRTY_SPANS
  }
  if (dirtyLocalRows.length !== dirtyMasks.length * 2 || dirtyLocalCols.length !== dirtyMasks.length * 2) {
    return resolveFullGridRenderTileDirtySpansV3(tile)
  }

  const contentMask = resolveDirtyContentMask(dirtyMasks)
  return {
    glyphSpans: [],
    rectSpans:
      tile.rectCount > 0 && (contentMask & RECT_DIRTY_MASK_V3) !== 0
        ? resolveGridRenderTileRectDirtySpansV3(tile, dirtyLocalRows, dirtyLocalCols, dirtyMasks)
        : [],
    textSpans:
      tile.textCount > 0 && (contentMask & TEXT_DIRTY_MASK_V3) !== 0
        ? resolveGridRenderTileTextDirtySpansV3(tile, dirtyLocalRows, dirtyLocalCols, dirtyMasks)
        : [],
  }
}

export function isFullGridRenderTileDirtySpanV3(span: GridRenderTileDirtySpan, count: number): boolean {
  return span.offset === 0 && span.length >= count
}

function resolveDirtyContentMask(masks: Uint32Array): number {
  let contentMask = 0
  for (const mask of masks) {
    contentMask |= mask
  }
  return contentMask
}

function resolveGridRenderTileRectDirtySpansV3(
  tile: GridRenderTile,
  dirtyLocalRows: Uint32Array,
  dirtyLocalCols: Uint32Array,
  dirtyMasks: Uint32Array,
): readonly GridRenderTileDirtySpan[] {
  const rowCount = tile.bounds.rowEnd - tile.bounds.rowStart + 1
  const colCount = tile.bounds.colEnd - tile.bounds.colStart + 1
  const cellCount = rowCount * colCount
  if (rowCount <= 0 || colCount <= 0 || tile.rectCount !== cellCount) {
    return [{ offset: 0, length: tile.rectCount }]
  }

  const spans: GridRenderTileDirtySpan[] = []
  for (let index = 0; index < dirtyMasks.length; index += 1) {
    const mask = dirtyMasks[index] ?? 0
    if ((mask & RECT_DIRTY_MASK_V3) === 0) {
      continue
    }
    const rowOffset = index * 2
    const colOffset = index * 2
    const rowStart = clampLocalIndex(dirtyLocalRows[rowOffset] ?? 0, rowCount)
    const rowEnd = clampLocalIndex(dirtyLocalRows[rowOffset + 1] ?? rowStart, rowCount)
    const colStart = clampLocalIndex(dirtyLocalCols[colOffset] ?? 0, colCount)
    const colEnd = clampLocalIndex(dirtyLocalCols[colOffset + 1] ?? colStart, colCount)
    if (rowEnd < rowStart || colEnd < colStart) {
      continue
    }
    for (let row = rowStart; row <= rowEnd; row += 1) {
      const offset = row * colCount + colStart
      spans.push({ offset, length: colEnd - colStart + 1 })
    }
  }

  return mergeDirtySpans(spans)
}

function resolveGridRenderTileTextDirtySpansV3(
  tile: GridRenderTile,
  dirtyLocalRows: Uint32Array,
  dirtyLocalCols: Uint32Array,
  dirtyMasks: Uint32Array,
): readonly GridRenderTileDirtySpan[] {
  if (tile.textRuns.length !== tile.textCount) {
    return [{ offset: 0, length: tile.textCount }]
  }
  if (tile.textRuns.some((run) => run.row === undefined || run.col === undefined)) {
    return [{ offset: 0, length: tile.textCount }]
  }

  const spans: GridRenderTileDirtySpan[] = []
  for (let runIndex = 0; runIndex < tile.textRuns.length; runIndex += 1) {
    const run = tile.textRuns[runIndex]
    const localRow = (run?.row ?? tile.bounds.rowStart) - tile.bounds.rowStart
    const localCol = (run?.col ?? tile.bounds.colStart) - tile.bounds.colStart
    if (isTextRunDirty(localRow, localCol, dirtyLocalRows, dirtyLocalCols, dirtyMasks)) {
      spans.push({ offset: runIndex, length: 1 })
    }
  }

  return spans.length > 0 ? mergeDirtySpans(spans) : [{ offset: 0, length: tile.textCount }]
}

function isTextRunDirty(
  localRow: number,
  localCol: number,
  dirtyLocalRows: Uint32Array,
  dirtyLocalCols: Uint32Array,
  dirtyMasks: Uint32Array,
): boolean {
  for (let index = 0; index < dirtyMasks.length; index += 1) {
    const mask = dirtyMasks[index] ?? 0
    if ((mask & TEXT_DIRTY_MASK_V3) === 0) {
      continue
    }
    const offset = index * 2
    const rowStart = dirtyLocalRows[offset] ?? 0
    const rowEnd = dirtyLocalRows[offset + 1] ?? rowStart
    const colStart = dirtyLocalCols[offset] ?? 0
    const colEnd = dirtyLocalCols[offset + 1] ?? colStart
    if (localRow >= rowStart && localRow <= rowEnd && localCol >= colStart && localCol <= colEnd) {
      return true
    }
  }
  return false
}

function clampLocalIndex(value: number, length: number): number {
  return Math.max(0, Math.min(length - 1, Math.floor(value)))
}

function mergeDirtySpans(spans: readonly GridRenderTileDirtySpan[]): readonly GridRenderTileDirtySpan[] {
  if (spans.length <= 1) {
    return spans
  }
  const sorted = [...spans].toSorted((left, right) => left.offset - right.offset || left.length - right.length)
  const merged: GridRenderTileDirtySpan[] = []
  for (const span of sorted) {
    const last = merged[merged.length - 1]
    if (!last) {
      merged.push(span)
      continue
    }
    const lastEnd = last.offset + last.length
    if (span.offset <= lastEnd) {
      merged[merged.length - 1] = {
        offset: last.offset,
        length: Math.max(lastEnd, span.offset + span.length) - last.offset,
      }
      continue
    }
    merged.push(span)
  }
  return merged
}
