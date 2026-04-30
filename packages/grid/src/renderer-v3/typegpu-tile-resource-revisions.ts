import type { TextDecorationRect, TextQuadRunSpan } from './line-text-quad-buffer.js'
import type { GridRenderTile } from './render-tile-source.js'
import { DirtyMaskV3 } from './tile-damage-index.js'
import type { createGlyphAtlas } from './typegpu-atlas-manager.js'
import type { TextInstanceVertexBuffer } from './typegpu-primitives.js'
import type { GpuBufferHandleV3 } from './gpu-buffer-arena.js'

const RECT_DIRTY_MASK_V3 =
  DirtyMaskV3.Style | DirtyMaskV3.Rect | DirtyMaskV3.Border | DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Freeze
const TEXT_DIRTY_MASK_V3 =
  DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Freeze
const TEXT_DECORATION_DIRTY_MASK_V3 = DirtyMaskV3.Value | DirtyMaskV3.Text

export interface TypeGpuTileTextRevisionKeyV3 {
  readonly tileId: number
  readonly textRunCount: number
  readonly valueSeq: number
  readonly styleSeq: number
  readonly textSeq: number
  readonly axisSeqX: number
  readonly axisSeqY: number
  readonly freezeSeq: number
  readonly batchSeq: number
}

export interface TypeGpuTileRectRevisionKeyV3 {
  readonly tileId: number
  readonly rectCount: number
  readonly valueSeq: number
  readonly styleSeq: number
  readonly axisSeqX: number
  readonly axisSeqY: number
  readonly freezeSeq: number
  readonly batchSeq: number
  readonly decorationRectCount: number
}

export function resolveGridTextTileRevisionKeyV3(tile: GridRenderTile): TypeGpuTileTextRevisionKeyV3 {
  return {
    axisSeqX: tile.version.axisX,
    axisSeqY: tile.version.axisY,
    batchSeq: tile.lastBatchId,
    freezeSeq: tile.version.freeze,
    styleSeq: tile.version.styles,
    textRunCount: tile.textCount,
    textSeq: tile.version.text,
    tileId: tile.tileId,
    valueSeq: tile.version.values,
  }
}

export function resolveGridRectTileRevisionKeyV3(input: {
  readonly tile: GridRenderTile
  readonly decorationRects?: readonly TextDecorationRect[] | undefined
}): TypeGpuTileRectRevisionKeyV3 {
  const decorationRects = input.decorationRects ?? []
  return {
    axisSeqX: input.tile.version.axisX,
    axisSeqY: input.tile.version.axisY,
    batchSeq: input.tile.lastBatchId,
    decorationRectCount: decorationRects.length,
    freezeSeq: input.tile.version.freeze,
    rectCount: input.tile.rectCount,
    styleSeq: input.tile.version.styles,
    tileId: input.tile.tileId,
    valueSeq: input.tile.version.values,
  }
}

export function areGridTextTileRevisionKeysEqualV3(
  left: TypeGpuTileTextRevisionKeyV3 | null | undefined,
  right: TypeGpuTileTextRevisionKeyV3 | null | undefined,
): boolean {
  return (
    left !== null &&
    left !== undefined &&
    right !== null &&
    right !== undefined &&
    left.tileId === right.tileId &&
    left.textRunCount === right.textRunCount &&
    left.valueSeq === right.valueSeq &&
    left.styleSeq === right.styleSeq &&
    left.textSeq === right.textSeq &&
    left.axisSeqX === right.axisSeqX &&
    left.axisSeqY === right.axisSeqY &&
    left.freezeSeq === right.freezeSeq &&
    left.batchSeq === right.batchSeq
  )
}

export function areGridRectTileRevisionKeysEqualV3(
  left: TypeGpuTileRectRevisionKeyV3 | null | undefined,
  right: TypeGpuTileRectRevisionKeyV3 | null | undefined,
): boolean {
  return (
    left !== null &&
    left !== undefined &&
    right !== null &&
    right !== undefined &&
    left.tileId === right.tileId &&
    left.rectCount === right.rectCount &&
    left.valueSeq === right.valueSeq &&
    left.styleSeq === right.styleSeq &&
    left.axisSeqX === right.axisSeqX &&
    left.axisSeqY === right.axisSeqY &&
    left.freezeSeq === right.freezeSeq &&
    left.batchSeq === right.batchSeq &&
    left.decorationRectCount === right.decorationRectCount
  )
}

export function resolveGridTileDirtyContentMaskV3(tile: Pick<GridRenderTile, 'dirtyMasks'>): number | null {
  const masks = tile.dirtyMasks
  if (!masks || masks.length === 0) {
    return null
  }
  let mask = 0
  for (const value of masks) {
    mask |= value
  }
  return mask
}

export function shouldSyncGridTextTileResourceV3(input: {
  readonly atlasGeometryVersion?: number | undefined
  readonly content: {
    readonly textAtlasGeometryVersion: number
    readonly textCount: number
    readonly textHandle: GpuBufferHandleV3<TextInstanceVertexBuffer> | null
    readonly textRunCount: number
    readonly textRevisionKey: TypeGpuTileTextRevisionKeyV3 | null
  }
  readonly missingGlyphDependencies?: boolean | undefined
  readonly textRevisionKey: TypeGpuTileTextRevisionKeyV3
  readonly tile: GridRenderTile
}): boolean {
  if (input.missingGlyphDependencies) {
    return true
  }
  if (areGridTextTileRevisionKeysEqualV3(input.content.textRevisionKey, input.textRevisionKey)) {
    return input.atlasGeometryVersion !== undefined && input.content.textAtlasGeometryVersion !== input.atlasGeometryVersion
  }
  if (!input.content.textRevisionKey) {
    return true
  }
  if (input.content.textRunCount !== input.tile.textCount) {
    return true
  }
  if (input.tile.textCount > 0 && !input.content.textHandle) {
    return true
  }
  const dirtyMask = resolveGridTileDirtyContentMaskV3(input.tile)
  return dirtyMask === null || (dirtyMask & TEXT_DIRTY_MASK_V3) !== 0
}

export function resolveMissingTextGlyphRunSpansV3(input: {
  readonly atlas: Pick<ReturnType<typeof createGlyphAtlas>, 'resolveGlyphRecord'>
  readonly content: {
    readonly textGlyphIds: readonly number[] | null
    readonly textGlyphPageIds: readonly number[] | null
    readonly textRunCount: number
    readonly textRunGlyphIds: readonly (readonly number[])[] | null
  }
}): readonly TextQuadRunSpan[] {
  const runCount = input.content.textRunCount
  if (runCount <= 0) {
    return []
  }

  const glyphIds = input.content.textGlyphIds
  const pageIds = input.content.textGlyphPageIds
  const runGlyphIds = input.content.textRunGlyphIds
  if (!glyphIds || glyphIds.length === 0) {
    return []
  }
  if (!pageIds || pageIds.length !== glyphIds.length || !runGlyphIds || runGlyphIds.length !== runCount) {
    return [{ length: runCount, offset: 0 }]
  }

  const expectedPageByGlyph = new Map<number, number>()
  for (let index = 0; index < glyphIds.length; index += 1) {
    const glyphId = glyphIds[index]
    const pageId = pageIds[index]
    if (glyphId === undefined || pageId === undefined) {
      return [{ length: runCount, offset: 0 }]
    }
    expectedPageByGlyph.set(glyphId, pageId)
  }

  const missing: TextQuadRunSpan[] = []
  for (let runIndex = 0; runIndex < runCount; runIndex += 1) {
    const runGlyphs = runGlyphIds[runIndex] ?? []
    let runMissing = false
    for (const glyphId of runGlyphs) {
      const expectedPageId = expectedPageByGlyph.get(glyphId)
      const glyphRecord = input.atlas.resolveGlyphRecord(glyphId)
      if (expectedPageId === undefined || !glyphRecord || glyphRecord.pageId !== expectedPageId) {
        runMissing = true
        break
      }
    }
    if (runMissing) {
      missing.push({ length: 1, offset: runIndex })
    }
  }
  return missing
}

export function shouldSyncGridRectTileResourceV3(input: {
  readonly content: {
    readonly decorationRects: readonly TextDecorationRect[] | null
    readonly rectCount: number
    readonly rectHandle: GpuBufferHandleV3 | null
    readonly rectRevisionKey: TypeGpuTileRectRevisionKeyV3 | null
  }
  readonly rectRevisionKey: TypeGpuTileRectRevisionKeyV3
  readonly tile: GridRenderTile
}): boolean {
  if (areGridRectTileRevisionKeysEqualV3(input.content.rectRevisionKey, input.rectRevisionKey)) {
    return false
  }
  if (!input.content.rectRevisionKey) {
    return true
  }
  if (input.content.rectCount !== input.tile.rectCount) {
    return true
  }
  if (input.tile.rectCount > 0 && !input.content.rectHandle) {
    return true
  }
  const dirtyMask = resolveGridTileDirtyContentMaskV3(input.tile)
  if (dirtyMask === null) {
    return true
  }
  if ((dirtyMask & RECT_DIRTY_MASK_V3) !== 0) {
    return true
  }
  if ((dirtyMask & TEXT_DECORATION_DIRTY_MASK_V3) === 0) {
    return false
  }
  return input.tile.textRuns.some((run) => run.underline || run.strike) || (input.content.decorationRects?.length ?? 0) > 0
}
