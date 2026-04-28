import type { GridTextItem } from '../gridTextScene.js'
import type { Rectangle } from '../gridTypes.js'
import { resolveCellTextLayoutV2 } from '../text/gridTextLayoutV2.js'
import { fontKeyToCssFont, type GridTextMetricsProvider } from '../text/gridTextMetrics.js'
import { CELL_TEXT_PADDING_X, CELL_TEXT_PADDING_Y, type FontKey, type ResolvedCellTextLayout } from '../text/gridTextPacket.js'
import {
  DEFAULT_TEXT_COLOR,
  DEFAULT_TEXT_FONT,
  DEFAULT_TEXT_HEIGHT,
  parseTextCssColor,
  parseTextFontSize,
  type GlyphAtlasLike,
} from './line-text-layout.js'

export interface TextQuadRun {
  readonly text: string
  readonly x: number
  readonly y: number
  readonly width?: number
  readonly height?: number
  readonly clipX?: number
  readonly clipY?: number
  readonly clipWidth?: number
  readonly clipHeight?: number
  readonly align?: 'left' | 'center' | 'right' | undefined
  readonly wrap?: boolean | undefined
  readonly font?: string
  readonly fontSize?: number
  readonly color?: string
  readonly underline?: boolean
  readonly strike?: boolean
}

export interface TextQuad {
  readonly atlasKey: string
  readonly glyphId: number
  readonly pageId: number
  readonly glyph: string
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly u0: number
  readonly v0: number
  readonly u1: number
  readonly v1: number
  readonly color: string
  readonly clipX: number
  readonly clipY: number
  readonly clipWidth: number
  readonly clipHeight: number
}

export interface TextDecorationRect {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly color: string
}

export interface TextQuadRunSpan {
  readonly offset: number
  readonly length: number
}

export interface TextQuadRunPayloadV3 {
  readonly signature: string
  readonly atlasVersion: number
  readonly floats: Float32Array
  readonly glyphIds: readonly number[]
  readonly pageIds: readonly number[]
  readonly quadCount: number
}

export interface BuildTextQuadsFromRunsWithSpansOptionsV3 {
  readonly previousRunPayloads?: readonly TextQuadRunPayloadV3[] | null | undefined
  readonly dirtyRunSpans?: readonly TextQuadRunSpan[] | undefined
}

const TEXT_INSTANCE_FLOAT_COUNT = 16

export function buildTextQuads(runs: readonly TextQuadRun[], atlas: GlyphAtlasLike): TextQuad[] {
  const quads: TextQuad[] = []
  for (const run of runs) {
    if (run.text.length === 0) {
      continue
    }

    const fontKey = resolveRunFontKey(run)
    const font = fontKeyToCssFont(fontKey)
    const layout = resolveRunTextLayout(run, atlas, fontKey)
    const clipRect = resolveRunClipRect(run, layout.textClipWorldRect)
    if (clipRect.width <= 0 || clipRect.height <= 0) {
      continue
    }

    for (const line of layout.lines) {
      if (line.text.length === 0) {
        continue
      }

      for (const glyph of line.glyphs) {
        const entry = atlas.intern(font, glyph.glyph)
        quads.push({
          atlasKey: entry.key,
          glyph: glyph.glyph,
          glyphId: entry.glyphId,
          pageId: entry.pageId,
          x: glyph.worldX - entry.originOffsetX,
          y: line.baselineWorldY - entry.baseline,
          width: entry.width,
          height: entry.height,
          u0: entry.u0,
          v0: entry.v0,
          u1: entry.u1,
          v1: entry.v1,
          color: run.color ?? DEFAULT_TEXT_COLOR,
          clipHeight: clipRect.height,
          clipWidth: clipRect.width,
          clipX: clipRect.x,
          clipY: clipRect.y,
        })
      }
    }
  }

  return quads
}

function packTextQuads(
  quads: readonly TextQuad[],
  targetBuffer?: Float32Array,
): { floats: Float32Array; glyphIds: readonly number[]; pageIds: readonly number[]; quadCount: number } {
  const floats =
    targetBuffer && targetBuffer.length >= quads.length * TEXT_INSTANCE_FLOAT_COUNT
      ? targetBuffer
      : new Float32Array(Math.max(1, quads.length) * TEXT_INSTANCE_FLOAT_COUNT)

  const glyphIds: number[] = []
  const pageIds: number[] = []
  quads.forEach((quad, index) => {
    const base = index * TEXT_INSTANCE_FLOAT_COUNT
    const [r, g, b, a] = parseTextCssColor(quad.color)
    floats[base + 0] = quad.x
    floats[base + 1] = quad.y
    floats[base + 2] = quad.width
    floats[base + 3] = quad.height
    floats[base + 4] = quad.u0
    floats[base + 5] = quad.v0
    floats[base + 6] = quad.u1
    floats[base + 7] = quad.v1
    floats[base + 8] = r
    floats[base + 9] = g
    floats[base + 10] = b
    floats[base + 11] = a
    floats[base + 12] = quad.clipX
    floats[base + 13] = quad.clipY
    floats[base + 14] = quad.clipX + quad.clipWidth
    floats[base + 15] = quad.clipY + quad.clipHeight
    glyphIds.push(quad.glyphId)
    pageIds.push(quad.pageId)
  })

  return { floats, glyphIds, pageIds, quadCount: quads.length }
}

export function buildTextQuadsFromScene(
  items: readonly GridTextItem[],
  atlas: GlyphAtlasLike,
  targetBuffer?: Float32Array,
): { floats: Float32Array; glyphIds: readonly number[]; pageIds: readonly number[]; quadCount: number } {
  const quads = buildTextQuads(
    items.map((item) => ({
      text: item.text,
      x: item.x,
      y: item.y,
      width: item.width,
      height: item.height,
      clipX: item.x + item.clipInsetLeft,
      clipY: item.y + item.clipInsetTop,
      clipWidth: Math.max(0, item.width - item.clipInsetLeft - item.clipInsetRight),
      clipHeight: Math.max(0, item.height - item.clipInsetTop - item.clipInsetBottom),
      align: item.align,
      wrap: item.wrap,
      font: item.font,
      fontSize: item.fontSize,
      color: item.color,
      underline: item.underline,
      strike: item.strike,
    })),
    atlas,
  )

  return packTextQuads(quads, targetBuffer)
}

export function buildTextQuadsFromRuns(
  runs: readonly TextQuadRun[],
  atlas: GlyphAtlasLike,
  targetBuffer?: Float32Array,
): { floats: Float32Array; glyphIds: readonly number[]; pageIds: readonly number[]; quadCount: number } {
  return packTextQuads(buildTextQuads(runs, atlas), targetBuffer)
}

export function buildTextQuadsFromRunsWithSpans(
  runs: readonly TextQuadRun[],
  atlas: GlyphAtlasLike,
  targetBuffer?: Float32Array,
  options: BuildTextQuadsFromRunsWithSpansOptionsV3 = {},
): {
  floats: Float32Array
  glyphIds: readonly number[]
  pageIds: readonly number[]
  quadCount: number
  runGlyphIds: readonly (readonly number[])[]
  runPayloads: readonly TextQuadRunPayloadV3[]
  runSpans: readonly TextQuadRunSpan[]
} {
  return buildTextQuadsFromRunsWithSpansInternal(runs, atlas, targetBuffer, options, true)
}

function buildTextQuadsFromRunsWithSpansInternal(
  runs: readonly TextQuadRun[],
  atlas: GlyphAtlasLike,
  targetBuffer: Float32Array | undefined,
  options: BuildTextQuadsFromRunsWithSpansOptionsV3,
  retryOnAtlasVersionChange: boolean,
): {
  floats: Float32Array
  glyphIds: readonly number[]
  pageIds: readonly number[]
  quadCount: number
  runGlyphIds: readonly (readonly number[])[]
  runPayloads: readonly TextQuadRunPayloadV3[]
  runSpans: readonly TextQuadRunSpan[]
} {
  const previousRunPayloads = options.previousRunPayloads ?? []
  const runPayloads: TextQuadRunPayloadV3[] = []
  const initialAtlasVersion = resolveAtlasVersion(atlas)
  let reusedPreviousPayload = false
  let quadCount = 0
  for (let index = 0; index < runs.length; index += 1) {
    const run = runs[index]!
    const signature = resolveTextQuadRunSignatureV3(run)
    const previousPayload = previousRunPayloads[index]
    if (
      previousPayload &&
      previousPayload.signature === signature &&
      previousPayload.atlasVersion === initialAtlasVersion &&
      !isTextRunDirty(index, options.dirtyRunSpans)
    ) {
      runPayloads.push(previousPayload)
      reusedPreviousPayload = true
      quadCount += previousPayload.quadCount
      continue
    }
    const runQuads = buildTextQuads([run], atlas)
    const packed = packTextQuads(runQuads)
    const payload: TextQuadRunPayloadV3 = {
      atlasVersion: resolveAtlasVersion(atlas),
      floats: packed.floats,
      glyphIds: packed.glyphIds,
      pageIds: packed.pageIds,
      quadCount: packed.quadCount,
      signature,
    }
    runPayloads.push(payload)
    quadCount += payload.quadCount
  }

  const finalAtlasVersion = resolveAtlasVersion(atlas)
  if (retryOnAtlasVersionChange && finalAtlasVersion !== initialAtlasVersion) {
    return buildTextQuadsFromRunsWithSpansInternal(
      runs,
      atlas,
      targetBuffer,
      {
        dirtyRunSpans: options.dirtyRunSpans,
        previousRunPayloads: reusedPreviousPayload ? [] : runPayloads,
      },
      false,
    )
  }

  const floats =
    targetBuffer && targetBuffer.length >= quadCount * TEXT_INSTANCE_FLOAT_COUNT
      ? targetBuffer
      : new Float32Array(Math.max(1, quadCount) * TEXT_INSTANCE_FLOAT_COUNT)
  const glyphIds: number[] = []
  const pageIds: number[] = []
  const runGlyphIds: number[][] = []
  const runSpans: TextQuadRunSpan[] = []

  let offset = 0
  for (const payload of runPayloads) {
    const floatCount = payload.quadCount * TEXT_INSTANCE_FLOAT_COUNT
    floats.set(payload.floats.subarray(0, floatCount), offset * TEXT_INSTANCE_FLOAT_COUNT)
    glyphIds.push(...payload.glyphIds)
    pageIds.push(...payload.pageIds)
    runGlyphIds.push([...payload.glyphIds])
    runSpans.push({ offset, length: payload.quadCount })
    offset += payload.quadCount
  }

  return {
    floats,
    glyphIds,
    pageIds,
    quadCount,
    runGlyphIds,
    runPayloads,
    runSpans,
  }
}

export function resolveTextQuadRunSignatureV3(run: TextQuadRun): string {
  return [
    'text-run-v3',
    run.text,
    run.x,
    run.y,
    run.width ?? '',
    run.height ?? '',
    run.clipX ?? '',
    run.clipY ?? '',
    run.clipWidth ?? '',
    run.clipHeight ?? '',
    run.align ?? '',
    run.wrap === true ? 1 : 0,
    run.font ?? '',
    run.fontSize ?? '',
    run.color ?? '',
    run.underline === true ? 1 : 0,
    run.strike === true ? 1 : 0,
  ].join('\u0001')
}

function isTextRunDirty(index: number, dirtyRunSpans: readonly TextQuadRunSpan[] | undefined): boolean {
  if (!dirtyRunSpans || dirtyRunSpans.length === 0) {
    return false
  }
  return dirtyRunSpans.some((span) => index >= span.offset && index < span.offset + span.length)
}

function resolveAtlasVersion(atlas: GlyphAtlasLike): number {
  return atlas.getVersion?.() ?? 0
}

export function buildTextDecorationRects(runs: readonly TextQuadRun[], atlas: GlyphAtlasLike): TextDecorationRect[] {
  const rects: TextDecorationRect[] = []
  for (const run of runs) {
    if (!run.underline && !run.strike) {
      continue
    }
    const layout = resolveRunTextLayout(run, atlas, resolveRunFontKey(run))
    const clipRect = resolveRunClipRect(run, layout.textClipWorldRect)
    const clipRight = clipRect.x + clipRect.width
    const clipBottom = clipRect.y + clipRect.height
    for (const decoration of layout.decorations) {
      const left = Math.max(decoration.x, clipRect.x)
      const right = Math.min(decoration.x + decoration.width, clipRight)
      const visibleWidth = right - left
      if (visibleWidth <= 0 || decoration.y < clipRect.y || decoration.y > clipBottom) {
        continue
      }
      rects.push({
        x: left,
        y: decoration.y,
        width: visibleWidth,
        height: decoration.height,
        color: run.color ?? DEFAULT_TEXT_COLOR,
      })
    }
  }

  return rects
}

export function buildTextDecorationRectsFromScene(items: readonly GridTextItem[], atlas: GlyphAtlasLike): TextDecorationRect[] {
  return buildTextDecorationRects(
    items.map((item) => ({
      text: item.text,
      x: item.x,
      y: item.y,
      width: item.width,
      height: item.height,
      clipX: item.x + item.clipInsetLeft,
      clipY: item.y + item.clipInsetTop,
      clipWidth: Math.max(0, item.width - item.clipInsetLeft - item.clipInsetRight),
      clipHeight: Math.max(0, item.height - item.clipInsetTop - item.clipInsetBottom),
      align: item.align,
      wrap: item.wrap,
      font: item.font,
      fontSize: item.fontSize,
      color: item.color,
      underline: item.underline,
      strike: item.strike,
    })),
    atlas,
  )
}

export function buildTextDecorationRectsFromRuns(runs: readonly TextQuadRun[], atlas: GlyphAtlasLike): TextDecorationRect[] {
  return buildTextDecorationRects(runs, atlas)
}

function resolveRunTextLayout(run: TextQuadRun, atlas: GlyphAtlasLike, fontKey: FontKey): ResolvedCellTextLayout {
  return resolveCellTextLayoutV2({
    cell: { col: -1, row: -1 },
    cellWorldRect: resolveRunCellWorldRect(run, atlas),
    color: run.color,
    fontKey,
    horizontalAlign: run.align,
    metrics: createAtlasTextMetricsProvider(atlas),
    overflow: run.wrap ? 'clip' : run.align === 'left' || run.align === undefined ? 'overflow' : 'clip',
    strike: run.strike,
    text: run.text,
    underline: run.underline,
    verticalAlign: run.wrap ? 'top' : 'middle',
    wrap: run.wrap,
  })
}

function resolveRunCellWorldRect(run: TextQuadRun, atlas: GlyphAtlasLike): Rectangle {
  const font = run.font ?? DEFAULT_TEXT_FONT
  const measuredAdvance = run.text.length === 0 ? 0 : measureAtlasText(font, run.text, atlas).advance
  return {
    x: run.x,
    y: run.y,
    width: Math.max(0, run.width ?? run.clipWidth ?? measuredAdvance + CELL_TEXT_PADDING_X * 2),
    height: Math.max(0, run.height ?? run.clipHeight ?? DEFAULT_TEXT_HEIGHT + CELL_TEXT_PADDING_Y * 2),
  }
}

function resolveRunClipRect(run: TextQuadRun, layoutClip: Rectangle): Rectangle {
  const explicitClip =
    run.clipX === undefined || run.clipY === undefined || run.clipWidth === undefined || run.clipHeight === undefined
      ? null
      : {
          x: run.clipX,
          y: run.clipY,
          width: Math.max(0, run.clipWidth),
          height: Math.max(0, run.clipHeight),
        }
  if (!explicitClip) {
    return layoutClip
  }
  const x0 = Math.max(layoutClip.x, explicitClip.x)
  const y0 = Math.max(layoutClip.y, explicitClip.y)
  const x1 = Math.min(layoutClip.x + layoutClip.width, explicitClip.x + explicitClip.width)
  const y1 = Math.min(layoutClip.y + layoutClip.height, explicitClip.y + explicitClip.height)
  return {
    x: x0,
    y: y0,
    width: Math.max(0, x1 - x0),
    height: Math.max(0, y1 - y0),
  }
}

function resolveRunFontKey(run: TextQuadRun): FontKey {
  const font = run.font ?? DEFAULT_TEXT_FONT
  const sizeCssPx = Math.max(1, run.fontSize ?? parseTextFontSize(font))
  const style = /\bitalic\b/i.test(font) ? 'italic' : 'normal'
  const weightMatch = font.match(/\b([1-9]00|normal|bold)\b/i)
  const rawWeight = weightMatch?.[1]?.toLowerCase()
  const weight = rawWeight === 'bold' ? 700 : rawWeight === 'normal' || rawWeight === undefined ? 400 : Number(rawWeight)
  const sizeMatch = font.match(/\b\d+(?:\.\d+)?px\s+(.+)$/i)
  const family = sizeMatch?.[1]?.trim() || 'sans-serif'
  return {
    dprBucket: 1,
    family,
    fontEpoch: 0,
    sizeCssPx,
    style,
    weight,
  }
}

function createAtlasTextMetricsProvider(atlas: GlyphAtlasLike): GridTextMetricsProvider {
  return {
    measure(text, fontKey) {
      const font = fontKeyToCssFont(fontKey)
      if (text.length === 0) {
        const sample = measureAtlasText(font, 'Mg', atlas)
        return {
          advance: 0,
          ascent: sample.ascent,
          descent: sample.descent,
          lineHeight: Math.max(fontKey.sizeCssPx * 1.2, sample.ascent + sample.descent),
        }
      }
      const measured = measureAtlasText(font, text, atlas)
      return {
        advance: measured.advance,
        ascent: measured.ascent,
        descent: measured.descent,
        lineHeight: Math.max(fontKey.sizeCssPx * 1.2, measured.ascent + measured.descent),
      }
    },
  }
}

function measureAtlasText(
  font: string,
  text: string,
  atlas: GlyphAtlasLike,
): {
  readonly advance: number
  readonly ascent: number
  readonly descent: number
} {
  let advance = 0
  let ascent = 1
  let descent = 1
  for (const glyph of splitGraphemes(text)) {
    const entry = atlas.intern(font, glyph)
    advance += entry.advance
    ascent = Math.max(ascent, entry.baseline)
    descent = Math.max(descent, entry.height - entry.baseline)
  }
  return { advance, ascent, descent }
}

function splitGraphemes(text: string): readonly string[] {
  if (typeof Intl !== 'undefined' && 'Segmenter' in Intl) {
    const segmenter = new Intl.Segmenter(undefined, { granularity: 'grapheme' })
    return [...segmenter.segment(text)].map((segment) => segment.segment)
  }
  return Array.from(text)
}
