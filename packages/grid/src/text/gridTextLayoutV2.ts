import { parseGpuColor } from '../gridGpuScene.js'
import type { Rectangle } from '../gridTypes.js'
import { createFallbackTextMetricsProvider, type GridTextMetricsProvider } from './gridTextMetrics.js'
import { segmentGraphemes, wrapTextByWidth } from './gridTextOverflow.js'
import {
  CELL_TEXT_PADDING_X,
  CELL_TEXT_PADDING_Y,
  type FontKey,
  type ResolvedCellTextLayout,
  type ResolvedGlyphPlacement,
  type ResolvedTextLine,
  type TextHorizontalAlign,
  type TextOverflowMode,
  type TextVerticalAlign,
} from './gridTextPacket.js'

export function resolveCellTextLayoutV2(input: {
  readonly cell: { readonly col: number; readonly row: number }
  readonly text: string
  readonly displayText?: string | undefined
  readonly fontKey: FontKey
  readonly color?: string | undefined
  readonly horizontalAlign?: TextHorizontalAlign | undefined
  readonly verticalAlign?: TextVerticalAlign | undefined
  readonly wrap?: boolean | undefined
  readonly overflow?: TextOverflowMode | undefined
  readonly underline?: boolean | undefined
  readonly strike?: boolean | undefined
  readonly cellWorldRect: Rectangle
  readonly generation?: number | undefined
  readonly metrics?: GridTextMetricsProvider | undefined
}): ResolvedCellTextLayout {
  const metrics = input.metrics ?? createFallbackTextMetricsProvider()
  const displayText = input.displayText ?? input.text
  const wrap = input.wrap ?? false
  const horizontalAlign = input.horizontalAlign ?? inferHorizontalAlign(displayText)
  const verticalAlign = input.verticalAlign ?? (wrap ? 'top' : 'middle')
  const overflow = input.overflow ?? (wrap || horizontalAlign !== 'left' ? 'clip' : 'overflow')
  const textClipWorldRect = insetRect(input.cellWorldRect, CELL_TEXT_PADDING_X, CELL_TEXT_PADDING_Y)
  const contentWidth = Math.max(0, textClipWorldRect.width)
  const rawLines = wrap ? wrapTextByWidth({ fontKey: input.fontKey, maxWidth: contentWidth, metrics, text: displayText }) : [displayText]
  const baseMetrics = metrics.measure('Mg', input.fontKey)
  const maxLineCount = Math.max(1, Math.floor(textClipWorldRect.height / baseMetrics.lineHeight))
  const visibleLines = rawLines.slice(0, maxLineCount)
  const contentHeight = visibleLines.length * baseMetrics.lineHeight
  const firstBaselineY =
    verticalAlign === 'top'
      ? textClipWorldRect.y + baseMetrics.ascent
      : verticalAlign === 'bottom'
        ? textClipWorldRect.y + textClipWorldRect.height - contentHeight + baseMetrics.ascent
        : textClipWorldRect.y + (textClipWorldRect.height - contentHeight) / 2 + baseMetrics.ascent
  const lines = visibleLines.map((line, index) =>
    resolveLine({
      baselineWorldY: firstBaselineY + index * baseMetrics.lineHeight,
      clip: textClipWorldRect,
      fontKey: input.fontKey,
      horizontalAlign,
      line,
      metrics,
    }),
  )
  return {
    cell: input.cell,
    cellWorldRect: input.cellWorldRect,
    color: parseGpuColor(input.color),
    decorations: buildDecorationRects({
      color: parseGpuColor(input.color),
      lines,
      strike: input.strike ?? false,
      underline: input.underline ?? false,
    }),
    displayText,
    fontKey: input.fontKey,
    generation: input.generation ?? 0,
    horizontalAlign,
    lines,
    overflow,
    overflowWorldRect:
      overflow === 'overflow'
        ? rect(textClipWorldRect.x, textClipWorldRect.y, Math.max(textClipWorldRect.width, maxLineAdvance(lines)), textClipWorldRect.height)
        : textClipWorldRect,
    text: input.text,
    textClipWorldRect,
    verticalAlign,
    wrap,
  }
}

function resolveLine(input: {
  readonly line: string
  readonly fontKey: FontKey
  readonly clip: Rectangle
  readonly horizontalAlign: TextHorizontalAlign
  readonly baselineWorldY: number
  readonly metrics: GridTextMetricsProvider
}): ResolvedTextLine {
  const measured = input.metrics.measure(input.line, input.fontKey)
  const worldX =
    input.horizontalAlign === 'right'
      ? input.clip.x + input.clip.width - measured.advance
      : input.horizontalAlign === 'center'
        ? input.clip.x + (input.clip.width - measured.advance) / 2
        : input.clip.x
  const glyphs = buildGlyphPlacements({
    baselineWorldY: input.baselineWorldY,
    fontKey: input.fontKey,
    line: input.line,
    metrics: input.metrics,
    worldX,
  })
  return {
    advance: measured.advance,
    ascent: measured.ascent,
    baselineWorldY: input.baselineWorldY,
    descent: measured.descent,
    glyphs,
    text: input.line,
    worldX,
  }
}

function buildGlyphPlacements(input: {
  readonly line: string
  readonly fontKey: FontKey
  readonly worldX: number
  readonly baselineWorldY: number
  readonly metrics: GridTextMetricsProvider
}): readonly ResolvedGlyphPlacement[] {
  const glyphs: ResolvedGlyphPlacement[] = []
  let cursor = input.worldX
  for (const glyph of segmentGraphemes(input.line)) {
    const measured = input.metrics.measure(glyph, input.fontKey)
    glyphs.push({
      advance: measured.advance,
      atlasGlyphKey: `${input.fontKey.family}:${input.fontKey.weight}:${input.fontKey.style}:${input.fontKey.sizeCssPx}:${input.fontKey.dprBucket}:${glyph}`,
      glyphId: glyph.codePointAt(0) ?? 0,
      height: measured.ascent + measured.descent,
      uvPadding: 1,
      width: measured.advance,
      worldX: cursor,
      worldY: input.baselineWorldY - measured.ascent,
    })
    cursor += measured.advance
  }
  return glyphs
}

function inferHorizontalAlign(text: string): TextHorizontalAlign {
  if (/^-?\(?[$€£¥]?\d/u.test(text.trim())) {
    return 'right'
  }
  if (/^(TRUE|FALSE)$/iu.test(text.trim())) {
    return 'center'
  }
  return 'left'
}

function buildDecorationRects(input: {
  readonly lines: readonly ResolvedTextLine[]
  readonly color: ReturnType<typeof parseGpuColor>
  readonly underline: boolean
  readonly strike: boolean
}) {
  return input.lines.flatMap((line) => {
    const rects = []
    if (input.underline) {
      rects.push({ color: input.color, height: 1, width: line.advance, x: line.worldX, y: line.baselineWorldY + line.descent * 0.35 })
    }
    if (input.strike) {
      rects.push({ color: input.color, height: 1, width: line.advance, x: line.worldX, y: line.baselineWorldY - line.ascent * 0.35 })
    }
    return rects
  })
}

function insetRect(rectangle: Rectangle, x: number, y: number): Rectangle {
  return rect(rectangle.x + x, rectangle.y + y, Math.max(0, rectangle.width - x * 2), Math.max(0, rectangle.height - y * 2))
}

function maxLineAdvance(lines: readonly ResolvedTextLine[]): number {
  return lines.reduce((max, line) => Math.max(max, line.advance), 0)
}

function rect(x: number, y: number, width: number, height: number): Rectangle {
  return { height: Math.max(0, height), width: Math.max(0, width), x, y }
}
