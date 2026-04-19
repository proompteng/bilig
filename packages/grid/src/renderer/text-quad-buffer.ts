import type { GridTextItem } from '../gridTextScene.js'
import type { GlyphAtlasEntry } from './glyph-atlas.js'

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
  readonly align?: 'left' | 'center' | 'right'
  readonly wrap?: boolean
  readonly font?: string
  readonly fontSize?: number
  readonly color?: string
  readonly underline?: boolean
  readonly strike?: boolean
}

export interface TextQuad {
  readonly atlasKey: string
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

interface GlyphAtlasLike {
  intern(font: string, glyph: string): GlyphAtlasEntry
}

const DEFAULT_FONT = '400 11px sans-serif'
const DEFAULT_COLOR = '#1f2937'
const DEFAULT_HEIGHT = 16
const DEFAULT_WIDTH = Number.POSITIVE_INFINITY
const HORIZONTAL_PADDING = 8
const WRAP_TOP_PADDING = 4
const WRAP_LINE_HEIGHT = 1.2

interface ResolvedLineLayout {
  readonly text: string
  readonly x: number
  readonly y: number
  readonly width: number
}

function resolveLineLayouts(run: TextQuadRun, atlas: GlyphAtlasLike, font: string, fontSize: number): ResolvedLineLayout[] {
  const lineHeight = Math.ceil(fontSize * WRAP_LINE_HEIGHT)
  const availableWidth = Math.max(0, (run.width ?? DEFAULT_WIDTH) - HORIZONTAL_PADDING * 2)
  const lines = wrapWords(run, atlas, font, availableWidth)
  let lineY = run.wrap ? run.y + WRAP_TOP_PADDING : run.y + Math.max(0, ((run.height ?? DEFAULT_HEIGHT) - fontSize) / 2)

  return lines.map((line) => {
    const lineWidth = measureTextWidth(line, font, atlas)
    const startX =
      (run.align ?? 'left') === 'right'
        ? run.x + (run.width ?? lineWidth) - HORIZONTAL_PADDING - lineWidth
        : (run.align ?? 'left') === 'center'
          ? run.x + ((run.width ?? lineWidth) - lineWidth) / 2
          : run.x + HORIZONTAL_PADDING
    const next = {
      text: line,
      x: startX,
      y: lineY,
      width: lineWidth,
    }
    lineY += run.wrap ? lineHeight : 0
    return next
  })
}

function wrapWords(run: TextQuadRun, atlas: GlyphAtlasLike, font: string, maxWidth: number): readonly string[] {
  if (!run.wrap || !Number.isFinite(maxWidth)) {
    return run.text.split('\n')
  }
  const lines: string[] = []
  for (const paragraph of run.text.split('\n')) {
    if (paragraph.length === 0) {
      lines.push('')
      continue
    }
    const words = paragraph.split(/(\s+)/).filter((segment) => segment.length > 0)
    let current = ''
    for (const word of words) {
      const candidate = current.length === 0 ? word : `${current}${word}`
      if (measureTextWidth(candidate, font, atlas) <= maxWidth) {
        current = candidate
        continue
      }
      if (current.length > 0) {
        lines.push(current.trimEnd())
      }
      current = ''
      if (measureTextWidth(word, font, atlas) <= maxWidth) {
        current = word.trimStart()
        continue
      }
      for (const segment of breakWord(word, font, atlas, maxWidth)) {
        lines.push(segment)
      }
    }
    if (current.length > 0) {
      lines.push(current.trimEnd())
    }
  }
  return lines
}

function breakWord(word: string, font: string, atlas: GlyphAtlasLike, maxWidth: number): string[] {
  const segments: string[] = []
  let current = ''
  for (const glyph of word) {
    const candidate = `${current}${glyph}`
    if (current.length > 0 && measureTextWidth(candidate, font, atlas) > maxWidth) {
      segments.push(current)
      current = glyph
      continue
    }
    current = candidate
  }
  if (current.length > 0) {
    segments.push(current)
  }
  return segments
}

function measureTextWidth(text: string, font: string, atlas: GlyphAtlasLike): number {
  let width = 0
  for (const glyph of text) {
    width += atlas.intern(font, glyph).advance
  }
  return width
}

export function buildTextQuads(runs: readonly TextQuadRun[], atlas: GlyphAtlasLike): TextQuad[] {
  const quads: TextQuad[] = []
  for (const run of runs) {
    const font = run.font ?? DEFAULT_FONT
    const fontSize = run.fontSize ?? 11
    const lineLayouts = resolveLineLayouts(run, atlas, font, fontSize)
    const measuredWidth = Math.max(0, ...lineLayouts.map((line) => line.width))
    const clipX = run.clipX ?? run.x
    const clipY = run.clipY ?? run.y
    const clipWidth =
      run.clipWidth ?? (Number.isFinite(run.width ?? DEFAULT_WIDTH) ? (run.width ?? DEFAULT_WIDTH) : measuredWidth + HORIZONTAL_PADDING * 2)
    const clipHeight = run.clipHeight ?? run.height ?? DEFAULT_HEIGHT

    for (const line of lineLayouts) {
      let cursorX = line.x
      for (const glyph of line.text) {
        const entry = atlas.intern(font, glyph)
        quads.push({
          atlasKey: entry.key,
          glyph,
          x: cursorX,
          y: line.y,
          width: entry.width,
          height: entry.height,
          u0: entry.u0,
          v0: entry.v0,
          u1: entry.u1,
          v1: entry.v1,
          color: run.color ?? DEFAULT_COLOR,
          clipX,
          clipY,
          clipWidth,
          clipHeight,
        })
        cursorX += entry.advance
      }
    }
  }
  return quads
}

export function buildTextDecorationRects(runs: readonly TextQuadRun[], atlas: GlyphAtlasLike): TextDecorationRect[] {
  const rects: TextDecorationRect[] = []
  for (const run of runs) {
    if (!run.underline && !run.strike) {
      continue
    }
    const font = run.font ?? DEFAULT_FONT
    const fontSize = run.fontSize ?? 11
    const color = run.color ?? DEFAULT_COLOR
    const lineThickness = Math.max(1, Math.round(fontSize / 14))
    const lineLayouts = resolveLineLayouts(run, atlas, font, fontSize)
    const measuredWidth = Math.max(0, ...lineLayouts.map((line) => line.width))
    const clipX = run.clipX ?? run.x
    const clipRight =
      clipX +
      (run.clipWidth ??
        (Number.isFinite(run.width ?? DEFAULT_WIDTH) ? (run.width ?? DEFAULT_WIDTH) : measuredWidth + HORIZONTAL_PADDING * 2))
    const clipY = run.clipY ?? run.y
    const clipBottom = clipY + (run.clipHeight ?? run.height ?? DEFAULT_HEIGHT)

    for (const line of lineLayouts) {
      const left = Math.max(line.x, clipX)
      const right = Math.min(line.x + line.width, clipRight)
      const visibleWidth = right - left
      if (visibleWidth <= 0) {
        continue
      }
      if (run.underline) {
        const underlineY = line.y + Math.max(1, fontSize * 0.36)
        if (underlineY >= clipY && underlineY <= clipBottom) {
          rects.push({
            x: left,
            y: underlineY,
            width: visibleWidth,
            height: lineThickness,
            color,
          })
        }
      }
      if (run.strike) {
        const strikeY = line.y - Math.max(1, fontSize * 0.18)
        if (strikeY >= clipY && strikeY <= clipBottom) {
          rects.push({
            x: left,
            y: strikeY,
            width: visibleWidth,
            height: lineThickness,
            color,
          })
        }
      }
    }
  }
  return rects
}

export function buildTextQuadsFromScene(items: readonly GridTextItem[], atlas: GlyphAtlasLike): TextQuad[] {
  return buildTextQuads(
    items.map((item) => ({
      text: item.text,
      x: item.x + item.clipInsetLeft,
      y: item.y + item.clipInsetTop,
      width: Math.max(0, item.width - item.clipInsetLeft - item.clipInsetRight),
      height: Math.max(0, item.height - item.clipInsetTop - item.clipInsetBottom),
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

export function buildTextDecorationRectsFromScene(items: readonly GridTextItem[], atlas: GlyphAtlasLike): TextDecorationRect[] {
  return buildTextDecorationRects(
    items.map((item) => ({
      text: item.text,
      x: item.x + item.clipInsetLeft,
      y: item.y + item.clipInsetTop,
      width: Math.max(0, item.width - item.clipInsetLeft - item.clipInsetRight),
      height: Math.max(0, item.height - item.clipInsetTop - item.clipInsetBottom),
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
