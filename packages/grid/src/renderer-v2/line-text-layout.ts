import type { GlyphAtlasEntry } from './typegpu-atlas-manager.js'

export interface TextLayoutRun {
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
}

export interface ResolvedTextLineLayout {
  readonly text: string
  readonly x: number
  readonly y: number
  readonly width: number
}

export interface ResolvedTextClipRect {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

export interface ResolvedTextDecorationRect {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly color: string
}

export interface GlyphAtlasLike {
  intern(font: string, glyph: string): GlyphAtlasEntry
}

export const DEFAULT_TEXT_FONT = '400 11px sans-serif'
export const DEFAULT_TEXT_COLOR = '#1f2933'
export const DEFAULT_TEXT_HEIGHT = 16
export const DEFAULT_TEXT_WIDTH = Number.POSITIVE_INFINITY
export const TEXT_HORIZONTAL_PADDING = 8
export const TEXT_WRAP_TOP_PADDING = 4
export const TEXT_WRAP_LINE_HEIGHT = 1.2

export function parseTextFontSize(font: string): number {
  const match = font.match(/(\d+(?:\.\d+)?)px/)
  return match ? Number(match[1]) : 11
}

export function parseTextCssColor(color: string): readonly [number, number, number, number] {
  if (color.startsWith('#')) {
    const hex = color.slice(1)
    if (hex.length === 3) {
      return [parseInt(hex[0]! + hex[0], 16) / 255, parseInt(hex[1]! + hex[1], 16) / 255, parseInt(hex[2]! + hex[2], 16) / 255, 1.0]
    }
    return [
      parseInt(hex.slice(0, 2), 16) / 255,
      parseInt(hex.slice(2, 4), 16) / 255,
      parseInt(hex.slice(4, 6), 16) / 255,
      hex.length === 8 ? parseInt(hex.slice(6, 8), 16) / 255 : 1.0,
    ]
  }
  if (color.startsWith('rgba')) {
    const match = color.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/)
    if (match) {
      return [Number(match[1]) / 255, Number(match[2]) / 255, Number(match[3]) / 255, match[4] ? Number(match[4]) : 1.0]
    }
  }
  return [0, 0, 0, 1]
}

export function resolveTextLineLayouts(run: TextLayoutRun, atlas: GlyphAtlasLike): ResolvedTextLineLayout[] {
  const font = run.font ?? DEFAULT_TEXT_FONT
  const fontSize = run.fontSize ?? parseTextFontSize(font)
  const lineHeight = Math.ceil(fontSize * TEXT_WRAP_LINE_HEIGHT)
  const availableWidth = Math.max(0, (run.width ?? DEFAULT_TEXT_WIDTH) - TEXT_HORIZONTAL_PADDING * 2)
  const lines = wrapTextLines(run, atlas, font, availableWidth)
  let lineY = run.wrap ? run.y + TEXT_WRAP_TOP_PADDING : run.y + Math.max(0, ((run.height ?? DEFAULT_TEXT_HEIGHT) - fontSize) / 2)

  return lines.map((line) => {
    const lineWidth = line.length === 0 ? 0 : atlas.intern(font, line).advance
    const startX =
      (run.align ?? 'left') === 'right'
        ? run.x + (run.width ?? lineWidth) - TEXT_HORIZONTAL_PADDING - lineWidth
        : (run.align ?? 'left') === 'center'
          ? run.x + ((run.width ?? lineWidth) - lineWidth) / 2
          : run.x + TEXT_HORIZONTAL_PADDING
    const resolvedLine = {
      text: line,
      width: lineWidth,
      x: startX,
      y: lineY,
    }
    if (run.wrap) {
      lineY += lineHeight
    }
    return resolvedLine
  })
}

export function resolveTextClipRect(run: TextLayoutRun, lineLayouts: readonly ResolvedTextLineLayout[]): ResolvedTextClipRect {
  const measuredWidth = Math.max(0, ...lineLayouts.map((line) => line.width))
  return {
    height: run.clipHeight ?? run.height ?? DEFAULT_TEXT_HEIGHT,
    width:
      run.clipWidth ??
      (Number.isFinite(run.width ?? DEFAULT_TEXT_WIDTH) ? (run.width ?? DEFAULT_TEXT_WIDTH) : measuredWidth + TEXT_HORIZONTAL_PADDING * 2),
    x: run.clipX ?? run.x,
    y: run.clipY ?? run.y,
  }
}

export function resolveTextDecorationRects(
  run: TextLayoutRun & {
    readonly color?: string
    readonly underline?: boolean
    readonly strike?: boolean
  },
  atlas: GlyphAtlasLike,
): ResolvedTextDecorationRect[] {
  if (!run.underline && !run.strike) {
    return []
  }

  const font = run.font ?? DEFAULT_TEXT_FONT
  const fontSize = run.fontSize ?? parseTextFontSize(font)
  const color = run.color ?? DEFAULT_TEXT_COLOR
  const lineThickness = Math.max(1, Math.round(fontSize / 14))
  const lineLayouts = resolveTextLineLayouts(run, atlas)
  const clipRect = resolveTextClipRect(run, lineLayouts)
  const clipRight = clipRect.x + clipRect.width
  const clipBottom = clipRect.y + clipRect.height
  const rects: ResolvedTextDecorationRect[] = []

  for (const line of lineLayouts) {
    const left = Math.max(line.x, clipRect.x)
    const right = Math.min(line.x + line.width, clipRight)
    const visibleWidth = right - left
    if (visibleWidth <= 0) {
      continue
    }
    if (run.underline) {
      const underlineY = line.y + Math.max(1, fontSize * 0.36)
      if (underlineY >= clipRect.y && underlineY <= clipBottom) {
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
      if (strikeY >= clipRect.y && strikeY <= clipBottom) {
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

  return rects
}

function measureTextWidth(text: string, font: string, atlas: GlyphAtlasLike): number {
  if (text.length === 0) {
    return 0
  }
  let width = 0
  for (const glyph of splitGraphemes(text)) {
    width += atlas.intern(font, glyph).advance
  }
  return width
}

function breakWord(word: string, font: string, atlas: GlyphAtlasLike, maxWidth: number): string[] {
  const segments: string[] = []
  let current = ''
  for (const glyph of splitGraphemes(word)) {
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

function wrapTextLines(run: TextLayoutRun, atlas: GlyphAtlasLike, font: string, maxWidth: number): readonly string[] {
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
      lines.push(...breakWord(word, font, atlas, maxWidth))
    }
    if (current.length > 0) {
      lines.push(current.trimEnd())
    }
  }

  return lines
}

function splitGraphemes(text: string): readonly string[] {
  if (typeof Intl !== 'undefined' && 'Segmenter' in Intl) {
    const segmenter = new Intl.Segmenter(undefined, { granularity: 'grapheme' })
    return [...segmenter.segment(text)].map((segment) => segment.segment)
  }
  return Array.from(text)
}
