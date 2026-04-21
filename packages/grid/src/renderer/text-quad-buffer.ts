import type { GridTextItem } from '../gridTextScene.js'
import {
  DEFAULT_TEXT_COLOR,
  DEFAULT_TEXT_FONT,
  parseTextCssColor,
  resolveTextClipRect,
  resolveTextDecorationRects,
  resolveTextLineLayouts,
  type GlyphAtlasLike,
} from './gridTextLayout.js'

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

const TEXT_INSTANCE_FLOAT_COUNT = 16

export function buildTextQuads(runs: readonly TextQuadRun[], atlas: GlyphAtlasLike): TextQuad[] {
  const quads: TextQuad[] = []
  for (const run of runs) {
    if (run.text.length === 0) {
      continue
    }

    const font = run.font ?? DEFAULT_TEXT_FONT
    const lineLayouts = resolveTextLineLayouts(run, atlas)
    const clipRect = resolveTextClipRect(run, lineLayouts)
    if (clipRect.width <= 0 || clipRect.height <= 0) {
      continue
    }

    for (const line of lineLayouts) {
      if (line.text.length === 0) {
        continue
      }

      const entry = atlas.intern(font, line.text)
      quads.push({
        atlasKey: entry.key,
        glyph: line.text,
        x: line.x - entry.originOffsetX,
        y: line.y,
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

  return quads
}

function packTextQuads(quads: readonly TextQuad[], targetBuffer?: Float32Array): { floats: Float32Array; quadCount: number } {
  const floats =
    targetBuffer && targetBuffer.length >= quads.length * TEXT_INSTANCE_FLOAT_COUNT
      ? targetBuffer
      : new Float32Array(Math.max(1, quads.length) * TEXT_INSTANCE_FLOAT_COUNT)

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
  })

  return { floats, quadCount: quads.length }
}

export function buildTextQuadsFromScene(
  items: readonly GridTextItem[],
  atlas: GlyphAtlasLike,
  targetBuffer?: Float32Array,
): { floats: Float32Array; quadCount: number } {
  const quads = buildTextQuads(
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

  return packTextQuads(quads, targetBuffer)
}

export function buildTextDecorationRects(runs: readonly TextQuadRun[], atlas: GlyphAtlasLike): TextDecorationRect[] {
  const rects: TextDecorationRect[] = []
  for (const run of runs) {
    rects.push(...resolveTextDecorationRects(run, atlas))
  }

  return rects
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
