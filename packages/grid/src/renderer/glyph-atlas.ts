export interface GlyphAtlasEntry {
  readonly key: string
  readonly font: string
  readonly glyph: string
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly advance: number
  readonly baseline: number
  readonly u0: number
  readonly v0: number
  readonly u1: number
  readonly v1: number
}

interface MutableGlyphAtlasEntry {
  key: string
  font: string
  glyph: string
  x: number
  y: number
  width: number
  height: number
  advance: number
  baseline: number
  u0: number
  v0: number
  u1: number
  v1: number
}

type AtlasCanvasLike = HTMLCanvasElement | OffscreenCanvas

function createAtlasCanvas(width: number, height: number): AtlasCanvasLike | null {
  if (typeof OffscreenCanvas !== 'undefined') {
    return new OffscreenCanvas(width, height)
  }
  if (typeof document !== 'undefined') {
    const canvas = document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    return canvas
  }
  return null
}

function parseFontSize(font: string): number {
  const match = font.match(/(\d+(?:\.\d+)?)px/)
  return match ? Number(match[1]) : 12
}

function measureGlyph(font: string, glyph: string): { width: number; advance: number; height: number; baseline: number } {
  const fontSize = parseFontSize(font)
  const canvas = createAtlasCanvas(Math.max(32, Math.ceil(fontSize * 2)), Math.max(32, Math.ceil(fontSize * 2)))
  const context = canvas?.getContext?.('2d') ?? null
  if (!context) {
    const width = Math.max(1, Math.ceil(fontSize * 0.6))
    const height = Math.max(1, Math.ceil(fontSize * 1.2))
    return {
      width,
      advance: width,
      height,
      baseline: Math.ceil(fontSize),
    }
  }
  context.font = font
  context.textBaseline = 'alphabetic'
  const metrics = context.measureText(glyph)
  const measuredWidth = Math.max(1, Math.ceil(metrics.actualBoundingBoxLeft + metrics.actualBoundingBoxRight || metrics.width))
  const measuredHeight = Math.max(1, Math.ceil(metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent || fontSize * 1.2))
  return {
    width: measuredWidth,
    advance: Math.max(1, Math.ceil(metrics.width)),
    height: measuredHeight,
    baseline: Math.max(1, Math.ceil(metrics.actualBoundingBoxAscent || fontSize)),
  }
}

export function createGlyphAtlas(input: { initialWidth?: number; initialHeight?: number; padding?: number } = {}) {
  const padding = Math.max(1, input.padding ?? 2)
  let width = Math.max(256, input.initialWidth ?? 1024)
  let height = Math.max(256, input.initialHeight ?? 1024)
  const canvas = createAtlasCanvas(width, height)
  let context = canvas?.getContext?.('2d') ?? null
  if (context) {
    context.textBaseline = 'alphabetic'
    context.fillStyle = '#ffffff'
    context.clearRect(0, 0, width, height)
  }
  const entries = new Map<string, MutableGlyphAtlasEntry>()
  let cursorX = padding
  let cursorY = padding
  let rowHeight = 0
  let version = 0

  const growAtlas = (minimumWidth: number, minimumHeight: number) => {
    if (!canvas) {
      return
    }
    let nextWidth = width
    let nextHeight = height
    while (nextWidth < minimumWidth) {
      nextWidth *= 2
    }
    while (nextHeight < minimumHeight) {
      nextHeight *= 2
    }
    if (canvas instanceof HTMLCanvasElement) {
      canvas.width = nextWidth
      canvas.height = nextHeight
    } else {
      ;(canvas).width = nextWidth
      ;(canvas).height = nextHeight
    }
    width = nextWidth
    height = nextHeight
    context = canvas.getContext('2d')
    if (!context) {
      return
    }
    context.textBaseline = 'alphabetic'
    context.fillStyle = '#ffffff'
    context.clearRect(0, 0, width, height)
    cursorX = padding
    cursorY = padding
    rowHeight = 0
    for (const entry of entries.values()) {
      if (cursorX + entry.width + padding > width) {
        cursorX = padding
        cursorY += rowHeight + padding
        rowHeight = 0
      }
      entry.x = cursorX
      entry.y = cursorY
      entry.u0 = entry.x / width
      entry.v0 = entry.y / height
      entry.u1 = (entry.x + entry.width) / width
      entry.v1 = (entry.y + entry.height) / height
      context.font = entry.font
      context.fillText(entry.glyph, entry.x, entry.y + entry.baseline)
      cursorX += entry.width + padding
      rowHeight = Math.max(rowHeight, entry.height)
    }
    version += 1
  }

  const intern = (font: string, glyph: string): GlyphAtlasEntry => {
    const key = `${font}:${glyph}`
    const existing = entries.get(key)
    if (existing) {
      return existing
    }
    const metrics = measureGlyph(font, glyph)
    if (cursorX + metrics.width + padding > width) {
      cursorX = padding
      cursorY += rowHeight + padding
      rowHeight = 0
    }
    if (cursorY + metrics.height + padding > height) {
      growAtlas(width, cursorY + metrics.height + padding)
    }
    const entry: MutableGlyphAtlasEntry = {
      key,
      font,
      glyph,
      x: cursorX,
      y: cursorY,
      width: metrics.width,
      height: metrics.height,
      advance: metrics.advance,
      baseline: metrics.baseline,
      u0: cursorX / width,
      v0: cursorY / height,
      u1: (cursorX + metrics.width) / width,
      v1: (cursorY + metrics.height) / height,
    }
    version += 1
    if (context) {
      context.font = font
      context.fillText(glyph, entry.x, entry.y + entry.baseline)
    }
    entries.set(key, entry)
    cursorX += entry.width + padding
    rowHeight = Math.max(rowHeight, entry.height)
    return entry
  }

  return {
    getCanvas(): AtlasCanvasLike | null {
      return canvas
    },
    getVersion(): number {
      return version
    },
    getSize(): { width: number; height: number } {
      return { width, height }
    },
    intern,
  }
}
