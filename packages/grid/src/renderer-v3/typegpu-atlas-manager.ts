export interface GlyphAtlasEntry {
  readonly key: string
  readonly font: string
  readonly glyph: string
  readonly x: number
  readonly y: number
  readonly originOffsetX: number
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
  originOffsetX: number
  width: number
  height: number
  advance: number
  baseline: number
  u0: number
  v0: number
  u1: number
  v1: number
}

export interface GlyphAtlasDirtyPageUpload {
  readonly pageId: number
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly byteSize: number
}

export interface GlyphAtlasDirtyPageStats {
  readonly atlasSeq: number
  readonly pageCount: number
  readonly dirtyPageCount: number
  readonly dirtyUploadBytes: number
}

type AtlasCanvasLike = HTMLCanvasElement | OffscreenCanvas
type AtlasContextLike = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D

const ATLAS_SCALE = 3
const ATLAS_DIRTY_PAGE_SIZE = 512
const ATLAS_PAGE_ID_STRIDE = 65536

function configureTextContext(context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D): void {
  context.textBaseline = 'alphabetic'
  context.imageSmoothingEnabled = true
  context.imageSmoothingQuality = 'high'
  if ('fontKerning' in context) {
    Reflect.set(context, 'fontKerning', 'normal')
  }
  if ('textRendering' in context) {
    Reflect.set(context, 'textRendering', 'geometricPrecision')
  }
}

function configureAtlasContext(
  context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  scale: number,
  logicalWidth: number,
  logicalHeight: number,
): void {
  context.setTransform(1, 0, 0, 1, 0, 0)
  context.scale(scale, scale)
  configureTextContext(context)
  context.fillStyle = '#ffffff'
  context.clearRect(0, 0, logicalWidth, logicalHeight)
}

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

function getAtlasContext(canvas: AtlasCanvasLike): AtlasContextLike | null {
  const context = canvas.getContext('2d') as AtlasContextLike | null
  if (!context || !('fillText' in context) || !('measureText' in context)) {
    return null
  }
  return context
}

function parseFontSize(font: string): number {
  const match = font.match(/(\d+(?:\.\d+)?)px/)
  return match ? Number(match[1]) : 12
}

const MEASURE_SCALE = 2

function measureGlyph(
  font: string,
  glyph: string,
): {
  width: number
  advance: number
  height: number
  baseline: number
  originOffsetX: number
} {
  const fontSize = parseFontSize(font)
  const canvas = createAtlasCanvas(
    Math.max(64, Math.ceil(fontSize * MEASURE_SCALE * 2)),
    Math.max(64, Math.ceil(fontSize * MEASURE_SCALE * 2)),
  )
  const context = canvas ? getAtlasContext(canvas) : null
  if (!context) {
    const width = Math.max(1, Math.ceil(fontSize * 0.6))
    return {
      width,
      advance: width,
      height: Math.max(1, Math.ceil(fontSize * 1.2)),
      baseline: Math.ceil(fontSize),
      originOffsetX: 0,
    }
  }

  configureTextContext(context)
  context.font = font
  const metrics = context.measureText(glyph)
  const originOffsetX = metrics.actualBoundingBoxLeft || 0
  const bboxWidth = originOffsetX + (metrics.actualBoundingBoxRight || 0)

  const measuredWidth = Math.max(1, Math.ceil(Math.max(bboxWidth, metrics.width || 0)))
  const measuredHeight = Math.max(1, Math.ceil(metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent || fontSize * 1.2))

  return {
    width: measuredWidth,
    advance: Math.max(1, metrics.width || measuredWidth),
    height: measuredHeight,
    baseline: Math.max(1, Math.ceil(metrics.actualBoundingBoxAscent || fontSize)),
    originOffsetX,
  }
}

export function createGlyphAtlas(input: { initialWidth?: number; initialHeight?: number; padding?: number } = {}) {
  const padding = Math.max(2, input.padding ?? 2)
  const scale = ATLAS_SCALE
  let width = Math.max(512, (input.initialWidth ?? 1024) * scale)
  let height = Math.max(512, (input.initialHeight ?? 1024) * scale)
  const canvas = createAtlasCanvas(width, height)
  let context = canvas ? getAtlasContext(canvas) : null

  if (context) {
    configureAtlasContext(context, scale, width / scale, height / scale)
  }

  const entries = new Map<string, MutableGlyphAtlasEntry>()
  let cursorX = padding
  let cursorY = padding
  let rowHeight = 0
  let version = 0
  let atlasSeq = 0
  const dirtyPages = new Map<number, GlyphAtlasDirtyPageUpload>()

  const redrawAll = () => {
    if (!context) return
    configureAtlasContext(context, scale, width / scale, height / scale)
    for (const entry of entries.values()) {
      context.font = entry.font
      context.fillText(entry.glyph, entry.x + entry.originOffsetX, entry.y + entry.baseline)
    }
  }

  const growAtlas = (minimumWidth: number, minimumHeight: number) => {
    if (!canvas) return
    let nextWidth = width
    let nextHeight = height
    while (nextWidth < minimumWidth) nextWidth *= 2
    while (nextHeight < minimumHeight) nextHeight *= 2

    canvas.width = nextWidth
    canvas.height = nextHeight

    width = nextWidth
    height = nextHeight
    const nextContext = getAtlasContext(canvas)
    if (!nextContext) return
    context = nextContext
    configureAtlasContext(nextContext, scale, width / scale, height / scale)

    // UVs need update
    for (const entry of entries.values()) {
      entry.u0 = (entry.x * scale) / width
      entry.v0 = (entry.y * scale) / height
      entry.u1 = ((entry.x + entry.width) * scale) / width
      entry.v1 = ((entry.y + entry.height) * scale) / height
    }

    redrawAll()
    version += 1
    atlasSeq += 1
    markAllPagesDirty()
  }

  const intern = (font: string, glyph: string): GlyphAtlasEntry => {
    const key = `${font}:${glyph}`
    const existing = entries.get(key)
    if (existing) return existing

    const metrics = measureGlyph(font, glyph)
    if (cursorX + metrics.width + padding > width / scale) {
      cursorX = padding
      cursorY += rowHeight + padding
      rowHeight = 0
    }

    if ((cursorY + metrics.height + padding) * scale > height) {
      growAtlas(width, (cursorY + metrics.height + padding) * scale)
    }

    const entry: MutableGlyphAtlasEntry = {
      key,
      font,
      glyph,
      x: cursorX,
      y: cursorY,
      originOffsetX: metrics.originOffsetX,
      width: metrics.width,
      height: metrics.height,
      advance: metrics.advance,
      baseline: metrics.baseline,
      u0: (cursorX * scale) / width,
      v0: (cursorY * scale) / height,
      u1: ((cursorX + metrics.width) * scale) / width,
      v1: ((cursorY + metrics.height) * scale) / height,
    }

    version += 1
    if (context) {
      context.font = font
      context.fillText(glyph, entry.x + entry.originOffsetX, entry.y + metrics.baseline)
    }
    markGlyphPageDirty(entry)

    entries.set(key, entry)
    cursorX += metrics.width + padding
    rowHeight = Math.max(rowHeight, metrics.height)
    return entry
  }

  const markGlyphPageDirty = (entry: Pick<MutableGlyphAtlasEntry, 'x' | 'y' | 'width' | 'height'>) => {
    const x = Math.max(0, Math.floor(entry.x * scale))
    const y = Math.max(0, Math.floor(entry.y * scale))
    const xEnd = Math.min(width, Math.ceil((entry.x + entry.width) * scale))
    const yEnd = Math.min(height, Math.ceil((entry.y + entry.height) * scale))
    markPhysicalRectDirty(x, y, xEnd - x, yEnd - y)
  }

  const markAllPagesDirty = () => {
    markPhysicalRectDirty(0, 0, width, height)
  }

  const markPhysicalRectDirty = (x: number, y: number, dirtyWidth: number, dirtyHeight: number) => {
    const x0 = Math.max(0, Math.floor(x))
    const y0 = Math.max(0, Math.floor(y))
    const x1 = Math.min(width, Math.ceil(x + dirtyWidth))
    const y1 = Math.min(height, Math.ceil(y + dirtyHeight))
    if (x1 <= x0 || y1 <= y0) {
      return
    }

    const pageX0 = Math.floor(x0 / ATLAS_DIRTY_PAGE_SIZE)
    const pageY0 = Math.floor(y0 / ATLAS_DIRTY_PAGE_SIZE)
    const pageX1 = Math.floor((x1 - 1) / ATLAS_DIRTY_PAGE_SIZE)
    const pageY1 = Math.floor((y1 - 1) / ATLAS_DIRTY_PAGE_SIZE)
    for (let pageY = pageY0; pageY <= pageY1; pageY += 1) {
      for (let pageX = pageX0; pageX <= pageX1; pageX += 1) {
        markPageDirty(pageX, pageY)
      }
    }
  }

  const markPageDirty = (pageX: number, pageY: number) => {
    const pageId = pageY * ATLAS_PAGE_ID_STRIDE + pageX
    const x = pageX * ATLAS_DIRTY_PAGE_SIZE
    const y = pageY * ATLAS_DIRTY_PAGE_SIZE
    const pageWidth = Math.max(0, Math.min(ATLAS_DIRTY_PAGE_SIZE, width - x))
    const pageHeight = Math.max(0, Math.min(ATLAS_DIRTY_PAGE_SIZE, height - y))
    if (pageWidth <= 0 || pageHeight <= 0) {
      return
    }
    if (!dirtyPages.has(pageId)) {
      dirtyPages.set(pageId, {
        byteSize: pageWidth * pageHeight * 4,
        height: pageHeight,
        pageId,
        width: pageWidth,
        x,
        y,
      })
    }
    atlasSeq += 1
  }

  return {
    drainDirtyPages(): readonly GlyphAtlasDirtyPageUpload[] {
      const pages = [...dirtyPages.values()].toSorted((left, right) => left.pageId - right.pageId)
      dirtyPages.clear()
      return pages
    },
    getCanvas(): AtlasCanvasLike | null {
      return canvas
    },
    getDirtyPageStats(): GlyphAtlasDirtyPageStats {
      let dirtyUploadBytes = 0
      for (const page of dirtyPages.values()) {
        dirtyUploadBytes += page.byteSize
      }
      return {
        atlasSeq,
        dirtyPageCount: dirtyPages.size,
        dirtyUploadBytes,
        pageCount: resolveAtlasPageCount(width, height),
      }
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

function resolveAtlasPageCount(width: number, height: number): number {
  return Math.ceil(width / ATLAS_DIRTY_PAGE_SIZE) * Math.ceil(height / ATLAS_DIRTY_PAGE_SIZE)
}
