import { useEffect, useRef, useState } from 'react'
import type { GridTextScene } from './gridTextScene.js'
import { resolveTextClipRect, resolveTextDecorationRects, resolveTextLineLayouts, type GlyphAtlasLike } from './renderer/gridTextLayout.js'
import type { GlyphAtlasEntry } from './renderer/glyph-atlas.js'

interface GridTextOverlayProps {
  readonly active: boolean
  readonly host: HTMLDivElement | null
  readonly scene: GridTextScene
}

interface SurfaceSize {
  readonly width: number
  readonly height: number
  readonly pixelWidth: number
  readonly pixelHeight: number
}

function noteCanvasSurfaceMount(kind: 'canvas' | 'dom'): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(
    window as Window & {
      __biligScrollPerf?: {
        noteCanvasSurfaceMount?: (kind: 'canvas' | 'dom') => void
      }
    }
  ).__biligScrollPerf?.noteCanvasSurfaceMount?.(kind)
}

function noteCanvasPaint(layer: string): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(window as Window & { __biligScrollPerf?: { noteCanvasPaint?: (layer: string) => void } }).__biligScrollPerf?.noteCanvasPaint?.(layer)
}

export function GridTextOverlay({ active, host, scene }: GridTextOverlayProps) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const [surfaceSize, setSurfaceSize] = useState<SurfaceSize>({
    width: 0,
    height: 0,
    pixelWidth: 0,
    pixelHeight: 0,
  })

  useEffect(() => {
    if (!host) {
      setSurfaceSize({ width: 0, height: 0, pixelWidth: 0, pixelHeight: 0 })
      return
    }

    const updateSurfaceSize = () => {
      const next = resolveSurfaceSize(host)
      setSurfaceSize((current) =>
        current.width === next.width &&
        current.height === next.height &&
        current.pixelWidth === next.pixelWidth &&
        current.pixelHeight === next.pixelHeight
          ? current
          : next,
      )
    }

    updateSurfaceSize()
    if (typeof ResizeObserver === 'undefined') {
      const frame = window.requestAnimationFrame(updateSurfaceSize)
      return () => {
        window.cancelAnimationFrame(frame)
      }
    }

    const observer = new ResizeObserver(() => {
      updateSurfaceSize()
    })
    observer.observe(host)
    return () => {
      observer.disconnect()
    }
  }, [host])

  useEffect(() => {
    if (!active) {
      return
    }
    noteCanvasSurfaceMount('canvas')
  }, [active])

  useEffect(() => {
    const canvas = canvasRef.current
    if (!active || !canvas) {
      return
    }
    noteCanvasPaint('text:overlay')
    configureCanvas(canvas, surfaceSize)
    const context = canvas.getContext('2d')
    if (!context) {
      return
    }
    drawScene(context, surfaceSize, scene)
  }, [active, scene, surfaceSize])

  if (!active || !host) {
    return null
  }

  return <canvas aria-hidden="true" className="pointer-events-none absolute inset-0 z-20" data-testid="grid-text-overlay" ref={canvasRef} />
}

function resolveSurfaceSize(host: HTMLElement): SurfaceSize {
  const width = Math.max(0, Math.floor(host.clientWidth))
  const height = Math.max(0, Math.floor(host.clientHeight))
  const dpr = Math.max(1, window.devicePixelRatio || 1)
  return {
    width,
    height,
    pixelWidth: Math.max(1, Math.floor(width * dpr)),
    pixelHeight: Math.max(1, Math.floor(height * dpr)),
  }
}

function configureCanvas(canvas: HTMLCanvasElement, surfaceSize: SurfaceSize): void {
  if (canvas.width !== surfaceSize.pixelWidth) {
    canvas.width = surfaceSize.pixelWidth
  }
  if (canvas.height !== surfaceSize.pixelHeight) {
    canvas.height = surfaceSize.pixelHeight
  }
  const cssWidth = `${surfaceSize.width}px`
  const cssHeight = `${surfaceSize.height}px`
  if (canvas.style.width !== cssWidth) {
    canvas.style.width = cssWidth
  }
  if (canvas.style.height !== cssHeight) {
    canvas.style.height = cssHeight
  }
}

function drawScene(context: CanvasRenderingContext2D, surfaceSize: SurfaceSize, scene: GridTextScene): void {
  context.save()
  context.setTransform(1, 0, 0, 1, 0, 0)
  context.clearRect(0, 0, surfaceSize.pixelWidth, surfaceSize.pixelHeight)
  context.scale(surfaceSize.pixelWidth / Math.max(1, surfaceSize.width), surfaceSize.pixelHeight / Math.max(1, surfaceSize.height))

  for (const item of scene.items) {
    drawTextItem(context, item)
  }

  context.restore()
}

function drawTextItem(context: CanvasRenderingContext2D, item: GridTextScene['items'][number]): void {
  const clipX = item.x + item.clipInsetLeft
  const clipY = item.y + item.clipInsetTop
  const clipWidth = Math.max(0, item.width - item.clipInsetLeft - item.clipInsetRight)
  const clipHeight = Math.max(0, item.height - item.clipInsetTop - item.clipInsetBottom)
  if (clipWidth <= 0 || clipHeight <= 0) {
    return
  }

  const atlas = createCanvasTextAtlas(context)
  const run = {
    align: item.align,
    clipHeight,
    clipWidth,
    clipX,
    clipY,
    color: item.color,
    font: item.font,
    fontSize: item.fontSize,
    height: clipHeight,
    strike: item.strike,
    text: item.text,
    underline: item.underline,
    width: clipWidth,
    wrap: item.wrap,
    x: clipX,
    y: clipY,
  }
  const lineLayouts = resolveTextLineLayouts(run, atlas)
  const clipRect = resolveTextClipRect(run, lineLayouts)
  const decorationRects = resolveTextDecorationRects(run, atlas)

  context.save()
  context.beginPath()
  context.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height)
  context.clip()
  context.font = item.font
  context.fillStyle = item.color
  context.textAlign = 'left'
  context.textBaseline = 'alphabetic'

  for (const line of lineLayouts) {
    if (line.text.length === 0) {
      continue
    }
    const entry = atlas.intern(item.font, line.text)
    context.fillText(line.text, line.x, line.y + entry.baseline)
  }

  context.fillStyle = item.color
  for (const rect of decorationRects) {
    context.fillRect(rect.x, rect.y, rect.width, rect.height)
  }

  context.restore()
}

function createCanvasTextAtlas(context: CanvasRenderingContext2D): GlyphAtlasLike {
  return {
    intern(font: string, glyph: string): GlyphAtlasEntry {
      context.font = font
      const metrics = context.measureText(glyph)
      const fontSize = parseCanvasFontSize(font)
      const originOffsetX = metrics.actualBoundingBoxLeft || 0
      const bboxWidth = originOffsetX + (metrics.actualBoundingBoxRight || 0)
      const width = Math.max(1, Math.ceil(Math.max(bboxWidth, metrics.width || 0)))
      const height = Math.max(1, Math.ceil(metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent || fontSize * 1.2))
      return {
        advance: Math.max(1, metrics.width || width),
        baseline: Math.max(1, Math.ceil(metrics.actualBoundingBoxAscent || fontSize)),
        font,
        glyph,
        height,
        key: `${font}:${glyph}`,
        originOffsetX,
        u0: 0,
        u1: 1,
        v0: 0,
        v1: 1,
        width,
        x: 0,
        y: 0,
      }
    },
  }
}

function parseCanvasFontSize(font: string): number {
  const match = font.match(/(\d+(?:\.\d+)?)px/)
  return match ? Number(match[1]) : 12
}
