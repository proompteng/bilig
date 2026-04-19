import { useEffect, useRef, useState } from 'react'
import type { GridTextScene } from './gridTextScene.js'

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

const HORIZONTAL_PADDING = 8
const WRAP_TOP_PADDING = 4
const WRAP_LINE_HEIGHT = 1.2

function noteTextSurface(kind: 'canvas' | 'dom'): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(window as Window & { __biligScrollPerf?: { noteTextSurface?: (kind: 'canvas' | 'dom') => void } }).__biligScrollPerf?.noteTextSurface?.(
    kind,
  )
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
    const canvas = canvasRef.current
    if (!active || !canvas) {
      return
    }
    noteTextSurface('canvas')
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

  context.save()
  context.beginPath()
  context.rect(clipX, clipY, clipWidth, clipHeight)
  context.clip()
  context.font = item.font
  context.fillStyle = item.color
  context.textAlign = item.align

  if (item.wrap) {
    context.textBaseline = 'top'
    drawWrappedText(context, item, clipX, clipY, clipWidth)
  } else {
    context.textBaseline = 'middle'
    drawSingleLineText(context, item, clipX, clipY, clipWidth, clipHeight)
  }

  context.restore()
}

function drawSingleLineText(
  context: CanvasRenderingContext2D,
  item: GridTextScene['items'][number],
  clipX: number,
  clipY: number,
  clipWidth: number,
  clipHeight: number,
): void {
  const textX =
    item.align === 'right'
      ? clipX + clipWidth - HORIZONTAL_PADDING
      : item.align === 'center'
        ? clipX + clipWidth / 2
        : clipX + HORIZONTAL_PADDING
  const textY = clipY + clipHeight / 2
  context.fillText(item.text, textX, textY)
  drawTextDecorations(context, item, textX, textY, item.text)
}

function drawWrappedText(
  context: CanvasRenderingContext2D,
  item: GridTextScene['items'][number],
  clipX: number,
  clipY: number,
  clipWidth: number,
): void {
  const availableWidth = Math.max(0, clipWidth - HORIZONTAL_PADDING * 2)
  if (availableWidth <= 0) {
    return
  }
  const lines = wrapText(context, item.text, availableWidth)
  const lineHeight = Math.ceil(item.fontSize * WRAP_LINE_HEIGHT)
  const textX =
    item.align === 'right'
      ? clipX + clipWidth - HORIZONTAL_PADDING
      : item.align === 'center'
        ? clipX + clipWidth / 2
        : clipX + HORIZONTAL_PADDING
  let textY = clipY + WRAP_TOP_PADDING
  for (const line of lines) {
    context.fillText(line, textX, textY)
    drawTextDecorations(context, item, textX, textY + item.fontSize / 2, line)
    textY += lineHeight
  }
}

function wrapText(context: CanvasRenderingContext2D, text: string, maxWidth: number): string[] {
  const paragraphs = text.split('\n')
  const lines: string[] = []
  for (const paragraph of paragraphs) {
    if (paragraph.length === 0) {
      lines.push('')
      continue
    }
    const words = paragraph.split(/(\s+)/).filter((part) => part.length > 0)
    let current = ''
    for (const word of words) {
      const candidate = current.length === 0 ? word : `${current}${word}`
      if (context.measureText(candidate).width <= maxWidth) {
        current = candidate
        continue
      }
      if (current.length > 0) {
        lines.push(current.trimEnd())
      }
      current = ''
      if (context.measureText(word).width <= maxWidth) {
        current = word.trimStart()
        continue
      }
      for (const segment of breakWord(context, word, maxWidth)) {
        lines.push(segment)
      }
    }
    if (current.length > 0) {
      lines.push(current.trimEnd())
    }
  }
  return lines
}

function breakWord(context: CanvasRenderingContext2D, word: string, maxWidth: number): string[] {
  const segments: string[] = []
  let current = ''
  for (const char of word) {
    const candidate = `${current}${char}`
    if (current.length > 0 && context.measureText(candidate).width > maxWidth) {
      segments.push(current)
      current = char
      continue
    }
    current = candidate
  }
  if (current.length > 0) {
    segments.push(current)
  }
  return segments
}

function drawTextDecorations(
  context: CanvasRenderingContext2D,
  item: GridTextScene['items'][number],
  textX: number,
  textY: number,
  text: string,
): void {
  if (!item.underline && !item.strike) {
    return
  }
  const metrics = context.measureText(text)
  const left = item.align === 'right' ? textX - metrics.width : item.align === 'center' ? textX - metrics.width / 2 : textX
  const right = left + metrics.width
  context.save()
  context.strokeStyle = item.color
  context.lineWidth = 1
  if (item.underline) {
    const underlineY = textY + item.fontSize * 0.32
    context.beginPath()
    context.moveTo(left, underlineY)
    context.lineTo(right, underlineY)
    context.stroke()
  }
  if (item.strike) {
    const strikeY = textY - item.fontSize * 0.18
    context.beginPath()
    context.moveTo(left, strikeY)
    context.lineTo(right, strikeY)
    context.stroke()
  }
  context.restore()
}
