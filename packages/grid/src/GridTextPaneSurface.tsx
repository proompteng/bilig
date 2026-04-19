import { memo, useEffect, useRef } from 'react'
import type { Rectangle } from './gridTypes.js'
import type { GridTextScene } from './gridTextScene.js'
import type { WorkbookGridScrollStore } from './workbookGridScrollStore.js'

interface GridTextPaneSurfaceProps {
  readonly paneId: string
  readonly active: boolean
  readonly frame: Rectangle
  readonly surfaceSize: {
    readonly width: number
    readonly height: number
  }
  readonly contentOffset: {
    readonly x: number
    readonly y: number
  }
  readonly scene: GridTextScene
  readonly scrollAxes?: {
    readonly x: boolean
    readonly y: boolean
  }
  readonly scrollTransformStore?: WorkbookGridScrollStore | null
}

const HORIZONTAL_PADDING = 8
const WRAP_TOP_PADDING = 4
const WRAP_LINE_HEIGHT = 1.2

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

export const GridTextPaneSurface = memo(function GridTextPaneSurface({
  paneId,
  active,
  frame,
  surfaceSize,
  contentOffset,
  scene,
  scrollAxes,
  scrollTransformStore = null,
}: GridTextPaneSurfaceProps) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null)

  useEffect(() => {
    if (!active) {
      return
    }
    noteCanvasSurfaceMount('canvas')
  }, [active])

  useEffect(() => {
    const canvas = canvasRef.current
    if (!canvas) {
      return
    }

    const applyContentOffset = () => {
      const snapshot = scrollTransformStore?.getSnapshot() ?? { tx: 0, ty: 0 }
      const nextX = contentOffset.x - (scrollAxes?.x ? snapshot.tx : 0)
      const nextY = contentOffset.y - (scrollAxes?.y ? snapshot.ty : 0)
      const nextTransform = `translate3d(${nextX}px, ${nextY}px, 0)`
      if (canvas.style.transform !== nextTransform) {
        canvas.style.transform = nextTransform
      }
    }

    applyContentOffset()
    if (!scrollTransformStore || (!scrollAxes?.x && !scrollAxes?.y)) {
      return
    }
    return scrollTransformStore.subscribe(applyContentOffset)
  }, [contentOffset.x, contentOffset.y, scrollAxes?.x, scrollAxes?.y, scrollTransformStore])

  useEffect(() => {
    const canvas = canvasRef.current
    if (!active || !canvas) {
      return
    }
    configureCanvas(canvas, surfaceSize)
    const context = canvas.getContext('2d')
    if (!context) {
      return
    }
    noteCanvasPaint(`text:${paneId}`)
    drawScene(context, surfaceSize, scene)
  }, [active, paneId, scene, surfaceSize])

  if (!active || frame.width <= 0 || frame.height <= 0 || surfaceSize.width <= 0 || surfaceSize.height <= 0) {
    return null
  }

  return (
    <div
      aria-hidden="true"
      className="pointer-events-none absolute z-20 overflow-hidden"
      style={{
        height: frame.height,
        left: frame.x,
        top: frame.y,
        width: frame.width,
      }}
    >
      <canvas
        className="absolute"
        data-pane-id={paneId}
        data-testid={`grid-text-pane-${paneId}`}
        ref={canvasRef}
        style={{
          left: 0,
          top: 0,
        }}
      />
    </div>
  )
})

function configureCanvas(
  canvas: HTMLCanvasElement,
  surfaceSize: {
    readonly width: number
    readonly height: number
  },
): void {
  const dpr = Math.max(1, window.devicePixelRatio || 1)
  const pixelWidth = Math.max(1, Math.floor(surfaceSize.width * dpr))
  const pixelHeight = Math.max(1, Math.floor(surfaceSize.height * dpr))
  if (canvas.width !== pixelWidth) {
    canvas.width = pixelWidth
  }
  if (canvas.height !== pixelHeight) {
    canvas.height = pixelHeight
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

function drawScene(
  context: CanvasRenderingContext2D,
  surfaceSize: {
    readonly width: number
    readonly height: number
  },
  scene: GridTextScene,
): void {
  context.save()
  context.setTransform(1, 0, 0, 1, 0, 0)
  context.clearRect(0, 0, context.canvas.width, context.canvas.height)
  context.scale(context.canvas.width / Math.max(1, surfaceSize.width), context.canvas.height / Math.max(1, surfaceSize.height))
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
  const measured = context.measureText(text)
  const actualWidth = measured.width
  const left = item.align === 'right' ? textX - actualWidth : item.align === 'center' ? textX - actualWidth / 2 : textX
  const lineWidth = Math.max(1, Math.round(item.fontSize / 14))
  context.save()
  context.strokeStyle = item.color
  context.lineWidth = lineWidth
  if (item.underline) {
    const underlineY = textY + Math.max(1, item.fontSize * 0.36)
    context.beginPath()
    context.moveTo(left, underlineY)
    context.lineTo(left + actualWidth, underlineY)
    context.stroke()
  }
  if (item.strike) {
    const strikeY = textY - Math.max(1, item.fontSize * 0.18)
    context.beginPath()
    context.moveTo(left, strikeY)
    context.lineTo(left + actualWidth, strikeY)
    context.stroke()
  }
  context.restore()
}
