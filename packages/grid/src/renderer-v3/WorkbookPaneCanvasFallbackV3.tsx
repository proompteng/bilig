import { memo, useCallback, useEffect, useLayoutEffect, useRef } from 'react'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import { WORKBOOK_DEFAULT_FONT_SIZE, WORKBOOK_FONT_SANS, workbookFontPointSizeToCssPx } from '../workbookTheme.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { TextQuadRun } from './line-text-quad-buffer.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from './rect-instance-buffer.js'
import { resolveTypeGpuV3DrawScrollSnapshot, resolveWorkbookPaneRendererGeometryV3 } from './workbook-pane-renderer-runtime.js'

type FallbackPane = GridHeaderPaneState | WorkbookRenderTilePaneState

export interface CanvasTextRunContext {
  fillStyle: string | CanvasGradient | CanvasPattern
  font: string
  lineWidth: number
  strokeStyle: string | CanvasGradient | CanvasPattern
  textAlign: CanvasTextAlign
  textBaseline: CanvasTextBaseline
  beginPath(): void
  clip(): void
  fillText(text: string, x: number, y: number): void
  lineTo(x: number, y: number): void
  measureText(text: string): { readonly width: number }
  moveTo(x: number, y: number): void
  rect(x: number, y: number, w: number, h: number): void
  restore(): void
  save(): void
  stroke(): void
}

export interface WorkbookPaneCanvasFallbackV3Props {
  readonly active: boolean
  readonly cameraStore?: GridCameraStore | null | undefined
  readonly drawText?: boolean | undefined
  readonly geometry: GridGeometrySnapshot | null
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly host: HTMLDivElement | null
  readonly layer?: 'fallback' | 'grid-floor' | undefined
  readonly overlay: DynamicGridOverlayBatchV3 | null
  readonly overlayBuilder?: ((geometry: GridGeometrySnapshot) => DynamicGridOverlayBatchV3 | null | undefined) | null | undefined
  readonly scrollTransformStore: WorkbookGridScrollStore | null
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}

function resolvePaneRenderOffset(
  pane: {
    readonly contentOffset: { readonly x: number; readonly y: number }
    readonly scrollAxes: { readonly x: boolean; readonly y: boolean }
  },
  scrollSnapshot: {
    readonly tx: number
    readonly ty: number
    readonly renderTx?: number | undefined
    readonly renderTy?: number | undefined
  },
): { readonly x: number; readonly y: number } {
  const renderTx = scrollSnapshot.renderTx ?? scrollSnapshot.tx
  const renderTy = scrollSnapshot.renderTy ?? scrollSnapshot.ty
  return {
    x: pane.contentOffset.x - (pane.scrollAxes.x ? renderTx : 0),
    y: pane.contentOffset.y - (pane.scrollAxes.y ? renderTy : 0),
  }
}

function colorFromFloats(r: number, g: number, b: number, a: number): string {
  return `rgba(${Math.round(r * 255)}, ${Math.round(g * 255)}, ${Math.round(b * 255)}, ${Math.max(0, Math.min(1, a))})`
}

function drawRectInstances(context: CanvasRenderingContext2D, rectInstances: Float32Array, rectCount: number): void {
  for (let index = 0; index < rectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    const x = rectInstances[offset + 0] ?? 0
    const y = rectInstances[offset + 1] ?? 0
    const width = rectInstances[offset + 2] ?? 0
    const height = rectInstances[offset + 3] ?? 0
    if (width <= 0 || height <= 0) {
      continue
    }
    const fillAlpha = rectInstances[offset + 7] ?? 0
    const borderAlpha = rectInstances[offset + 11] ?? 0
    context.fillStyle =
      fillAlpha > 0
        ? colorFromFloats(rectInstances[offset + 4] ?? 0, rectInstances[offset + 5] ?? 0, rectInstances[offset + 6] ?? 0, fillAlpha)
        : colorFromFloats(rectInstances[offset + 8] ?? 0, rectInstances[offset + 9] ?? 0, rectInstances[offset + 10] ?? 0, borderAlpha)
    context.fillRect(x, y, width, height)
  }
}

export function drawTextRuns(context: CanvasTextRunContext, textRuns: readonly TextQuadRun[]): void {
  for (const run of textRuns) {
    const width = run.width ?? 0
    const height = run.height ?? 0
    const clipX = run.clipX ?? run.x
    const clipY = run.clipY ?? run.y
    const clipWidth = run.clipWidth ?? width
    const clipHeight = run.clipHeight ?? height
    if (!run.text || clipWidth <= 0 || clipHeight <= 0) {
      continue
    }
    context.save()
    context.beginPath()
    context.rect(clipX, clipY, clipWidth, clipHeight)
    context.clip()
    context.fillStyle = run.color ?? '#111827'
    context.font = run.font?.trim()
      ? run.font
      : `400 ${run.fontSize ?? workbookFontPointSizeToCssPx(WORKBOOK_DEFAULT_FONT_SIZE)}px ${WORKBOOK_FONT_SANS}`
    context.textBaseline = 'middle'
    context.textAlign = run.align ?? 'left'
    const textX = run.align === 'right' ? run.x + width - 6 : run.align === 'center' ? run.x + width / 2 : run.x + 6
    const textY = run.y + height / 2
    context.fillText(run.text, textX, textY)
    if (run.underline || run.strike) {
      const metrics = context.measureText(run.text)
      const lineWidth = Math.min(metrics.width, Math.max(0, clipWidth))
      const startX = run.align === 'right' ? textX - lineWidth : run.align === 'center' ? textX - lineWidth / 2 : textX
      const lineY = run.strike ? textY : run.y + height - 4
      context.strokeStyle = run.color ?? '#111827'
      context.lineWidth = 1
      context.beginPath()
      context.moveTo(startX, lineY)
      context.lineTo(startX + lineWidth, lineY)
      context.stroke()
    }
    context.restore()
  }
}

function drawPane(
  context: CanvasRenderingContext2D,
  pane: FallbackPane,
  scrollSnapshot: WorkbookGridScrollSnapshot,
  drawText: boolean,
): void {
  if (pane.frame.width <= 0 || pane.frame.height <= 0) {
    return
  }
  const offset = resolvePaneRenderOffset(pane, scrollSnapshot)
  const rectInstances = 'tile' in pane ? pane.tile.rectInstances : pane.rectInstances
  const rectCount = 'tile' in pane ? pane.tile.rectCount : pane.rectCount
  context.save()
  context.beginPath()
  context.rect(pane.frame.x, pane.frame.y, pane.frame.width, pane.frame.height)
  context.clip()
  context.translate(pane.frame.x + offset.x, pane.frame.y + offset.y)
  drawRectInstances(context, rectInstances, rectCount)
  if (drawText) {
    drawTextRuns(context, 'tile' in pane ? pane.tile.textRuns : pane.textRuns)
  }
  context.restore()
}

function drawOverlay(context: CanvasRenderingContext2D, overlay: DynamicGridOverlayBatchV3 | null): void {
  if (!overlay || overlay.rectCount <= 0) {
    return
  }
  context.save()
  drawRectInstances(context, overlay.rectInstances, overlay.rectCount)
  context.restore()
}

export function resolveWorkbookPaneCanvasFallbackFrame(input: {
  readonly cameraStore?: GridCameraStore | null | undefined
  readonly geometry: GridGeometrySnapshot | null
  readonly layer?: 'fallback' | 'grid-floor' | undefined
  readonly overlay: DynamicGridOverlayBatchV3 | null
  readonly overlayBuilder?: ((geometry: GridGeometrySnapshot) => DynamicGridOverlayBatchV3 | null | undefined) | null | undefined
  readonly scrollTransformStore: WorkbookGridScrollStore | null
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}): {
  readonly geometry: GridGeometrySnapshot | null
  readonly overlay: DynamicGridOverlayBatchV3 | null
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
} {
  const geometry = resolveWorkbookPaneRendererGeometryV3({
    cameraStore: input.cameraStore,
    geometry: input.geometry,
  })
  const shouldDrawOverlay = input.layer !== 'grid-floor'
  return {
    geometry,
    overlay: shouldDrawOverlay ? (input.overlayBuilder && geometry ? (input.overlayBuilder(geometry) ?? null) : input.overlay) : null,
    scrollSnapshot: resolveTypeGpuV3DrawScrollSnapshot({
      fallback: input.scrollTransformStore?.getSnapshot() ?? { tx: 0, ty: 0 },
      geometry,
      panes: input.tilePanes,
    }),
  }
}

export const WorkbookPaneCanvasFallbackV3 = memo(function WorkbookPaneCanvasFallbackV3({
  active,
  cameraStore,
  drawText = true,
  geometry,
  headerPanes,
  host,
  layer = 'fallback',
  overlay,
  overlayBuilder,
  scrollTransformStore,
  tilePanes,
}: WorkbookPaneCanvasFallbackV3Props) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null)

  const draw = useCallback(() => {
    const canvas = canvasRef.current
    if (!active || !canvas || !host) {
      return
    }
    const width = Math.max(0, Math.floor(host.clientWidth))
    const height = Math.max(0, Math.floor(host.clientHeight))
    if (width <= 0 || height <= 0) {
      return
    }
    const dpr = Math.max(1, window.devicePixelRatio || 1)
    const pixelWidth = Math.max(1, Math.ceil(width * dpr))
    const pixelHeight = Math.max(1, Math.ceil(height * dpr))
    if (canvas.width !== pixelWidth) {
      canvas.width = pixelWidth
    }
    if (canvas.height !== pixelHeight) {
      canvas.height = pixelHeight
    }
    const context = canvas.getContext('2d')
    if (!context) {
      return
    }
    context.setTransform(dpr, 0, 0, dpr, 0, 0)
    context.clearRect(0, 0, width, height)
    const frame = resolveWorkbookPaneCanvasFallbackFrame({
      cameraStore,
      geometry,
      layer,
      overlay,
      overlayBuilder,
      scrollTransformStore,
      tilePanes,
    })
    tilePanes.forEach((pane) => drawPane(context, pane, frame.scrollSnapshot, drawText))
    headerPanes.forEach((pane) => drawPane(context, pane, frame.scrollSnapshot, drawText))
    drawOverlay(context, frame.overlay)
  }, [active, cameraStore, drawText, geometry, headerPanes, host, layer, overlay, overlayBuilder, scrollTransformStore, tilePanes])

  useLayoutEffect(() => {
    if (!active || !host) {
      return
    }
    draw()
  }, [active, draw, host])

  useEffect(() => {
    if (!active || !host) {
      return
    }
    let frame = 0
    const scheduleDraw = () => {
      if (frame !== 0) {
        return
      }
      frame = window.requestAnimationFrame(() => {
        frame = 0
        draw()
      })
    }
    const unsubscribeCamera = cameraStore?.subscribe(scheduleDraw)
    const unsubscribeScroll = scrollTransformStore?.subscribe(scheduleDraw)
    const resizeObserver = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(scheduleDraw)
    scheduleDraw()
    resizeObserver?.observe(host)
    return () => {
      unsubscribeCamera?.()
      unsubscribeScroll?.()
      resizeObserver?.disconnect()
      if (frame !== 0) {
        window.cancelAnimationFrame(frame)
      }
    }
  }, [active, cameraStore, draw, host, scrollTransformStore])

  if (!active || !host) {
    return null
  }
  const isGridFloor = layer === 'grid-floor'

  return (
    <canvas
      aria-hidden="true"
      className={isGridFloor ? 'pointer-events-none absolute inset-0 z-[12]' : 'pointer-events-none absolute inset-0 z-[5]'}
      data-pane-renderer={isGridFloor ? 'workbook-pane-renderer-v3-grid-floor' : 'workbook-pane-renderer-v3-fallback'}
      data-renderer-mode={isGridFloor ? 'canvas2d-v3-grid-floor' : 'canvas2d-v3-fallback'}
      data-testid={isGridFloor ? 'grid-pane-renderer-floor' : 'grid-pane-renderer-fallback'}
      data-v3-draw-text={drawText ? 'true' : 'false'}
      data-v3-header-pane-count={headerPanes.length}
      data-v3-header-text-run-count={headerPanes.reduce((total, pane) => total + pane.textRuns.length, 0)}
      data-v3-layer-role={layer}
      data-v3-overlay-enabled={isGridFloor ? 'false' : 'true'}
      data-v3-text-run-count={tilePanes.reduce((total, pane) => total + pane.tile.textRuns.length, 0)}
      data-v3-tile-pane-count={tilePanes.length}
      ref={canvasRef}
      style={{ contain: 'strict', height: '100%', width: '100%' }}
    />
  )
})
