import type { Rectangle } from '../gridTypes.js'
import type { WorkbookGridScrollSnapshot } from '../workbookGridScrollStore.js'
import type { WorkbookPaneBufferCache } from './pane-buffer-cache.js'
import type { WorkbookRenderPaneState } from './pane-scene-types.js'
import {
  WORKBOOK_RECT_INSTANCE_LAYOUT,
  WORKBOOK_TEXT_INSTANCE_LAYOUT,
  WORKBOOK_UNIT_QUAD_LAYOUT,
  type TypeGpuRendererArtifacts,
  updateTypeGpuSurfaceUniform,
} from './typegpu-backend.js'
import { noteGridDrawFrame, noteTypeGpuDrawCall, noteTypeGpuPaneDraw, noteTypeGpuSubmit } from './grid-render-counters.js'
import { WORKBOOK_DYNAMIC_OVERLAY_BUFFER_KEY, ensurePaneSurfaceBindings, resolveWorkbookPaneBufferKey } from './typegpu-buffer-pool.js'
import type { DynamicGridOverlayBatchV3 } from '../renderer-v3/dynamic-overlay-batch.js'

export interface TypeGpuDrawSurface {
  readonly width: number
  readonly height: number
  readonly pixelWidth: number
  readonly pixelHeight: number
  readonly dpr: number
}

export function drawTypeGpuPanes(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly panes: readonly WorkbookRenderPaneState[]
  readonly overlay?: DynamicGridOverlayBatchV3 | null | undefined
  readonly surface: TypeGpuDrawSurface
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
}): void {
  const commandEncoder = input.artifacts.device.createCommandEncoder()
  const pass = commandEncoder.beginRenderPass({
    colorAttachments: [
      {
        view: input.artifacts.context.getCurrentTexture().createView(),
        clearValue: { r: 0, g: 0, b: 0, a: 0 },
        loadOp: 'clear',
        storeOp: 'store',
      },
    ],
  })

  input.panes.forEach((pane) => {
    const paneCache = input.paneBuffers.get(resolveWorkbookPaneBufferKey(pane))
    const scissorRect = resolveClampedScissorRect(pane.frame, input.surface)
    if (!scissorRect) {
      return
    }

    pass.setScissorRect(scissorRect.x, scissorRect.y, scissorRect.width, scissorRect.height)
    const paneOrigin = resolvePaneOrigin(pane)
    const paneRenderOffset = resolvePaneRenderOffset(pane, input.scrollSnapshot)

    if (paneCache.rectCount > 0 || paneCache.textCount > 0) {
      ensurePaneSurfaceBindings(input.artifacts, paneCache)
      updateTypeGpuSurfaceUniform(paneCache.surfaceUniform!, input.surface, paneOrigin, paneRenderOffset)
    }

    if (paneCache.rectCount > 0 && paneCache.rectBuffer && paneCache.surfaceBindGroup) {
      const rectRenderer = input.artifacts.rectPipeline.with(pass).with(paneCache.surfaceBindGroup)
      rectRenderer
        .with(WORKBOOK_UNIT_QUAD_LAYOUT, input.artifacts.quadBuffer)
        .with(WORKBOOK_RECT_INSTANCE_LAYOUT, paneCache.rectBuffer)
        .draw(6, paneCache.rectCount)
      noteTypeGpuDrawCall(1)
    }

    if (paneCache.textCount > 0 && paneCache.textBuffer && paneCache.textBindGroup) {
      const textRenderer = input.artifacts.textPipeline.with(pass).with(paneCache.textBindGroup)
      textRenderer
        .with(WORKBOOK_UNIT_QUAD_LAYOUT, input.artifacts.quadBuffer)
        .with(WORKBOOK_TEXT_INSTANCE_LAYOUT, paneCache.textBuffer)
        .draw(6, paneCache.textCount)
      noteTypeGpuDrawCall(1)
    }
    noteTypeGpuPaneDraw(1)
  })

  drawTypeGpuOverlay({
    artifacts: input.artifacts,
    overlay: input.overlay ?? null,
    paneBuffers: input.paneBuffers,
    pass,
    surface: input.surface,
  })

  pass.end()
  input.artifacts.device.queue.submit([commandEncoder.finish()])
  noteTypeGpuSubmit()
  noteGridDrawFrame(performance.now())
}

function drawTypeGpuOverlay(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly pass: GPURenderPassEncoder
  readonly overlay: DynamicGridOverlayBatchV3 | null
  readonly surface: TypeGpuDrawSurface
}): void {
  if (!input.overlay || input.overlay.rectCount === 0) {
    return
  }
  const overlayCache = input.paneBuffers.peek(WORKBOOK_DYNAMIC_OVERLAY_BUFFER_KEY)
  if (!overlayCache?.rectBuffer || overlayCache.rectCount <= 0) {
    return
  }

  input.pass.setScissorRect(0, 0, input.surface.pixelWidth, input.surface.pixelHeight)
  ensurePaneSurfaceBindings(input.artifacts, overlayCache)
  updateTypeGpuSurfaceUniform(overlayCache.surfaceUniform!, input.surface, { x: 0, y: 0 }, { x: 0, y: 0 })
  if (!overlayCache.surfaceBindGroup) {
    return
  }
  const rectRenderer = input.artifacts.rectPipeline.with(input.pass).with(overlayCache.surfaceBindGroup)
  rectRenderer
    .with(WORKBOOK_UNIT_QUAD_LAYOUT, input.artifacts.quadBuffer)
    .with(WORKBOOK_RECT_INSTANCE_LAYOUT, overlayCache.rectBuffer)
    .draw(6, overlayCache.rectCount)
  noteTypeGpuDrawCall(1)
  noteTypeGpuPaneDraw(1)
}

function resolveClampedScissorRect(
  frame: Rectangle,
  surface: TypeGpuDrawSurface,
): { x: number; y: number; width: number; height: number } | null {
  const dpr = surface.dpr
  const x0 = Math.max(0, Math.min(surface.pixelWidth, Math.floor(frame.x * dpr)))
  const y0 = Math.max(0, Math.min(surface.pixelHeight, Math.floor(frame.y * dpr)))
  const x1 = Math.max(x0, Math.min(surface.pixelWidth, Math.ceil((frame.x + frame.width) * dpr)))
  const y1 = Math.max(y0, Math.min(surface.pixelHeight, Math.ceil((frame.y + frame.height) * dpr)))
  if (x0 >= x1 || y0 >= y1) return null
  return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 }
}

function resolvePaneOrigin(pane: WorkbookRenderPaneState): { x: number; y: number } {
  return {
    x: pane.frame.x,
    y: pane.frame.y,
  }
}

function resolvePaneRenderOffset(
  pane: WorkbookRenderPaneState,
  scrollSnapshot: {
    readonly tx: number
    readonly ty: number
    readonly renderTx?: number | undefined
    readonly renderTy?: number | undefined
  },
): { x: number; y: number } {
  const renderTx = scrollSnapshot.renderTx ?? scrollSnapshot.tx
  const renderTy = scrollSnapshot.renderTy ?? scrollSnapshot.ty
  return {
    x: pane.contentOffset.x - (pane.scrollAxes.x ? renderTx : 0),
    y: pane.contentOffset.y - (pane.scrollAxes.y ? renderTy : 0),
  }
}
