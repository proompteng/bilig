import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { Rectangle } from '../gridTypes.js'
import type { WorkbookGridScrollSnapshot } from '../workbookGridScrollStore.js'
import {
  WORKBOOK_RECT_INSTANCE_LAYOUT,
  WORKBOOK_TEXT_INSTANCE_LAYOUT,
  WORKBOOK_UNIT_QUAD_LAYOUT,
  type TypeGpuRendererArtifacts,
  updateTypeGpuSurfaceUniform,
} from './typegpu-primitives.js'
import { noteGridDrawFrame, noteTypeGpuDrawCall, noteTypeGpuPaneDraw, noteTypeGpuSubmit } from '../grid-render-counters.js'
import {
  WORKBOOK_DYNAMIC_OVERLAY_CHROME_LAYER_KEY_V3,
  WORKBOOK_DYNAMIC_OVERLAY_FILL_LAYER_KEY_V3,
  ensureLayerSurfaceBindingsV3,
  resolveWorkbookHeaderLayerKeyV3,
  type TypeGpuLayerResourceCacheV3,
} from './typegpu-layer-buffer-pool.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import {
  ensureTilePlacementSurfaceBindingsV3,
  resolveWorkbookTileContentBufferKeyV3,
  resolveWorkbookTilePlacementBufferKeyV3,
  type TypeGpuTileResourceCacheV3,
} from './typegpu-tile-buffer-pool.js'

export interface TypeGpuTileDrawSurface {
  readonly width: number
  readonly height: number
  readonly pixelWidth: number
  readonly pixelHeight: number
  readonly dpr: number
}

const TYPEGPU_WORKBOOK_CLEAR_COLOR = { r: 1, g: 1, b: 1, a: 1 } as const

export function hasCompleteTypeGpuBodyTileContentV3(input: {
  readonly drawText?: boolean | undefined
  readonly surface?: TypeGpuTileDrawSurface | undefined
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
  readonly tileResources: Pick<TypeGpuTileResourceCacheV3, 'peekContent'>
}): boolean {
  if (input.tilePanes.length === 0) {
    return false
  }
  for (const pane of input.tilePanes) {
    if (pane.paneId !== 'body' && !pane.paneId.startsWith('body:')) {
      continue
    }
    if (!isPaneDrawVisible(pane)) {
      continue
    }
    if (input.surface && !resolveClampedScissorRect(pane.frame, input.surface)) {
      continue
    }
    const content = input.tileResources.peekContent(resolveWorkbookTileContentBufferKeyV3(pane))
    if (!content) {
      return false
    }
    if (content.rectCount === 0 && content.textCount === 0) {
      return false
    }
    if (
      (pane.tile.rectCount > 0 && (!content.rectHandle || content.rectCount < pane.tile.rectCount)) ||
      ((input.drawText ?? true) && pane.tile.textCount > 0 && (!content.textHandle || content.textCount === 0))
    ) {
      return false
    }
  }
  return true
}

function isPaneDrawVisible(pane: WorkbookRenderTilePaneState | GridHeaderPaneState): boolean {
  return pane.drawVisible !== false
}

export function drawTypeGpuTilePanesV3(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly drawText?: boolean | undefined
  readonly layerResources: TypeGpuLayerResourceCacheV3
  readonly tileResources: TypeGpuTileResourceCacheV3
  readonly headerPanes?: readonly GridHeaderPaneState[] | undefined
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
  readonly overlay?: DynamicGridOverlayBatchV3 | null | undefined
  readonly surface: TypeGpuTileDrawSurface
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
}): boolean {
  if (!hasCompleteTypeGpuBodyTileContentV3(input) || !hasDrawableTypeGpuBodyPaneFramesV3(input)) {
    return false
  }

  const commandEncoder = input.artifacts.device.createCommandEncoder()
  const pass = commandEncoder.beginRenderPass({
    colorAttachments: [
      {
        view: input.artifacts.context.getCurrentTexture().createView(),
        clearValue: TYPEGPU_WORKBOOK_CLEAR_COLOR,
        loadOp: 'clear',
        storeOp: 'store',
      },
    ],
  })

  input.tilePanes.forEach((pane) => {
    drawTypeGpuTilePaneRects({
      artifacts: input.artifacts,
      pane,
      pass,
      scrollSnapshot: input.scrollSnapshot,
      surface: input.surface,
      tileResources: input.tileResources,
    })
  })

  drawTypeGpuHeaderPaneRects({
    artifacts: input.artifacts,
    headerPanes: input.headerPanes ?? [],
    layerResources: input.layerResources,
    pass,
    scrollSnapshot: input.scrollSnapshot,
    surface: input.surface,
  })

  drawTypeGpuOverlayLayer({
    artifacts: input.artifacts,
    layerKey: WORKBOOK_DYNAMIC_OVERLAY_FILL_LAYER_KEY_V3,
    layerResources: input.layerResources,
    pass,
    surface: input.surface,
  })

  if (input.drawText ?? true) {
    input.tilePanes.forEach((pane) => {
      drawTypeGpuTilePaneText({
        artifacts: input.artifacts,
        pane,
        pass,
        scrollSnapshot: input.scrollSnapshot,
        surface: input.surface,
        tileResources: input.tileResources,
      })
    })
    drawTypeGpuHeaderPaneText({
      artifacts: input.artifacts,
      headerPanes: input.headerPanes ?? [],
      layerResources: input.layerResources,
      pass,
      scrollSnapshot: input.scrollSnapshot,
      surface: input.surface,
    })
  }

  drawTypeGpuOverlayLayer({
    artifacts: input.artifacts,
    layerKey: WORKBOOK_DYNAMIC_OVERLAY_CHROME_LAYER_KEY_V3,
    layerResources: input.layerResources,
    pass,
    surface: input.surface,
  })

  pass.end()
  input.artifacts.device.queue.submit([commandEncoder.finish()])
  noteTypeGpuSubmit()
  noteGridDrawFrame(performance.now())
  return true
}

export function hasDrawableTypeGpuBodyPaneFramesV3(input: {
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
  readonly surface: TypeGpuTileDrawSurface
}): boolean {
  let hasDrawableBodyPane = false
  for (const pane of input.tilePanes) {
    if (pane.paneId !== 'body' && !pane.paneId.startsWith('body:')) {
      continue
    }
    if (!isPaneDrawVisible(pane)) {
      continue
    }
    if (resolveClampedScissorRect(pane.frame, input.surface)) {
      hasDrawableBodyPane = true
    }
  }
  return hasDrawableBodyPane
}

function drawTypeGpuTilePaneRects(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly pane: WorkbookRenderTilePaneState
  readonly pass: GPURenderPassEncoder
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly surface: TypeGpuTileDrawSurface
  readonly tileResources: TypeGpuTileResourceCacheV3
}): void {
  if (!isPaneDrawVisible(input.pane)) {
    return
  }
  const scissorRect = resolveClampedScissorRect(input.pane.frame, input.surface)
  if (!scissorRect) {
    return
  }
  const content = input.tileResources.peekContent(resolveWorkbookTileContentBufferKeyV3(input.pane))
  const placement = input.tileResources.getPlacement(resolveWorkbookTilePlacementBufferKeyV3(input.pane))
  if (!content || content.rectCount <= 0 || !content.rectHandle) {
    return
  }
  ensureTilePlacementSurfaceBindingsV3(input.artifacts, placement)
  updateTypeGpuSurfaceUniform(
    placement.surfaceUniform!,
    input.surface,
    resolvePaneOrigin(input.pane),
    resolvePaneRenderOffset(input.pane, input.scrollSnapshot),
  )
  if (!placement.surfaceBindGroup) {
    return
  }
  input.pass.setScissorRect(scissorRect.x, scissorRect.y, scissorRect.width, scissorRect.height)
  const rectRenderer = input.artifacts.rectPipeline.with(input.pass).with(placement.surfaceBindGroup)
  rectRenderer
    .with(WORKBOOK_UNIT_QUAD_LAYOUT, input.artifacts.quadBuffer)
    .with(WORKBOOK_RECT_INSTANCE_LAYOUT, content.rectHandle.buffer)
    .draw(6, content.rectCount)
  noteTypeGpuDrawCall(1)
  noteTypeGpuPaneDraw(1)
}

function drawTypeGpuTilePaneText(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly pane: WorkbookRenderTilePaneState
  readonly pass: GPURenderPassEncoder
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly surface: TypeGpuTileDrawSurface
  readonly tileResources: TypeGpuTileResourceCacheV3
}): void {
  if (!isPaneDrawVisible(input.pane)) {
    return
  }
  const scissorRect = resolveClampedScissorRect(input.pane.frame, input.surface)
  if (!scissorRect) {
    return
  }
  const content = input.tileResources.peekContent(resolveWorkbookTileContentBufferKeyV3(input.pane))
  const placement = input.tileResources.getPlacement(resolveWorkbookTilePlacementBufferKeyV3(input.pane))
  if (!content || content.textCount <= 0 || !content.textHandle) {
    return
  }
  ensureTilePlacementSurfaceBindingsV3(input.artifacts, placement)
  updateTypeGpuSurfaceUniform(
    placement.surfaceUniform!,
    input.surface,
    resolvePaneOrigin(input.pane),
    resolvePaneRenderOffset(input.pane, input.scrollSnapshot),
  )
  if (!placement.textBindGroup) {
    return
  }
  input.pass.setScissorRect(scissorRect.x, scissorRect.y, scissorRect.width, scissorRect.height)
  const textRenderer = input.artifacts.textPipeline.with(input.pass).with(placement.textBindGroup)
  textRenderer
    .with(WORKBOOK_UNIT_QUAD_LAYOUT, input.artifacts.quadBuffer)
    .with(WORKBOOK_TEXT_INSTANCE_LAYOUT, content.textHandle.buffer)
    .draw(6, content.textCount)
  noteTypeGpuDrawCall(1)
  noteTypeGpuPaneDraw(1)
}

function drawTypeGpuOverlayLayer(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly layerKey: string
  readonly layerResources: TypeGpuLayerResourceCacheV3
  readonly pass: GPURenderPassEncoder
  readonly surface: TypeGpuTileDrawSurface
}): void {
  const overlayCache = input.layerResources.peek(input.layerKey)
  if (!overlayCache?.rectHandle || overlayCache.rectCount <= 0) {
    return
  }

  input.pass.setScissorRect(0, 0, input.surface.pixelWidth, input.surface.pixelHeight)
  ensureLayerSurfaceBindingsV3(input.artifacts, overlayCache)
  updateTypeGpuSurfaceUniform(overlayCache.surfaceUniform!, input.surface, { x: 0, y: 0 }, { x: 0, y: 0 })
  if (!overlayCache.surfaceBindGroup) {
    return
  }
  const rectRenderer = input.artifacts.rectPipeline.with(input.pass).with(overlayCache.surfaceBindGroup)
  rectRenderer
    .with(WORKBOOK_UNIT_QUAD_LAYOUT, input.artifacts.quadBuffer)
    .with(WORKBOOK_RECT_INSTANCE_LAYOUT, overlayCache.rectHandle.buffer)
    .draw(6, overlayCache.rectCount)
  noteTypeGpuDrawCall(1)
  noteTypeGpuPaneDraw(1)
}

function drawTypeGpuHeaderPaneRects(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly layerResources: TypeGpuLayerResourceCacheV3
  readonly pass: GPURenderPassEncoder
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly surface: TypeGpuTileDrawSurface
}): void {
  input.headerPanes.forEach((pane) => {
    drawTypeGpuHeaderPaneLayer({
      ...input,
      pane,
      phase: 'rects',
    })
  })
}

function drawTypeGpuHeaderPaneText(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly layerResources: TypeGpuLayerResourceCacheV3
  readonly pass: GPURenderPassEncoder
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly surface: TypeGpuTileDrawSurface
}): void {
  input.headerPanes.forEach((pane) => {
    drawTypeGpuHeaderPaneLayer({
      ...input,
      pane,
      phase: 'text',
    })
  })
}

function drawTypeGpuHeaderPaneLayer(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly layerResources: TypeGpuLayerResourceCacheV3
  readonly pane: GridHeaderPaneState
  readonly pass: GPURenderPassEncoder
  readonly phase: 'rects' | 'text'
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly surface: TypeGpuTileDrawSurface
}): void {
  if (!isPaneDrawVisible(input.pane)) {
    return
  }
  const scissorRect = resolveClampedScissorRect(input.pane.frame, input.surface)
  if (!scissorRect) {
    return
  }
  const paneCache = input.layerResources.peek(resolveWorkbookHeaderLayerKeyV3(input.pane))
  if (!paneCache) {
    return
  }
  const hasDrawableRects = input.phase === 'rects' && paneCache.rectCount > 0 && paneCache.rectHandle
  const hasDrawableText = input.phase === 'text' && paneCache.textCount > 0 && paneCache.textHandle
  if (!hasDrawableRects && !hasDrawableText) {
    return
  }
  ensureLayerSurfaceBindingsV3(input.artifacts, paneCache)
  updateTypeGpuSurfaceUniform(
    paneCache.surfaceUniform!,
    input.surface,
    resolvePaneOrigin(input.pane),
    resolvePaneRenderOffset(input.pane, input.scrollSnapshot),
  )
  input.pass.setScissorRect(scissorRect.x, scissorRect.y, scissorRect.width, scissorRect.height)
  if (input.phase === 'rects' && paneCache.rectHandle && paneCache.surfaceBindGroup) {
    const rectRenderer = input.artifacts.rectPipeline.with(input.pass).with(paneCache.surfaceBindGroup)
    rectRenderer
      .with(WORKBOOK_UNIT_QUAD_LAYOUT, input.artifacts.quadBuffer)
      .with(WORKBOOK_RECT_INSTANCE_LAYOUT, paneCache.rectHandle.buffer)
      .draw(6, paneCache.rectCount)
    noteTypeGpuDrawCall(1)
    noteTypeGpuPaneDraw(1)
    return
  }
  if (input.phase === 'text' && paneCache.textHandle && paneCache.textBindGroup) {
    const textRenderer = input.artifacts.textPipeline.with(input.pass).with(paneCache.textBindGroup)
    textRenderer
      .with(WORKBOOK_UNIT_QUAD_LAYOUT, input.artifacts.quadBuffer)
      .with(WORKBOOK_TEXT_INSTANCE_LAYOUT, paneCache.textHandle.buffer)
      .draw(6, paneCache.textCount)
    noteTypeGpuDrawCall(1)
    noteTypeGpuPaneDraw(1)
  }
}

function resolveClampedScissorRect(
  frame: Rectangle,
  surface: TypeGpuTileDrawSurface,
): { x: number; y: number; width: number; height: number } | null {
  const dpr = surface.dpr
  const x0 = Math.max(0, Math.min(surface.pixelWidth, Math.floor(frame.x * dpr)))
  const y0 = Math.max(0, Math.min(surface.pixelHeight, Math.floor(frame.y * dpr)))
  const x1 = Math.max(x0, Math.min(surface.pixelWidth, Math.ceil((frame.x + frame.width) * dpr)))
  const y1 = Math.max(y0, Math.min(surface.pixelHeight, Math.ceil((frame.y + frame.height) * dpr)))
  if (x0 >= x1 || y0 >= y1) return null
  return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 }
}

function resolvePaneOrigin(pane: { readonly frame: Rectangle }): { x: number; y: number } {
  return {
    x: pane.frame.x,
    y: pane.frame.y,
  }
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
): { x: number; y: number } {
  const renderTx = scrollSnapshot.renderTx ?? scrollSnapshot.tx
  const renderTy = scrollSnapshot.renderTy ?? scrollSnapshot.ty
  return {
    x: pane.contentOffset.x - (pane.scrollAxes.x ? renderTx : 0),
    y: pane.contentOffset.y - (pane.scrollAxes.y ? renderTy : 0),
  }
}
