import type { GridGpuCounters } from './grid-render-contract.js'

export interface GridRenderDebugSnapshot {
  readonly frameMs: number
  readonly inputToDrawMs: number
  readonly gpu: GridGpuCounters
}

export function formatRenderDebugHud(snapshot: GridRenderDebugSnapshot): readonly string[] {
  return [
    `frame ${snapshot.frameMs.toFixed(2)}ms`,
    `input ${snapshot.inputToDrawMs.toFixed(2)}ms`,
    `submits ${snapshot.gpu.submitCount}`,
    `draws ${snapshot.gpu.drawCalls}`,
    `uploads ${(snapshot.gpu.vertexUploadBytes / 1024).toFixed(1)}KB`,
    `overlay ${(snapshot.gpu.overlayUploadBytes / 1024).toFixed(1)}KB`,
    `allocs ${snapshot.gpu.bufferAllocations}`,
    `atlas ${(snapshot.gpu.atlasUploadBytes / 1024).toFixed(1)}KB`,
    `tiles miss ${snapshot.gpu.tileMisses}`,
  ]
}

export function isRenderDebugSnapshotInsideBudget(input: {
  readonly snapshot: GridRenderDebugSnapshot
  readonly maxFrameMs: number
  readonly maxInputToDrawMs: number
  readonly maxVertexUploadBytes: number
  readonly maxBufferAllocations: number
  readonly maxTileMisses: number
}): boolean {
  return (
    input.snapshot.frameMs <= input.maxFrameMs &&
    input.snapshot.inputToDrawMs <= input.maxInputToDrawMs &&
    input.snapshot.gpu.vertexUploadBytes <= input.maxVertexUploadBytes &&
    input.snapshot.gpu.bufferAllocations <= input.maxBufferAllocations &&
    input.snapshot.gpu.tileMisses <= input.maxTileMisses
  )
}
