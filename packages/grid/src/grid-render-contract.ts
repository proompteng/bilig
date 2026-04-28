import type { Viewport } from '@bilig/protocol'

export interface GridAxisSnapshot {
  readonly index: number
  readonly offset: number
  readonly size: number
}

export interface GridCameraSnapshot {
  readonly scrollLeft: number
  readonly scrollTop: number
  readonly tx: number
  readonly ty: number
  readonly visibleViewport: Viewport
  readonly residentViewport: Viewport
  readonly viewportWidth: number
  readonly viewportHeight: number
  readonly dpr: number
  readonly velocityX: number
  readonly velocityY: number
  readonly updatedAt: number
}

export interface GridRenderFrame {
  readonly camera: GridCameraSnapshot
  readonly inputAt: number
  readonly frameAt: number
}

export interface GridGpuCounters {
  readonly configureCount: number
  readonly submitCount: number
  readonly drawCalls: number
  readonly paneDraws: number
  readonly uniformWriteBytes: number
  readonly vertexUploadBytes: number
  readonly overlayUploadBytes: number
  readonly bufferAllocations: number
  readonly bufferAllocationBytes: number
  readonly atlasUploadBytes: number
  readonly atlasDirtyPages: number
  readonly atlasDirtyPageUploadBytes: number
  readonly surfaceResizes: number
  readonly tileMisses: number
  readonly tileCacheEvictions: number
  readonly tileCacheEntriesScanned: number
  readonly tileCacheSorts: number
  readonly tileCacheStaleHits: number
  readonly tileCacheStaleLookups: number
  readonly tileCacheVisibleMarks: number
  readonly textAtlasGeometryResyncs: number
  readonly textAtlasGeometryRetries: number
  readonly textGlyphDependencies: number
  readonly textPageDependencies: number
  readonly textRunPayloadRebuilds: number
  readonly textRunPayloadReuses: number
}

export interface GridRenderStats {
  readonly inputToDrawMs: readonly number[]
  readonly gpu: GridGpuCounters
}
