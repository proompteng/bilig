import type { GridGpuCounters, GridResidentTileKey } from './grid-render-contract.js'

type ScrollPerfCounterSink = Partial<{
  noteTypeGpuConfigure: () => void
  noteTypeGpuSubmit: () => void
  noteTypeGpuDrawCall: (count: number) => void
  noteTypeGpuPaneDraw: (count: number) => void
  noteTypeGpuUniformWrite: (bytes: number, label: string) => void
  noteTypeGpuBufferWrite: (bytes: number, label: string) => void
  noteTypeGpuOverlayWrite: (bytes: number) => void
  noteTypeGpuBufferAllocation: (bytes: number, label: string) => void
  noteTypeGpuAtlasUpload: (bytes: number) => void
  noteTypeGpuSurfaceResize: (width: number, height: number, dpr: number) => void
  noteTypeGpuTileMiss: (tileKey: GridResidentTileKey | string) => void
  noteTypeGpuTileCacheEviction: (count: number) => void
  noteTypeGpuTileCacheSort: (count: number) => void
  noteTypeGpuTileCacheStaleLookup: (scannedEntries: number, hit: boolean) => void
  noteTypeGpuTileCacheVisibleMark: (count: number) => void
  noteTypeGpuScenePacketApplied: (packetKey: GridResidentTileKey | string) => void
  noteGridScrollInput: (timestamp: number) => void
  noteGridDrawFrame: (timestamp: number) => void
}>

export const EMPTY_GRID_GPU_COUNTERS: GridGpuCounters = Object.freeze({
  atlasUploadBytes: 0,
  bufferAllocationBytes: 0,
  bufferAllocations: 0,
  configureCount: 0,
  drawCalls: 0,
  paneDraws: 0,
  scenePacketsApplied: 0,
  submitCount: 0,
  surfaceResizes: 0,
  tileMisses: 0,
  tileCacheEvictions: 0,
  tileCacheEntriesScanned: 0,
  tileCacheSorts: 0,
  tileCacheStaleHits: 0,
  tileCacheStaleLookups: 0,
  tileCacheVisibleMarks: 0,
  uniformWriteBytes: 0,
  vertexUploadBytes: 0,
  overlayUploadBytes: 0,
})

function getCounterSink(): ScrollPerfCounterSink | null {
  if (typeof window === 'undefined') {
    return null
  }
  return (window as Window & { __biligScrollPerf?: ScrollPerfCounterSink }).__biligScrollPerf ?? null
}

export function noteTypeGpuConfigure(): void {
  getCounterSink()?.noteTypeGpuConfigure?.()
}

export function noteTypeGpuSubmit(): void {
  getCounterSink()?.noteTypeGpuSubmit?.()
}

export function noteTypeGpuDrawCall(count = 1): void {
  getCounterSink()?.noteTypeGpuDrawCall?.(count)
}

export function noteTypeGpuPaneDraw(count = 1): void {
  getCounterSink()?.noteTypeGpuPaneDraw?.(count)
}

export function noteTypeGpuUniformWrite(bytes: number, label: string): void {
  getCounterSink()?.noteTypeGpuUniformWrite?.(bytes, label)
}

export function noteTypeGpuBufferWrite(bytes: number, label: string): void {
  getCounterSink()?.noteTypeGpuBufferWrite?.(bytes, label)
  if (label.startsWith('overlay:')) {
    getCounterSink()?.noteTypeGpuOverlayWrite?.(bytes)
  }
}

export function noteTypeGpuBufferAllocation(bytes: number, label: string): void {
  getCounterSink()?.noteTypeGpuBufferAllocation?.(bytes, label)
}

export function noteTypeGpuAtlasUpload(bytes: number): void {
  getCounterSink()?.noteTypeGpuAtlasUpload?.(bytes)
}

export function noteTypeGpuSurfaceResize(width: number, height: number, dpr: number): void {
  getCounterSink()?.noteTypeGpuSurfaceResize?.(width, height, dpr)
}

export function noteTypeGpuTileMiss(tileKey: GridResidentTileKey | string): void {
  getCounterSink()?.noteTypeGpuTileMiss?.(tileKey)
}

export function noteTypeGpuTileCacheEviction(count = 1): void {
  getCounterSink()?.noteTypeGpuTileCacheEviction?.(count)
}

export function noteTypeGpuTileCacheSort(count = 1): void {
  getCounterSink()?.noteTypeGpuTileCacheSort?.(count)
}

export function noteTypeGpuTileCacheStaleLookup(scannedEntries: number, hit: boolean): void {
  getCounterSink()?.noteTypeGpuTileCacheStaleLookup?.(scannedEntries, hit)
}

export function noteTypeGpuTileCacheVisibleMark(count: number): void {
  getCounterSink()?.noteTypeGpuTileCacheVisibleMark?.(count)
}

export function noteTypeGpuScenePacketApplied(packetKey: GridResidentTileKey | string): void {
  getCounterSink()?.noteTypeGpuScenePacketApplied?.(packetKey)
}

export function noteGridScrollInput(timestamp = performance.now()): void {
  getCounterSink()?.noteGridScrollInput?.(timestamp)
}

export function noteGridDrawFrame(timestamp = performance.now()): void {
  getCounterSink()?.noteGridDrawFrame?.(timestamp)
}
