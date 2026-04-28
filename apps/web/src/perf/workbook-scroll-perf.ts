export interface WorkbookScrollPerfFixture {
  readonly id: string
  readonly materializedCellCount: number
  readonly sheetName: string
}

interface WorkbookScrollPerfCounters {
  viewportSubscriptions: number
  fullPatches: number
  fullPatchBroadcasts: Record<string, number>
  damagePatches: number
  damageCells: number
  rendererDeltaBatches: number
  rendererDeltaMutations: number
  rendererDeltaApplyMs: number
  dirtyTilesMarked: number
  rendererTileInterestBatches: number
  rendererTileExactHits: number
  rendererTileStaleHits: number
  rendererTileMisses: number
  rendererVisibleDirtyTiles: number
  rendererWarmDirtyTiles: number
  visibleWindowChanges: number
  headerPaneBuilds: number
  reactCommits: number
  canvasSurfaceMounts: number
  domSurfaceMounts: number
  canvasPaints: Record<string, number>
  surfaceCommits: Record<string, number>
  typeGpuConfigures: number
  typeGpuSubmits: number
  typeGpuDrawCalls: number
  typeGpuPaneDraws: number
  typeGpuUniformWriteBytes: number
  typeGpuVertexUploadBytes: number
  typeGpuOverlayUploadBytes: number
  typeGpuBufferAllocations: number
  typeGpuBufferAllocationBytes: number
  typeGpuAtlasUploadBytes: number
  typeGpuAtlasDirtyPages: number
  typeGpuAtlasDirtyPageUploadBytes: number
  typeGpuSurfaceResizes: number
  typeGpuTileMisses: number
  typeGpuTileCacheEvictions: number
  typeGpuTileCacheEntriesScanned: number
  typeGpuTileCacheSorts: number
  typeGpuTileCacheStaleHits: number
  typeGpuTileCacheStaleLookups: number
  typeGpuTileCacheVisibleMarks: number
}

interface WorkbookScrollPerfSamples {
  readonly frameMs: number[]
  readonly longTasksMs: number[]
  readonly inputToDrawMs: number[]
}

interface WorkbookScrollPerfSummary {
  readonly min: number
  readonly median: number
  readonly p95: number
  readonly p99: number
  readonly max: number
}

export interface WorkbookScrollPerfReport {
  readonly workload: string
  readonly fixture: WorkbookScrollPerfFixture | null
  readonly samples: WorkbookScrollPerfSamples
  readonly summary: {
    readonly frameMs: WorkbookScrollPerfSummary
    readonly inputToDrawMs: WorkbookScrollPerfSummary
    readonly longTasksMs: WorkbookScrollPerfSummary
  }
  readonly counters: WorkbookScrollPerfCounters
}

type BenchmarkState = 'idle' | 'loading' | 'ready' | 'error'
const WARMUP_FRAME_COUNT = 12

class WorkbookScrollPerfCollector {
  private readonly totalCounters: WorkbookScrollPerfCounters = {
    viewportSubscriptions: 0,
    fullPatches: 0,
    fullPatchBroadcasts: {},
    damagePatches: 0,
    damageCells: 0,
    dirtyTilesMarked: 0,
    rendererDeltaApplyMs: 0,
    rendererDeltaBatches: 0,
    rendererDeltaMutations: 0,
    rendererTileExactHits: 0,
    rendererTileInterestBatches: 0,
    rendererTileMisses: 0,
    rendererTileStaleHits: 0,
    rendererVisibleDirtyTiles: 0,
    rendererWarmDirtyTiles: 0,
    visibleWindowChanges: 0,
    headerPaneBuilds: 0,
    reactCommits: 0,
    canvasSurfaceMounts: 0,
    domSurfaceMounts: 0,
    canvasPaints: {},
    surfaceCommits: {},
    typeGpuAtlasUploadBytes: 0,
    typeGpuAtlasDirtyPages: 0,
    typeGpuAtlasDirtyPageUploadBytes: 0,
    typeGpuBufferAllocationBytes: 0,
    typeGpuBufferAllocations: 0,
    typeGpuConfigures: 0,
    typeGpuDrawCalls: 0,
    typeGpuPaneDraws: 0,
    typeGpuSubmits: 0,
    typeGpuSurfaceResizes: 0,
    typeGpuTileMisses: 0,
    typeGpuTileCacheEvictions: 0,
    typeGpuTileCacheEntriesScanned: 0,
    typeGpuTileCacheSorts: 0,
    typeGpuTileCacheStaleHits: 0,
    typeGpuTileCacheStaleLookups: 0,
    typeGpuTileCacheVisibleMarks: 0,
    typeGpuUniformWriteBytes: 0,
    typeGpuVertexUploadBytes: 0,
    typeGpuOverlayUploadBytes: 0,
  }
  private baselineCounters: WorkbookScrollPerfCounters | null = null
  private frameSamples: number[] = []
  private longTaskSamples: number[] = []
  private inputToDrawSamples: number[] = []
  private workload = 'idle'
  private fixture: WorkbookScrollPerfFixture | null = null
  private benchmarkState: BenchmarkState = 'idle'
  private benchmarkError: string | null = null
  private rafHandle: number | null = null
  private lastFrameAt: number | null = null
  private lastScrollInputAt: number | null = null
  private observer: PerformanceObserver | null = null
  private warmupFramesRemaining = 0

  getBenchmarkState(): { state: BenchmarkState; error: string | null; fixture: WorkbookScrollPerfFixture | null } {
    return {
      state: this.benchmarkState,
      error: this.benchmarkError,
      fixture: this.fixture,
    }
  }

  setBenchmarkState(state: BenchmarkState, error: string | null = null): void {
    this.benchmarkState = state
    this.benchmarkError = error
  }

  setFixture(fixture: WorkbookScrollPerfFixture): void {
    this.fixture = fixture
    this.benchmarkState = 'ready'
    this.benchmarkError = null
  }

  noteViewportSubscription(): void {
    this.totalCounters.viewportSubscriptions += 1
  }

  noteViewportPatch(input: { full: boolean; damageCount: number }): void {
    if (input.full) {
      this.totalCounters.fullPatches += 1
      return
    }
    this.totalCounters.damagePatches += 1
    this.totalCounters.damageCells += input.damageCount
  }

  noteViewportPatchBroadcast(reason: string): void {
    this.totalCounters.fullPatchBroadcasts[reason] = (this.totalCounters.fullPatchBroadcasts[reason] ?? 0) + 1
  }

  noteRendererDeltaApply(input: { mutationCount: number; dirtyTileCount: number; durationMs: number }): void {
    this.totalCounters.rendererDeltaBatches += 1
    this.totalCounters.rendererDeltaMutations += input.mutationCount
    this.totalCounters.dirtyTilesMarked += input.dirtyTileCount
    this.totalCounters.rendererDeltaApplyMs += input.durationMs
  }

  noteRendererTileReadiness(input: {
    readonly exactHits: number
    readonly staleHits: number
    readonly misses: number
    readonly visibleDirtyTiles: number
    readonly warmDirtyTiles: number
  }): void {
    this.totalCounters.rendererTileInterestBatches += 1
    this.totalCounters.rendererTileExactHits += input.exactHits
    this.totalCounters.rendererTileStaleHits += input.staleHits
    this.totalCounters.rendererTileMisses += input.misses
    this.totalCounters.rendererVisibleDirtyTiles += input.visibleDirtyTiles
    this.totalCounters.rendererWarmDirtyTiles += input.warmDirtyTiles
  }

  noteVisibleWindowChange(): void {
    this.totalCounters.visibleWindowChanges += 1
  }

  noteHeaderPaneBuild(): void {
    this.totalCounters.headerPaneBuilds += 1
  }

  noteSurfaceCommit(surface: string): void {
    this.totalCounters.reactCommits += 1
    this.totalCounters.surfaceCommits[surface] = (this.totalCounters.surfaceCommits[surface] ?? 0) + 1
  }

  noteCanvasSurfaceMount(kind: 'canvas' | 'dom'): void {
    if (kind === 'canvas') {
      this.totalCounters.canvasSurfaceMounts += 1
      return
    }
    this.totalCounters.domSurfaceMounts += 1
  }

  noteCanvasPaint(layer: string): void {
    this.totalCounters.canvasPaints[layer] = (this.totalCounters.canvasPaints[layer] ?? 0) + 1
  }

  noteTypeGpuConfigure(): void {
    this.totalCounters.typeGpuConfigures += 1
  }

  noteTypeGpuSubmit(): void {
    this.totalCounters.typeGpuSubmits += 1
  }

  noteTypeGpuDrawCall(count: number): void {
    this.totalCounters.typeGpuDrawCalls += count
  }

  noteTypeGpuPaneDraw(count: number): void {
    this.totalCounters.typeGpuPaneDraws += count
  }

  noteTypeGpuUniformWrite(bytes: number): void {
    this.totalCounters.typeGpuUniformWriteBytes += bytes
  }

  noteTypeGpuBufferWrite(bytes: number): void {
    this.totalCounters.typeGpuVertexUploadBytes += bytes
  }

  noteTypeGpuOverlayWrite(bytes: number): void {
    this.totalCounters.typeGpuOverlayUploadBytes += bytes
  }

  noteTypeGpuBufferAllocation(bytes: number): void {
    this.totalCounters.typeGpuBufferAllocations += 1
    this.totalCounters.typeGpuBufferAllocationBytes += bytes
  }

  noteTypeGpuAtlasUpload(bytes: number): void {
    this.totalCounters.typeGpuAtlasUploadBytes += bytes
  }

  noteTypeGpuAtlasDirtyPageUpload(bytes: number, pageCount: number): void {
    this.totalCounters.typeGpuAtlasDirtyPageUploadBytes += bytes
    this.totalCounters.typeGpuAtlasDirtyPages += pageCount
  }

  noteTypeGpuSurfaceResize(): void {
    this.totalCounters.typeGpuSurfaceResizes += 1
  }

  noteTypeGpuTileMiss(): void {
    this.totalCounters.typeGpuTileMisses += 1
  }

  noteTypeGpuTileCacheEviction(count: number): void {
    this.totalCounters.typeGpuTileCacheEvictions += count
  }

  noteTypeGpuTileCacheSort(count: number): void {
    this.totalCounters.typeGpuTileCacheSorts += count
  }

  noteTypeGpuTileCacheStaleLookup(scannedEntries: number, hit: boolean): void {
    this.totalCounters.typeGpuTileCacheStaleLookups += 1
    this.totalCounters.typeGpuTileCacheEntriesScanned += scannedEntries
    if (hit) {
      this.totalCounters.typeGpuTileCacheStaleHits += 1
    }
  }

  noteTypeGpuTileCacheVisibleMark(count: number): void {
    this.totalCounters.typeGpuTileCacheVisibleMarks += count
  }

  noteGridScrollInput(timestamp: number): void {
    if (this.warmupFramesRemaining > 0) {
      return
    }
    this.lastScrollInputAt = timestamp
  }

  noteGridDrawFrame(timestamp: number): void {
    if (this.warmupFramesRemaining > 0 || this.lastScrollInputAt === null) {
      return
    }
    this.inputToDrawSamples.push(Math.max(0, timestamp - this.lastScrollInputAt))
    this.lastScrollInputAt = null
  }

  startSampling(workload: string): void {
    this.stopSampling()
    this.workload = workload
    this.frameSamples = []
    this.longTaskSamples = []
    this.inputToDrawSamples = []
    this.baselineCounters = null
    this.lastFrameAt = null
    this.lastScrollInputAt = null
    this.warmupFramesRemaining = WARMUP_FRAME_COUNT
    this.installLongTaskObserver()
    this.scheduleFrame()
  }

  stopSampling(): WorkbookScrollPerfReport | null {
    if (this.rafHandle !== null) {
      window.cancelAnimationFrame(this.rafHandle)
      this.rafHandle = null
    }
    this.observer?.disconnect()
    this.observer = null
    if (this.baselineCounters === null) {
      return null
    }
    const report: WorkbookScrollPerfReport = {
      workload: this.workload,
      fixture: this.fixture,
      samples: {
        frameMs: [...this.frameSamples],
        inputToDrawMs: [...this.inputToDrawSamples],
        longTasksMs: [...this.longTaskSamples],
      },
      summary: {
        frameMs: summarizeNumbers(this.frameSamples),
        inputToDrawMs: summarizeNumbers(this.inputToDrawSamples),
        longTasksMs: summarizeNumbers(this.longTaskSamples),
      },
      counters: subtractCounters(this.totalCounters, this.baselineCounters),
    }
    this.baselineCounters = null
    this.frameSamples = []
    this.longTaskSamples = []
    this.inputToDrawSamples = []
    this.lastFrameAt = null
    this.lastScrollInputAt = null
    this.warmupFramesRemaining = 0
    return report
  }

  private scheduleFrame(): void {
    this.rafHandle = window.requestAnimationFrame((timestamp) => {
      if (this.lastFrameAt !== null) {
        if (this.warmupFramesRemaining > 0) {
          this.warmupFramesRemaining -= 1
          if (this.warmupFramesRemaining === 0) {
            this.baselineCounters = cloneCounters(this.totalCounters)
            this.frameSamples = []
            this.longTaskSamples = []
            this.inputToDrawSamples = []
            this.lastScrollInputAt = null
          }
        } else {
          this.frameSamples.push(timestamp - this.lastFrameAt)
        }
      }
      this.lastFrameAt = timestamp
      this.scheduleFrame()
    })
  }

  private installLongTaskObserver(): void {
    if (typeof PerformanceObserver === 'undefined') {
      return
    }
    try {
      this.observer = new PerformanceObserver((list) => {
        if (this.warmupFramesRemaining > 0) {
          return
        }
        for (const entry of list.getEntries()) {
          this.longTaskSamples.push(entry.duration)
        }
      })
      this.observer.observe({ entryTypes: ['longtask'] })
    } catch {
      this.observer = null
    }
  }
}

function cloneCounters(counters: WorkbookScrollPerfCounters): WorkbookScrollPerfCounters {
  return {
    ...counters,
    canvasPaints: { ...counters.canvasPaints },
    fullPatchBroadcasts: { ...counters.fullPatchBroadcasts },
    surfaceCommits: { ...counters.surfaceCommits },
  }
}

function subtractCounters(counters: WorkbookScrollPerfCounters, baseline: WorkbookScrollPerfCounters): WorkbookScrollPerfCounters {
  return {
    viewportSubscriptions: counters.viewportSubscriptions - baseline.viewportSubscriptions,
    fullPatches: counters.fullPatches - baseline.fullPatches,
    fullPatchBroadcasts: subtractRecordCounters(counters.fullPatchBroadcasts, baseline.fullPatchBroadcasts),
    damagePatches: counters.damagePatches - baseline.damagePatches,
    damageCells: counters.damageCells - baseline.damageCells,
    dirtyTilesMarked: counters.dirtyTilesMarked - baseline.dirtyTilesMarked,
    rendererDeltaApplyMs: counters.rendererDeltaApplyMs - baseline.rendererDeltaApplyMs,
    rendererDeltaBatches: counters.rendererDeltaBatches - baseline.rendererDeltaBatches,
    rendererDeltaMutations: counters.rendererDeltaMutations - baseline.rendererDeltaMutations,
    rendererTileExactHits: counters.rendererTileExactHits - baseline.rendererTileExactHits,
    rendererTileInterestBatches: counters.rendererTileInterestBatches - baseline.rendererTileInterestBatches,
    rendererTileMisses: counters.rendererTileMisses - baseline.rendererTileMisses,
    rendererTileStaleHits: counters.rendererTileStaleHits - baseline.rendererTileStaleHits,
    rendererVisibleDirtyTiles: counters.rendererVisibleDirtyTiles - baseline.rendererVisibleDirtyTiles,
    rendererWarmDirtyTiles: counters.rendererWarmDirtyTiles - baseline.rendererWarmDirtyTiles,
    visibleWindowChanges: counters.visibleWindowChanges - baseline.visibleWindowChanges,
    headerPaneBuilds: counters.headerPaneBuilds - baseline.headerPaneBuilds,
    reactCommits: counters.reactCommits - baseline.reactCommits,
    canvasSurfaceMounts: counters.canvasSurfaceMounts - baseline.canvasSurfaceMounts,
    domSurfaceMounts: counters.domSurfaceMounts - baseline.domSurfaceMounts,
    canvasPaints: subtractRecordCounters(counters.canvasPaints, baseline.canvasPaints),
    surfaceCommits: subtractRecordCounters(counters.surfaceCommits, baseline.surfaceCommits),
    typeGpuAtlasUploadBytes: counters.typeGpuAtlasUploadBytes - baseline.typeGpuAtlasUploadBytes,
    typeGpuAtlasDirtyPages: counters.typeGpuAtlasDirtyPages - baseline.typeGpuAtlasDirtyPages,
    typeGpuAtlasDirtyPageUploadBytes: counters.typeGpuAtlasDirtyPageUploadBytes - baseline.typeGpuAtlasDirtyPageUploadBytes,
    typeGpuBufferAllocationBytes: counters.typeGpuBufferAllocationBytes - baseline.typeGpuBufferAllocationBytes,
    typeGpuBufferAllocations: counters.typeGpuBufferAllocations - baseline.typeGpuBufferAllocations,
    typeGpuConfigures: counters.typeGpuConfigures - baseline.typeGpuConfigures,
    typeGpuDrawCalls: counters.typeGpuDrawCalls - baseline.typeGpuDrawCalls,
    typeGpuPaneDraws: counters.typeGpuPaneDraws - baseline.typeGpuPaneDraws,
    typeGpuSubmits: counters.typeGpuSubmits - baseline.typeGpuSubmits,
    typeGpuSurfaceResizes: counters.typeGpuSurfaceResizes - baseline.typeGpuSurfaceResizes,
    typeGpuTileMisses: counters.typeGpuTileMisses - baseline.typeGpuTileMisses,
    typeGpuTileCacheEvictions: counters.typeGpuTileCacheEvictions - baseline.typeGpuTileCacheEvictions,
    typeGpuTileCacheEntriesScanned: counters.typeGpuTileCacheEntriesScanned - baseline.typeGpuTileCacheEntriesScanned,
    typeGpuTileCacheSorts: counters.typeGpuTileCacheSorts - baseline.typeGpuTileCacheSorts,
    typeGpuTileCacheStaleHits: counters.typeGpuTileCacheStaleHits - baseline.typeGpuTileCacheStaleHits,
    typeGpuTileCacheStaleLookups: counters.typeGpuTileCacheStaleLookups - baseline.typeGpuTileCacheStaleLookups,
    typeGpuTileCacheVisibleMarks: counters.typeGpuTileCacheVisibleMarks - baseline.typeGpuTileCacheVisibleMarks,
    typeGpuUniformWriteBytes: counters.typeGpuUniformWriteBytes - baseline.typeGpuUniformWriteBytes,
    typeGpuVertexUploadBytes: counters.typeGpuVertexUploadBytes - baseline.typeGpuVertexUploadBytes,
    typeGpuOverlayUploadBytes: counters.typeGpuOverlayUploadBytes - baseline.typeGpuOverlayUploadBytes,
  }
}

function subtractRecordCounters(
  counters: Readonly<Record<string, number>>,
  baseline: Readonly<Record<string, number>>,
): Record<string, number> {
  const keys = new Set([...Object.keys(counters), ...Object.keys(baseline)])
  const next: Record<string, number> = {}
  for (const key of keys) {
    next[key] = (counters[key] ?? 0) - (baseline[key] ?? 0)
  }
  return next
}

function summarizeNumbers(values: readonly number[]): WorkbookScrollPerfSummary {
  if (values.length === 0) {
    return {
      min: 0,
      median: 0,
      p95: 0,
      p99: 0,
      max: 0,
    }
  }
  const sorted = [...values].toSorted((left, right) => left - right)
  return {
    min: sorted[0] ?? 0,
    median: quantile(sorted, 0.5),
    p95: quantile(sorted, 0.95),
    p99: quantile(sorted, 0.99),
    max: sorted.at(-1) ?? 0,
  }
}

function quantile(sorted: readonly number[], percentile: number): number {
  if (sorted.length === 0) {
    return 0
  }
  const index = Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * percentile) - 1))
  return sorted[index] ?? 0
}

declare global {
  interface Window {
    __biligScrollPerf?: WorkbookScrollPerfCollector
  }
}

export function getWorkbookScrollPerfCollector(): WorkbookScrollPerfCollector | null {
  if (typeof window === 'undefined') {
    return null
  }
  window.__biligScrollPerf ??= new WorkbookScrollPerfCollector()
  return window.__biligScrollPerf
}
