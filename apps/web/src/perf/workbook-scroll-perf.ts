export interface WorkbookScrollPerfFixture {
  readonly id: string
  readonly materializedCellCount: number
  readonly sheetName: string
}

interface WorkbookScrollPerfCounters {
  viewportSubscriptions: number
  fullPatches: number
  damagePatches: number
  damageCells: number
  reactCommits: number
  canvasSurfaceMounts: number
  domSurfaceMounts: number
}

interface WorkbookScrollPerfSamples {
  readonly frameMs: number[]
  readonly longTasksMs: number[]
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
    damagePatches: 0,
    damageCells: 0,
    reactCommits: 0,
    canvasSurfaceMounts: 0,
    domSurfaceMounts: 0,
  }
  private baselineCounters: WorkbookScrollPerfCounters | null = null
  private frameSamples: number[] = []
  private longTaskSamples: number[] = []
  private workload = 'idle'
  private fixture: WorkbookScrollPerfFixture | null = null
  private benchmarkState: BenchmarkState = 'idle'
  private benchmarkError: string | null = null
  private rafHandle: number | null = null
  private lastFrameAt: number | null = null
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

  noteReactCommit(): void {
    this.totalCounters.reactCommits += 1
  }

  noteTextSurface(kind: 'canvas' | 'dom'): void {
    if (kind === 'canvas') {
      this.totalCounters.canvasSurfaceMounts += 1
      return
    }
    this.totalCounters.domSurfaceMounts += 1
  }

  startSampling(workload: string): void {
    this.stopSampling()
    this.workload = workload
    this.frameSamples = []
    this.longTaskSamples = []
    this.baselineCounters = cloneCounters(this.totalCounters)
    this.lastFrameAt = null
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
        longTasksMs: [...this.longTaskSamples],
      },
      summary: {
        frameMs: summarizeNumbers(this.frameSamples),
        longTasksMs: summarizeNumbers(this.longTaskSamples),
      },
      counters: subtractCounters(this.totalCounters, this.baselineCounters),
    }
    this.baselineCounters = null
    this.frameSamples = []
    this.longTaskSamples = []
    this.lastFrameAt = null
    this.warmupFramesRemaining = 0
    return report
  }

  private scheduleFrame(): void {
    this.rafHandle = window.requestAnimationFrame((timestamp) => {
      if (this.lastFrameAt !== null) {
        if (this.warmupFramesRemaining > 0) {
          this.warmupFramesRemaining -= 1
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
  return { ...counters }
}

function subtractCounters(counters: WorkbookScrollPerfCounters, baseline: WorkbookScrollPerfCounters): WorkbookScrollPerfCounters {
  return {
    viewportSubscriptions: counters.viewportSubscriptions - baseline.viewportSubscriptions,
    fullPatches: counters.fullPatches - baseline.fullPatches,
    damagePatches: counters.damagePatches - baseline.damagePatches,
    damageCells: counters.damageCells - baseline.damageCells,
    reactCommits: counters.reactCommits - baseline.reactCommits,
    canvasSurfaceMounts: counters.canvasSurfaceMounts - baseline.canvasSurfaceMounts,
    domSurfaceMounts: counters.domSurfaceMounts - baseline.domSurfaceMounts,
  }
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
