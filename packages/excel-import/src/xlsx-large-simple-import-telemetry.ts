export type LargeSimpleXlsxImportPhase =
  | 'zip-setup'
  | 'worksheet-scan'
  | 'shared-string-resolution'
  | 'metadata-parsing'
  | 'style-parsing'
  | 'zip-source-release'
  | 'public-snapshot-materialization'

export interface LargeSimpleXlsxImportPhaseTelemetry {
  readonly phase: LargeSimpleXlsxImportPhase
  readonly elapsedMs: number
  readonly rssBytes?: number
  readonly heapUsedBytes?: number
  readonly zipSourceBytesBeforeRelease?: number
  readonly zipSourceBytesAfterRelease?: number
  readonly ownedSourceBytesBeforeRelease?: number
  readonly ownedSourceBytesAfterRelease?: number
}

interface PhaseAccumulator {
  elapsedMs: number
  rssBytes?: number
  heapUsedBytes?: number
  zipSourceBytesBeforeRelease?: number
  zipSourceBytesAfterRelease?: number
  ownedSourceBytesBeforeRelease?: number
  ownedSourceBytesAfterRelease?: number
}

export interface LargeSimpleXlsxImportPhaseEvidence {
  readonly zipSourceBytesBeforeRelease?: number
  readonly zipSourceBytesAfterRelease?: number
  readonly ownedSourceBytesBeforeRelease?: number
  readonly ownedSourceBytesAfterRelease?: number
}

export class LargeSimpleXlsxImportPhaseRecorder {
  private readonly phases = new Map<LargeSimpleXlsxImportPhase, PhaseAccumulator>()
  private readonly order: LargeSimpleXlsxImportPhase[] = []

  start(): number {
    return nowMs()
  }

  finish(phase: LargeSimpleXlsxImportPhase, startedAtMs: number, evidence: LargeSimpleXlsxImportPhaseEvidence = {}): void {
    const elapsedMs = Math.max(0, Math.round(nowMs() - startedAtMs))
    let accumulator = this.phases.get(phase)
    if (!accumulator) {
      accumulator = { elapsedMs: 0 }
      this.phases.set(phase, accumulator)
      this.order.push(phase)
    }
    accumulator.elapsedMs += elapsedMs
    const memory = readMemoryUsage()
    if (memory?.rssBytes !== undefined) {
      accumulator.rssBytes = Math.max(accumulator.rssBytes ?? 0, memory.rssBytes)
    }
    if (memory?.heapUsedBytes !== undefined) {
      accumulator.heapUsedBytes = Math.max(accumulator.heapUsedBytes ?? 0, memory.heapUsedBytes)
    }
    if (evidence.zipSourceBytesBeforeRelease !== undefined) {
      accumulator.zipSourceBytesBeforeRelease = evidence.zipSourceBytesBeforeRelease
    }
    if (evidence.zipSourceBytesAfterRelease !== undefined) {
      accumulator.zipSourceBytesAfterRelease = evidence.zipSourceBytesAfterRelease
    }
    if (evidence.ownedSourceBytesBeforeRelease !== undefined) {
      accumulator.ownedSourceBytesBeforeRelease = evidence.ownedSourceBytesBeforeRelease
    }
    if (evidence.ownedSourceBytesAfterRelease !== undefined) {
      accumulator.ownedSourceBytesAfterRelease = evidence.ownedSourceBytesAfterRelease
    }
  }

  entries(): LargeSimpleXlsxImportPhaseTelemetry[] {
    return this.order.map((phase) => {
      const entry = this.phases.get(phase)!
      return {
        phase,
        elapsedMs: entry.elapsedMs,
        ...(entry.rssBytes !== undefined ? { rssBytes: entry.rssBytes } : {}),
        ...(entry.heapUsedBytes !== undefined ? { heapUsedBytes: entry.heapUsedBytes } : {}),
        ...(entry.zipSourceBytesBeforeRelease !== undefined ? { zipSourceBytesBeforeRelease: entry.zipSourceBytesBeforeRelease } : {}),
        ...(entry.zipSourceBytesAfterRelease !== undefined ? { zipSourceBytesAfterRelease: entry.zipSourceBytesAfterRelease } : {}),
        ...(entry.ownedSourceBytesBeforeRelease !== undefined
          ? { ownedSourceBytesBeforeRelease: entry.ownedSourceBytesBeforeRelease }
          : {}),
        ...(entry.ownedSourceBytesAfterRelease !== undefined ? { ownedSourceBytesAfterRelease: entry.ownedSourceBytesAfterRelease } : {}),
      }
    })
  }
}

function nowMs(): number {
  return globalThis.performance?.now?.() ?? Date.now()
}

function readMemoryUsage(): { readonly rssBytes?: number; readonly heapUsedBytes?: number } | undefined {
  const processLike = (globalThis as { readonly process?: { readonly memoryUsage?: () => unknown } }).process
  const usage = processLike?.memoryUsage?.()
  if (!isMemoryUsage(usage)) {
    return undefined
  }
  return {
    rssBytes: usage.rss,
    heapUsedBytes: usage.heapUsed,
  }
}

function isMemoryUsage(value: unknown): value is { readonly rss: number; readonly heapUsed: number } {
  return (
    typeof value === 'object' &&
    value !== null &&
    'rss' in value &&
    typeof value.rss === 'number' &&
    'heapUsed' in value &&
    typeof value.heapUsed === 'number'
  )
}
