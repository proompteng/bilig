import type { NumericSummary } from './stats.js'

export const DEFAULT_COMPETITIVE_WARMUP_COUNT = 2
export const DEFAULT_COMPETITIVE_SAMPLE_COUNT = 5
export const HYPERFORMULA_LICENSE_KEY = 'gpl-v3'

export interface ComparativeBenchmarkSuiteOptions {
  sampleCount?: number
  warmupCount?: number
}

export interface ComparativeMeasuredEngineResult {
  status: 'supported'
  elapsedMs: NumericSummary
  memoryDeltaBytes: ComparativeMemorySummary
  engineCounters?: Record<string, NumericSummary>
  verification: Record<string, unknown>
}

export interface ComparativeUnsupportedEngineResult {
  status: 'unsupported'
  reason: string
  evidence: readonly string[]
}

export interface ComparativeMemorySummary {
  rssBytes: NumericSummary
  heapUsedBytes: NumericSummary
  heapTotalBytes: NumericSummary
  externalBytes: NumericSummary
  arrayBuffersBytes: NumericSummary
}
