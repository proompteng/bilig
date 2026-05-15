import type { PublicWorkbookCorpusWorkerOptions } from './public-workbook-corpus-footprint.ts'
import type {
  PublicWorkbookCorpusCase,
  PublicWorkbookVerificationPhase,
  PublicWorkbookVerificationPhaseTiming,
} from './public-workbook-corpus-types.ts'

interface VerificationRuntimeMetrics {
  readonly startedAt: number
  readonly phaseTimings: PublicWorkbookVerificationPhaseTiming[]
}

export function startVerificationRuntimeMetrics(): VerificationRuntimeMetrics {
  return {
    startedAt: performance.now(),
    phaseTimings: [],
  }
}

export async function timeVerificationPhase<T>(
  metrics: VerificationRuntimeMetrics,
  workerOptions: PublicWorkbookCorpusWorkerOptions,
  phase: PublicWorkbookVerificationPhase,
  fn: () => T | Promise<T>,
): Promise<T> {
  workerOptions.onPhase?.(phase)
  const startedAt = performance.now()
  try {
    return await fn()
  } finally {
    metrics.phaseTimings.push({ phase, elapsedMs: roundElapsedMs(performance.now() - startedAt) })
  }
}

export function withVerificationRuntimeMetrics(
  corpusCase: PublicWorkbookCorpusCase,
  metrics: VerificationRuntimeMetrics,
  peakRssBytes?: number,
): PublicWorkbookCorpusCase {
  return withPeakRssBytes(
    {
      ...corpusCase,
      elapsedMs: roundElapsedMs(performance.now() - metrics.startedAt),
      phaseTimings: metrics.phaseTimings,
    },
    peakRssBytes,
  )
}

export function withPeakRssBytes(corpusCase: PublicWorkbookCorpusCase, peakRssBytes: number): PublicWorkbookCorpusCase
export function withPeakRssBytes(corpusCase: PublicWorkbookCorpusCase, peakRssBytes?: number): PublicWorkbookCorpusCase
export function withPeakRssBytes(corpusCase: PublicWorkbookCorpusCase, peakRssBytes?: number): PublicWorkbookCorpusCase {
  if (peakRssBytes === undefined || peakRssBytes <= 0) {
    return corpusCase
  }
  return {
    ...corpusCase,
    peakRssBytes: Math.max(corpusCase.peakRssBytes ?? 0, Math.trunc(peakRssBytes)),
  }
}

function roundElapsedMs(value: number): number {
  return Math.max(0, Math.round(value))
}
