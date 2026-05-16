import { readFileSync } from 'node:fs'

type PublicBenchmarkEvidence = {
  readonly workpaperVsHyperFormula: {
    readonly overall: {
      readonly comparableCount: number
      readonly workpaperWins: number
      readonly worstP95RatioWorkload: string
      readonly worstWorkpaperToHyperFormulaP95Ratio: number
    }
  }
}

export type BenchmarkDiscoveryEvidence = {
  readonly comparableCount: number
  readonly meanWinHeadline: string
  readonly meanWinSentencePrefix: string
  readonly p95HoldoutWorkload: string
  readonly p95HoldoutRatio: string
}

let cachedBenchmarkDiscoveryEvidence: BenchmarkDiscoveryEvidence | undefined

export function getBenchmarkDiscoveryEvidence(): BenchmarkDiscoveryEvidence {
  cachedBenchmarkDiscoveryEvidence ??= readBenchmarkDiscoveryEvidence()
  return cachedBenchmarkDiscoveryEvidence
}

function readBenchmarkDiscoveryEvidence(): BenchmarkDiscoveryEvidence {
  const evidence = parsePublicBenchmarkEvidence(readFileSync(new URL('../docs/public-evidence.json', import.meta.url), 'utf8'))
  const overall = evidence.workpaperVsHyperFormula.overall
  const p95HoldoutRatio = `${overall.worstWorkpaperToHyperFormulaP95Ratio.toFixed(3)}x`

  return {
    comparableCount: overall.comparableCount,
    meanWinHeadline: `${overall.workpaperWins}/${overall.comparableCount}`,
    meanWinSentencePrefix: `${overall.workpaperWins} of ${overall.comparableCount}`,
    p95HoldoutWorkload: overall.worstP95RatioWorkload,
    p95HoldoutRatio,
  }
}

function parsePublicBenchmarkEvidence(rawEvidence: string): PublicBenchmarkEvidence {
  const parsedEvidence: unknown = JSON.parse(rawEvidence)

  if (!isPublicBenchmarkEvidence(parsedEvidence)) {
    throw new Error('docs/public-evidence.json is missing WorkPaper benchmark evidence')
  }

  return parsedEvidence
}

function isPublicBenchmarkEvidence(value: unknown): value is PublicBenchmarkEvidence {
  if (!isRecord(value) || !isRecord(value.workpaperVsHyperFormula)) {
    return false
  }

  const { overall } = value.workpaperVsHyperFormula
  return (
    isRecord(overall) &&
    Number.isFinite(overall.comparableCount) &&
    Number.isFinite(overall.workpaperWins) &&
    typeof overall.worstP95RatioWorkload === 'string' &&
    Number.isFinite(overall.worstWorkpaperToHyperFormulaP95Ratio)
  )
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}
