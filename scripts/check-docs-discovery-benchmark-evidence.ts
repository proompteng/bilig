import { readFileSync } from 'node:fs'

type PublicBenchmarkEvidence = {
  readonly package: {
    readonly version: string
  }
  readonly workpaperVsHyperFormula: {
    readonly overall: {
      readonly comparableCount: number
      readonly workpaperWins: number
      readonly worstP95RatioWorkload: string
      readonly worstWorkpaperToHyperFormulaP95Ratio: number
    }
    readonly meanAndP95WinCount: number
  }
}

export type BenchmarkDiscoveryEvidence = {
  readonly comparableCount: number
  readonly meanAndP95Headline: string
  readonly meanWinHeadline: string
  readonly meanWinSentencePrefix: string
  readonly packageVersion: string
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
  const p95HoldoutRatio = `${overall.worstWorkpaperToHyperFormulaP95Ratio.toFixed(3).replace(/0+$/u, '').replace(/\.$/u, '')}x`

  return {
    comparableCount: overall.comparableCount,
    meanAndP95Headline: `${evidence.workpaperVsHyperFormula.meanAndP95WinCount}/${overall.comparableCount}`,
    meanWinHeadline: `${overall.workpaperWins}/${overall.comparableCount}`,
    meanWinSentencePrefix: `${overall.workpaperWins} of ${overall.comparableCount}`,
    packageVersion: evidence.package.version,
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
  if (!isRecord(value) || !isRecord(value.package) || !isRecord(value.workpaperVsHyperFormula)) {
    return false
  }

  const { overall } = value.workpaperVsHyperFormula
  return (
    typeof value.package.version === 'string' &&
    isRecord(overall) &&
    Number.isFinite(value.workpaperVsHyperFormula.meanAndP95WinCount) &&
    Number.isFinite(overall.comparableCount) &&
    Number.isFinite(overall.workpaperWins) &&
    typeof overall.worstP95RatioWorkload === 'string' &&
    Number.isFinite(overall.worstWorkpaperToHyperFormulaP95Ratio)
  )
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}
