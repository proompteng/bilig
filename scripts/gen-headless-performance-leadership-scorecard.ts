#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'

import type { CompetitiveArtifact, DominanceGoalStatus } from './bilig-dominance-scorecard-types.ts'
import { parseCompetitiveArtifact } from './bilig-dominance-scorecard-parsers.ts'
import type { ExtraHeadlessComparisonEngineSummary } from './headless-comparison-engine-summary.ts'
import { isFiniteNumber, readJsonObject } from './json-scorecard-helpers.ts'
import { formatJsonForRepo } from './scorecard-format.ts'
import { parseWorkPaperTrueCalcExtraComparisonEngineSummary } from './workpaper-vs-truecalc-artifact.ts'
import { parseWorkPaperXlsxCalcExtraComparisonEngineSummary } from './workpaper-vs-xlsx-calc-artifact.ts'

export interface BuildHeadlessPerformanceLeadershipScorecardInput {
  readonly competitiveArtifact: CompetitiveArtifact
  readonly competitiveArtifactPath: string
  readonly extraComparisonEngines?: readonly ExtraHeadlessComparisonEngineSummary[]
}

export interface HeadlessPerformanceLeadershipScorecard {
  readonly schemaVersion: 1
  readonly objective: string
  readonly goalStatus: DominanceGoalStatus
  readonly claimPolicy: {
    readonly blanketHeadlessPerformanceLeadershipClaimAllowed: boolean
    readonly requiredForBlanketClaim: readonly string[]
    readonly unmetRequirements: readonly string[]
  }
  readonly sourceArtifacts: {
    readonly primaryCompetitiveBenchmark: {
      readonly comparisonTarget: 'HyperFormula'
      readonly generatedAt: string
      readonly hyperFormulaCommit: string
      readonly hyperFormulaVersion: string
      readonly path: string
    }
    readonly extraCompetitiveBenchmarks: readonly ExtraHeadlessComparisonEngineSummary[]
  }
  readonly summary: {
    readonly comparableWorkloadCount: number
    readonly comparisonEngineCount: number
    readonly comparisonEngines: readonly string[]
    readonly eligibleFamilyCount: number
    readonly eligibleFamilies: readonly string[]
    readonly excludedFamilies: readonly string[]
    readonly meanAndP95WinCount: number
    readonly meanGeomeanRatio: number
    readonly meanWinCount: number
    readonly p95GeomeanRatio: number
    readonly p95Holdouts: readonly HeadlessPerformanceRatioHoldout[]
    readonly p95WinCount: number
    readonly tenXMeanAndP95WorkloadCountAgainstHyperFormula: number
    readonly limitedComparisonEngines: readonly HeadlessPerformanceLimitedComparisonEngine[]
    readonly workbookWideComparisonEngineCount: number
    readonly workbookWideComparisonEngines: readonly string[]
    readonly worstMeanRatio: number
    readonly worstMeanRatioWorkload: string
    readonly worstP95Ratio: number
    readonly worstP95RatioWorkload: string
  }
  readonly completionAudit: HeadlessPerformanceLeadershipAudit
}

export interface HeadlessPerformanceRatioHoldout {
  readonly comparisonTarget: 'HyperFormula'
  readonly p95Ratio: number
  readonly workload: string
}

export interface HeadlessPerformanceLimitedComparisonEngine {
  readonly comparableWorkloadCount: number
  readonly coverageTier: Exclude<ExtraHeadlessComparisonEngineSummary['coverageTier'], 'workbook-wide'>
  readonly engineName: string
}

export interface HeadlessPerformanceLeadershipAudit {
  readonly allCriteriaPassed: boolean
  readonly criteria: readonly HeadlessPerformanceLeadershipCriterion[]
  readonly unmetRequirements: readonly string[]
}

export interface HeadlessPerformanceLeadershipCriterion {
  readonly evidence: readonly string[]
  readonly gaps: readonly string[]
  readonly id: string
  readonly passed: boolean
  readonly requirement: string
}

export const rootDir = resolve(new URL('..', import.meta.url).pathname)
export const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'headless-performance-leadership-scorecard.json')
const competitiveArtifactPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-hyperformula.json')
const trueCalcArtifactPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-truecalc.json')
const xlsxCalcArtifactPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-xlsx-calc.json')
const requiredWorkbookWideComparisonEngineCount = 2

function main(): void {
  const isCheckMode = process.argv.includes('--check')
  const scorecard = buildHeadlessPerformanceLeadershipScorecard(loadHeadlessPerformanceLeadershipScorecardInput())
  const serializedScorecard = formatJsonForRepo({
    rootDir,
    serializedJson: `${JSON.stringify(scorecard, null, 2)}\n`,
    tempPrefix: 'headless-performance-leadership-scorecard',
  })

  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(
        `Missing generated headless performance leadership scorecard at ${outputPath}. Run: pnpm headless:performance:generate`,
      )
    }
    const currentScorecard = readFileSync(outputPath, 'utf8')
    if (currentScorecard !== serializedScorecard) {
      throw new Error('Generated headless performance leadership scorecard is out of date. Run: pnpm headless:performance:generate')
    }
  } else {
    mkdirSync(dirname(outputPath), { recursive: true })
    writeFileSync(outputPath, serializedScorecard)
  }

  console.log(
    JSON.stringify(
      {
        mode: isCheckMode ? 'check' : 'write',
        outputPath,
        goalStatus: scorecard.goalStatus,
        comparisonEngines: scorecard.summary.comparisonEngines,
        meanAndP95WinCount: scorecard.summary.meanAndP95WinCount,
        comparableWorkloadCount: scorecard.summary.comparableWorkloadCount,
        p95Holdouts: scorecard.summary.p95Holdouts.map((entry) => entry.workload),
      },
      null,
      2,
    ),
  )
}

export function loadHeadlessPerformanceLeadershipScorecardInput(): BuildHeadlessPerformanceLeadershipScorecardInput {
  return {
    competitiveArtifact: parseCompetitiveArtifact(readJsonObject(competitiveArtifactPath)),
    competitiveArtifactPath: toRepoPath(competitiveArtifactPath),
    extraComparisonEngines: [
      parseWorkPaperTrueCalcExtraComparisonEngineSummary(readJsonObject(trueCalcArtifactPath), toRepoPath(trueCalcArtifactPath)),
      parseWorkPaperXlsxCalcExtraComparisonEngineSummary(readJsonObject(xlsxCalcArtifactPath), toRepoPath(xlsxCalcArtifactPath)),
    ],
  }
}

export function buildHeadlessPerformanceLeadershipScorecard(
  input: BuildHeadlessPerformanceLeadershipScorecardInput,
): HeadlessPerformanceLeadershipScorecard {
  const eligibleFamilies = input.competitiveArtifact.families.filter((family) => family.scorecardEligible)
  const excludedFamilies = input.competitiveArtifact.families.filter((family) => !family.scorecardEligible)
  const eligibleWorkloads = new Set(eligibleFamilies.flatMap((family) => family.workloads ?? []))
  const comparableResults = input.competitiveArtifact.results.filter(
    (result) => result.comparable && (eligibleWorkloads.size === 0 || eligibleWorkloads.has(result.workload)),
  )
  const meanWinCount = comparableResults.filter(
    (result) => (result.comparison?.workpaperToHyperFormulaMeanRatio ?? Number.POSITIVE_INFINITY) < 1,
  ).length
  const p95WinCount = comparableResults.filter(
    (result) => (result.comparison?.workpaperToHyperFormulaP95Ratio ?? Number.POSITIVE_INFINITY) < 1,
  ).length
  const meanAndP95WinCount = comparableResults.filter((result) => {
    const comparison = result.comparison
    return comparison !== undefined && comparison.workpaperToHyperFormulaMeanRatio < 1 && comparison.workpaperToHyperFormulaP95Ratio < 1
  }).length
  const p95Holdouts = comparableResults
    .filter((result) => (result.comparison?.workpaperToHyperFormulaP95Ratio ?? Number.POSITIVE_INFINITY) >= 1)
    .map((result) => ({
      comparisonTarget: 'HyperFormula' as const,
      p95Ratio: result.comparison?.workpaperToHyperFormulaP95Ratio ?? Number.POSITIVE_INFINITY,
      workload: result.workload,
    }))
  const tenXMeanAndP95WorkloadCountAgainstHyperFormula = comparableResults.filter((result) => {
    const comparison = result.comparison
    return (
      comparison !== undefined && comparison.workpaperToHyperFormulaMeanRatio <= 0.1 && comparison.workpaperToHyperFormulaP95Ratio <= 0.1
    )
  }).length
  const extraComparisonEngines = [...(input.extraComparisonEngines ?? [])]
  const comparisonEngines = ['HyperFormula', ...extraComparisonEngines.map((engine) => engine.engineName)]
  const workbookWideComparisonEngines = [
    'HyperFormula',
    ...extraComparisonEngines.filter((engine) => engine.coverageTier === 'workbook-wide').map((engine) => engine.engineName),
  ]
  const limitedComparisonEngines = extraComparisonEngines
    .filter(
      (
        engine,
      ): engine is ExtraHeadlessComparisonEngineSummary & {
        coverageTier: Exclude<ExtraHeadlessComparisonEngineSummary['coverageTier'], 'workbook-wide'>
      } => engine.coverageTier !== 'workbook-wide',
    )
    .map((engine) => ({
      comparableWorkloadCount: engine.comparableWorkloadCount,
      coverageTier: engine.coverageTier,
      engineName: engine.engineName,
    }))
  const summary = {
    comparableWorkloadCount: comparableResults.length,
    comparisonEngineCount: comparisonEngines.length,
    comparisonEngines,
    eligibleFamilyCount: eligibleFamilies.length,
    eligibleFamilies: eligibleFamilies.map((family) => family.family),
    excludedFamilies: excludedFamilies.map((family) => family.family),
    meanAndP95WinCount,
    meanGeomeanRatio: input.competitiveArtifact.scorecard.directionalMeanRatioGeomean,
    meanWinCount,
    p95GeomeanRatio: input.competitiveArtifact.scorecard.directionalP95RatioGeomean,
    p95Holdouts,
    p95WinCount,
    tenXMeanAndP95WorkloadCountAgainstHyperFormula,
    limitedComparisonEngines,
    workbookWideComparisonEngineCount: workbookWideComparisonEngines.length,
    workbookWideComparisonEngines,
    worstMeanRatio: input.competitiveArtifact.scorecard.worstWorkpaperToHyperFormulaMeanRatio,
    worstMeanRatioWorkload: input.competitiveArtifact.scorecard.worstMeanRatioWorkload,
    worstP95Ratio: input.competitiveArtifact.scorecard.worstWorkpaperToHyperFormulaP95Ratio,
    worstP95RatioWorkload: input.competitiveArtifact.scorecard.worstP95RatioWorkload,
  }
  const completionAudit = buildCompletionAudit(input, summary, extraComparisonEngines)

  return {
    schemaVersion: 1,
    objective:
      'Make the bilig headless WorkPaper engine the fastest verified headless spreadsheet engine across diverse workbook benchmark families.',
    goalStatus: completionAudit.allCriteriaPassed ? 'achieved' : 'active-not-achieved',
    claimPolicy: {
      blanketHeadlessPerformanceLeadershipClaimAllowed: completionAudit.allCriteriaPassed,
      requiredForBlanketClaim: completionAudit.criteria.map((entry) => entry.requirement),
      unmetRequirements: completionAudit.unmetRequirements,
    },
    sourceArtifacts: {
      primaryCompetitiveBenchmark: {
        comparisonTarget: 'HyperFormula',
        generatedAt: input.competitiveArtifact.generatedAt,
        hyperFormulaCommit: input.competitiveArtifact.engines.hyperformula.commit,
        hyperFormulaVersion: input.competitiveArtifact.engines.hyperformula.version,
        path: input.competitiveArtifactPath,
      },
      extraCompetitiveBenchmarks: extraComparisonEngines,
    },
    summary,
    completionAudit,
  }
}

function buildCompletionAudit(
  input: BuildHeadlessPerformanceLeadershipScorecardInput,
  summary: HeadlessPerformanceLeadershipScorecard['summary'],
  extraComparisonEngines: readonly ExtraHeadlessComparisonEngineSummary[],
): HeadlessPerformanceLeadershipAudit {
  const eligibleWorkloads = new Set(
    input.competitiveArtifact.families.filter((family) => family.scorecardEligible).flatMap((family) => family.workloads ?? []),
  )
  const comparableResults = input.competitiveArtifact.results.filter(
    (result) => result.comparable && (eligibleWorkloads.size === 0 || eligibleWorkloads.has(result.workload)),
  )
  const malformedComparisons = comparableResults.filter((result) => result.comparison === undefined).map((result) => result.workload)
  const workbookWideExtraEngines = extraComparisonEngines.filter((engine) => engine.coverageTier === 'workbook-wide')
  const undercoveredWorkbookWideEngines = workbookWideExtraEngines.filter(
    (engine) =>
      engine.comparableWorkloadCount < summary.comparableWorkloadCount ||
      engine.meanWinCount < engine.comparableWorkloadCount ||
      engine.p95WinCount < engine.comparableWorkloadCount ||
      engine.meanAndP95WinCount < engine.comparableWorkloadCount,
  )
  const losingLimitedExtraEngines = extraComparisonEngines.filter(
    (engine) => engine.coverageTier !== 'workbook-wide' && engine.meanAndP95WinCount < engine.comparableWorkloadCount,
  )
  const criteria = [
    criterion({
      id: 'reproducible-artifact',
      requirement: 'Headless benchmark evidence must be generated from a checked-in reproducible artifact and check command.',
      evidence: [
        `primary artifact: ${input.competitiveArtifactPath}`,
        `generated at: ${input.competitiveArtifact.generatedAt}`,
        'check command: pnpm headless:performance:check',
        'timing refresh command: pnpm workpaper:bench:competitive:generate',
      ],
      gaps: [
        ...(summary.comparableWorkloadCount === 0 ? ['no comparable headless workloads are present'] : []),
        ...(summary.comparableWorkloadCount !== input.competitiveArtifact.scorecard.comparableCount
          ? ['scorecard comparable count does not match comparable result rows']
          : []),
        ...malformedComparisons.map((workload) => `missing comparison ratios for workload: ${workload}`),
      ],
    }),
    criterion({
      id: 'competitor-coverage',
      requirement:
        'Broad headless leadership must compare against at least two direct workbook-wide headless spreadsheet engines; scalar formula lanes are tracked separately.',
      evidence: [
        `comparison engines: ${summary.comparisonEngines.join(', ')}`,
        `workbook-wide engines: ${summary.workbookWideComparisonEngines.join(', ')}`,
        `limited engines: ${
          summary.limitedComparisonEngines
            .map((engine) => `${engine.engineName} (${engine.coverageTier}, ${String(engine.comparableWorkloadCount)} workloads)`)
            .join(', ') || 'none'
        }`,
      ],
      gaps:
        summary.workbookWideComparisonEngineCount >= requiredWorkbookWideComparisonEngineCount
          ? []
          : [
              `only ${summary.workbookWideComparisonEngines.join(
                ', ',
              )} is workbook-wide; add at least one more direct workbook-wide headless spreadsheet engine before broad headless leadership claims`,
            ],
    }),
    criterion({
      id: 'workload-family-breadth',
      requirement: 'Eligible benchmark families must cover diverse workbook build, recalc, structural, read, aggregate, and lookup shapes.',
      evidence: [
        `eligible families: ${summary.eligibleFamilies.join(', ')}`,
        `excluded families: ${summary.excludedFamilies.join(', ') || 'none'}`,
      ],
      gaps: [
        ...(summary.eligibleFamilyCount < 10
          ? [`only ${String(summary.eligibleFamilyCount)} eligible benchmark families are covered`]
          : []),
        ...input.competitiveArtifact.families
          .filter((family) => family.scorecardEligible && family.comparableCount === 0)
          .map((family) => `eligible family has no comparable workloads: ${family.family}`),
      ],
    }),
    criterion({
      id: 'per-workload-mean-and-p95-wins',
      requirement:
        'Every comparable headless workload must be faster on both mean and p95 latency against every direct headless competitor.',
      evidence: [
        `HyperFormula mean wins: ${String(summary.meanWinCount)}/${String(summary.comparableWorkloadCount)}`,
        `HyperFormula p95 wins: ${String(summary.p95WinCount)}/${String(summary.comparableWorkloadCount)}`,
        `HyperFormula mean+p95 wins: ${String(summary.meanAndP95WinCount)}/${String(summary.comparableWorkloadCount)}`,
      ],
      gaps: [
        ...(summary.meanAndP95WinCount === summary.comparableWorkloadCount
          ? []
          : [
              `${String(summary.meanAndP95WinCount)}/${String(
                summary.comparableWorkloadCount,
              )} comparable workloads win both mean and p95; p95 holdouts: ${
                summary.p95Holdouts.map((entry) => entry.workload).join(', ') || 'none'
              }`,
            ]),
        ...undercoveredWorkbookWideEngines.map(
          (engine) =>
            `${engine.engineName} workbook-wide comparison is incomplete: covers ${String(engine.comparableWorkloadCount)}/${String(
              summary.comparableWorkloadCount,
            )} comparable workloads and has ${String(engine.meanAndP95WinCount)}/${String(engine.comparableWorkloadCount)} mean+p95 wins`,
        ),
        ...losingLimitedExtraEngines.map(
          (engine) =>
            `${engine.engineName} ${engine.coverageTier} comparison has ${String(engine.meanAndP95WinCount)}/${String(
              engine.comparableWorkloadCount,
            )} mean+p95 wins`,
        ),
      ],
    }),
    criterion({
      id: 'multiple-reporting-integrity',
      requirement: 'Performance multiple reporting must include geomean ratios, worst-case rows, holdouts, and 10x workload count.',
      evidence: [
        `mean geomean ratio: ${String(summary.meanGeomeanRatio)}`,
        `p95 geomean ratio: ${String(summary.p95GeomeanRatio)}`,
        `worst mean row: ${summary.worstMeanRatioWorkload} (${String(summary.worstMeanRatio)})`,
        `worst p95 row: ${summary.worstP95RatioWorkload} (${String(summary.worstP95Ratio)})`,
        `10x mean+p95 workloads against HyperFormula: ${String(summary.tenXMeanAndP95WorkloadCountAgainstHyperFormula)}`,
      ],
      gaps: [
        ...finiteMetricGap('mean geomean ratio', summary.meanGeomeanRatio),
        ...finiteMetricGap('p95 geomean ratio', summary.p95GeomeanRatio),
        ...finiteMetricGap('worst mean ratio', summary.worstMeanRatio),
        ...finiteMetricGap('worst p95 ratio', summary.worstP95Ratio),
      ],
    }),
  ]
  const unmetRequirements = criteria.filter((entry) => !entry.passed).map((entry) => `${entry.id}: ${entry.gaps.join('; ')}`)
  return {
    allCriteriaPassed: unmetRequirements.length === 0,
    criteria,
    unmetRequirements,
  }
}

function criterion(args: {
  readonly evidence: readonly string[]
  readonly gaps: readonly string[]
  readonly id: string
  readonly requirement: string
}): HeadlessPerformanceLeadershipCriterion {
  const gaps = [...args.gaps]
  return {
    evidence: [...args.evidence],
    gaps,
    id: args.id,
    passed: gaps.length === 0,
    requirement: args.requirement,
  }
}

function finiteMetricGap(label: string, value: number): string[] {
  return isFiniteNumber(value) ? [] : [`${label} is not finite`]
}

function toRepoPath(path: string): string {
  return path.startsWith(`${rootDir}/`) ? path.slice(rootDir.length + 1) : path
}

if (import.meta.main) {
  main()
}
