import { EXPANDED_COMPARATIVE_WORKLOADS, type ExpandedComparativeBenchmarkWorkload } from './expanded-competitive-workloads.js'
import type { ExpandedComparativeBenchmarkResult } from './benchmark-workpaper-vs-hyperformula-expanded.js'

export type ExpandedCompetitiveFamily =
  | 'build'
  | 'rebuild'
  | 'runtime-restore'
  | 'config-toggle'
  | 'dirty-execution'
  | 'batch-edit'
  | 'structural-rows'
  | 'structural-columns'
  | 'range-read'
  | 'overlapping-aggregate'
  | 'sliding-window-aggregate'
  | 'conditional-aggregation'
  | 'lookup-exact'
  | 'lookup-after-write'
  | 'lookup-approximate'
  | 'lookup-approximate-after-write'
  | 'lookup-text'
  | 'dynamic-array'

export const EXPANDED_COMPARATIVE_FAMILY_ORDER = [
  'build',
  'rebuild',
  'runtime-restore',
  'config-toggle',
  'dirty-execution',
  'batch-edit',
  'structural-rows',
  'structural-columns',
  'range-read',
  'overlapping-aggregate',
  'sliding-window-aggregate',
  'conditional-aggregation',
  'lookup-exact',
  'lookup-after-write',
  'lookup-approximate',
  'lookup-approximate-after-write',
  'lookup-text',
  'dynamic-array',
] as const satisfies readonly ExpandedCompetitiveFamily[]

const EXPANDED_COMPARATIVE_FAMILY_METADATA = {
  build: { scorecardEligible: true },
  rebuild: { scorecardEligible: true },
  'runtime-restore': { scorecardEligible: true },
  'config-toggle': {
    scorecardEligible: false,
    exclusionReason: 'Control-only rebuild toggle; not evidence of broad competitive victory.',
  },
  'dirty-execution': { scorecardEligible: true },
  'batch-edit': { scorecardEligible: true },
  'structural-rows': { scorecardEligible: true },
  'structural-columns': { scorecardEligible: true },
  'range-read': { scorecardEligible: true },
  'overlapping-aggregate': { scorecardEligible: true },
  'sliding-window-aggregate': { scorecardEligible: true },
  'conditional-aggregation': { scorecardEligible: true },
  'lookup-exact': { scorecardEligible: true },
  'lookup-after-write': { scorecardEligible: true },
  'lookup-approximate': { scorecardEligible: true },
  'lookup-approximate-after-write': { scorecardEligible: true },
  'lookup-text': { scorecardEligible: true },
  'dynamic-array': {
    scorecardEligible: false,
    exclusionReason: 'Leadership-only support lane; not an apples-to-apples performance scorecard input.',
  },
} as const satisfies Record<ExpandedCompetitiveFamily, { scorecardEligible: boolean; exclusionReason?: string }>

export const EXPANDED_COMPARATIVE_FAMILY_GROUPS = {
  build: [
    'build-dense-literals',
    'build-mixed-content',
    'build-parser-cache-row-templates',
    'build-parser-cache-mixed-templates',
    'build-many-sheets',
  ],
  rebuild: ['rebuild-and-recalculate'],
  'runtime-restore': ['rebuild-runtime-from-snapshot'],
  'config-toggle': ['rebuild-config-toggle'],
  'dirty-execution': ['single-edit-chain', 'single-edit-fanout', 'partial-recompute-mixed-frontier', 'single-formula-edit-recalc'],
  'batch-edit': [
    'batch-edit-single-column',
    'batch-edit-multi-column',
    'batch-edit-single-column-with-undo',
    'batch-suspended-single-column',
    'batch-suspended-multi-column',
  ],
  'structural-rows': ['structural-insert-rows', 'structural-delete-rows', 'structural-move-rows'],
  'structural-columns': ['structural-insert-columns', 'structural-delete-columns', 'structural-move-columns'],
  'range-read': ['range-read-dense'],
  'overlapping-aggregate': ['aggregate-overlapping-ranges'],
  'sliding-window-aggregate': ['aggregate-overlapping-sliding-window'],
  'conditional-aggregation': ['conditional-aggregation-reused-ranges', 'conditional-aggregation-criteria-cell-edit'],
  'lookup-exact': ['lookup-no-column-index', 'lookup-with-column-index'],
  'lookup-after-write': ['lookup-with-column-index-after-column-write', 'lookup-with-column-index-after-batch-write'],
  'lookup-approximate': ['lookup-approximate-sorted'],
  'lookup-approximate-after-write': ['lookup-approximate-sorted-after-column-write'],
  'lookup-text': ['lookup-text-exact'],
  'dynamic-array': ['dynamic-array-filter'],
} as const satisfies Record<ExpandedCompetitiveFamily, readonly ExpandedComparativeBenchmarkWorkload[]>

export const EXPANDED_COMPARATIVE_WORKLOAD_FAMILY = buildExpandedComparativeWorkloadFamilyMap(EXPANDED_COMPARATIVE_FAMILY_GROUPS)

export interface ExpandedCompetitiveFamilySummary {
  family: ExpandedCompetitiveFamily
  workloads: readonly ExpandedComparativeBenchmarkWorkload[]
  scorecardEligible: boolean
  exclusionReason: string | null
  resultCount: number
  comparableCount: number
  leadershipCount: number
  workpaperWins: number
  hyperformulaWins: number
  meanSpeedupGeomean: number | null
}

export interface ExpandedCompetitiveScorecardSummary {
  eligibleFamilies: readonly ExpandedCompetitiveFamily[]
  excludedFamilies: readonly ExpandedCompetitiveFamily[]
  comparableCount: number
  workpaperWins: number
  hyperformulaWins: number
}

export interface ExpandedCompetitiveFamilyReport {
  suite: 'workpaper-vs-hyperformula-expanded'
  families: readonly ExpandedCompetitiveFamilySummary[]
  scorecard: ExpandedCompetitiveScorecardSummary
}

export function getExpandedCompetitiveFamily(workload: ExpandedComparativeBenchmarkWorkload): ExpandedCompetitiveFamily {
  return EXPANDED_COMPARATIVE_WORKLOAD_FAMILY[workload]
}

export function groupExpandedCompetitiveBenchmarkResultsByFamily(
  results: readonly ExpandedComparativeBenchmarkResult[],
): Record<ExpandedCompetitiveFamily, readonly ExpandedComparativeBenchmarkResult[]> {
  return {
    build: results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'build'),
    rebuild: results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'rebuild'),
    'runtime-restore': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'runtime-restore'),
    'config-toggle': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'config-toggle'),
    'dirty-execution': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'dirty-execution'),
    'batch-edit': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'batch-edit'),
    'structural-rows': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'structural-rows'),
    'structural-columns': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'structural-columns'),
    'range-read': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'range-read'),
    'overlapping-aggregate': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'overlapping-aggregate'),
    'sliding-window-aggregate': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'sliding-window-aggregate'),
    'conditional-aggregation': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'conditional-aggregation'),
    'lookup-exact': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'lookup-exact'),
    'lookup-after-write': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'lookup-after-write'),
    'lookup-approximate': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'lookup-approximate'),
    'lookup-approximate-after-write': results.filter(
      (result) => getExpandedCompetitiveFamily(result.workload) === 'lookup-approximate-after-write',
    ),
    'lookup-text': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'lookup-text'),
    'dynamic-array': results.filter((result) => getExpandedCompetitiveFamily(result.workload) === 'dynamic-array'),
  }
}

export function summarizeExpandedCompetitiveFamilies(
  results: readonly ExpandedComparativeBenchmarkResult[],
): ExpandedCompetitiveFamilySummary[] {
  const groupedResults = groupExpandedCompetitiveBenchmarkResultsByFamily(results)
  return EXPANDED_COMPARATIVE_FAMILY_ORDER.map((family) => {
    const familyResults = groupedResults[family]
    const comparableResults = familyResults.filter((result) => result.comparable)
    const workpaperWins = comparableResults.filter((result) => result.comparison.fasterEngine === 'workpaper').length
    const hyperformulaWins = comparableResults.length - workpaperWins
    const metadata = EXPANDED_COMPARATIVE_FAMILY_METADATA[family]
    return {
      family,
      workloads: EXPANDED_COMPARATIVE_FAMILY_GROUPS[family],
      scorecardEligible: metadata.scorecardEligible,
      exclusionReason: 'exclusionReason' in metadata ? metadata.exclusionReason : null,
      resultCount: familyResults.length,
      comparableCount: comparableResults.length,
      leadershipCount: familyResults.length - comparableResults.length,
      workpaperWins,
      hyperformulaWins,
      meanSpeedupGeomean:
        comparableResults.length === 0 ? null : geometricMean(comparableResults.map((result) => result.comparison.meanSpeedup)),
    }
  })
}

export function buildExpandedCompetitiveFamilyReport(
  results: readonly ExpandedComparativeBenchmarkResult[],
): ExpandedCompetitiveFamilyReport {
  const families = summarizeExpandedCompetitiveFamilies(results)
  const scorecardFamilies = families.filter((family) => family.scorecardEligible)
  return {
    suite: 'workpaper-vs-hyperformula-expanded',
    families,
    scorecard: {
      eligibleFamilies: scorecardFamilies.map((family) => family.family),
      excludedFamilies: families.filter((family) => !family.scorecardEligible).map((family) => family.family),
      comparableCount: scorecardFamilies.reduce((sum, family) => sum + family.comparableCount, 0),
      workpaperWins: scorecardFamilies.reduce((sum, family) => sum + family.workpaperWins, 0),
      hyperformulaWins: scorecardFamilies.reduce((sum, family) => sum + family.hyperformulaWins, 0),
    },
  }
}

export function formatExpandedCompetitiveFamilyReport(results: readonly ExpandedComparativeBenchmarkResult[]): string {
  return JSON.stringify(buildExpandedCompetitiveFamilyReport(results), null, 2)
}

function buildExpandedComparativeWorkloadFamilyMap(
  groups: Record<ExpandedCompetitiveFamily, readonly ExpandedComparativeBenchmarkWorkload[]>,
): Record<ExpandedComparativeBenchmarkWorkload, ExpandedCompetitiveFamily> {
  const familyByWorkload: Partial<Record<ExpandedComparativeBenchmarkWorkload, ExpandedCompetitiveFamily>> = {}
  const missingWorkloads = new Set(EXPANDED_COMPARATIVE_WORKLOADS)
  for (const family of EXPANDED_COMPARATIVE_FAMILY_ORDER) {
    const workloads = groups[family]
    for (const workload of workloads) {
      if (familyByWorkload[workload] !== undefined) {
        throw new Error(`Expanded competitive workload ${workload} is assigned to multiple families`)
      }
      familyByWorkload[workload] = family
      missingWorkloads.delete(workload)
    }
  }

  if (missingWorkloads.size !== 0) {
    throw new Error(`Expanded competitive workload ${missingWorkloads.values().next().value} is missing from family coverage`)
  }
  if (!isCompleteExpandedComparativeWorkloadFamilyMap(familyByWorkload)) {
    throw new Error('Expanded competitive workload family coverage is incomplete')
  }

  return familyByWorkload
}

function isCompleteExpandedComparativeWorkloadFamilyMap(
  value: Partial<Record<ExpandedComparativeBenchmarkWorkload, ExpandedCompetitiveFamily>>,
): value is Record<ExpandedComparativeBenchmarkWorkload, ExpandedCompetitiveFamily> {
  return EXPANDED_COMPARATIVE_WORKLOADS.every((workload) => value[workload] !== undefined)
}

function geometricMean(values: readonly number[]): number {
  if (values.length === 0) {
    return Number.NaN
  }
  const totalLog = values.reduce((sum, value) => {
    if (value <= 0) {
      throw new Error(`Cannot compute geomean for non-positive value: ${String(value)}`)
    }
    return sum + Math.log(value)
  }, 0)
  return Math.exp(totalLog / values.length)
}
