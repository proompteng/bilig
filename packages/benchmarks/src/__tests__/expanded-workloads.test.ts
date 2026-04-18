import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import {
  EXPANDED_COMPARATIVE_WORKLOADS,
  buildExpandedComparativeBenchmarkReport,
  type ExpandedComparativeBenchmarkWorkload,
} from '../benchmark-workpaper-vs-hyperformula-expanded.js'
import {
  EXPANDED_COMPARATIVE_FAMILY_GROUPS,
  EXPANDED_COMPARATIVE_FAMILY_ORDER,
  formatExpandedCompetitiveFamilyReport,
  EXPANDED_COMPARATIVE_WORKLOAD_FAMILY,
  type ExpandedCompetitiveFamily,
} from '../report-competitive-families.js'

const expectedExpandedWorkloads: ExpandedComparativeBenchmarkWorkload[] = [
  'build-dense-literals',
  'build-mixed-content',
  'build-parser-cache-row-templates',
  'build-parser-cache-mixed-templates',
  'build-many-sheets',
  'rebuild-and-recalculate',
  'rebuild-config-toggle',
  'rebuild-runtime-from-snapshot',
  'single-edit-chain',
  'single-edit-fanout',
  'partial-recompute-mixed-frontier',
  'single-formula-edit-recalc',
  'batch-edit-single-column',
  'batch-edit-multi-column',
  'batch-edit-single-column-with-undo',
  'batch-suspended-single-column',
  'batch-suspended-multi-column',
  'structural-insert-rows',
  'structural-delete-rows',
  'structural-move-rows',
  'structural-insert-columns',
  'structural-delete-columns',
  'structural-move-columns',
  'range-read-dense',
  'aggregate-overlapping-ranges',
  'aggregate-overlapping-sliding-window',
  'conditional-aggregation-reused-ranges',
  'conditional-aggregation-criteria-cell-edit',
  'lookup-no-column-index',
  'lookup-with-column-index',
  'lookup-with-column-index-after-column-write',
  'lookup-with-column-index-after-batch-write',
  'lookup-approximate-sorted',
  'lookup-approximate-sorted-after-column-write',
  'lookup-text-exact',
  'dynamic-array-filter',
]

const benchmarkDir = dirname(fileURLToPath(import.meta.url))
const expandedBaselinePath = join(benchmarkDir, '..', '..', 'baselines', 'workpaper-vs-hyperformula-expanded.json')

function familyEligibility(family: ExpandedCompetitiveFamily): { scorecardEligible: boolean; exclusionReason: string | null } {
  switch (family) {
    case 'build':
    case 'rebuild':
    case 'runtime-restore':
    case 'config-toggle':
      return family === 'config-toggle'
        ? {
            scorecardEligible: false,
            exclusionReason: 'Control-only rebuild toggle; not evidence of broad competitive victory.',
          }
        : { scorecardEligible: true, exclusionReason: null }
    case 'dirty-execution':
    case 'batch-edit':
    case 'structural-rows':
    case 'structural-columns':
    case 'range-read':
    case 'overlapping-aggregate':
    case 'sliding-window-aggregate':
    case 'conditional-aggregation':
    case 'lookup-exact':
    case 'lookup-after-write':
    case 'lookup-approximate':
    case 'lookup-approximate-after-write':
    case 'lookup-text':
    case 'dynamic-array':
      return {
        scorecardEligible: family !== 'dynamic-array',
        exclusionReason:
          family === 'dynamic-array' ? 'Leadership-only support lane; not an apples-to-apples performance scorecard input.' : null,
      }
  }
}

function readExpandedBaselineWorkloads(path: string): string[] {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  if (
    parsed === null ||
    typeof parsed !== 'object' ||
    !('results' in parsed) ||
    !Array.isArray(parsed.results) ||
    !parsed.results.every(
      (result) => result !== null && typeof result === 'object' && 'workload' in result && typeof result.workload === 'string',
    )
  ) {
    throw new Error(`Unexpected expanded baseline format: ${path}`)
  }
  return parsed.results.map((result) => result.workload)
}

describe('expanded comparative benchmark workloads', () => {
  it('enumerates the full expanded workload inventory without duplicates', () => {
    expect(new Set(EXPANDED_COMPARATIVE_WORKLOADS).size).toBe(EXPANDED_COMPARATIVE_WORKLOADS.length)
    expect(EXPANDED_COMPARATIVE_WORKLOADS).toEqual(expectedExpandedWorkloads)
  })

  it('checked-in expanded baseline covers every expanded workload exactly once', () => {
    expect(readExpandedBaselineWorkloads(expandedBaselinePath)).toEqual(expectedExpandedWorkloads)
  })

  it('assigns every expanded workload to exactly one family', () => {
    const seen = new Map<ExpandedComparativeBenchmarkWorkload, ExpandedCompetitiveFamily>()

    for (const family of EXPANDED_COMPARATIVE_FAMILY_ORDER) {
      const workloads = EXPANDED_COMPARATIVE_FAMILY_GROUPS[family]
      for (const workload of workloads) {
        expect(EXPANDED_COMPARATIVE_WORKLOAD_FAMILY[workload]).toBe(family)
        expect(seen.has(workload)).toBe(false)
        seen.set(workload, family)
      }
    }

    expect(seen.size).toBe(expectedExpandedWorkloads.length)
    expect([...seen.keys()].toSorted()).toEqual([...expectedExpandedWorkloads].toSorted())
  })

  it('formats an expanded family report with a machine-readable top-level families array', () => {
    const parsed: unknown = JSON.parse(formatExpandedCompetitiveFamilyReport([]))

    expect(parsed).toEqual({
      suite: 'workpaper-vs-hyperformula-expanded',
      families: EXPANDED_COMPARATIVE_FAMILY_ORDER.map((family) => ({
        family,
        workloads: EXPANDED_COMPARATIVE_FAMILY_GROUPS[family],
        ...familyEligibility(family),
        resultCount: 0,
        comparableCount: 0,
        leadershipCount: 0,
        workpaperWins: 0,
        hyperformulaWins: 0,
        meanSpeedupGeomean: null,
      })),
      scorecard: {
        eligibleFamilies: EXPANDED_COMPARATIVE_FAMILY_ORDER.filter((family) => familyEligibility(family).scorecardEligible),
        excludedFamilies: EXPANDED_COMPARATIVE_FAMILY_ORDER.filter((family) => !familyEligibility(family).scorecardEligible),
        comparableCount: 0,
        workpaperWins: 0,
        hyperformulaWins: 0,
      },
    })
  })

  it('builds an expanded benchmark report with attached family summaries', () => {
    expect(buildExpandedComparativeBenchmarkReport([])).toEqual({
      suite: 'workpaper-vs-hyperformula-expanded',
      results: [],
      families: EXPANDED_COMPARATIVE_FAMILY_ORDER.map((family) => ({
        family,
        workloads: EXPANDED_COMPARATIVE_FAMILY_GROUPS[family],
        ...familyEligibility(family),
        resultCount: 0,
        comparableCount: 0,
        leadershipCount: 0,
        workpaperWins: 0,
        hyperformulaWins: 0,
        meanSpeedupGeomean: null,
      })),
      scorecard: {
        eligibleFamilies: EXPANDED_COMPARATIVE_FAMILY_ORDER.filter((family) => familyEligibility(family).scorecardEligible),
        excludedFamilies: EXPANDED_COMPARATIVE_FAMILY_ORDER.filter((family) => !familyEligibility(family).scorecardEligible),
        comparableCount: 0,
        workpaperWins: 0,
        hyperformulaWins: 0,
      },
    })
  })
})
