import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import {
  EXPANDED_COMPARATIVE_WORKLOADS,
  type ExpandedComparativeBenchmarkWorkload,
} from '../benchmark-workpaper-vs-hyperformula-expanded.js'

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
})
