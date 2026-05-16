import { describe, expect, it } from 'vitest'
import {
  assertRangeAggregateRunsUseFastPath,
  runIsolatedEditBenchmark,
  runIsolatedRangeAggregateBenchmark,
} from '../../../../scripts/bench-contracts.ts'
import { parseBenchToleranceMultiplier } from '../../../../scripts/bench-tolerance.ts'

describe('bench contracts runner', () => {
  it('parses benchmark tolerance overrides without truncating malformed input', () => {
    expect(parseBenchToleranceMultiplier(undefined, false)).toBe(1)
    expect(parseBenchToleranceMultiplier(undefined, true)).toBe(1.5)
    expect(parseBenchToleranceMultiplier('1.25', false)).toBe(1.25)
    expect(() => parseBenchToleranceMultiplier('1abc', false)).toThrow('BILIG_BENCH_TOLERANCE must be a positive number, got 1abc')
    expect(() => parseBenchToleranceMultiplier('0', false)).toThrow('BILIG_BENCH_TOLERANCE must be a positive number, got 0')
  })

  it('runs isolated benchmarks through node and tsx', async () => {
    const result = await runIsolatedEditBenchmark(100)

    expect(result.elapsedMs).toBeGreaterThanOrEqual(0)
    expect(result.metrics.recalcMs).toBeGreaterThanOrEqual(0)
    expect(result.memory.after.heapUsedBytes).toBeGreaterThan(0)
  }, 30_000)

  it('keeps range aggregate contract output on the direct aggregate fast path', async () => {
    const result = await runIsolatedRangeAggregateBenchmark(64, 100)

    expect(result.metrics.jsFormulaCount).toBe(0)
    expect(result.performanceCounters.directAggregateDeltaApplications).toBe(100)
    expect(result.performanceCounters.directAggregateScanCells).toBe(0)
    assertRangeAggregateRunsUseFastPath('range aggregate test benchmark', [result], 100)
  }, 30_000)

  it('rejects range aggregate runs that miss the wasm and direct aggregate fast paths', () => {
    expect(() =>
      assertRangeAggregateRunsUseFastPath(
        'range aggregate test benchmark',
        [
          {
            elapsedMs: 1,
            metrics: {
              recalcMs: 1,
              wasmFormulaCount: 0,
              jsFormulaCount: 1,
            },
            performanceCounters: {
              directAggregateDeltaApplications: 0,
              directAggregateDeltaOnlyRecalcSkips: 0,
              directAggregateScanEvaluations: 1,
              directAggregateScanCells: 64,
            },
            memory: {
              before: {
                rssBytes: 1,
                heapUsedBytes: 1,
                heapTotalBytes: 1,
                externalBytes: 1,
                arrayBuffersBytes: 1,
              },
              after: {
                rssBytes: 1,
                heapUsedBytes: 1,
                heapTotalBytes: 1,
                externalBytes: 1,
                arrayBuffersBytes: 1,
              },
              delta: {
                rssBytes: 0,
                heapUsedBytes: 0,
                heapTotalBytes: 0,
                externalBytes: 0,
                arrayBuffersBytes: 0,
              },
            },
          },
        ],
        100,
      ),
    ).toThrow(/supported aggregate fast path/)
  })
})
