import { describe, expect, it } from 'vitest'
import { runIsolatedEditBenchmark } from '../../../../scripts/bench-contracts.ts'

describe('bench contracts runner', () => {
  it('runs isolated benchmarks through node and tsx', async () => {
    const result = await runIsolatedEditBenchmark(100)

    expect(result.elapsedMs).toBeGreaterThanOrEqual(0)
    expect(result.metrics.recalcMs).toBeGreaterThanOrEqual(0)
    expect(result.memory.after.heapUsedBytes).toBeGreaterThan(0)
  })
})
