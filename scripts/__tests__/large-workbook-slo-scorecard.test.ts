import { describe, expect, it } from 'vitest'

import { buildLargeWorkbookSloScorecard } from '../gen-large-workbook-slo-scorecard.ts'

describe('large workbook SLO scorecard', () => {
  it('maps benchmark-contract output into checked large-workbook and worker-runtime SLOs', () => {
    const scorecard = buildLargeWorkbookSloScorecard(buildReportFixture())

    expect(scorecard.summary.coveredLargeWorkbookRows).toEqual([100_000, 250_000])
    expect(scorecard.summary.allSloBudgetsPassed).toBe(true)
    expect(scorecard.measurements.map((measurement) => measurement.id)).toEqual([
      'load100k',
      'load250k',
      'workerWarmStart100k',
      'workerWarmStart250k',
      'workerVisibleEdit10k',
      'workerReconnectCatchUp100Pending',
    ])
    expect(scorecard.measurements.find((measurement) => measurement.id === 'workerVisibleEdit10k')).toMatchObject({
      category: 'ui-responsiveness',
      actualP95: 4,
      budgetP95: 16,
      passed: true,
    })
  })

  it('rejects reports that do not cover both 100k and 250k workbook sessions', () => {
    const report = buildReportFixture()
    delete report.results.load250k

    expect(() => buildLargeWorkbookSloScorecard(report)).toThrow('Missing large workbook SLO benchmark result: load250k')
  })
})

function buildReportFixture() {
  return {
    baseBudgets: {
      load100kP95Ms: 1500,
      load250kP95Ms: 1500,
      workerWarmStart100kP95Ms: 500,
      workerWarmStart250kP95Ms: 700,
      workerVisibleEdit10kP95Ms: 16,
      workerReconnectCatchUp100PendingP95Ms: 2000,
    },
    budgets: {
      load100kP95Ms: 1500,
      load250kP95Ms: 1500,
      workerWarmStart100kP95Ms: 500,
      workerWarmStart250kP95Ms: 700,
      workerVisibleEdit10kP95Ms: 16,
      workerReconnectCatchUp100PendingP95Ms: 2000,
    },
    toleranceMultiplier: 1,
    sampleCounts: {
      load100k: 5,
      load250k: 3,
      workerWarmStart100k: 3,
      workerWarmStart250k: 3,
      workerVisibleEdit10k: 5,
      workerReconnectCatchUp100Pending: 3,
    },
    results: {
      load100k: benchmarkResult('dense-mixed-100k', 100_000, 230),
      load250k: benchmarkResult('dense-mixed-250k', 250_000, 600),
      workerWarmStart100k: benchmarkResult('dense-mixed-100k', 100_000, 12),
      workerWarmStart250k: benchmarkResult('dense-mixed-250k', 250_000, 17),
      workerVisibleEdit10k: {
        materializedCells: 10_000,
        visiblePatchMs: numericSummary(4),
      },
      workerReconnectCatchUp100Pending: {
        materializedCells: 10_000,
        pendingMutationCount: 100,
        catchUpMs: numericSummary(270),
      },
    },
  }
}

function benchmarkResult(corpusCaseId: string, materializedCells: number, p95: number) {
  return {
    corpusCaseId,
    materializedCells,
    elapsedMs: numericSummary(p95),
  }
}

function numericSummary(p95: number) {
  return {
    samples: [p95],
    min: p95,
    median: p95,
    p95,
    max: p95,
    mean: p95,
  }
}
