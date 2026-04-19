import { writeFile } from 'node:fs/promises'
import { expect, test } from '@playwright/test'
import {
  gotoWorkbookShell,
  performHorizontalGridBrowse,
  startWorkbookScrollPerf,
  stopWorkbookScrollPerf,
  waitForBenchmarkCorpus,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

test.describe('web app scroll performance', () => {
  test.setTimeout(90_000)

  test('keeps horizontal browse inside one resident window smooth and free of subscription churn', async ({ page }, testInfo) => {
    await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
    await waitForWorkbookReady(page)
    const benchmarkState = await waitForBenchmarkCorpus(page)

    expect(benchmarkState.fixture?.id).toBe('wide-mixed-250k')

    await startWorkbookScrollPerf(page, 'wide-250k-main-body')
    await performHorizontalGridBrowse(page, { distancePx: 4_096, steps: 180 })
    const report = await stopWorkbookScrollPerf(page)

    if (!report) {
      throw new Error('scroll performance report was not available')
    }

    await writeFile(testInfo.outputPath('scroll-perf-wide-250k-main-body.json'), JSON.stringify(report, null, 2), 'utf8')

    expect(report.fixture?.id).toBe('wide-mixed-250k')
    expect(report.samples.frameMs.length).toBeGreaterThan(120)
    expect(report.summary.frameMs.p95).toBeLessThan(20)
    expect(report.summary.frameMs.p99).toBeLessThan(30)
    expect(report.summary.longTasksMs.max).toBeLessThan(50)
    expect(report.counters.viewportSubscriptions).toBe(0)
    expect(report.counters.fullPatches).toBe(0)
    expect(report.counters.damagePatches).toBe(0)
    expect(report.counters.domSurfaceMounts).toBe(0)

    await expect(page.getByTestId('grid-text-overlay')).toHaveJSProperty('tagName', 'CANVAS')
    await expect(page.locator('[data-testid="grid-text-overlay"] span')).toHaveCount(0)
  })
})
