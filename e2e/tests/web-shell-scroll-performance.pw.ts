import { writeFile } from 'node:fs/promises'
import { expect, test } from '@playwright/test'
import type { Page } from '@playwright/test'
import {
  clickProductCell,
  gotoWorkbookShell,
  performDiagonalGridBrowse,
  performHorizontalGridBrowse,
  performVerticalGridBrowse,
  remoteSyncEnabled,
  settleWorkbookScrollPerf,
  stopWorkbookScrollPerf,
  warmStartWorkbookScrollPerf,
  waitForBenchmarkCorpus,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

function expectQuietShell(
  report: NonNullable<Awaited<ReturnType<typeof stopWorkbookScrollPerf>>>,
  options: {
    readonly maxSurfaceCommits?: number
  } = {},
) {
  const maxSurfaceCommits = options.maxSurfaceCommits ?? 0
  expect(report.counters.surfaceCommits.formulaBar ?? 0).toBeLessThanOrEqual(maxSurfaceCommits)
  expect(report.counters.surfaceCommits.statusBar ?? 0).toBeLessThanOrEqual(maxSurfaceCommits)
  expect(report.counters.surfaceCommits.sheetTabs ?? 0).toBeLessThanOrEqual(maxSurfaceCommits)
}

function expectSmoothBrowse(
  report: NonNullable<Awaited<ReturnType<typeof stopWorkbookScrollPerf>>>,
  options: {
    readonly p99Max?: number
    readonly longTaskMax?: number
  } = {},
) {
  expect(report.samples.frameMs.length).toBeGreaterThan(120)
  expect(report.summary.frameMs.p95).toBeLessThan(20)
  expect(report.summary.frameMs.p99).toBeLessThan(options.p99Max ?? 30)
  expect(report.summary.longTasksMs.max).toBeLessThan(options.longTaskMax ?? 50)
  expect(report.counters.viewportSubscriptions).toBe(0)
  expect(report.counters.domSurfaceMounts).toBe(0)
}

async function expectTypeGpuSteadyScroll(page: Page, report: NonNullable<Awaited<ReturnType<typeof stopWorkbookScrollPerf>>>) {
  const supportsWebGpu = await page.evaluate(() => 'gpu' in navigator)
  if (!supportsWebGpu) {
    return
  }
  expect(report.counters.typeGpuConfigures).toBe(0)
  expect(report.counters.typeGpuSurfaceResizes).toBe(0)
  expect(report.counters.typeGpuBufferAllocations).toBe(0)
  const hasSceneChurn =
    report.counters.fullPatches > 0 || report.counters.headerPaneBuilds > 0 || report.counters.typeGpuScenePacketsApplied > 0
  if (!hasSceneChurn) {
    expect(report.counters.typeGpuVertexUploadBytes).toBe(0)
  }
  expect(report.counters.typeGpuSubmits).toBeGreaterThan(0)
}

test.describe('@browser-perf web app scroll performance', () => {
  test.setTimeout(120_000)

  test('keeps horizontal browse inside one resident window smooth and free of data-canvas redraw churn', async ({ page }, testInfo) => {
    await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
    await waitForWorkbookReady(page)
    const benchmarkState = await waitForBenchmarkCorpus(page)

    expect(benchmarkState.fixture?.id).toBe('wide-mixed-250k')

    await settleWorkbookScrollPerf(page, 80)
    await warmStartWorkbookScrollPerf(page, 'wide-250k-main-body')
    await performHorizontalGridBrowse(page, { distancePx: 4_096, steps: 180 })
    const report = await stopWorkbookScrollPerf(page)

    if (!report) {
      throw new Error('scroll performance report was not available')
    }

    await writeFile(testInfo.outputPath('scroll-perf-wide-250k-main-body.json'), JSON.stringify(report, null, 2), 'utf8')

    expect(report.fixture?.id).toBe('wide-mixed-250k')
    expectSmoothBrowse(report, { longTaskMax: 60 })
    expectQuietShell(report, { maxSurfaceCommits: 1 })
    expect(report.counters.damagePatches).toBe(0)
    expect(report.counters.scenePacketRefreshes).toBe(0)
    expect(report.counters.canvasPaints['text:body'] ?? 0).toBeLessThanOrEqual(1)
    expect(report.counters.canvasPaints['gpu:body'] ?? 0).toBeLessThanOrEqual(1)
    await expectTypeGpuSteadyScroll(page, report)
    const supportsWebGpu = await page.evaluate(() => 'gpu' in navigator)
    if (supportsWebGpu) {
      await expect(page.getByTestId('grid-pane-renderer')).toHaveJSProperty('tagName', 'CANVAS')
      await expect(page.getByTestId('grid-text-pane-body')).toHaveCount(0)
    } else {
      await expect(page.getByTestId('grid-text-pane-body')).toHaveJSProperty('tagName', 'CANVAS')
      await expect(page.getByTestId('grid-text-pane-top-body')).toHaveJSProperty('tagName', 'CANVAS')
      await expect(page.getByTestId('grid-text-pane-left-body')).toHaveJSProperty('tagName', 'CANVAS')
    }
  })

  test('keeps frozen-pane browse smooth without repainting resident body or frozen data panes inside a tile window', async ({
    page,
  }, testInfo) => {
    await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-frozen-250k')
    await waitForWorkbookReady(page)
    const benchmarkState = await waitForBenchmarkCorpus(page)

    expect(benchmarkState.fixture?.id).toBe('wide-mixed-frozen-250k')

    await settleWorkbookScrollPerf(page, 40)
    await warmStartWorkbookScrollPerf(page, 'wide-250k-frozen-panes')
    await performHorizontalGridBrowse(page, { distancePx: 3_072, steps: 160 })
    const report = await stopWorkbookScrollPerf(page)

    if (!report) {
      throw new Error('scroll performance report was not available')
    }

    await writeFile(testInfo.outputPath('scroll-perf-wide-250k-frozen.json'), JSON.stringify(report, null, 2), 'utf8')

    expect(report.fixture?.id).toBe('wide-mixed-frozen-250k')
    expectSmoothBrowse(report, { p99Max: 35, longTaskMax: 60 })
    expectQuietShell(report, { maxSurfaceCommits: 4 })
    expect(report.counters.damagePatches).toBe(0)
    expect(report.counters.scenePacketRefreshes).toBe(0)
    expect(report.counters.canvasPaints['text:body'] ?? 0).toBeLessThanOrEqual(2)
    expect(report.counters.canvasPaints['text:top'] ?? 0).toBeLessThanOrEqual(2)
    expect(report.counters.canvasPaints['text:left'] ?? 0).toBeLessThanOrEqual(2)
    expect(report.counters.canvasPaints['gpu:body'] ?? 0).toBeLessThanOrEqual(2)
    expect(report.counters.canvasPaints['gpu:top'] ?? 0).toBeLessThanOrEqual(2)
    expect(report.counters.canvasPaints['gpu:left'] ?? 0).toBeLessThanOrEqual(2)
    await expectTypeGpuSteadyScroll(page, report)
  })

  test('keeps deep vertical browse smooth inside one resident window', async ({ page }, testInfo) => {
    await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
    await waitForWorkbookReady(page)
    const benchmarkState = await waitForBenchmarkCorpus(page)

    expect(benchmarkState.fixture?.id).toBe('wide-mixed-250k')

    await settleWorkbookScrollPerf(page, 80)
    await warmStartWorkbookScrollPerf(page, 'wide-250k-vertical-main-body')
    await performVerticalGridBrowse(page, { distancePx: 440, steps: 140 })
    const report = await stopWorkbookScrollPerf(page)

    if (!report) {
      throw new Error('scroll performance report was not available')
    }

    await writeFile(testInfo.outputPath('scroll-perf-wide-250k-vertical.json'), JSON.stringify(report, null, 2), 'utf8')

    expect(report.fixture?.id).toBe('wide-mixed-250k')
    expectSmoothBrowse(report, { longTaskMax: 60 })
    expectQuietShell(report, { maxSurfaceCommits: 1 })
    expect(report.counters.damagePatches).toBe(0)
    expect(report.counters.scenePacketRefreshes).toBe(0)
    await expectTypeGpuSteadyScroll(page, report)
  })

  test('keeps diagonal browse smooth inside one resident window', async ({ page }, testInfo) => {
    await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
    await waitForWorkbookReady(page)
    const benchmarkState = await waitForBenchmarkCorpus(page)

    expect(benchmarkState.fixture?.id).toBe('wide-mixed-250k')

    await settleWorkbookScrollPerf(page, 80)
    await warmStartWorkbookScrollPerf(page, 'wide-250k-diagonal-main-body')
    await performDiagonalGridBrowse(page, { deltaX: 2_048, deltaY: 440, steps: 160 })
    const report = await stopWorkbookScrollPerf(page)

    if (!report) {
      throw new Error('scroll performance report was not available')
    }

    await writeFile(testInfo.outputPath('scroll-perf-wide-250k-diagonal.json'), JSON.stringify(report, null, 2), 'utf8')

    expect(report.fixture?.id).toBe('wide-mixed-250k')
    expectSmoothBrowse(report, { longTaskMax: 60 })
    expectQuietShell(report, { maxSurfaceCommits: 1 })
    expect(report.counters.damagePatches).toBe(0)
    expect(report.counters.scenePacketRefreshes).toBe(0)
    await expectTypeGpuSteadyScroll(page, report)
  })

  test('keeps variable-width browse smooth without resubscribing or remounting grid text surfaces', async ({ page }, testInfo) => {
    await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-variable-250k')
    await waitForWorkbookReady(page)
    const benchmarkState = await waitForBenchmarkCorpus(page)

    expect(benchmarkState.fixture?.id).toBe('wide-mixed-variable-250k')

    await settleWorkbookScrollPerf(page, 40)
    await warmStartWorkbookScrollPerf(page, 'wide-250k-variable-widths')
    await performHorizontalGridBrowse(page, { distancePx: 3_840, steps: 180 })
    const report = await stopWorkbookScrollPerf(page)

    if (!report) {
      throw new Error('scroll performance report was not available')
    }

    await writeFile(testInfo.outputPath('scroll-perf-wide-250k-variable.json'), JSON.stringify(report, null, 2), 'utf8')

    expect(report.fixture?.id).toBe('wide-mixed-variable-250k')
    expectSmoothBrowse(report)
    expectQuietShell(report)
    expect(report.counters.damagePatches).toBe(0)
    expect(report.counters.scenePacketRefreshes).toBe(0)
    expect(report.counters.canvasSurfaceMounts).toBe(0)
    expect(report.counters.canvasPaints['text:body'] ?? 0).toBe(0)
    expect(report.counters.canvasPaints['gpu:body'] ?? 0).toBe(0)
    await expectTypeGpuSteadyScroll(page, report)
  })

  test('keeps shell surfaces quiet and coalesces visible collaborator patch churn while browsing', async ({ page }, testInfo) => {
    test.skip(!remoteSyncEnabled, 'requires Zero-backed browser sync')
    const documentId = `playwright-zero-scroll-patches-${Date.now()}`
    const mirrorPage = await page.context().newPage()
    const viewport = page.viewportSize()
    if (viewport) {
      await mirrorPage.setViewportSize(viewport)
    }

    try {
      await Promise.all([
        gotoWorkbookShell(page, `/?document=${encodeURIComponent(documentId)}&benchmarkCorpus=wide-mixed-250k`),
        gotoWorkbookShell(mirrorPage, `/?document=${encodeURIComponent(documentId)}&benchmarkCorpus=wide-mixed-250k`),
      ])
      await Promise.all([waitForWorkbookReady(page), waitForWorkbookReady(mirrorPage)])
      await Promise.all([waitForBenchmarkCorpus(page), waitForBenchmarkCorpus(mirrorPage)])
      await settleWorkbookScrollPerf(page, 40)

      const emitRemoteEdits = async () => {
        const formulaInput = mirrorPage.getByTestId('formula-input')
        const cells: ReadonlyArray<readonly [number, number, string]> = [
          [0, 4, '101'],
          [1, 5, '102'],
          [2, 6, '103'],
          [3, 7, '104'],
          [4, 8, '105'],
          [5, 9, '106'],
        ]
        await clickProductCell(mirrorPage, 0, 0)
        const applyCell = async (index: number): Promise<void> => {
          const entry = cells[index]
          if (!entry) {
            return
          }
          const [columnIndex, rowIndex, value] = entry
          await clickProductCell(mirrorPage, columnIndex, rowIndex)
          await formulaInput.fill(value)
          await formulaInput.press('Enter')
          await applyCell(index + 1)
        }
        await applyCell(0)
      }

      await warmStartWorkbookScrollPerf(page, 'wide-250k-browse-with-visible-patches')
      await Promise.all([performHorizontalGridBrowse(page, { distancePx: 2_560, steps: 140 }), emitRemoteEdits()])
      await settleWorkbookScrollPerf(page, 20)
      const report = await stopWorkbookScrollPerf(page)

      if (!report) {
        throw new Error('scroll performance report was not available')
      }

      await writeFile(testInfo.outputPath('scroll-perf-wide-250k-visible-patches.json'), JSON.stringify(report, null, 2), 'utf8')

      expect(report.fixture?.id).toBe('wide-mixed-250k')
      test.skip(
        report.counters.damagePatches === 0 && report.counters.scenePacketRefreshes === 0,
        'remote edits did not arrive during the sampling window',
      )
      expect(report.summary.frameMs.p95).toBeLessThan(20)
      expect(report.summary.frameMs.p99).toBeLessThan(30)
      expect(report.summary.longTasksMs.max).toBeLessThan(50)
      expect(report.counters.viewportSubscriptions).toBe(0)
      expect(report.counters.fullPatches).toBe(0)
      expect(report.counters.damagePatches).toBeGreaterThan(0)
      expect(report.counters.damagePatches).toBeLessThanOrEqual(6)
      expect(report.counters.scenePacketRefreshes).toBeGreaterThan(0)
      expect(report.counters.scenePacketRefreshes).toBeLessThanOrEqual(6)
      expect(report.counters.canvasPaints['text:body'] ?? 0).toBeLessThanOrEqual(6)
      expect(report.counters.canvasPaints['gpu:body'] ?? 0).toBeLessThanOrEqual(6)
      expectQuietShell(report)
    } finally {
      await mirrorPage.close().catch(() => undefined)
    }
  })
})
