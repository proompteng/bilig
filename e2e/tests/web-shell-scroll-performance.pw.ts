import { writeFile } from 'node:fs/promises'
import { expect, test } from '@playwright/test'
import type { Page } from '@playwright/test'
import {
  clickProductCell,
  createTestDocumentId,
  dragProductColumnResize,
  dragProductRowResize,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductFillHandleDragPoints,
  gotoWorkbookShell,
  performDiagonalGridBrowse,
  performHorizontalGridBrowse,
  performVerticalGridBrowse,
  PRODUCT_HEADER_HEIGHT,
  remoteSyncEnabled,
  resetGridScroll,
  settleWorkbookScrollPerf,
  stopWorkbookScrollPerf,
  warmStartWorkbookScrollPerf,
  waitForBenchmarkCorpus,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

type ScrollPerfReport = NonNullable<Awaited<ReturnType<typeof stopWorkbookScrollPerf>>>
type ScrollPerfCounters = ScrollPerfReport['counters']
const remoteSyncTest = remoteSyncEnabled ? test : test.skip.bind(test)

function readCounter(counters: ScrollPerfCounters, key: keyof ScrollPerfCounters): number {
  return counters[key] ?? 0
}

function summarizeSamples(samples: readonly number[]): {
  readonly p95: number
  readonly p99: number
} {
  if (samples.length === 0) {
    return { p95: 0, p99: 0 }
  }
  const sorted = [...samples].toSorted((left, right) => left - right)
  return {
    p95: sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * 0.95) - 1))] ?? 0,
    p99: sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * 0.99) - 1))] ?? 0,
  }
}

function sumRecordCounters(counters: Readonly<Record<string, number>>): number {
  return Object.values(counters).reduce((sum, value) => sum + value, 0)
}

function expectQuietShell(
  report: ScrollPerfReport,
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
  report: ScrollPerfReport,
  options: {
    readonly p99Max?: number
    readonly longTaskMax?: number
    readonly ignoreInitialSamples?: number
    readonly maxViewportSubscriptions?: number
  } = {},
) {
  const frameSamples = report.samples.frameMs.slice(options.ignoreInitialSamples ?? 0)
  const frameSummary = summarizeSamples(frameSamples)
  expect(report.samples.frameMs.length).toBeGreaterThan(120)
  expect(frameSummary.p95).toBeLessThan(20)
  expect(frameSummary.p99).toBeLessThan(options.p99Max ?? 30)
  expect(report.summary.longTasksMs.max).toBeLessThan(options.longTaskMax ?? 50)
  expect(report.counters.viewportSubscriptions).toBeLessThanOrEqual(options.maxViewportSubscriptions ?? 0)
  expect(report.counters.domSurfaceMounts).toBe(0)
}

async function expectTypeGpuSteadyScroll(page: Page, report: ScrollPerfReport) {
  const supportsWebGpu = await page.evaluate(() => 'gpu' in navigator)
  if (!supportsWebGpu) {
    return
  }
  expect(readCounter(report.counters, 'typeGpuConfigures')).toBe(0)
  expect(readCounter(report.counters, 'typeGpuSurfaceResizes')).toBe(0)
  expect(report.counters.fullPatches).toBe(0)
  expect(sumRecordCounters(report.counters.fullPatchBroadcasts)).toBe(0)
  expect(readCounter(report.counters, 'typeGpuTileCacheSorts')).toBe(0)
  expect(readCounter(report.counters, 'rendererTileMisses')).toBe(0)
  expectNoTypeGpuTextAtlasGeometryChurn(report)
  const hasTileChurn =
    report.counters.headerPaneBuilds > 0 ||
    readCounter(report.counters, 'rendererTileInterestBatches') > 0 ||
    readCounter(report.counters, 'rendererVisibleDirtyTiles') > 0
  if (!hasTileChurn) {
    expect(readCounter(report.counters, 'typeGpuBufferAllocations')).toBe(0)
    expect(readCounter(report.counters, 'typeGpuVertexUploadBytes')).toBe(0)
    expectNoTypeGpuTextPayloadChurn(report)
  }
  if ('typeGpuSubmits' in report.counters) {
    expect(readCounter(report.counters, 'typeGpuSubmits')).toBeGreaterThan(0)
  }
}

function expectNoTypeGpuTextAtlasGeometryChurn(report: ScrollPerfReport) {
  expect(readCounter(report.counters, 'typeGpuTextAtlasGeometryRetries')).toBe(0)
  expect(readCounter(report.counters, 'typeGpuTextAtlasGeometryResyncs')).toBe(0)
}

function expectNoTypeGpuTextPayloadChurn(report: ScrollPerfReport) {
  expect(readCounter(report.counters, 'typeGpuTextRunPayloadRebuilds')).toBe(0)
  expect(readCounter(report.counters, 'typeGpuTextRunPayloadReuses')).toBe(0)
  expect(readCounter(report.counters, 'typeGpuTextGlyphDependencies')).toBe(0)
  expect(readCounter(report.counters, 'typeGpuTextPageDependencies')).toBe(0)
}

function expectNoRendererMutationChurn(report: ScrollPerfReport) {
  expect(report.counters.fullPatches).toBe(0)
  expect(report.counters.damagePatches).toBe(0)
  expect(readCounter(report.counters, 'rendererDeltaBatches')).toBe(0)
  expect(readCounter(report.counters, 'dirtyTilesMarked')).toBe(0)
  expect(readCounter(report.counters, 'rendererVisibleDirtyTiles')).toBe(0)
}

function expectNoTypeGpuDataTileUpload(report: ScrollPerfReport) {
  const totalVertexUploadBytes = readCounter(report.counters, 'typeGpuVertexUploadBytes')
  const overlayUploadBytes = readCounter(report.counters, 'typeGpuOverlayUploadBytes')
  expect(Math.max(0, totalVertexUploadBytes - overlayUploadBytes)).toBe(0)
}

function expectBoundedVisibleMutation(
  report: ScrollPerfReport,
  options: {
    readonly minDamagePatches?: number
    readonly maxDamagePatches?: number
    readonly maxRendererDeltaBatches?: number
    readonly maxRendererDeltaMutations?: number
    readonly maxRendererVisibleDirtyTiles?: number
    readonly maxRendererWarmDirtyTiles?: number
    readonly mutationToVisibleP95Max?: number
    readonly frameP95Max?: number
    readonly frameP99Max?: number
    readonly longTaskMax?: number
  } = {},
) {
  expect(report.summary.frameMs.p95).toBeLessThan(options.frameP95Max ?? 20)
  expect(report.summary.frameMs.p99).toBeLessThan(options.frameP99Max ?? 30)
  expect(report.summary.longTasksMs.max).toBeLessThan(options.longTaskMax ?? 50)
  expect(report.counters.viewportSubscriptions).toBe(0)
  expect(report.counters.fullPatches).toBe(0)
  expect(sumRecordCounters(report.counters.fullPatchBroadcasts)).toBe(0)
  expect(report.counters.damagePatches).toBeGreaterThanOrEqual(options.minDamagePatches ?? 0)
  expect(report.counters.damagePatches).toBeLessThanOrEqual(options.maxDamagePatches ?? 4)
  expect(readCounter(report.counters, 'rendererDeltaBatches')).toBeGreaterThan(0)
  expect(readCounter(report.counters, 'rendererDeltaBatches')).toBeLessThanOrEqual(options.maxRendererDeltaBatches ?? 4)
  expect(readCounter(report.counters, 'rendererDeltaMutations')).toBeGreaterThan(0)
  expect(readCounter(report.counters, 'rendererDeltaMutations')).toBeLessThanOrEqual(
    options.maxRendererDeltaMutations ?? Number.MAX_SAFE_INTEGER,
  )
  expect(readCounter(report.counters, 'dirtyTilesMarked')).toBeGreaterThan(0)
  expect(readCounter(report.counters, 'rendererVisibleDirtyTiles')).toBeGreaterThan(0)
  expect(readCounter(report.counters, 'rendererVisibleDirtyTiles')).toBeLessThanOrEqual(
    options.maxRendererVisibleDirtyTiles ?? Number.MAX_SAFE_INTEGER,
  )
  expect(readCounter(report.counters, 'rendererWarmDirtyTiles')).toBeLessThanOrEqual(
    options.maxRendererWarmDirtyTiles ?? Number.MAX_SAFE_INTEGER,
  )
  expect(report.samples.mutationToVisibleMs.length).toBeGreaterThan(0)
  expect(report.summary.mutationToVisibleMs.p95).toBeLessThan(options.mutationToVisibleP95Max ?? 100)
  expect(readCounter(report.counters, 'rendererTileMisses')).toBeLessThanOrEqual(readCounter(report.counters, 'rendererVisibleDirtyTiles'))
  expect(readCounter(report.counters, 'typeGpuTileCacheSorts')).toBe(0)
  expectNoTypeGpuTextAtlasGeometryChurn(report)
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
    await performHorizontalGridBrowse(page, { distancePx: 3_072, steps: 80 })
    await resetGridScroll(page)
    await settleWorkbookScrollPerf(page, 40)
    await warmStartWorkbookScrollPerf(page, 'wide-250k-frozen-panes')
    await performHorizontalGridBrowse(page, { distancePx: 3_072, steps: 160 })
    const report = await stopWorkbookScrollPerf(page)

    if (!report) {
      throw new Error('scroll performance report was not available')
    }

    await writeFile(testInfo.outputPath('scroll-perf-wide-250k-frozen.json'), JSON.stringify(report, null, 2), 'utf8')

    expect(report.fixture?.id).toBe('wide-mixed-frozen-250k')
    expectSmoothBrowse(report, { ignoreInitialSamples: 10, p99Max: 35, longTaskMax: 60, maxViewportSubscriptions: 2 })
    expectQuietShell(report, { maxSurfaceCommits: 4 })
    expect(report.counters.damagePatches).toBe(0)
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
    expect(report.counters.canvasSurfaceMounts).toBe(0)
    expect(report.counters.canvasPaints['text:body'] ?? 0).toBe(0)
    expect(report.counters.canvasPaints['gpu:body'] ?? 0).toBe(0)
    await expectTypeGpuSteadyScroll(page, report)
  })

  test('keeps tile-boundary browse bounded without blanking the typegpu renderer', async ({ page }, testInfo) => {
    await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
    await waitForWorkbookReady(page)
    const benchmarkState = await waitForBenchmarkCorpus(page)

    expect(benchmarkState.fixture?.id).toBe('wide-mixed-250k')

    await settleWorkbookScrollPerf(page, 80)
    await warmStartWorkbookScrollPerf(page, 'wide-250k-tile-boundary')
    await performDiagonalGridBrowse(page, { deltaX: 13_520, deltaY: 760, steps: 180 })
    const report = await stopWorkbookScrollPerf(page)

    if (!report) {
      throw new Error('scroll performance report was not available')
    }

    await writeFile(testInfo.outputPath('scroll-perf-wide-250k-tile-boundary.json'), JSON.stringify(report, null, 2), 'utf8')

    expect(report.fixture?.id).toBe('wide-mixed-250k')
    expect(report.samples.frameMs.length).toBeGreaterThan(120)
    expect(report.summary.frameMs.p95).toBeLessThan(24)
    expect(report.summary.longTasksMs.max).toBeLessThan(60)
    expect(readCounter(report.counters, 'typeGpuTileMisses')).toBe(0)
    expect(readCounter(report.counters, 'rendererTileMisses')).toBe(0)
    expectNoTypeGpuTextAtlasGeometryChurn(report)
    expect(report.counters.viewportSubscriptions).toBeLessThanOrEqual(12)
  })

  test('keeps resized and editing-overlay scroll inside frame budgets', async ({ page }, testInfo) => {
    await page.setViewportSize({ width: 840, height: 620 })
    await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
    await waitForWorkbookReady(page)
    const benchmarkState = await waitForBenchmarkCorpus(page)

    expect(benchmarkState.fixture?.id).toBe('wide-mixed-250k')

    await clickProductCell(page, 1, 1)
    await page.keyboard.press('F2')
    await expect(page.getByTestId('cell-editor-input')).toBeVisible()
    await page.setViewportSize({ width: 920, height: 680 })
    await settleWorkbookScrollPerf(page, 40)
    await warmStartWorkbookScrollPerf(page, 'wide-250k-editing-resize-scroll')
    await performDiagonalGridBrowse(page, { deltaX: 1_536, deltaY: 330, steps: 140 })
    const report = await stopWorkbookScrollPerf(page)

    if (!report) {
      throw new Error('scroll performance report was not available')
    }

    await writeFile(testInfo.outputPath('scroll-perf-wide-250k-editing-resize.json'), JSON.stringify(report, null, 2), 'utf8')

    expect(report.fixture?.id).toBe('wide-mixed-250k')
    expect(report.samples.frameMs.length).toBeGreaterThan(100)
    expect(report.summary.frameMs.p95).toBeLessThan(24)
    expect(report.summary.longTasksMs.max).toBeLessThan(60)
    expect(readCounter(report.counters, 'typeGpuConfigures')).toBe(0)
    expectNoTypeGpuTextAtlasGeometryChurn(report)
    if ('typeGpuSubmits' in report.counters) {
      expect(readCounter(report.counters, 'typeGpuSubmits')).toBeGreaterThan(0)
    }
  })

  test('keeps committed column resize bounded with controlled V3 dirty-tile writes', async ({ page }, testInfo) => {
    await page.setViewportSize({ width: 920, height: 680 })
    await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
    await waitForWorkbookReady(page)
    const benchmarkState = await waitForBenchmarkCorpus(page)

    expect(benchmarkState.fixture?.id).toBe('wide-mixed-250k')

    await settleWorkbookScrollPerf(page, 40)
    await warmStartWorkbookScrollPerf(page, 'wide-250k-commit-column-resize')
    await dragProductColumnResize(page, 1, 64)
    await settleWorkbookScrollPerf(page, 24)
    const report = await stopWorkbookScrollPerf(page)

    if (!report) {
      throw new Error('scroll performance report was not available')
    }

    await writeFile(testInfo.outputPath('scroll-perf-wide-250k-column-resize-commit.json'), JSON.stringify(report, null, 2), 'utf8')

    expect(report.fixture?.id).toBe('wide-mixed-250k')
    expectBoundedVisibleMutation(report, {
      maxDamagePatches: 16,
      maxRendererDeltaBatches: 6,
      maxRendererDeltaMutations: 80,
      maxRendererVisibleDirtyTiles: 64,
      maxRendererWarmDirtyTiles: 24,
      mutationToVisibleP95Max: 80,
      frameP95Max: 30,
      longTaskMax: 70,
    })
    expectQuietShell(report, { maxSurfaceCommits: 4 })
  })

  test('keeps committed row resize bounded with controlled V3 dirty-tile writes', async ({ page }, testInfo) => {
    await page.setViewportSize({ width: 920, height: 680 })
    await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
    await waitForWorkbookReady(page)
    const benchmarkState = await waitForBenchmarkCorpus(page)

    expect(benchmarkState.fixture?.id).toBe('wide-mixed-250k')

    await settleWorkbookScrollPerf(page, 40)
    await warmStartWorkbookScrollPerf(page, 'wide-250k-commit-row-resize')
    await dragProductRowResize(page, 1, 64)
    await settleWorkbookScrollPerf(page, 24)
    const report = await stopWorkbookScrollPerf(page)

    if (!report) {
      throw new Error('scroll performance report was not available')
    }

    await writeFile(testInfo.outputPath('scroll-perf-wide-250k-row-resize-commit.json'), JSON.stringify(report, null, 2), 'utf8')

    expect(report.fixture?.id).toBe('wide-mixed-250k')
    expectBoundedVisibleMutation(report, {
      maxDamagePatches: 16,
      maxRendererDeltaBatches: 6,
      maxRendererDeltaMutations: 80,
      maxRendererVisibleDirtyTiles: 64,
      maxRendererWarmDirtyTiles: 24,
      mutationToVisibleP95Max: 80,
      frameP95Max: 30,
      longTaskMax: 70,
    })
    expectQuietShell(report, { maxSurfaceCommits: 4 })
  })

  test('keeps column resize preview overlay-only with no renderer mutation or data-tile upload', async ({ page }, testInfo) => {
    await page.setViewportSize({ width: 920, height: 680 })
    await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
    await waitForWorkbookReady(page)
    const benchmarkState = await waitForBenchmarkCorpus(page)

    expect(benchmarkState.fixture?.id).toBe('wide-mixed-250k')

    const gridLocator = page.getByTestId('sheet-grid')
    await expect(gridLocator).toBeVisible()
    const grid = await gridLocator.boundingBox()
    if (!grid) {
      throw new Error('sheet grid is not visible')
    }

    const columnLeft = await getProductColumnLeft(page, 1)
    const columnWidth = await getProductColumnWidth(page, 1)
    const edgeX = grid.x + columnLeft + columnWidth - 1
    const edgeY = grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2)

    await settleWorkbookScrollPerf(page, 40)
    await warmStartWorkbookScrollPerf(page, 'wide-250k-column-resize-preview')
    try {
      await page.mouse.move(edgeX, edgeY)
      await page.mouse.down()
      await page.mouse.move(edgeX + 56, edgeY, { steps: 24 })
      await settleWorkbookScrollPerf(page, 16)
      const report = await stopWorkbookScrollPerf(page)

      if (!report) {
        throw new Error('scroll performance report was not available')
      }

      await writeFile(testInfo.outputPath('scroll-perf-wide-250k-column-resize-preview.json'), JSON.stringify(report, null, 2), 'utf8')

      expect(report.fixture?.id).toBe('wide-mixed-250k')
      expect(report.summary.frameMs.p95).toBeLessThan(20)
      expect(report.summary.longTasksMs.max).toBeLessThan(50)
      expect(report.counters.viewportSubscriptions).toBe(0)
      expectNoRendererMutationChurn(report)
      expectNoTypeGpuDataTileUpload(report)
      expect(readCounter(report.counters, 'typeGpuBufferAllocations')).toBe(0)
      expect(readCounter(report.counters, 'rendererTileMisses')).toBe(0)
      expectQuietShell(report)
    } finally {
      await page.mouse.up().catch(() => undefined)
    }
  })

  test('keeps fill-handle preview overlay-only with no renderer mutation or data-tile upload', async ({ page }, testInfo) => {
    await page.setViewportSize({ width: 920, height: 680 })
    await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
    await waitForWorkbookReady(page)
    const benchmarkState = await waitForBenchmarkCorpus(page)

    expect(benchmarkState.fixture?.id).toBe('wide-mixed-250k')

    const nameBox = page.getByTestId('name-box')
    const formulaInput = page.getByTestId('formula-input')
    await nameBox.fill('F6')
    await nameBox.press('Enter')
    await formulaInput.fill('7')
    await formulaInput.press('Enter')
    await expect(page.getByTestId('status-selection')).toContainText('!F6')

    const { sourceX, sourceY, targetX, targetY } = await getProductFillHandleDragPoints(page, 5, 5, 8, 5)

    await settleWorkbookScrollPerf(page, 40)
    await warmStartWorkbookScrollPerf(page, 'wide-250k-fill-preview')
    try {
      await page.mouse.move(sourceX, sourceY)
      await page.mouse.down()
      await page.mouse.move(targetX, targetY, { steps: 24 })
      await settleWorkbookScrollPerf(page, 16)
      const report = await stopWorkbookScrollPerf(page)

      if (!report) {
        throw new Error('scroll performance report was not available')
      }

      await writeFile(testInfo.outputPath('scroll-perf-wide-250k-fill-preview.json'), JSON.stringify(report, null, 2), 'utf8')

      expect(report.fixture?.id).toBe('wide-mixed-250k')
      expect(report.summary.frameMs.p95).toBeLessThan(20)
      expect(report.summary.longTasksMs.max).toBeLessThan(50)
      expect(report.counters.viewportSubscriptions).toBe(0)
      expectNoRendererMutationChurn(report)
      expectNoTypeGpuDataTileUpload(report)
      expect(readCounter(report.counters, 'typeGpuBufferAllocations')).toBe(0)
      expect(readCounter(report.counters, 'rendererTileMisses')).toBe(0)
    } finally {
      await page.mouse.up().catch(() => undefined)
    }
  })

  test('keeps visible edit commits bounded to dirty V3 tiles', async ({ page }, testInfo) => {
    await page.setViewportSize({ width: 920, height: 680 })
    await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
    await waitForWorkbookReady(page)
    const benchmarkState = await waitForBenchmarkCorpus(page)

    expect(benchmarkState.fixture?.id).toBe('wide-mixed-250k')

    const nameBox = page.getByTestId('name-box')
    const formulaInput = page.getByTestId('formula-input')
    await nameBox.fill('F6')
    await nameBox.press('Enter')
    await expect(page.getByTestId('status-selection')).toContainText('!F6')
    await page.getByTestId('sheet-grid').focus()
    await settleWorkbookScrollPerf(page, 40)
    await warmStartWorkbookScrollPerf(page, 'wide-250k-visible-edit-commit')
    await settleWorkbookScrollPerf(page, 16)
    await formulaInput.fill('7777777')
    await formulaInput.press('Enter')
    await settleWorkbookScrollPerf(page, 24)
    const report = await stopWorkbookScrollPerf(page)

    if (!report) {
      throw new Error('scroll performance report was not available')
    }

    await writeFile(testInfo.outputPath('scroll-perf-wide-250k-visible-edit.json'), JSON.stringify(report, null, 2), 'utf8')

    expect(report.fixture?.id).toBe('wide-mixed-250k')
    expectBoundedVisibleMutation(report, {
      maxDamagePatches: 2,
      maxRendererDeltaBatches: 2,
      maxRendererDeltaMutations: 10,
      maxRendererVisibleDirtyTiles: 10,
      maxRendererWarmDirtyTiles: 4,
      mutationToVisibleP95Max: 50,
    })
    expect(readCounter(report.counters, 'rendererWarmDirtyTiles')).toBeLessThanOrEqual(readCounter(report.counters, 'dirtyTilesMarked'))
    expect(readCounter(report.counters, 'typeGpuBufferAllocations')).toBe(0)
    expectQuietShell(report, { maxSurfaceCommits: 4 })
    await nameBox.fill('F6')
    await nameBox.press('Enter')
    await expect(formulaInput).toHaveValue('7777777')
  })

  remoteSyncTest('keeps shell surfaces quiet and coalesces visible collaborator patch churn while browsing', async ({ page }, testInfo) => {
    const documentId = createTestDocumentId('playwright-zero-scroll-patches')
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
      expectBoundedVisibleMutation(report, {
        maxDamagePatches: 6,
        maxRendererDeltaBatches: 6,
        maxRendererDeltaMutations: 40,
        maxRendererVisibleDirtyTiles: 24,
        maxRendererWarmDirtyTiles: 12,
        minDamagePatches: 1,
      })
      expect(report.counters.canvasPaints['text:body'] ?? 0).toBeLessThanOrEqual(6)
      expect(report.counters.canvasPaints['gpu:body'] ?? 0).toBeLessThanOrEqual(6)
      expectQuietShell(report)
    } finally {
      await mirrorPage.close().catch(() => undefined)
    }
  })
})
