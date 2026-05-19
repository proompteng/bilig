import { writeFile } from 'node:fs/promises'
import { expect, test, type Locator, type Page, type TestInfo } from '@playwright/test'
import { ISOLATED_WORKBOOK_PANE_RENDERER_PATH } from '../../apps/web/src/root-route.js'
import {
  clickProductCell,
  createTestDocumentId,
  dragProductBodySelection,
  dragProductHeaderSelection,
  getProductColumnLeft,
  getProductColumnWidth,
  gotoWorkbookShell,
  pickToolbarPresetColor,
  PRIMARY_MODIFIER,
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  settleWorkbookScrollPerf,
  stopWorkbookScrollPerf,
  warmStartWorkbookScrollPerf,
  waitForBenchmarkCorpus,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

interface ReadbackPoint {
  readonly r: number
  readonly g: number
  readonly b: number
  readonly a: number
}

interface TypeGpuReadbackSummary {
  readonly ready: boolean
  readonly hasGpu: boolean
  readonly width: number
  readonly height: number
  readonly sequence: number
  readonly points: {
    readonly headerFill: ReadbackPoint
    readonly bodyFill: ReadbackPoint
    readonly selectionBorder: ReadbackPoint
    readonly selectionFill: ReadbackPoint
    readonly valueFill: ReadbackPoint
    readonly bodyWhite: ReadbackPoint
  }
  readonly darkPixelCounts: {
    readonly header: number
    readonly body: number
    readonly number: number
  }
}

interface ReadbackInspectorPoint {
  readonly name: string
  readonly x: number
  readonly y: number
}

interface ReadbackInspectorRegion {
  readonly name: string
  readonly x0: number
  readonly y0: number
  readonly x1: number
  readonly y1: number
  readonly threshold?: number
}

interface DynamicReadbackResult {
  readonly ready: boolean
  readonly hasGpu: boolean
  readonly width: number
  readonly height: number
  readonly sequence: number
  readonly points: Record<string, ReadbackPoint>
  readonly darkPixelCounts: Record<string, number>
  readonly opaquePixelCounts: Record<string, number>
}

function selectedRangeFillProbe(columnIndex: number, rowIndex: number): ReadbackInspectorPoint {
  return {
    name: `${columnIndex}:${rowIndex}`,
    x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * columnIndex + Math.floor(PRODUCT_COLUMN_WIDTH / 2),
    y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * rowIndex + Math.floor(PRODUCT_ROW_HEIGHT / 2),
  }
}

function allReadbackPointsMatch(result: DynamicReadbackResult, predicate: (point: ReadbackPoint) => boolean): boolean {
  return Object.values(result.points).every(predicate)
}

function isCornflowerBlueFill(point: ReadbackPoint): boolean {
  return point.a === 255 && point.b > 215 && point.g > 175 && point.r > 135 && point.b > point.g && point.g > point.r
}

function isThemeGreenFill(point: ReadbackPoint): boolean {
  return point.a === 255 && point.g > 135 && point.r < 95 && point.b < 125 && point.g > point.r + 55 && point.g > point.b + 35
}

interface NativeTextExpectation {
  readonly name: string
  readonly text: string
  readonly exact?: boolean | undefined
}

interface NativeTextLayerInspection {
  readonly gridAttrs: Record<string, string>
  readonly matches: Record<string, boolean>
  readonly rendererAttrs: Record<string, string>
  readonly rowHeaderRunCount: number
  readonly sampleTexts: readonly string[]
  readonly visibleRunCount: number
}

interface RendererPresentationSample {
  readonly backendStatus: string | null
  readonly canvasProofLayer: string | null
  readonly editorInputs: number
  readonly editorOverlays: number
  readonly fallbackCanvases: number
  readonly frameProofStatus: string | null
  readonly headerPaneCount: number
  readonly hasPresentedFrame: string | null
  readonly hasPresentedVisibleFrame: string | null
  readonly nativeTextLayerMounted: boolean
  readonly nativeTextRunCount: number
  readonly rowHeaderRunCount: number
  readonly selection: string | null
}

function isResizeGuidePixel(point: ReadbackPoint): boolean {
  return point.a > 150 && point.g > point.r && point.r < 180 && point.b < 180
}

function expectNear(actual: number, expected: number, tolerance: number): void {
  expect(Math.abs(actual - expected)).toBeLessThanOrEqual(tolerance)
}

async function waitForTypeGpuRenderer(page: Page): Promise<void> {
  await page.waitForSelector('[data-testid="grid-pane-renderer"]', { state: 'attached', timeout: 15_000 })
}

async function collectRendererPresentationSamplesDuring(
  page: Page,
  action: () => Promise<void>,
  sampleCount = 24,
): Promise<readonly RendererPresentationSample[]> {
  const samplesPromise = page.evaluate(
    async ({ count, rowMarkerWidth }) => {
      const samples: RendererPresentationSample[] = []
      await new Promise<void>((resolve) => {
        let index = 0
        const sampleNextFrame = () => {
          const canvas = document.querySelector('[data-testid="grid-pane-renderer"]')
          const grid = document.querySelector('[data-testid="sheet-grid"]')
          const gridRect = grid instanceof HTMLElement ? grid.getBoundingClientRect() : null
          const nativeTextLayer = document.querySelector('[data-testid="grid-native-text-layer"]')
          const nativeTextRuns = [...document.querySelectorAll('[data-native-text-run]')]
          const rowHeaderRunCount = nativeTextRuns.filter((run) => {
            if (!gridRect || !/^\d+$/.test(run.textContent ?? '')) {
              return false
            }
            const rect = run.getBoundingClientRect()
            return rect.left >= gridRect.left && rect.right <= gridRect.left + rowMarkerWidth + 1
          }).length
          samples.push({
            backendStatus: canvas?.getAttribute('data-v3-backend-status') ?? null,
            canvasProofLayer: canvas?.getAttribute('data-v3-canvas-proof-layer') ?? null,
            editorInputs: document.querySelectorAll('[data-testid="cell-editor-input"]').length,
            editorOverlays: document.querySelectorAll('[data-testid="cell-editor-overlay"]').length,
            fallbackCanvases: document.querySelectorAll('[data-testid="grid-pane-renderer-fallback"]').length,
            frameProofStatus: canvas?.getAttribute('data-v3-frame-proof-status') ?? null,
            headerPaneCount: Number(canvas?.getAttribute('data-v3-header-pane-count') ?? '0'),
            hasPresentedFrame: canvas?.getAttribute('data-v3-has-presented-frame') ?? null,
            hasPresentedVisibleFrame: canvas?.getAttribute('data-v3-has-presented-visible-frame') ?? null,
            nativeTextLayerMounted: nativeTextLayer instanceof HTMLElement,
            nativeTextRunCount: nativeTextRuns.length,
            rowHeaderRunCount,
            selection: document.querySelector('[data-testid="status-selection"]')?.textContent ?? null,
          })
          index += 1
          if (index >= count) {
            resolve()
            return
          }
          requestAnimationFrame(sampleNextFrame)
        }
        requestAnimationFrame(sampleNextFrame)
      })
      return samples
    },
    { count: sampleCount, rowMarkerWidth: PRODUCT_ROW_MARKER_WIDTH },
  )
  await action()
  return samplesPromise
}

async function exerciseClickAwayEditCommit(
  page: Page,
  input: {
    readonly address: string
    readonly awayCol: number
    readonly awayRow: number
    readonly awaySelection: string
    readonly col: number
    readonly formulaInput: Locator
    readonly renderer: Locator
    readonly row: number
    readonly text: string
  },
): Promise<void> {
  await clickProductCell(page, input.col, input.row)
  await expect(page.getByTestId('status-selection')).toHaveText(`Sheet1!${input.address}`)

  const firstCharacter = input.text.charAt(0)
  await page.keyboard.press(firstCharacter)

  const cellEditor = page.getByTestId('cell-editor-input')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue(firstCharacter)
  await expect(page.getByTestId('cell-editor-overlay')).toHaveCount(1)

  await page.keyboard.type(input.text.slice(1))
  await expect(cellEditor).toHaveValue(input.text)
  await expect
    .poll(async () => await cellEditor.evaluate((editor) => (editor instanceof HTMLTextAreaElement ? editor.selectionStart : -1)))
    .toBe(input.text.length)

  const samples = await collectRendererPresentationSamplesDuring(page, () => clickProductCell(page, input.awayCol, input.awayRow))
  expect(samples.every((sample) => sample.editorInputs <= 1)).toBe(true)
  expect(samples.every((sample) => sample.editorOverlays <= 1)).toBe(true)
  expect(samples.every((sample) => sample.headerPaneCount > 0)).toBe(true)
  expect(samples.every((sample) => sample.nativeTextLayerMounted)).toBe(true)
  expect(samples.every((sample) => sample.nativeTextRunCount > 0)).toBe(true)
  expect(samples.every((sample) => sample.rowHeaderRunCount >= 10)).toBe(true)
  expect(samples.filter((sample) => sample.fallbackCanvases !== 0)).toEqual([])
  expect(samples.filter((sample) => sample.canvasProofLayer === 'mounted')).toEqual([])

  await expect(page.getByTestId('status-selection')).toHaveText(input.awaySelection)
  await expect(cellEditor).toHaveCount(0)
  await expect(page.getByTestId('cell-editor-overlay')).toHaveCount(0)
  await expect(page.getByTestId('grid-pane-renderer-fallback')).toHaveCount(0)

  await clickProductCell(page, input.col, input.row)
  await expect(input.formulaInput).toHaveValue(input.text)
  await expect.poll(async () => await input.renderer.getAttribute('data-v3-frame-proof-status')).toBe('presented')
  await expect(input.renderer).toHaveAttribute('data-v3-canvas-proof-layer', 'not-mounted')
}

async function dragProductSelectedContentLane(page: Page, startColumn: number, startRow: number, targetColumn: number, targetRow: number) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startLeft = await getProductColumnLeft(page, startColumn)
  const startWidth = await getProductColumnWidth(page, startColumn)
  const sourceX = grid.x + startLeft + Math.min(32, Math.floor(startWidth * 0.35))
  const sourceY = grid.y + PRODUCT_HEADER_HEIGHT + startRow * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  const targetLeft = await getProductColumnLeft(page, targetColumn)
  const targetWidth = await getProductColumnWidth(page, targetColumn)
  const targetX = grid.x + targetLeft + Math.floor(targetWidth / 2)
  const targetY = grid.y + PRODUCT_HEADER_HEIGHT + targetRow * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)

  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 12 })
  await page.mouse.up()
}

test('@browser-webgpu isolated workbook pane renderer draws grid content through typegpu', async ({ page }, testInfo) => {
  await page.setViewportSize({ width: 640, height: 480 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, ISOLATED_WORKBOOK_PANE_RENDERER_PATH)
  await page.waitForSelector('[data-testid="isolated-pane-renderer-route"]', { timeout: 15_000 })
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () => Boolean((window as Window & { __biligGpuReadback?: { readonly ready: boolean } }).__biligGpuReadback?.ready),
    undefined,
    { timeout: 15_000 },
  )

  const summary = await page.evaluate(() => {
    return (window as Window & { __biligGpuReadback?: TypeGpuReadbackSummary }).__biligGpuReadback ?? null
  })

  expect(summary).not.toBeNull()
  expect(summary?.hasGpu).toBe(true)
  expect(summary?.width).toBe(640)
  expect(summary?.height).toBe(480)
  expect(summary?.sequence).toBeGreaterThan(0)
  expect(summary?.points.headerFill).toMatchObject({ r: 243, g: 242, b: 238, a: 255 })
  expect(summary?.points.bodyFill).toMatchObject({ r: 255, g: 255, b: 255, a: 255 })
  expect(summary?.points.selectionBorder.a ?? 0).toBeGreaterThan(150)
  expect(summary?.points.selectionBorder.g ?? 0).toBeGreaterThan(summary?.points.selectionBorder.r ?? 0)
  expect(summary?.points.bodyWhite).toMatchObject({ r: 255, g: 255, b: 255, a: 255 })
  const textRuns = await waitForVisibleNativeTextRuns(
    page,
    [
      { name: 'header', text: 'Region', exact: true },
      { name: 'body', text: 'North', exact: true },
      { name: 'number', text: '168', exact: true },
    ],
    (runs) => runs.matches.header && runs.matches.body && runs.matches.number,
  )
  expect(textRuns.matches.header).toBe(true)
  expect(textRuns.matches.body).toBe(true)
  expect(textRuns.matches.number).toBe(true)

  await saveReadbackArtifact(page, testInfo, 'isolated-pane-renderer-readback.png', 'isolated-pane-renderer-readback')
})

test('main workbook shell mounts typegpu-v3 as the only grid renderer', async ({ page }) => {
  await page.setViewportSize({ width: 960, height: 720 })
  await gotoWorkbookShell(page)
  await waitForWorkbookReady(page)

  await expect(page.getByTestId('grid-pane-renderer')).toHaveAttribute('data-renderer-mode', 'typegpu-v3')
  await expect(page.getByTestId('grid-pane-renderer')).toHaveAttribute('data-pane-renderer', 'workbook-pane-renderer-v3')
  await expect(page.locator('[data-pane-renderer="workbook-pane-renderer"]')).toHaveCount(0)
  await expect(page.getByTestId('grid-pane-text-overlay')).toHaveCount(0)
})

test('@browser-webgpu @browser-serial main workbook shell grid renders and updates through typegpu', async ({ page }, testInfo) => {
  const rangeFillPoint = {
    x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 2 + 24,
    y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 2 + Math.floor(PRODUCT_ROW_HEIGHT / 2),
  }
  const rangeBorderPoint = {
    x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH + 50,
    y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT,
  }
  const topHeaderSelectionFillPoint = {
    x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH + 20,
    y: Math.floor(PRODUCT_HEADER_HEIGHT / 2),
  }

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-grid-updates'))}`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )
  await waitForReadbackSequence(page, 0)

  const initialProbe = {
    points: [
      { name: 'unselectedHeaderFill', x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH + 20, y: 12 },
      {
        name: 'bodyBlank',
        x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 4 + 14,
        y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 4 + Math.floor(PRODUCT_ROW_HEIGHT / 2),
      },
    ],
    regions: [
      { name: 'columnHeaderText', x0: 176, y0: 4, x1: 228, y1: 18 },
      { name: 'rowHeaderText', x0: 10, y0: 48, x1: 36, y1: 66 },
    ],
  } as const

  const initialReadback = await waitForReadback(page, initialProbe, (result) => {
    return result.opaquePixelCounts.columnHeaderText > 100 && result.opaquePixelCounts.rowHeaderText > 100
  })
  const initialTextRuns = await waitForVisibleNativeTextRuns(page, [{ name: 'columnHeaderA', text: 'A', exact: true }], (runs) => {
    return runs.matches.columnHeaderA && runs.rowHeaderRunCount > 5
  })

  expect(initialReadback.hasGpu).toBe(true)
  expect(initialReadback.width).toBeGreaterThan(400)
  expect(initialReadback.height).toBeGreaterThan(250)
  expect(initialReadback.points.bodyBlank).toMatchObject({ r: 0, g: 0, b: 0, a: 0 })
  expect(initialReadback.opaquePixelCounts.columnHeaderText).toBeGreaterThan(100)
  expect(initialReadback.opaquePixelCounts.rowHeaderText).toBeGreaterThan(100)
  expect(initialTextRuns.matches.columnHeaderA).toBe(true)
  expect(initialTextRuns.rowHeaderRunCount).toBeGreaterThan(5)

  await clickProductCell(page, 2, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C4')

  await page.getByTestId('formula-input').fill('123')
  await page.getByTestId('formula-input').press('Enter')
  await expect(page.getByTestId('formula-input')).toHaveValue('123')
  await expect(page.getByTestId('formula-resolved-value')).toHaveText('123')
  await waitForReadbackSequence(page, initialReadback.sequence)

  const valueProbe = {
    points: [],
    regions: [
      {
        name: 'c4ValueText',
        x0: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 2 + 8,
        y0: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 3 + 4,
        x1: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 3 - 8,
        y1: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 4 - 4,
      },
    ],
  } as const

  const valueReadback = await inspectGpuReadback(page, valueProbe)
  const valueTextRuns = await waitForVisibleNativeTextRuns(page, [{ name: 'c4ValueText', text: '123', exact: true }], (runs) => {
    return runs.matches.c4ValueText
  })
  expect(valueTextRuns.matches.c4ValueText).toBe(true)

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await clickProductCell(page, 2, 2, { shift: true })
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')
  await waitForReadbackSequence(page, valueReadback.sequence)

  const rangeProbe = {
    points: [
      { name: 'rangeFill', x: rangeFillPoint.x, y: rangeFillPoint.y },
      { name: 'rangeBorder', x: rangeBorderPoint.x, y: rangeBorderPoint.y },
      { name: 'topHeaderSelectionFill', x: topHeaderSelectionFillPoint.x, y: topHeaderSelectionFillPoint.y },
    ],
    regions: [
      {
        name: 'fillHandleRegion',
        x0: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 3 - 12,
        y0: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 3 - 12,
        x1: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 3 + 12,
        y1: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 3 + 12,
      },
    ],
  } as const

  const rangeReadback = await waitForReadback(page, rangeProbe, (result) => {
    return result.points.rangeBorder.a > 150 && result.darkPixelCounts.fillHandleRegion > 4 && result.points.topHeaderSelectionFill.a > 0
  })

  expect(rangeReadback.points.rangeFill.a).toBeLessThanOrEqual(25)
  expect(rangeReadback.points.rangeBorder.a).toBeGreaterThan(150)
  expect(rangeReadback.darkPixelCounts.fillHandleRegion).toBeGreaterThan(4)
  expect(rangeReadback.points.topHeaderSelectionFill.a).toBeGreaterThan(0)

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-readback.png', 'main-workbook-grid-readback')
})

test('@browser-webgpu @browser-serial main workbook shell keeps row headers visible after click-away edit', async ({ page }, testInfo) => {
  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-click-away-edit-row-headers'))}`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )
  await waitForReadbackSequence(page, 0)

  const rowHeaderProbe = {
    points: [],
    regions: [
      {
        name: 'rowHeaderText',
        x0: 0,
        y0: PRODUCT_HEADER_HEIGHT,
        x1: PRODUCT_ROW_MARKER_WIDTH,
        y1: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 30,
      },
    ],
  } as const
  const initialReadback = await waitForReadback(page, rowHeaderProbe, (result) => result.opaquePixelCounts.rowHeaderText > 1_000)
  const initialRowHeaderRuns = await waitForVisibleNativeTextRuns(page, [], (runs) => runs.rowHeaderRunCount > 20)
  expect(initialRowHeaderRuns.rowHeaderRunCount).toBeGreaterThan(20)

  await clickProductCell(page, 1, 24)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B25')
  await page.keyboard.press('a')

  const cellEditor = page.getByTestId('cell-editor-input')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('a')
  await expect
    .poll(async () => await cellEditor.evaluate((input) => (input instanceof HTMLTextAreaElement ? input.selectionStart : -1)))
    .toBe(1)
  const pressRemainingText = async (remainingCharacters: readonly string[], previousText: string): Promise<void> => {
    const [character, ...rest] = remainingCharacters
    if (!character) {
      return
    }
    const nextText = `${previousText}${character}`
    await cellEditor.press(character)
    await expect(cellEditor).toHaveValue(nextText)
    await expect
      .poll(async () => await cellEditor.evaluate((input) => (input instanceof HTMLTextAreaElement ? input.selectionStart : -1)))
      .toBe(nextText.length)
    await pressRemainingText(rest, nextText)
  }
  await pressRemainingText(['b', 'c', 'd', 'e', 'f'], 'a')

  await clickProductCell(page, 3, 25)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D26')
  await expect(cellEditor).toHaveCount(0)
  await waitForReadbackSequence(page, initialReadback.sequence)

  await clickProductCell(page, 1, 24)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B25')
  await expect(page.getByLabel('Formula')).toHaveValue('abcdef')
  await clickProductCell(page, 3, 25)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D26')

  const committedProbe = {
    points: [],
    regions: [
      {
        name: 'rowHeaderText',
        x0: 0,
        y0: PRODUCT_HEADER_HEIGHT,
        x1: PRODUCT_ROW_MARKER_WIDTH,
        y1: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 30,
      },
      {
        name: 'b25Text',
        x0: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH + 8,
        y0: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 24 + 4,
        x1: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 2 - 8,
        y1: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 25 - 4,
      },
    ],
  } as const
  const committedReadback = await waitForReadback(page, committedProbe, (result) => result.opaquePixelCounts.rowHeaderText > 1_000)
  const committedTextRuns = await waitForVisibleNativeTextRuns(page, [{ name: 'b25Text', text: 'abcdef', exact: true }], (runs) => {
    return runs.rowHeaderRunCount > 20 && runs.matches.b25Text
  })

  expect(committedReadback.opaquePixelCounts.rowHeaderText).toBeGreaterThan(1_000)
  expect(committedTextRuns.rowHeaderRunCount).toBeGreaterThan(20)
  expect(committedTextRuns.matches.b25Text).toBe(true)

  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-click-away-edit-row-headers-readback.png',
    'main-workbook-click-away-edit-row-headers-readback',
  )
})

test('@browser-webgpu @browser-serial main workbook shell keeps the live typegpu layer stable during click-away edit commits', async ({
  page,
}) => {
  await page.setViewportSize({ width: 960, height: 720 })
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-click-away-edit-no-proof-blink'))}`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)

  const renderer = page.getByTestId('grid-pane-renderer')
  await expect.poll(async () => await renderer.getAttribute('data-v3-frame-proof-status')).toBe('presented')
  await expect(renderer).toHaveAttribute('data-v3-canvas-proof-layer', 'not-mounted')
  await expect(page.getByTestId('grid-pane-renderer-fallback')).toHaveCount(0)

  const formulaInput = page.getByTestId('formula-input')
  await exerciseClickAwayEditCommit(page, {
    address: 'B1',
    awayCol: 2,
    awayRow: 0,
    awaySelection: 'Sheet1!C1',
    col: 1,
    formulaInput,
    renderer,
    row: 0,
    text: 'abc',
  })
  await exerciseClickAwayEditCommit(page, {
    address: 'B2',
    awayCol: 2,
    awayRow: 1,
    awaySelection: 'Sheet1!C2',
    col: 1,
    formulaInput,
    renderer,
    row: 1,
    text: 'def',
  })
})

test('@browser-webgpu @browser-perf main workbook shell keeps resident typegpu content visible while selection moves', async ({
  page,
}, testInfo) => {
  test.slow()
  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(
    page,
    `/?document=${encodeURIComponent(createTestDocumentId('typegpu-selection-no-flash'))}&benchmarkCorpus=wide-mixed-250k`,
  )
  await waitForWorkbookReady(page)
  await waitForBenchmarkCorpus(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const probe = {
    points: [],
    regions: [
      { name: 'columnHeaderText', x0: 60, y0: 4, x1: 220, y1: 18 },
      { name: 'rowHeaderText', x0: 6, y0: 48, x1: 36, y1: 140 },
      { name: 'bodyGrid', x0: 70, y0: 48, x1: 360, y1: 150 },
    ],
  } as const

  const initialReadback = await waitForReadback(page, probe, (result) => {
    return (
      result.opaquePixelCounts.columnHeaderText > 100 &&
      result.opaquePixelCounts.rowHeaderText > 100 &&
      result.opaquePixelCounts.bodyGrid > 400
    )
  })
  await waitForVisibleNativeTextRuns(page, [{ name: 'columnHeaderA', text: 'A', exact: true }], (runs) => {
    return runs.matches.columnHeaderA && runs.rowHeaderRunCount > 5
  })
  await warmStartWorkbookScrollPerf(page, 'typegpu-selection-overlay-only')
  await settleWorkbookScrollPerf(page, 16)

  const selectionTargets = [
    { col: 2, row: 3, address: '!C4' },
    { col: 5, row: 6, address: '!F7' },
    { col: 1, row: 8, address: '!B9' },
    { col: 3, row: 4, address: '!D5' },
  ] as const
  const sampleSelectionTarget = async (
    targetIndex: number,
    lastSequence: number,
    frames: ReadonlyArray<DynamicReadbackResult>,
  ): Promise<ReadonlyArray<DynamicReadbackResult>> => {
    const target = selectionTargets[targetIndex]
    if (!target) {
      return frames
    }
    await clickProductCell(page, target.col, target.row)
    await expect(page.getByTestId('status-selection')).toContainText(target.address)
    await page.waitForTimeout(50)
    const frame = await inspectGpuReadback(page, probe)
    return await sampleSelectionTarget(targetIndex + 1, Math.max(lastSequence, frame.sequence), [...frames, frame])
  }
  const selectionFrames = await sampleSelectionTarget(0, initialReadback.sequence, [])
  await settleWorkbookScrollPerf(page, 4)
  const perfReport = await stopWorkbookScrollPerf(page)

  expect(selectionFrames.length).toBe(4)
  for (const frame of selectionFrames) {
    expect(frame.opaquePixelCounts.columnHeaderText).toBeGreaterThan(100)
    expect(frame.opaquePixelCounts.rowHeaderText).toBeGreaterThan(100)
    expect(frame.opaquePixelCounts.bodyGrid).toBeGreaterThan(400)
  }
  const finalSelectionTextRuns = await waitForVisibleNativeTextRuns(page, [{ name: 'columnHeaderA', text: 'A', exact: true }], (runs) => {
    return runs.matches.columnHeaderA && runs.rowHeaderRunCount > 5
  })
  expect(finalSelectionTextRuns.matches.columnHeaderA).toBe(true)
  expect(finalSelectionTextRuns.rowHeaderRunCount).toBeGreaterThan(5)
  expect(perfReport).not.toBeNull()
  expect(perfReport?.counters.headerPaneBuilds).toBeLessThanOrEqual(1)
  expect(perfReport?.counters.typeGpuBufferAllocations).toBe(0)
  expect(perfReport?.counters.typeGpuTileMisses).toBe(0)
  expect(perfReport?.counters.rendererTileMisses).toBe(0)

  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-selection-no-flash-readback.png',
    'main-workbook-grid-selection-no-flash-readback',
  )
})

test('@browser-webgpu @browser-perf main workbook shell keeps header labels and body text visible while scrolling through typegpu', async ({
  page,
}, testInfo) => {
  test.slow()
  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
  await waitForWorkbookReady(page)
  await waitForBenchmarkCorpus(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const initialReadback = await waitForReadback(
    page,
    {
      points: [],
      regions: [
        { name: 'columnHeaderText', x0: 60, y0: 4, x1: 220, y1: 18 },
        { name: 'rowHeaderText', x0: 6, y0: 48, x1: 36, y1: 140 },
        { name: 'bodyText', x0: 70, y0: 48, x1: 280, y1: 120 },
      ],
    },
    (result) => result.opaquePixelCounts.columnHeaderText > 100 && result.opaquePixelCounts.rowHeaderText > 100,
  )
  await waitForVisibleNativeTextRuns(page, [{ name: 'columnHeaderA', text: 'A', exact: true }], (runs) => {
    return runs.matches.columnHeaderA && runs.rowHeaderRunCount > 5 && runs.visibleRunCount > 20
  })

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    if (!(viewport instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not a div')
    }
    viewport.scrollLeft = 1_768
    viewport.scrollTop = 418
    viewport.dispatchEvent(new Event('scroll'))
  })
  await waitForReadbackSequence(page, initialReadback.sequence)

  const scrolledReadback = await waitForReadback(
    page,
    {
      points: [],
      regions: [
        { name: 'columnHeaderText', x0: 60, y0: 4, x1: 220, y1: 18 },
        { name: 'rowHeaderText', x0: 6, y0: 48, x1: 36, y1: 140 },
        { name: 'bodyText', x0: 70, y0: 48, x1: 280, y1: 120 },
      ],
    },
    (result) => result.opaquePixelCounts.columnHeaderText > 100 && result.opaquePixelCounts.rowHeaderText > 100,
  )
  const scrolledTextRuns = await waitForVisibleNativeTextRuns(page, [{ name: 'columnHeader', text: 'E', exact: true }], (runs) => {
    return runs.rowHeaderRunCount > 5 && runs.visibleRunCount > 20
  })

  expect(scrolledReadback.sequence).toBeGreaterThan(initialReadback.sequence)
  expect(scrolledReadback.opaquePixelCounts.columnHeaderText).toBeGreaterThan(100)
  expect(scrolledReadback.opaquePixelCounts.rowHeaderText).toBeGreaterThan(100)
  expect(scrolledTextRuns.rowHeaderRunCount).toBeGreaterThan(5)
  expect(scrolledTextRuns.visibleRunCount).toBeGreaterThan(20)

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-scrolled-readback.png', 'main-workbook-grid-scrolled-readback')
})

test('@browser-webgpu @browser-deep main workbook shell keeps typegpu grid lines exactly aligned after diagonal scroll', async ({
  page,
}, testInfo) => {
  const scrollLeft = PRODUCT_COLUMN_WIDTH * 4 + 17
  const scrollTop = PRODUCT_ROW_HEIGHT * 5 + 9
  const visibleStartCol = Math.floor(scrollLeft / PRODUCT_COLUMN_WIDTH)
  const visibleStartRow = Math.floor(scrollTop / PRODUCT_ROW_HEIGHT)
  const verticalLineAfterCol = visibleStartCol + 3
  const horizontalLineAfterRow = visibleStartRow + 7
  const verticalLineX = PRODUCT_ROW_MARKER_WIDTH + (verticalLineAfterCol + 1) * PRODUCT_COLUMN_WIDTH - scrollLeft - 1
  const horizontalLineY = PRODUCT_HEADER_HEIGHT + (horizontalLineAfterRow + 1) * PRODUCT_ROW_HEIGHT - scrollTop - 1
  const bodyProbeY = PRODUCT_HEADER_HEIGHT + 180
  const bodyProbeX = PRODUCT_ROW_MARKER_WIDTH + 360

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const initialSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })

  await page.getByTestId('grid-scroll-viewport').evaluate(
    (viewport, target) => {
      if (!(viewport instanceof HTMLDivElement)) {
        throw new Error('grid scroll viewport is not a div')
      }
      viewport.scrollLeft = target.left
      viewport.scrollTop = target.top
      viewport.dispatchEvent(new Event('scroll'))
    },
    { left: scrollLeft, top: scrollTop },
  )
  await waitForReadbackSequence(page, initialSequence)

  const readback = await waitForReadback(
    page,
    {
      points: [
        { name: 'headerVerticalLine', x: verticalLineX, y: Math.floor(PRODUCT_HEADER_HEIGHT / 2) },
        { name: 'bodyVerticalLine', x: verticalLineX, y: bodyProbeY },
        { name: 'bodyVerticalBlank', x: verticalLineX - 3, y: bodyProbeY },
        { name: 'rowHeaderHorizontalLine', x: Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2), y: horizontalLineY },
        { name: 'bodyHorizontalLine', x: bodyProbeX, y: horizontalLineY },
        { name: 'bodyHorizontalBlank', x: bodyProbeX, y: horizontalLineY - 3 },
      ],
      regions: [],
    },
    (result) =>
      result.points.headerVerticalLine.a > 150 &&
      result.points.bodyVerticalLine.a > 150 &&
      result.points.rowHeaderHorizontalLine.a > 150 &&
      result.points.bodyHorizontalLine.a > 150,
  )

  expect(readback.points.headerVerticalLine.a).toBeGreaterThan(150)
  expect(readback.points.bodyVerticalLine.a).toBeGreaterThan(150)
  expect(readback.points.bodyVerticalBlank.a).toBeLessThan(50)
  expect(readback.points.rowHeaderHorizontalLine.a).toBeGreaterThan(150)
  expect(readback.points.bodyHorizontalLine.a).toBeGreaterThan(150)
  expect(readback.points.bodyHorizontalBlank.a).toBeLessThan(50)
  expect(verticalLineX).toBeGreaterThan(PRODUCT_ROW_MARKER_WIDTH)
  expect(horizontalLineY).toBeGreaterThan(PRODUCT_HEADER_HEIGHT)

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-exact-scroll-readback.png', 'main-workbook-grid-exact-scroll-readback')
})

test('@browser-webgpu @browser-deep main workbook shell draws typegpu resize guides at exact geometry positions', async ({
  page,
}, testInfo) => {
  const columnGuideX = PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH - 1
  const rowGuideY = PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT - 1

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const grid = await page.getByTestId('sheet-grid').boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const initialSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })

  await page.mouse.move(grid.x + columnGuideX, grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2))
  await waitForReadbackSequence(page, initialSequence)
  const columnReadback = await waitForReadback(
    page,
    {
      points: [
        { name: 'columnGuideHeader', x: columnGuideX, y: Math.floor(PRODUCT_HEADER_HEIGHT / 2) },
        { name: 'columnGuideBody', x: columnGuideX, y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 4 },
        { name: 'columnGuideAdjacent', x: columnGuideX - 3, y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 4 },
      ],
      regions: [],
    },
    (result) => isResizeGuidePixel(result.points.columnGuideHeader) && isResizeGuidePixel(result.points.columnGuideBody),
  )

  expect(isResizeGuidePixel(columnReadback.points.columnGuideHeader)).toBe(true)
  expect(isResizeGuidePixel(columnReadback.points.columnGuideBody)).toBe(true)
  expect(columnReadback.points.columnGuideAdjacent.a).toBeLessThan(80)

  await page.mouse.move(grid.x + Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2), grid.y + rowGuideY)
  await page.mouse.down()
  await waitForReadbackSequence(page, columnReadback.sequence)
  const rowReadback = await waitForReadback(
    page,
    {
      points: [
        { name: 'rowGuideHeader', x: Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2), y: rowGuideY },
        { name: 'rowGuideBody', x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 3, y: rowGuideY },
        { name: 'rowGuideAdjacent', x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 3, y: rowGuideY - 3 },
      ],
      regions: [],
    },
    (result) => isResizeGuidePixel(result.points.rowGuideHeader) && isResizeGuidePixel(result.points.rowGuideBody),
  )

  expect(isResizeGuidePixel(rowReadback.points.rowGuideHeader)).toBe(true)
  expect(isResizeGuidePixel(rowReadback.points.rowGuideBody)).toBe(true)
  expect(rowReadback.points.rowGuideAdjacent.a).toBeLessThan(80)
  await page.mouse.up()

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-resize-guide-readback.png', 'main-workbook-grid-resize-guide-readback')
})

test('@browser-webgpu @browser-deep main workbook shell keeps DOM editor overlay aligned to typegpu geometry while scrolling', async ({
  page,
}, testInfo) => {
  const targetCol = 3
  const targetRow = 7
  const scrollLeft = 37
  const scrollTop = 13
  const expectedLocalX = PRODUCT_ROW_MARKER_WIDTH + targetCol * PRODUCT_COLUMN_WIDTH - scrollLeft
  const expectedLocalY = PRODUCT_HEADER_HEIGHT + targetRow * PRODUCT_ROW_HEIGHT - scrollTop

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const grid = await page.getByTestId('sheet-grid').boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  await clickProductCell(page, targetCol, targetRow)
  await page.keyboard.press('F2')
  await expect(page.getByTestId('cell-editor-input')).toBeVisible()

  const initialSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })

  await page.getByTestId('grid-scroll-viewport').evaluate(
    (viewport, target) => {
      if (!(viewport instanceof HTMLDivElement)) {
        throw new Error('grid scroll viewport is not a div')
      }
      viewport.scrollLeft = target.left
      viewport.scrollTop = target.top
      viewport.dispatchEvent(new Event('scroll'))
    },
    { left: scrollLeft, top: scrollTop },
  )
  await waitForReadbackSequence(page, initialSequence)

  const dpr = await page.evaluate(() => window.devicePixelRatio || 1)
  const tolerance = Math.max(1, 1 / dpr)
  const expectedViewportRect = {
    x: grid.x + expectedLocalX,
    y: grid.y + expectedLocalY,
    width: PRODUCT_COLUMN_WIDTH,
    height: PRODUCT_ROW_HEIGHT,
  }

  await expect
    .poll(
      async () => {
        const box = await page.getByTestId('cell-editor-overlay').boundingBox()
        if (!box) {
          return Number.POSITIVE_INFINITY
        }
        return Math.max(
          Math.abs(box.x - expectedViewportRect.x),
          Math.abs(box.y - expectedViewportRect.y),
          Math.abs(box.width - expectedViewportRect.width),
          Math.abs(box.height - expectedViewportRect.height),
        )
      },
      { timeout: 15_000 },
    )
    .toBeLessThanOrEqual(tolerance)

  const editorBox = await page.getByTestId('cell-editor-overlay').boundingBox()
  if (!editorBox) {
    throw new Error('cell editor overlay is not visible')
  }
  expectNear(editorBox.x, expectedViewportRect.x, tolerance)
  expectNear(editorBox.y, expectedViewportRect.y, tolerance)
  expectNear(editorBox.width, expectedViewportRect.width, tolerance)
  expectNear(editorBox.height, expectedViewportRect.height, tolerance)

  const readback = await waitForReadback(
    page,
    {
      points: [
        { name: 'activeCellTopBorder', x: expectedLocalX + Math.floor(PRODUCT_COLUMN_WIDTH / 2), y: expectedLocalY },
        { name: 'activeCellLeftBorder', x: expectedLocalX, y: expectedLocalY + Math.floor(PRODUCT_ROW_HEIGHT / 2) },
      ],
      regions: [],
    },
    (result) => result.points.activeCellTopBorder.a > 150 && result.points.activeCellLeftBorder.a > 150,
  )

  expect(readback.points.activeCellTopBorder.a).toBeGreaterThan(150)
  expect(readback.points.activeCellLeftBorder.a).toBeGreaterThan(150)

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-editor-overlay-readback.png', 'main-workbook-grid-editor-overlay-readback')
})

test('@browser-webgpu @browser-deep main workbook shell refreshes typegpu resident packets after style-only changes', async ({
  page,
}, testInfo) => {
  const fillPoint = {
    x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH + Math.floor(PRODUCT_COLUMN_WIDTH / 2),
    y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2),
  }

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-style-refresh'))}`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const initialReadback = await waitForReadback(
    page,
    {
      points: [{ name: 'cellFill', ...fillPoint }],
      regions: [],
    },
    (result) => result.points.cellFill.a === 0,
  )

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toContainText('!B2')
  await pickToolbarPresetColor(page, 'Fill color', 'light cornflower blue 3')
  await waitForReadbackSequence(page, initialReadback.sequence)
  const afterStyleSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })
  await clickProductCell(page, 2, 2)
  await waitForReadbackSequence(page, afterStyleSequence)

  const styledReadback = await waitForReadback(
    page,
    {
      points: [{ name: 'cellFill', ...fillPoint }],
      regions: [],
    },
    (result) => result.points.cellFill.a > 200,
  )

  expect(styledReadback.points.cellFill.a).toBe(255)
  expect(styledReadback.points.cellFill.r).toBeGreaterThan(170)
  expect(styledReadback.points.cellFill.r).toBeLessThan(215)
  expect(styledReadback.points.cellFill.g).toBeGreaterThan(190)
  expect(styledReadback.points.cellFill.g).toBeLessThan(230)
  expect(styledReadback.points.cellFill.b).toBeGreaterThan(220)
  expect(styledReadback.points.cellFill.b).toBeLessThan(255)
  expect(styledReadback.points.cellFill.b).toBeGreaterThan(styledReadback.points.cellFill.g)
  expect(styledReadback.points.cellFill.g).toBeGreaterThan(styledReadback.points.cellFill.r)

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-style-refresh-readback.png', 'main-workbook-grid-style-refresh-readback')
})

test('@browser-webgpu @browser-deep selected range fill changes stay visually authoritative while selected', async ({ page }, testInfo) => {
  const points = [
    { ...selectedRangeFillProbe(1, 1), name: 'topLeft' },
    { ...selectedRangeFillProbe(2, 4), name: 'middle' },
    { ...selectedRangeFillProbe(3, 8), name: 'bottomRight' },
  ]

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-selected-fill-refresh'))}`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  await dragProductBodySelection(page, 1, 1, 3, 8)
  await expect(page.getByTestId('status-selection')).toContainText('!B2:D9')
  await pickToolbarPresetColor(page, 'Fill color', 'light cornflower blue 2')
  const blueReadback = await waitForReadback(
    page,
    {
      points,
      regions: [],
    },
    (result) => allReadbackPointsMatch(result, isCornflowerBlueFill),
  )

  await expect(page.getByTestId('status-selection')).toContainText('!B2:D9')
  await pickToolbarPresetColor(page, 'Fill color', 'theme green')
  const greenReadback = await waitForReadback(
    page,
    {
      points,
      regions: [],
    },
    (result) => result.sequence > blueReadback.sequence && allReadbackPointsMatch(result, isThemeGreenFill),
  )

  for (const [name, point] of Object.entries(greenReadback.points)) {
    expect(point.g, `${name} green channel`).toBeGreaterThan(point.b)
    expect(point.g, `${name} green channel`).toBeGreaterThan(point.r)
  }

  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-selected-fill-refresh-readback.png',
    'main-workbook-grid-selected-fill-refresh-readback',
  )
})

test('@browser-webgpu @browser-deep moved content delete preserves selected fill without fallback flashes', async ({ page }, testInfo) => {
  const movedText = 'move-fill-delete-proof'
  const points = [
    { ...selectedRangeFillProbe(1, 1), name: 'topLeft' },
    { ...selectedRangeFillProbe(2, 4), name: 'middle' },
    { ...selectedRangeFillProbe(3, 8), name: 'bottomRight' },
  ]

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-move-fill-delete-no-flash'))}`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const formulaInput = page.getByTestId('formula-input')
  const grid = page.getByTestId('sheet-grid')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill(movedText)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(movedText)
  await dragProductSelectedContentLane(page, 1, 1, 3, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4')

  await clickProductCell(page, 1, 1)
  await expect(formulaInput).toHaveValue('')
  await clickProductCell(page, 3, 3)
  await expect(formulaInput).toHaveValue(movedText)

  await dragProductBodySelection(page, 1, 1, 3, 8)
  await expect(page.getByTestId('status-selection')).toContainText('!B2:D9')
  await pickToolbarPresetColor(page, 'Fill color', 'theme green')
  const greenReadback = await waitForReadback(
    page,
    {
      points,
      regions: [],
    },
    (result) => allReadbackPointsMatch(result, isThemeGreenFill),
  )

  const deleteSamples = await collectRendererPresentationSamplesDuring(page, () => grid.press('Delete'), 30)
  expect(deleteSamples.filter((sample) => sample.fallbackCanvases !== 0)).toEqual([])
  expect(deleteSamples.filter((sample) => sample.canvasProofLayer === 'mounted')).toEqual([])

  const afterDeleteReadback = await waitForReadback(
    page,
    {
      points,
      regions: [],
    },
    (result) => result.sequence >= greenReadback.sequence && allReadbackPointsMatch(result, isThemeGreenFill),
  )
  expect(afterDeleteReadback.sequence).toBeGreaterThanOrEqual(greenReadback.sequence)

  await clickProductCell(page, 3, 3)
  await expect(formulaInput).toHaveValue('')
  await expect
    .poll(async () =>
      page.evaluate((text) => [...document.querySelectorAll('[data-native-text-run]')].some((run) => run.textContent === text), movedText),
    )
    .toBe(false)

  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-move-fill-delete-no-flash-readback.png',
    'main-workbook-grid-move-fill-delete-no-flash-readback',
  )
})

test('@browser-webgpu @browser-deep name-box range fill presents the current frame before success', async ({ page }, testInfo) => {
  const points = [
    { ...selectedRangeFillProbe(3, 3), name: 'topLeft' },
    { ...selectedRangeFillProbe(4, 5), name: 'middle' },
    { ...selectedRangeFillProbe(5, 7), name: 'bottomRight' },
  ]

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-name-box-fill-refresh'))}&persist=0`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const nameBox = page.getByTestId('name-box')
  await nameBox.fill('D4:F8')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4:F8')

  await pickToolbarPresetColor(page, 'Fill color', 'theme green')
  const greenReadback = await waitForReadback(
    page,
    {
      points,
      regions: [],
    },
    (result) => allReadbackPointsMatch(result, isThemeGreenFill),
  )
  const rendererProof = await page.getByTestId('grid-pane-renderer').evaluate((canvas) => ({
    frameProofStatus: canvas.getAttribute('data-v3-frame-proof-status'),
    projectedRenderRevision: canvas.getAttribute('data-v3-projected-render-revision'),
    tileSceneRevision: canvas.getAttribute('data-v3-tile-scene-revision'),
    visibleProjectedRenderRevision: canvas.getAttribute('data-v3-visible-projected-render-revision'),
    visibleRenderRevision: canvas.getAttribute('data-v3-visible-render-revision'),
  }))
  const gridProjectedRevision = await page.getByTestId('sheet-grid').getAttribute('data-render-projected-revision')

  expect(rendererProof.frameProofStatus).toBe('presented')
  expect(rendererProof.projectedRenderRevision).toBe(gridProjectedRevision)
  expect(rendererProof.visibleProjectedRenderRevision).toBe(gridProjectedRevision)
  expect(rendererProof.visibleRenderRevision).toBe(rendererProof.tileSceneRevision)
  expect(greenReadback.sequence).toBeGreaterThan(0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4:F8')
  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-name-box-fill-refresh-readback.png',
    'main-workbook-grid-name-box-fill-refresh-readback',
  )
})

test('@browser-webgpu @browser-deep large range fill remains applied after scrolling into uncached rows', async ({ page }, testInfo) => {
  const initialPoints = [
    { ...selectedRangeFillProbe(3, 3), name: 'topLeft' },
    { ...selectedRangeFillProbe(4, 8), name: 'middle' },
    { ...selectedRangeFillProbe(5, 12), name: 'bottomRight' },
  ]
  const scrolledPoints = [
    { ...selectedRangeFillProbe(3, 5), name: 'scrolledD' },
    { ...selectedRangeFillProbe(4, 8), name: 'scrolledE' },
    { ...selectedRangeFillProbe(5, 11), name: 'scrolledF' },
  ]

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-large-range-fill-scroll'))}&persist=0`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const nameBox = page.getByTestId('name-box')
  await nameBox.fill('D4:F240')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4:F240')
  await pickToolbarPresetColor(page, 'Fill color', 'theme green')
  const initialReadback = await waitForReadback(
    page,
    {
      points: initialPoints,
      regions: [],
    },
    (result) => allReadbackPointsMatch(result, isThemeGreenFill),
  )

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport, rowHeight) => {
    if (!(viewport instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not a div')
    }
    viewport.scrollTop = rowHeight * 160
    viewport.dispatchEvent(new Event('scroll'))
  }, PRODUCT_ROW_HEIGHT)
  await waitForReadbackSequence(page, initialReadback.sequence)

  const scrolledReadback = await waitForReadback(
    page,
    {
      points: scrolledPoints,
      regions: [],
    },
    (result) => result.sequence > initialReadback.sequence && allReadbackPointsMatch(result, isThemeGreenFill),
  )
  expect(scrolledReadback.sequence).toBeGreaterThan(initialReadback.sequence)

  await clickProductCell(page, 4, 8)
  await pickToolbarPresetColor(page, 'Fill color', 'light cornflower blue 2')
  const overrideReadback = await waitForReadback(
    page,
    {
      points: scrolledPoints,
      regions: [],
    },
    (result) =>
      result.sequence > scrolledReadback.sequence &&
      isThemeGreenFill(result.points.scrolledD) &&
      isCornflowerBlueFill(result.points.scrolledE) &&
      isThemeGreenFill(result.points.scrolledF),
  )
  expect(overrideReadback.sequence).toBeGreaterThan(scrolledReadback.sequence)
  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-large-range-fill-scroll-readback.png',
    'main-workbook-grid-large-range-fill-scroll-readback',
  )
})

test('@browser-webgpu @browser-deep keyboard fill down paints uncached rows before sync catch-up', async ({ page }, testInfo) => {
  const scrolledPoints = [
    { ...selectedRangeFillProbe(1, 5), name: 'bVisibleTop' },
    { ...selectedRangeFillProbe(1, 8), name: 'bVisibleMiddle' },
    { ...selectedRangeFillProbe(1, 11), name: 'bVisibleBottom' },
  ]

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-keyboard-filldown-scroll'))}&persist=0`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const formulaInput = page.getByTestId('formula-input')
  const nameBox = page.getByTestId('name-box')
  const grid = page.getByTestId('sheet-grid')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill('filldown-overlay')
  await formulaInput.press('Enter')
  await clickProductCell(page, 1, 1)
  await pickToolbarPresetColor(page, 'Fill color', 'theme green')

  await nameBox.fill('B2:B20000')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:B20000')
  await grid.press(`${PRIMARY_MODIFIER}+D`)

  const filledSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })
  await page.getByTestId('grid-scroll-viewport').evaluate((viewport, rowHeight) => {
    if (!(viewport instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not a div')
    }
    viewport.scrollTop = rowHeight * 160
    viewport.dispatchEvent(new Event('scroll'))
  }, PRODUCT_ROW_HEIGHT)
  await waitForReadbackSequence(page, filledSequence)

  const readback = await waitForReadback(
    page,
    {
      points: scrolledPoints,
      regions: [],
    },
    (result) => result.sequence > filledSequence && allReadbackPointsMatch(result, isThemeGreenFill),
  )
  expect(readback.sequence).toBeGreaterThan(filledSequence)
  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-keyboard-filldown-scroll-readback.png',
    'main-workbook-grid-keyboard-filldown-scroll-readback',
  )
})

test('@browser-webgpu @browser-deep full-column fill changes repaint visible cells without waiting for sync catch-up', async ({
  page,
}, testInfo) => {
  const points = [
    { ...selectedRangeFillProbe(1, 1), name: 'b2' },
    { ...selectedRangeFillProbe(2, 8), name: 'c9' },
    { ...selectedRangeFillProbe(4, 20), name: 'e21' },
  ]

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-column-fill-fast'))}&persist=0`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  await dragProductHeaderSelection(page, 'column', 1, 4)
  await expect(page.getByTestId('status-selection')).toContainText('!B:E')
  const startedAt = Date.now()
  await pickToolbarPresetColor(page, 'Fill color', 'theme green')
  const readback = await waitForReadback(
    page,
    {
      points,
      regions: [],
    },
    (result) => allReadbackPointsMatch(result, isThemeGreenFill),
  )

  expect(Date.now() - startedAt).toBeLessThan(3_000)
  expect(readback.sequence).toBeGreaterThan(0)

  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-full-column-fill-fast-readback.png',
    'main-workbook-grid-full-column-fill-fast-readback',
  )
})

test('@browser-webgpu @browser-perf main workbook shell keeps typegpu content visible after hover-driven scroll', async ({
  page,
}, testInfo) => {
  test.slow()
  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
  await waitForWorkbookReady(page)
  await waitForBenchmarkCorpus(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const grid = await page.getByTestId('sheet-grid').boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  await page.mouse.move(grid.x + PRODUCT_ROW_MARKER_WIDTH + 180, grid.y + PRODUCT_HEADER_HEIGHT + 96)

  const initialSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    if (!(viewport instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not a div')
    }
    viewport.scrollLeft = 2_600
    viewport.scrollTop = 520
    viewport.dispatchEvent(new Event('scroll'))
  })
  await waitForReadbackSequence(page, initialSequence)

  const readback = await waitForReadback(
    page,
    {
      points: [],
      regions: [
        { name: 'columnHeaderText', x0: PRODUCT_ROW_MARKER_WIDTH, y0: 0, x1: 360, y1: PRODUCT_HEADER_HEIGHT },
        { name: 'rowHeaderText', x0: 0, y0: PRODUCT_HEADER_HEIGHT, x1: PRODUCT_ROW_MARKER_WIDTH, y1: 220 },
        { name: 'bodyText', x0: PRODUCT_ROW_MARKER_WIDTH, y0: PRODUCT_HEADER_HEIGHT, x1: 420, y1: 240 },
      ],
    },
    (result) => result.opaquePixelCounts.columnHeaderText > 100 && result.opaquePixelCounts.rowHeaderText > 100,
  )
  const hoverScrollTextRuns = await waitForVisibleNativeTextRuns(page, [], (runs) => {
    return runs.rowHeaderRunCount > 5 && runs.visibleRunCount > 20
  })

  expect(readback.opaquePixelCounts.columnHeaderText).toBeGreaterThan(100)
  expect(readback.opaquePixelCounts.rowHeaderText).toBeGreaterThan(100)
  expect(hoverScrollTextRuns.rowHeaderRunCount).toBeGreaterThan(5)
  expect(hoverScrollTextRuns.visibleRunCount).toBeGreaterThan(20)

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-hover-scroll-readback.png', 'main-workbook-grid-hover-scroll-readback')
})

test('@browser-webgpu @browser-perf main workbook shell keeps typegpu text visible across tile boundary scroll and resize', async ({
  page,
}, testInfo) => {
  test.slow()
  await page.setViewportSize({ width: 900, height: 680 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
  await waitForWorkbookReady(page)
  await waitForBenchmarkCorpus(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const initialSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })

  await page.getByTestId('grid-scroll-viewport').evaluate(
    (viewport, scrollTarget) => {
      if (!(viewport instanceof HTMLDivElement)) {
        throw new Error('grid scroll viewport is not a div')
      }
      viewport.scrollLeft = scrollTarget.left
      viewport.scrollTop = scrollTarget.top
      viewport.dispatchEvent(new Event('scroll'))
    },
    { left: PRODUCT_COLUMN_WIDTH * 130, top: PRODUCT_ROW_HEIGHT * 34 },
  )
  await waitForReadbackSequence(page, initialSequence)

  await page.setViewportSize({ width: 960, height: 720 })
  await waitForReadbackSequence(page, initialSequence + 1)

  const readback = await waitForReadback(
    page,
    {
      points: [{ name: 'blankBody', x: PRODUCT_ROW_MARKER_WIDTH + 20, y: PRODUCT_HEADER_HEIGHT + 40 }],
      regions: [
        { name: 'canvasDark', x0: 0, y0: 0, x1: 960, y1: 720, threshold: 220 },
        { name: 'columnHeaderText', x0: PRODUCT_ROW_MARKER_WIDTH, y0: 0, x1: 960, y1: PRODUCT_HEADER_HEIGHT },
        { name: 'rowHeaderText', x0: 0, y0: PRODUCT_HEADER_HEIGHT, x1: PRODUCT_ROW_MARKER_WIDTH, y1: 720 },
        { name: 'bodyText', x0: PRODUCT_ROW_MARKER_WIDTH, y0: PRODUCT_HEADER_HEIGHT, x1: 960, y1: 720 },
      ],
    },
    (result) => result.darkPixelCounts.canvasDark > 200,
  )

  expect(readback.darkPixelCounts.canvasDark).toBeGreaterThan(200)
  const tileBoundaryTextRuns = await waitForVisibleNativeTextRuns(page, [], (runs) => {
    return runs.rowHeaderRunCount > 5 && runs.visibleRunCount > 20
  })
  expect(tileBoundaryTextRuns.rowHeaderRunCount).toBeGreaterThan(5)
  expect(tileBoundaryTextRuns.visibleRunCount).toBeGreaterThan(20)
  expect(readback.points.blankBody.a).toBeGreaterThanOrEqual(0)

  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-tile-boundary-resize-readback.png',
    'main-workbook-grid-tile-boundary-resize-readback',
  )
})

async function installTypeGpuReadbackHarness(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const globalWindow = window as Window & {
      __biligGpuReadbackState?: {
        bgra: Uint8Array
        bytesPerRow: number
        hasGpu: boolean
        height: number
        ready: boolean
        sequence: number
        width: number
      }
      __biligGpuReadback?: TypeGpuReadbackSummary
      __biligTypeGpuHarnessInstalled?: boolean
      __biligGpuReadbackInspector?: {
        readonly isReady: () => boolean
        readonly getSequence: () => number
        readonly getSize: () => { readonly width: number; readonly height: number }
        readonly samplePoints: (
          points: readonly { readonly name: string; readonly x: number; readonly y: number }[],
        ) => Record<string, ReadbackPoint>
        readonly countDarkPixels: (
          regions: readonly {
            readonly name: string
            readonly x0: number
            readonly y0: number
            readonly x1: number
            readonly y1: number
            readonly threshold?: number
          }[],
        ) => Record<string, number>
        readonly countOpaquePixels: (
          regions: readonly {
            readonly name: string
            readonly x0: number
            readonly y0: number
            readonly x1: number
            readonly y1: number
            readonly threshold?: number
          }[],
        ) => Record<string, number>
      }
    }

    const readbackState = globalWindow.__biligGpuReadbackState ?? {
      bgra: new Uint8Array(0),
      bytesPerRow: 0,
      hasGpu: Boolean(navigator.gpu),
      height: 0,
      ready: false,
      sequence: 0,
      width: 0,
    }
    globalWindow.__biligGpuReadbackState = readbackState
    readbackState.bgra = new Uint8Array(0)
    readbackState.bytesPerRow = 0
    readbackState.hasGpu = Boolean(navigator.gpu)
    readbackState.height = 0
    readbackState.ready = false
    readbackState.sequence = 0
    readbackState.width = 0

    const pointAt = (x: number, y: number): ReadbackPoint => {
      if (!readbackState.ready || x < 0 || y < 0 || x >= readbackState.width || y >= readbackState.height) {
        return { r: 0, g: 0, b: 0, a: 0 }
      }
      const offset = y * readbackState.bytesPerRow + x * 4
      return {
        r: readbackState.bgra[offset + 2] ?? 0,
        g: readbackState.bgra[offset + 1] ?? 0,
        b: readbackState.bgra[offset + 0] ?? 0,
        a: readbackState.bgra[offset + 3] ?? 0,
      }
    }

    const countDarkPixels = (x0: number, y0: number, x1: number, y1: number, threshold = 120): number => {
      let count = 0
      for (let y = Math.max(0, y0); y < Math.min(readbackState.height, y1); y += 1) {
        for (let x = Math.max(0, x0); x < Math.min(readbackState.width, x1); x += 1) {
          const point = pointAt(x, y)
          if (point.a > 0 && point.r < threshold && point.g < threshold && point.b < threshold) {
            count += 1
          }
        }
      }
      return count
    }

    const countOpaquePixels = (x0: number, y0: number, x1: number, y1: number, threshold = 1): number => {
      let count = 0
      for (let y = Math.max(0, y0); y < Math.min(readbackState.height, y1); y += 1) {
        for (let x = Math.max(0, x0); x < Math.min(readbackState.width, x1); x += 1) {
          if (pointAt(x, y).a >= threshold) {
            count += 1
          }
        }
      }
      return count
    }

    globalWindow.__biligGpuReadbackInspector = {
      countDarkPixels(regions) {
        return Object.fromEntries(
          regions.map((region) => [region.name, countDarkPixels(region.x0, region.y0, region.x1, region.y1, region.threshold)]),
        )
      },
      countOpaquePixels(regions) {
        return Object.fromEntries(
          regions.map((region) => [region.name, countOpaquePixels(region.x0, region.y0, region.x1, region.y1, region.threshold)]),
        )
      },
      getSequence() {
        return readbackState.sequence
      },
      getSize() {
        return { height: readbackState.height, width: readbackState.width }
      },
      isReady() {
        return readbackState.ready
      },
      samplePoints(points) {
        return Object.fromEntries(points.map((point) => [point.name, pointAt(point.x, point.y)]))
      },
    }

    globalWindow.__biligGpuReadback = {
      ready: readbackState.ready,
      hasGpu: readbackState.hasGpu,
      width: readbackState.width,
      height: readbackState.height,
      sequence: readbackState.sequence,
      points: {
        headerFill: { r: 0, g: 0, b: 0, a: 0 },
        bodyFill: { r: 0, g: 0, b: 0, a: 0 },
        selectionBorder: { r: 0, g: 0, b: 0, a: 0 },
        selectionFill: { r: 0, g: 0, b: 0, a: 0 },
        valueFill: { r: 0, g: 0, b: 0, a: 0 },
        bodyWhite: { r: 0, g: 0, b: 0, a: 0 },
      },
      darkPixelCounts: {
        header: 0,
        body: 0,
        number: 0,
      },
    }

    if (globalWindow.__biligTypeGpuHarnessInstalled) {
      return
    }

    if (!navigator.gpu) {
      return
    }

    globalWindow.__biligTypeGpuHarnessInstalled = true

    const functionKind = 'function'
    const isCanvasContextConfigure = (value: unknown): value is (this: GPUCanvasContext, descriptor: GPUCanvasConfiguration) => void =>
      typeof value === functionKind
    const isCanvasContextGetCurrentTexture = (value: unknown): value is (this: GPUCanvasContext) => GPUTexture =>
      typeof value === functionKind
    const readbackCanvasId = 'gpu-readback-canvas'
    const originalConfigure = Object.getOwnPropertyDescriptor(GPUCanvasContext.prototype, 'configure')?.value
    if (!isCanvasContextConfigure(originalConfigure)) {
      return
    }
    GPUCanvasContext.prototype.configure = function configureWithCopySrc(descriptor: GPUCanvasConfiguration) {
      return originalConfigure.call(this, {
        ...descriptor,
        usage: (descriptor.usage ?? GPUTextureUsage.RENDER_ATTACHMENT) | GPUTextureUsage.COPY_SRC,
      })
    }

    const originalRequestAdapter = navigator.gpu.requestAdapter.bind(navigator.gpu)
    navigator.gpu.requestAdapter = async (...adapterArgs) => {
      const adapter = await originalRequestAdapter(...adapterArgs)
      if (!adapter) {
        return adapter
      }

      const originalRequestDevice = adapter.requestDevice.bind(adapter)
      adapter.requestDevice = async (...deviceArgs) => {
        const device = await originalRequestDevice(...deviceArgs)
        let lastTexture: GPUTexture | null = null
        let lastWidth = 0
        let lastHeight = 0
        let lastFallbackTexture: GPUTexture | null = null
        let lastFallbackWidth = 0
        let lastFallbackHeight = 0
        let readbackSerial = 0
        let committedReadbackSerial = 0

        const originalGetCurrentTexture = Object.getOwnPropertyDescriptor(GPUCanvasContext.prototype, 'getCurrentTexture')?.value
        if (!isCanvasContextGetCurrentTexture(originalGetCurrentTexture)) {
          return device
        }
        GPUCanvasContext.prototype.getCurrentTexture = function recordCurrentTexture() {
          const texture = originalGetCurrentTexture.call(this)
          if (this.canvas instanceof HTMLCanvasElement) {
            const testId = this.canvas.getAttribute('data-testid')
            const dataPaneRenderer = this.canvas.getAttribute('data-pane-renderer')
            const tracked = testId === 'grid-pane-renderer' || dataPaneRenderer === 'workbook-pane-renderer'
            lastFallbackTexture = texture
            lastFallbackWidth = this.canvas.width
            lastFallbackHeight = this.canvas.height
            if (!tracked) {
              return texture
            }
            lastTexture = texture
            lastWidth = this.canvas.width
            lastHeight = this.canvas.height
          }
          return texture
        }

        const originalSubmit = device.queue.submit.bind(device.queue)
        device.queue.submit = (buffers: Iterable<GPUCommandBuffer>) => {
          const commandBuffers = Array.from(buffers)
          const targetTexture = lastTexture ?? lastFallbackTexture
          const targetWidth = lastTexture ? lastWidth : lastFallbackWidth
          const targetHeight = lastTexture ? lastHeight : lastFallbackHeight
          if (targetTexture && targetWidth > 0 && targetHeight > 0) {
            readbackSerial += 1
            const serial = readbackSerial
            const bytesPerRow = Math.ceil((targetWidth * 4) / 256) * 256
            const buffer = device.createBuffer({
              size: bytesPerRow * targetHeight,
              usage: GPUBufferUsage.COPY_DST | GPUBufferUsage.MAP_READ,
            })
            const encoder = device.createCommandEncoder()
            encoder.copyTextureToBuffer(
              { texture: targetTexture },
              { buffer, bytesPerRow, rowsPerImage: targetHeight },
              { width: targetWidth, height: targetHeight, depthOrArrayLayers: 1 },
            )
            const result = originalSubmit([...commandBuffers, encoder.finish()])
            void buffer
              .mapAsync(GPUMapMode.READ)
              .then(() => {
                const mapped = new Uint8Array(buffer.getMappedRange())
                if (serial <= committedReadbackSerial) {
                  return mapped
                }
                const bgra = new Uint8Array(mapped)
                committedReadbackSerial = serial
                readbackState.bgra = bgra
                readbackState.bytesPerRow = bytesPerRow
                readbackState.hasGpu = true
                readbackState.height = targetHeight
                readbackState.ready = true
                readbackState.sequence += 1
                readbackState.width = targetWidth
                globalWindow.__biligGpuReadback = buildReadbackSummary({
                  width: targetWidth,
                  height: targetHeight,
                  bytesPerRow,
                  bgra,
                  hasGpu: true,
                  sequence: readbackState.sequence,
                })
                renderReadbackCanvas({ width: targetWidth, height: targetHeight, bytesPerRow, bgra })
                return bgra
              })
              .finally(() => {
                try {
                  buffer.unmap()
                } catch (error) {
                  console.warn('Ignoring GPU buffer unmap error', error)
                }
                buffer.destroy()
              })
            return result
          }
          return originalSubmit(commandBuffers)
        }

        return device
      }

      return adapter
    }

    function buildReadbackSummary(input: {
      readonly width: number
      readonly height: number
      readonly bytesPerRow: number
      readonly bgra: Uint8Array
      readonly hasGpu: boolean
      readonly sequence: number
    }): TypeGpuReadbackSummary {
      const samplePoint = (x: number, y: number): ReadbackPoint => {
        const offset = y * input.bytesPerRow + x * 4
        return {
          r: input.bgra[offset + 2] ?? 0,
          g: input.bgra[offset + 1] ?? 0,
          b: input.bgra[offset + 0] ?? 0,
          a: input.bgra[offset + 3] ?? 0,
        }
      }

      const sampleDarkPixels = (x0: number, y0: number, x1: number, y1: number): number => {
        let count = 0
        for (let y = y0; y < y1; y += 1) {
          for (let x = x0; x < x1; x += 1) {
            const point = samplePoint(x, y)
            if (point.a > 0 && point.r < 120 && point.g < 120 && point.b < 120) {
              count += 1
            }
          }
        }
        return count
      }

      return {
        ready: true,
        hasGpu: input.hasGpu,
        width: input.width,
        height: input.height,
        sequence: input.sequence,
        points: {
          headerFill: samplePoint(20, 12),
          bodyFill: samplePoint(60, 40),
          selectionBorder: samplePoint(200, 68),
          selectionFill: samplePoint(260, 100),
          valueFill: samplePoint(520, 140),
          bodyWhite: samplePoint(400, 300),
        },
        darkPixelCounts: {
          header: sampleDarkPixels(80, 4, 120, 18),
          body: sampleDarkPixels(58, 48, 110, 66),
          number: sampleDarkPixels(532, 48, 620, 70),
        },
      }
    }

    function renderReadbackCanvas(input: {
      readonly width: number
      readonly height: number
      readonly bytesPerRow: number
      readonly bgra: Uint8Array
    }): void {
      const existing = globalWindow.document.getElementById(readbackCanvasId)
      existing?.remove()

      const rgba = new Uint8ClampedArray(input.width * input.height * 4)
      for (let y = 0; y < input.height; y += 1) {
        const rowOffset = y * input.bytesPerRow
        for (let x = 0; x < input.width; x += 1) {
          const src = rowOffset + x * 4
          const dst = (y * input.width + x) * 4
          rgba[dst + 0] = input.bgra[src + 2] ?? 0
          rgba[dst + 1] = input.bgra[src + 1] ?? 0
          rgba[dst + 2] = input.bgra[src + 0] ?? 0
          rgba[dst + 3] = input.bgra[src + 3] ?? 0
        }
      }

      const canvas = globalWindow.document.createElement('canvas')
      canvas.id = readbackCanvasId
      canvas.width = input.width
      canvas.height = input.height
      canvas.style.position = 'fixed'
      canvas.style.left = '0'
      canvas.style.top = '0'
      canvas.style.zIndex = '99999'
      canvas.style.pointerEvents = 'none'
      const context = canvas.getContext('2d')
      if (!context) {
        return
      }
      context.putImageData(new ImageData(rgba, input.width, input.height), 0, 0)
      globalWindow.document.body.appendChild(canvas)
    }
  })
}

async function inspectGpuReadback(
  page: Page,
  input: {
    readonly points: readonly ReadbackInspectorPoint[]
    readonly regions: readonly ReadbackInspectorRegion[]
  },
): Promise<DynamicReadbackResult> {
  const result = await page.evaluate(({ points, regions }) => {
    const inspector = (
      window as Window & {
        __biligGpuReadbackInspector?: {
          readonly isReady: () => boolean
          readonly getSequence: () => number
          readonly getSize: () => { readonly width: number; readonly height: number }
          readonly samplePoints: (points: readonly ReadbackInspectorPoint[]) => Record<string, ReadbackPoint>
          readonly countDarkPixels: (regions: readonly ReadbackInspectorRegion[]) => Record<string, number>
          readonly countOpaquePixels: (regions: readonly ReadbackInspectorRegion[]) => Record<string, number>
        }
        __biligGpuReadback?: { readonly hasGpu: boolean }
      }
    ).__biligGpuReadbackInspector
    const hasGpu = Boolean((window as Window & { __biligGpuReadback?: { readonly hasGpu: boolean } }).__biligGpuReadback?.hasGpu)

    if (!inspector) {
      return {
        ready: false,
        hasGpu,
        width: 0,
        height: 0,
        sequence: 0,
        points: {},
        darkPixelCounts: {},
        opaquePixelCounts: {},
      }
    }

    const size = inspector.getSize()
    return {
      ready: inspector.isReady(),
      hasGpu,
      width: size.width,
      height: size.height,
      sequence: inspector.getSequence(),
      points: inspector.samplePoints(points),
      darkPixelCounts: inspector.countDarkPixels(regions),
      opaquePixelCounts: inspector.countOpaquePixels(regions),
    }
  }, input)

  expect(result.ready).toBe(true)
  return result
}

async function inspectVisibleNativeTextRuns(
  page: Page,
  textExpectations: readonly NativeTextExpectation[],
): Promise<NativeTextLayerInspection> {
  return await page.evaluate(
    (payload) => {
      const textLayer = document.querySelector('[data-testid="grid-native-text-layer"]')
      const sheetGrid = document.querySelector('[data-testid="sheet-grid"]')
      const renderer = document.querySelector('[data-testid="grid-pane-renderer"]')
      const { rowHeaderWidth } = payload
      const viewportElement = textLayer instanceof HTMLElement ? textLayer : sheetGrid instanceof HTMLElement ? sheetGrid : null
      const viewportRect = viewportElement?.getBoundingClientRect() ?? null
      const runs = [...document.querySelectorAll('[data-native-text-run]')].flatMap((node) => {
        const rect = node.getBoundingClientRect()
        if (
          !viewportRect ||
          rect.right <= viewportRect.left ||
          rect.left >= viewportRect.right ||
          rect.bottom <= viewportRect.top ||
          rect.top >= viewportRect.bottom
        ) {
          return []
        }
        return [
          {
            relativeLeft: rect.left - viewportRect.left,
            text: node.textContent ?? '',
          },
        ]
      })
      return {
        gridAttrs: Object.fromEntries(
          [...(sheetGrid?.attributes ?? [])]
            .filter((attribute) => attribute.name.includes('render') || attribute.name.includes('revision'))
            .map((attribute) => [attribute.name, attribute.value]),
        ),
        matches: Object.fromEntries(
          payload.textExpectations.map((expectation) => [
            expectation.name,
            runs.some((run) => (expectation.exact ? run.text === expectation.text : run.text.includes(expectation.text))),
          ]),
        ),
        rendererAttrs: Object.fromEntries(
          [...(renderer?.attributes ?? [])]
            .filter((attribute) => attribute.name.includes('v3'))
            .map((attribute) => [attribute.name, attribute.value]),
        ),
        rowHeaderRunCount: runs.filter((run) => run.relativeLeft >= 0 && run.relativeLeft < rowHeaderWidth && /^\d+$/.test(run.text))
          .length,
        sampleTexts: runs.map((run) => run.text).slice(0, 120),
        visibleRunCount: runs.length,
      }
    },
    { textExpectations, rowHeaderWidth: PRODUCT_ROW_MARKER_WIDTH },
  )
}

async function waitForVisibleNativeTextRuns(
  page: Page,
  expectations: readonly NativeTextExpectation[],
  predicate: (runs: NativeTextLayerInspection) => boolean,
): Promise<NativeTextLayerInspection> {
  let lastResult: NativeTextLayerInspection | null = null
  try {
    await expect
      .poll(
        async () => {
          lastResult = await inspectVisibleNativeTextRuns(page, expectations)
          return predicate(lastResult)
        },
        { timeout: 10_000 },
      )
      .toBe(true)
  } catch (error) {
    throw new Error(`${error instanceof Error ? error.message : String(error)}\nLast native text runs: ${JSON.stringify(lastResult)}`, {
      cause: error,
    })
  }
  if (!lastResult) {
    throw new Error('expected native text run result')
  }
  return lastResult
}

async function waitForReadback(
  page: Page,
  input: {
    readonly points: readonly ReadbackInspectorPoint[]
    readonly regions: readonly ReadbackInspectorRegion[]
  },
  predicate: (result: DynamicReadbackResult) => boolean,
): Promise<DynamicReadbackResult> {
  let lastResult: DynamicReadbackResult | null = null
  try {
    await expect
      .poll(
        async () => {
          lastResult = await inspectGpuReadback(page, input)
          return lastResult.ready && lastResult.hasGpu && predicate(lastResult)
        },
        { timeout: 30_000 },
      )
      .toBe(true)
  } catch (error) {
    throw new Error(`${error instanceof Error ? error.message : String(error)}\nLast readback: ${JSON.stringify(lastResult)}`, {
      cause: error,
    })
  }
  if (!lastResult) {
    throw new Error('expected readback result')
  }
  return lastResult
}

async function waitForReadbackSequence(page: Page, previousSequence: number): Promise<void> {
  await page.waitForFunction(
    (sequence) => {
      const inspector = (window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } })
        .__biligGpuReadbackInspector
      return (inspector?.getSequence() ?? 0) > sequence
    },
    previousSequence,
    { timeout: 30_000 },
  )
}

async function saveReadbackArtifact(page: Page, testInfo: TestInfo, fileName: string, attachmentName: string): Promise<void> {
  const outputPath = testInfo.outputPath(fileName)
  const dataUrl = await page.evaluate(() => {
    const canvas = document.getElementById('gpu-readback-canvas')
    if (!(canvas instanceof HTMLCanvasElement)) {
      return null
    }
    return canvas.toDataURL('image/png')
  })
  if (!dataUrl) {
    throw new Error('gpu readback canvas unavailable')
  }
  await writeFile(outputPath, dataUrl.replace(/^data:image\/png;base64,/, ''), 'base64')
  await testInfo.attach(attachmentName, {
    path: outputPath,
    contentType: 'image/png',
  })
}
