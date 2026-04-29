import { createHash } from 'node:crypto'
import { writeFile } from 'node:fs/promises'
import { expect, type Locator, type Page, type TestInfo } from '@playwright/test'

export const PRODUCT_ROW_MARKER_WIDTH = 46
export const PRODUCT_COLUMN_WIDTH = 104
export const PRODUCT_HEADER_HEIGHT = 24
export const PRODUCT_ROW_HEIGHT = 22
export const PRIMARY_MODIFIER = process.platform === 'darwin' ? 'Meta' : 'Control'
export const remoteSyncEnabled = process.env['BILIG_E2E_REMOTE_SYNC'] !== '0'

interface ToolbarSyncAction {
  readonly label: string
  readonly apply: (page: Page) => Promise<void>
}

function delay(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms))
}

function parseDimensionOverrides(raw: string | null): Record<string, number> {
  if (!raw) {
    return {}
  }
  const parsed: unknown = JSON.parse(raw)
  if (typeof parsed !== 'object' || parsed === null) {
    return {}
  }
  const entries = Object.entries(parsed).filter((entry): entry is [string, number] => typeof entry[1] === 'number')
  return Object.fromEntries(entries)
}

export async function getProductColumnWidth(page: Page, columnIndex: number) {
  const grid = page.getByTestId('sheet-grid')
  const [defaultWidthRaw, overridesRaw] = await Promise.all([
    grid.getAttribute('data-default-column-width'),
    grid.getAttribute('data-column-width-overrides'),
  ])
  const defaultWidth = Number(defaultWidthRaw ?? String(PRODUCT_COLUMN_WIDTH))
  const overrides = parseDimensionOverrides(overridesRaw)
  return overrides[String(columnIndex)] ?? defaultWidth
}

export async function getProductRowHeight(page: Page, rowIndex: number) {
  const grid = page.getByTestId('sheet-grid')
  const [defaultHeightRaw, overridesRaw] = await Promise.all([
    grid.getAttribute('data-default-row-height'),
    grid.getAttribute('data-row-height-overrides'),
  ])
  const defaultHeight = Number(defaultHeightRaw ?? String(PRODUCT_ROW_HEIGHT))
  const overrides = parseDimensionOverrides(overridesRaw)
  return overrides[String(rowIndex)] ?? defaultHeight
}

export async function getProductRowTop(page: Page, rowIndex: number) {
  const heights = await Promise.all(Array.from({ length: rowIndex }, (_, index) => getProductRowHeight(page, index)))
  return heights.reduce((total, height) => total + height, 0)
}

export async function getProductColumnLeft(page: Page, columnIndex: number) {
  const widths = await Promise.all(Array.from({ length: columnIndex }, (_, index) => getProductColumnWidth(page, index)))
  return PRODUCT_ROW_MARKER_WIDTH + widths.reduce((total, width) => total + width, 0)
}

export async function getBox(locator: Locator) {
  await expect(locator).toBeVisible()
  const box = await locator.boundingBox()
  if (!box) {
    throw new Error('locator is not visible')
  }
  return box
}

export async function clickProductCell(
  page: Page,
  columnIndex: number,
  rowIndex: number,
  options?: {
    shift?: boolean
  },
) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex)
  const columnWidth = await getProductColumnWidth(page, columnIndex)
  if (options?.shift) {
    await page.keyboard.down('Shift')
  }
  try {
    await page.mouse.click(
      grid.x + columnLeft + Math.floor(columnWidth / 2),
      grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2),
    )
  } finally {
    if (options?.shift) {
      await page.keyboard.up('Shift')
    }
  }
}

export async function dragProductColumnResize(page: Page, columnIndex: number, deltaX: number) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex)
  const columnWidth = await getProductColumnWidth(page, columnIndex)
  const edgeX = grid.x + columnLeft + columnWidth - 1
  const edgeY = grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2)

  await page.mouse.move(edgeX, edgeY)
  await page.mouse.down()
  await page.mouse.move(edgeX + deltaX, edgeY, { steps: 10 })
  await page.mouse.up()
}

export async function doubleClickProductColumnResizeHandle(page: Page, columnIndex: number) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex)
  const columnWidth = await getProductColumnWidth(page, columnIndex)
  const edgeX = grid.x + columnLeft + columnWidth - 1
  const headerY = grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2)
  await page.mouse.move(edgeX, headerY)
  await page.mouse.dblclick(edgeX, headerY)
}

export async function dragProductHeaderSelection(page: Page, axis: 'column' | 'row', startIndex: number, endIndex: number) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startColumnLeft = axis === 'column' ? await getProductColumnLeft(page, startIndex) : 0
  const endColumnLeft = axis === 'column' ? await getProductColumnLeft(page, endIndex) : 0
  const startColumnWidth = axis === 'column' ? await getProductColumnWidth(page, startIndex) : 0
  const endColumnWidth = axis === 'column' ? await getProductColumnWidth(page, endIndex) : 0
  const startX =
    axis === 'column' ? grid.x + startColumnLeft + Math.floor(startColumnWidth / 2) : grid.x + Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2)
  const startY =
    axis === 'column'
      ? grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2)
      : grid.y + PRODUCT_HEADER_HEIGHT + startIndex * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  const endX =
    axis === 'column' ? grid.x + endColumnLeft + Math.floor(endColumnWidth / 2) : grid.x + Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2)
  const endY =
    axis === 'column'
      ? grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2)
      : grid.y + PRODUCT_HEADER_HEIGHT + endIndex * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)

  await page.mouse.move(startX, startY)
  await page.mouse.down()
  await page.mouse.move(endX, endY, { steps: 8 })
  await page.mouse.up()
}

export async function clickGridRightEdge(page: Page, rowIndex = 2) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const x = grid.x + grid.width - 3
  const y = grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  await page.mouse.click(x, y)
}

export function getToolbarButton(page: Page, label: string): Locator {
  return page.getByRole('button', { name: label, exact: true })
}

function getToolbarSelect(page: Page, label: string): Locator {
  return page.getByRole('combobox', { name: label, exact: true })
}

export async function expectToolbarColor(locator: Locator, value: string) {
  await expect(locator).toHaveAttribute('data-current-color', value.toLowerCase())
}

async function expectToolbarSelectValue(page: Page, label: string, value: string) {
  await expect(getToolbarSelect(page, label)).toHaveAttribute('data-current-value', value)
}

async function selectToolbarOption(page: Page, label: string, optionLabel: string, expectedValue = optionLabel) {
  await getToolbarSelect(page, label).click()
  await page.getByRole('option', { name: optionLabel, exact: true }).click()
  await expectToolbarSelectValue(page, label, expectedValue)
}

async function setColorInput(locator: Locator, value: string) {
  await locator.evaluate((element, nextValue) => {
    if (!(element instanceof HTMLInputElement)) {
      throw new Error('color input is not an HTMLInputElement')
    }
    const input = element
    const descriptor = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')
    descriptor?.set?.call(input, String(nextValue))
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new Event('change', { bubbles: true }))
  }, value)
}

async function setToolbarCustomColor(page: Page, controlLabel: 'Fill color' | 'Text color', value: string) {
  await getToolbarButton(page, controlLabel).click()
  await page.getByLabel(`Open custom ${controlLabel.toLowerCase()} picker`).click()
  await setColorInput(
    page.getByLabel(controlLabel === 'Fill color' ? 'Custom fill color' : 'Custom text color', {
      exact: true,
    }),
    value,
  )
}

export async function pickToolbarPresetColor(page: Page, controlLabel: 'Fill color' | 'Text color', swatchLabel: string) {
  await getToolbarButton(page, controlLabel).click()
  await page.getByLabel(`${controlLabel} ${swatchLabel}`).click()
}

export async function pickToolbarBorderPreset(
  page: Page,
  presetLabel: 'All borders' | 'Outer borders' | 'Left border' | 'Top border' | 'Right border' | 'Bottom border' | 'Clear borders',
) {
  await getToolbarButton(page, 'Borders').click()
  await page.getByRole('button', { name: presetLabel }).click()
}

export const TOOLBAR_SYNC_ACTIONS: readonly ToolbarSyncAction[] = [
  {
    label: 'number-format-accounting',
    apply: async (activePage) => await selectToolbarOption(activePage, 'Number format', 'Accounting', 'accounting'),
  },
  {
    label: 'font-size-14',
    apply: async (activePage) => await selectToolbarOption(activePage, 'Font size', '14'),
  },
  { label: 'bold', apply: async (activePage) => await activePage.getByLabel('Bold').click() },
  {
    label: 'italic',
    apply: async (activePage) => await activePage.getByLabel('Italic').click(),
  },
  {
    label: 'underline',
    apply: async (activePage) => await activePage.getByLabel('Underline').click(),
  },
  {
    label: 'fill-color',
    apply: async (activePage) => await setToolbarCustomColor(activePage, 'Fill color', '#dbeafe'),
  },
  {
    label: 'text-color',
    apply: async (activePage) => await setToolbarCustomColor(activePage, 'Text color', '#7c2d12'),
  },
  {
    label: 'align-left',
    apply: async (activePage) => await activePage.getByLabel('Align left').click(),
  },
  {
    label: 'align-center',
    apply: async (activePage) => await activePage.getByLabel('Align center').click(),
  },
  {
    label: 'align-right',
    apply: async (activePage) => await activePage.getByLabel('Align right').click(),
  },
  {
    label: 'border-all',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'All borders'),
  },
  {
    label: 'border-outer',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'Outer borders'),
  },
  {
    label: 'border-left',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'Left border'),
  },
  {
    label: 'border-top',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'Top border'),
  },
  {
    label: 'border-right',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'Right border'),
  },
  {
    label: 'border-bottom',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'Bottom border'),
  },
  {
    label: 'border-clear',
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, 'Clear borders'),
  },
  { label: 'wrap', apply: async (activePage) => await activePage.getByLabel('Wrap').click() },
  {
    label: 'clear-style',
    apply: async (activePage) => await activePage.getByLabel('Clear style').click(),
  },
  {
    label: 'number-format-general',
    apply: async (activePage) => await selectToolbarOption(activePage, 'Number format', 'General', 'general'),
  },
]

export async function selectToolbarActionRange(page: Page) {
  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await clickProductCell(page, 2, 2, { shift: true })
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')
}

export async function clickProductBodyOffset(page: Page, offsetX: number, rowIndex = 0) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  await page.mouse.click(
    grid.x + PRODUCT_ROW_MARKER_WIDTH + offsetX,
    grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2),
  )
}

export async function clickProductCellUpperHalf(page: Page, columnIndex: number, rowIndex: number) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex)
  const columnWidth = await getProductColumnWidth(page, columnIndex)
  await page.mouse.click(
    grid.x + columnLeft + Math.floor(columnWidth / 2),
    grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT + 4,
  )
}

export async function dragProductBodySelection(page: Page, startColumn: number, startRow: number, endColumn: number, endRow: number) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startLeft = await getProductColumnLeft(page, startColumn)
  const startWidth = await getProductColumnWidth(page, startColumn)
  const endLeft = await getProductColumnLeft(page, endColumn)
  const endWidth = await getProductColumnWidth(page, endColumn)

  const startX = grid.x + startLeft + Math.floor(startWidth / 2)
  const startY = grid.y + PRODUCT_HEADER_HEIGHT + startRow * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  const endX = grid.x + endLeft + Math.floor(endWidth / 2)
  const endY = grid.y + PRODUCT_HEADER_HEIGHT + endRow * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  await page.mouse.move(startX, startY)
  await page.mouse.down()
  await page.mouse.move(endX, endY, { steps: 8 })
  await page.mouse.up()
}

export async function seedToolbarActionRange(page: Page) {
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  await nameBox.fill('B2')
  await nameBox.press('Enter')
  await formulaInput.fill('1234.5')
  await formulaInput.press('Enter')

  await nameBox.fill('C2')
  await nameBox.press('Enter')
  await formulaInput.fill('6789.125')
  await formulaInput.press('Enter')

  await nameBox.fill('B3')
  await nameBox.press('Enter')
  await formulaInput.fill('42.25')
  await formulaInput.press('Enter')

  await nameBox.fill('C3')
  await nameBox.press('Enter')
  await formulaInput.fill('-7.5')
  await formulaInput.press('Enter')
}

async function captureGridRangeScreenshot(page: Page, startColumn: number, startRow: number, endColumn: number, endRow: number) {
  await page.bringToFront()
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const minColumn = Math.min(startColumn, endColumn)
  const maxColumn = Math.max(startColumn, endColumn)
  const minRow = Math.min(startRow, endRow)
  const maxRow = Math.max(startRow, endRow)
  const startLeft = await getProductColumnLeft(page, minColumn)
  const endLeft = await getProductColumnLeft(page, maxColumn)
  const endWidth = await getProductColumnWidth(page, maxColumn)

  return await page.screenshot({
    animations: 'disabled',
    caret: 'hide',
    clip: {
      x: Math.round(grid.x + startLeft),
      y: Math.round(grid.y + PRODUCT_HEADER_HEIGHT + minRow * PRODUCT_ROW_HEIGHT),
      width: Math.round(endLeft + endWidth - startLeft),
      height: Math.round((maxRow - minRow + 1) * PRODUCT_ROW_HEIGHT),
    },
  })
}

async function compareScreenshotPixels(page: Page, left: Buffer, right: Buffer) {
  return await page.evaluate(
    async ({ leftDataUrl, rightDataUrl, channelTolerance }) => {
      const [leftImage, rightImage] = await Promise.all([
        new Promise<HTMLImageElement>((resolve, reject) => {
          const image = new Image()
          image.addEventListener('load', () => resolve(image), { once: true })
          image.addEventListener('error', () => reject(new Error('Failed to decode left screenshot data URL')), { once: true })
          image.src = leftDataUrl
        }),
        new Promise<HTMLImageElement>((resolve, reject) => {
          const image = new Image()
          image.addEventListener('load', () => resolve(image), { once: true })
          image.addEventListener('error', () => reject(new Error('Failed to decode right screenshot data URL')), { once: true })
          image.src = rightDataUrl
        }),
      ])
      if (leftImage.naturalWidth !== rightImage.naturalWidth || leftImage.naturalHeight !== rightImage.naturalHeight) {
        return {
          equal: false,
          diffPixels: Number.POSITIVE_INFINITY,
          width: leftImage.naturalWidth,
          height: leftImage.naturalHeight,
        }
      }

      const width = leftImage.naturalWidth
      const height = leftImage.naturalHeight
      const canvas = document.createElement('canvas')
      canvas.width = width
      canvas.height = height
      const context = canvas.getContext('2d')
      if (!context) {
        throw new Error('Missing 2d context for screenshot comparison')
      }

      context.clearRect(0, 0, width, height)
      context.drawImage(leftImage, 0, 0)
      const leftPixels = context.getImageData(0, 0, width, height).data
      context.clearRect(0, 0, width, height)
      context.drawImage(rightImage, 0, 0)
      const rightPixels = context.getImageData(0, 0, width, height).data

      let diffPixels = 0
      for (let index = 0; index < leftPixels.length; index += 4) {
        if (
          Math.abs(leftPixels[index] - rightPixels[index]) > channelTolerance ||
          Math.abs(leftPixels[index + 1] - rightPixels[index + 1]) > channelTolerance ||
          Math.abs(leftPixels[index + 2] - rightPixels[index + 2]) > channelTolerance ||
          Math.abs(leftPixels[index + 3] - rightPixels[index + 3]) > channelTolerance
        ) {
          diffPixels += 1
        }
      }

      return { equal: diffPixels === 0, diffPixels, width, height }
    },
    {
      leftDataUrl: `data:image/png;base64,${left.toString('base64')}`,
      rightDataUrl: `data:image/png;base64,${right.toString('base64')}`,
      channelTolerance: 2,
    },
  )
}

async function pollMatchingGridRangeScreenshots(
  primaryPage: Page,
  mirrorPage: Page,
  startedAt: number,
  timeoutMs: number,
  maxDiffPixels: number,
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
): Promise<{
  primaryBuffer: Buffer
  mirrorBuffer: Buffer
  diffPixels: number
  matched: boolean
}> {
  const [primaryBuffer, mirrorBuffer] = await Promise.all([
    captureGridRangeScreenshot(primaryPage, startColumn, startRow, endColumn, endRow),
    captureGridRangeScreenshot(mirrorPage, startColumn, startRow, endColumn, endRow),
  ])
  const comparison = await compareScreenshotPixels(primaryPage, primaryBuffer, mirrorBuffer)
  const matched = comparison.equal || comparison.diffPixels <= maxDiffPixels
  if (matched || Date.now() - startedAt > timeoutMs) {
    return {
      primaryBuffer,
      mirrorBuffer,
      diffPixels: comparison.diffPixels,
      matched,
    }
  }

  await delay(50)
  return await pollMatchingGridRangeScreenshots(
    primaryPage,
    mirrorPage,
    startedAt,
    timeoutMs,
    maxDiffPixels,
    startColumn,
    startRow,
    endColumn,
    endRow,
  )
}

export async function expectMatchingGridRangeScreenshots(
  primaryPage: Page,
  mirrorPage: Page,
  actionLabel: string,
  testInfo: TestInfo,
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
  timeoutMs = 1_500,
  maxDiffPixels = 8,
) {
  const startedAt = Date.now()
  const result = await pollMatchingGridRangeScreenshots(
    primaryPage,
    mirrorPage,
    startedAt,
    timeoutMs,
    maxDiffPixels,
    startColumn,
    startRow,
    endColumn,
    endRow,
  )
  if (result.matched) {
    return Date.now() - startedAt
  }

  const primaryHash = createHash('sha256').update(result.primaryBuffer).digest('hex')
  const mirrorHash = createHash('sha256').update(result.mirrorBuffer).digest('hex')
  await writeFile(testInfo.outputPath(`multiplayer-${actionLabel}-range-primary.png`), result.primaryBuffer)
  await writeFile(testInfo.outputPath(`multiplayer-${actionLabel}-range-mirror.png`), result.mirrorBuffer)

  throw new Error(
    `multiplayer grid range screenshots diverged for ${actionLabel} after ${timeoutMs}ms (primary=${primaryHash}, mirror=${mirrorHash}, diffPixels=${result.diffPixels}, maxDiffPixels=${maxDiffPixels})`,
  )
}

export async function waitForWorkbookReady(page: Page) {
  await expect(page.getByTestId('formula-bar')).toBeVisible({ timeout: 15_000 })
  await expect(page.getByTestId('sheet-grid')).toBeVisible({ timeout: 15_000 })
  await expect(page.getByTestId('status-sync')).toHaveText(
    /^(Saved|Saving…|Sync pending|Local saved|Local only|Offline|Sync issue|Read only)$/,
    {
      timeout: 15_000,
    },
  )
}

export async function gotoWorkbookShell(page: Page, path = '/', timeoutMs = 15_000) {
  async function attempt(deadline: number, lastError: unknown): Promise<void> {
    if (Date.now() >= deadline) {
      throw lastError instanceof Error ? lastError : new Error(`Timed out navigating to ${path}`)
    }

    try {
      await page.goto(path)
      return
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      if (!message.includes('ERR_CONNECTION_REFUSED')) {
        throw error
      }
      await delay(250)
      await attempt(deadline, error)
    }
  }

  await attempt(Date.now() + timeoutMs, null)
}

export async function openZeroWorkbookPage(page: Page, documentId: string) {
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(documentId)}`)
  await waitForWorkbookReady(page)
  await selectToolbarActionRange(page)
}

export async function waitForBenchmarkCorpus(page: Page, timeoutMs = 60_000) {
  await page.waitForFunction(
    () => {
      const collector = (window as Window & { __biligScrollPerf?: { getBenchmarkState?: () => { state: string; error: string | null } } })
        .__biligScrollPerf
      const state = collector?.getBenchmarkState?.()
      return state?.state === 'ready' || state?.state === 'error'
    },
    undefined,
    { timeout: timeoutMs },
  )

  const benchmarkState = await page.evaluate(() => {
    const collector = (
      window as Window & {
        __biligScrollPerf?: {
          getBenchmarkState?: () => {
            state: string
            error: string | null
            fixture: { id: string; materializedCellCount: number; sheetName: string } | null
          }
        }
      }
    ).__biligScrollPerf
    return collector?.getBenchmarkState?.() ?? null
  })

  if (!benchmarkState) {
    throw new Error('benchmark corpus state was not available')
  }
  if (benchmarkState.state === 'error') {
    throw new Error(benchmarkState.error ?? 'benchmark corpus installation failed')
  }
  return benchmarkState
}

export async function startWorkbookScrollPerf(
  page: Page,
  workload: string,
  options: {
    readonly primeRenderer?: boolean
  } = {},
) {
  await page.bringToFront()
  await settleWorkbookScrollPerf(page, 2)
  if (options.primeRenderer ?? true) {
    await primeWorkbookGridScrollRenderer(page)
  }
  await page.evaluate((nextWorkload) => {
    ;(window as Window & { __biligScrollPerf?: { startSampling?: (workload: string) => void } }).__biligScrollPerf?.startSampling?.(
      nextWorkload,
    )
  }, workload)
}

async function primeWorkbookGridScrollRenderer(page: Page) {
  await page.getByTestId('grid-scroll-viewport').evaluate(async (element) => {
    if (!(element instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not an HTMLDivElement')
    }
    const startLeft = element.scrollLeft
    const startTop = element.scrollTop
    element.scrollLeft = startLeft + 1
    element.scrollTop = startTop + 1
    element.dispatchEvent(new Event('scroll'))
    await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
    element.scrollLeft = startLeft
    element.scrollTop = startTop
    element.dispatchEvent(new Event('scroll'))
    await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
  })
}

export async function warmStartWorkbookScrollPerf(page: Page, workload: string, warmupFrames = 12, maxAttempts = 8) {
  const quietFrames = Math.max(warmupFrames, 96)
  const runWarmup = async (attempt: number): Promise<void> => {
    await startWorkbookScrollPerf(page, `${workload}:warmup:${String(attempt)}`, { primeRenderer: attempt === 1 })
    await settleWorkbookScrollPerf(page, warmupFrames + quietFrames + 2)
    const warmupReport = await stopWorkbookScrollPerf(page)
    if (!warmupReport) {
      throw new Error('warmup performance report was not available')
    }
    const surfaceCommits = Object.values(warmupReport.counters.surfaceCommits ?? {})
    const hasSurfaceCommitNoise = surfaceCommits.some((count) => count > 0)
    const hasTypeGpuWarmupNoise =
      warmupReport.counters.typeGpuAtlasUploadBytes > 0 ||
      warmupReport.counters.typeGpuBufferAllocations > 0 ||
      warmupReport.counters.typeGpuConfigures > 0 ||
      warmupReport.counters.typeGpuSurfaceResizes > 0 ||
      warmupReport.counters.typeGpuVertexUploadBytes > 0
    const hasRenderNoise =
      warmupReport.counters.viewportSubscriptions > 0 ||
      warmupReport.counters.fullPatches > 0 ||
      warmupReport.counters.headerPaneBuilds > 0 ||
      warmupReport.counters.reactCommits > 0 ||
      warmupReport.counters.canvasSurfaceMounts > 0 ||
      warmupReport.counters.domSurfaceMounts > 0 ||
      hasSurfaceCommitNoise ||
      hasTypeGpuWarmupNoise
    if (!hasRenderNoise) {
      return
    }
    if (attempt >= maxAttempts) {
      throw new Error(`scroll performance never reached a steady state for ${workload}`)
    }
    await runWarmup(attempt + 1)
  }
  await runWarmup(1)
  await startWorkbookScrollPerf(page, workload, { primeRenderer: false })
}

export async function resetGridScroll(page: Page, input: { left?: number; top?: number } = {}) {
  await page.getByTestId('grid-scroll-viewport').evaluate(async (element, position) => {
    if (!(element instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not an HTMLDivElement')
    }
    element.scrollLeft = position.left ?? 0
    element.scrollTop = position.top ?? 0
    element.dispatchEvent(new Event('scroll'))
    await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
  }, input)
}

export async function settleWorkbookScrollPerf(page: Page, frames = 4) {
  await page.evaluate(async (frameCount) => {
    await Array.from({ length: frameCount }).reduce<Promise<void>>(async (previous) => {
      await previous
      await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
    }, Promise.resolve())
  }, frames)
}

export async function stopWorkbookScrollPerf(page: Page) {
  return await page.evaluate(() => {
    return (
      (
        window as Window & {
          __biligScrollPerf?: {
            stopSampling?: () => {
              workload: string
              fixture: { id: string; materializedCellCount: number; sheetName: string } | null
              samples: { frameMs: number[]; inputToDrawMs: number[]; longTasksMs: number[] }
              summary: {
                frameMs: { min: number; median: number; p95: number; p99: number; max: number }
                inputToDrawMs: { min: number; median: number; p95: number; p99: number; max: number }
                longTasksMs: { min: number; median: number; p95: number; p99: number; max: number }
              }
              counters: {
                viewportSubscriptions: number
                fullPatches: number
                damagePatches: number
                damageCells: number
                rendererTileInterestBatches: number
                rendererTileExactHits: number
                rendererTileStaleHits: number
                rendererTileMisses: number
                rendererVisibleDirtyTiles: number
                rendererWarmDirtyTiles: number
                visibleWindowChanges: number
                headerPaneBuilds: number
                reactCommits: number
                canvasSurfaceMounts: number
                domSurfaceMounts: number
                canvasPaints: Record<string, number>
                surfaceCommits: Record<string, number>
                typeGpuAtlasUploadBytes: number
                typeGpuAtlasDirtyPages: number
                typeGpuAtlasDirtyPageUploadBytes: number
                typeGpuBufferAllocationBytes: number
                typeGpuBufferAllocations: number
                typeGpuConfigures: number
                typeGpuDrawCalls: number
                typeGpuPaneDraws: number
                typeGpuSubmits: number
                typeGpuSurfaceResizes: number
                typeGpuTileMisses: number
                typeGpuUniformWriteBytes: number
                typeGpuVertexUploadBytes: number
              }
            } | null
          }
        }
      ).__biligScrollPerf?.stopSampling?.() ?? null
    )
  })
}

export async function performHorizontalGridBrowse(page: Page, input: { distancePx: number; steps?: number }) {
  await performGridBrowse(page, { deltaX: input.distancePx, deltaY: 0, ...(input.steps ? { steps: input.steps } : {}) })
}

export async function performVerticalGridBrowse(page: Page, input: { distancePx: number; steps?: number }) {
  await performGridBrowse(page, { deltaX: 0, deltaY: input.distancePx, ...(input.steps ? { steps: input.steps } : {}) })
}

export async function performDiagonalGridBrowse(page: Page, input: { deltaX: number; deltaY: number; steps?: number }) {
  await performGridBrowse(page, input)
}

async function performGridBrowse(page: Page, input: { deltaX: number; deltaY: number; steps?: number }) {
  await page.getByTestId('grid-scroll-viewport').evaluate(
    async (element, options) => {
      if (!(element instanceof HTMLDivElement)) {
        throw new Error('grid scroll viewport is not an HTMLDivElement')
      }
      const viewport = element
      const steps = Math.max(1, options.steps ?? 120)
      const startLeft = viewport.scrollLeft
      const startTop = viewport.scrollTop
      const advance = async (step: number): Promise<void> => {
        if (step > steps) {
          return
        }
        viewport.scrollLeft = startLeft + (options.deltaX * step) / steps
        viewport.scrollTop = startTop + (options.deltaY * step) / steps
        viewport.dispatchEvent(new Event('scroll'))
        await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
        await advance(step + 1)
      }
      await advance(1)
    },
    { deltaX: input.deltaX, deltaY: input.deltaY, ...(input.steps ? { steps: input.steps } : {}) },
  )
}

export async function runToolbarSyncActions(
  page: Page,
  mirrorPage: Page,
  actions: readonly ToolbarSyncAction[],
  testInfo: TestInfo,
  index = 0,
): Promise<void> {
  const action = actions[index]
  if (!action) {
    return
  }

  await action.apply(page)
  await selectToolbarActionRange(page)
  await selectToolbarActionRange(mirrorPage)
  const elapsed = await expectMatchingGridRangeScreenshots(page, mirrorPage, action.label, testInfo, 1, 1, 2, 2, 1_500)
  expect(elapsed).toBeLessThanOrEqual(1_500)
  await runToolbarSyncActions(page, mirrorPage, actions, testInfo, index + 1)
}
