import { writeFile } from 'node:fs/promises'
import { expect, test, type Page, type TestInfo } from '@playwright/test'
import { ISOLATED_WORKBOOK_PANE_RENDERER_PATH } from '../../apps/web/src/root-route.js'
import {
  clickProductCell,
  gotoWorkbookShell,
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
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
}

test('isolated workbook pane renderer draws grid content through typegpu', async ({ page }, testInfo) => {
  await page.setViewportSize({ width: 640, height: 480 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, ISOLATED_WORKBOOK_PANE_RENDERER_PATH)
  await page.waitForSelector('[data-testid="isolated-pane-renderer-route"]', { timeout: 15_000 })
  await page.waitForSelector('[data-testid="grid-pane-renderer"]', { timeout: 15_000 })
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
  expect(summary?.darkPixelCounts.header).toBeGreaterThan(15)
  expect(summary?.darkPixelCounts.body).toBeGreaterThan(40)
  expect(summary?.darkPixelCounts.number).toBeGreaterThan(40)

  await saveReadbackArtifact(page, testInfo, 'isolated-pane-renderer-readback.png', 'isolated-pane-renderer-readback')
})

test('main workbook shell grid renders and updates through typegpu', async ({ page }, testInfo) => {
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
  await gotoWorkbookShell(page)
  await waitForWorkbookReady(page)
  await page.waitForSelector('[data-testid="grid-pane-renderer"]', { timeout: 15_000 })
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
      { name: 'bodyBlank', x: PRODUCT_ROW_MARKER_WIDTH + 14, y: PRODUCT_HEADER_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2) },
    ],
    regions: [
      { name: 'columnHeaderText', x0: 176, y0: 4, x1: 228, y1: 18 },
      { name: 'rowHeaderText', x0: 10, y0: 48, x1: 36, y1: 66 },
    ],
  } as const

  const initialReadback = await waitForReadback(page, initialProbe, (result) => {
    return result.darkPixelCounts.columnHeaderText > 10 && result.darkPixelCounts.rowHeaderText > 5
  })

  expect(initialReadback.hasGpu).toBe(true)
  expect(initialReadback.width).toBeGreaterThan(400)
  expect(initialReadback.height).toBeGreaterThan(250)
  expect(initialReadback.points.bodyBlank).toMatchObject({ r: 0, g: 0, b: 0, a: 0 })
  expect(initialReadback.darkPixelCounts.columnHeaderText).toBeGreaterThan(10)
  expect(initialReadback.darkPixelCounts.rowHeaderText).toBeGreaterThan(5)

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

  const valueReadback = await waitForReadback(page, valueProbe, (result) => {
    return result.darkPixelCounts.c4ValueText > 0
  })

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
    regions: [],
  } as const

  const rangeReadback = await waitForReadback(page, rangeProbe, (result) => {
    return result.points.rangeFill.a > 0
  })

  expect(rangeReadback.points.rangeFill).toMatchObject(premultiplyReadbackPoint({ r: 33, g: 86, b: 58, a: 20 }))
  expect(rangeReadback.points.rangeBorder.a).toBeGreaterThan(150)
  expect(rangeReadback.points.topHeaderSelectionFill.a).toBeGreaterThan(0)

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-readback.png', 'main-workbook-grid-readback')
})

async function installTypeGpuReadbackHarness(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const globalWindow = window as Window & {
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
      }
    }

    const readbackState = {
      bgra: new Uint8Array(0),
      bytesPerRow: 0,
      hasGpu: Boolean(navigator.gpu),
      height: 0,
      ready: false,
      sequence: 0,
      width: 0,
    }

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

    globalWindow.__biligGpuReadbackInspector = {
      countDarkPixels(regions) {
        return Object.fromEntries(
          regions.map((region) => [region.name, countDarkPixels(region.x0, region.y0, region.x1, region.y1, region.threshold)]),
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

    if (globalWindow.__biligTypeGpuHarnessInstalled) {
      return
    }
    globalWindow.__biligTypeGpuHarnessInstalled = true
    globalWindow.__biligGpuReadback = {
      ready: false,
      hasGpu: Boolean(navigator.gpu),
      width: 0,
      height: 0,
      sequence: 0,
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

    if (!navigator.gpu) {
      return
    }

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
        let readbackPending = false

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
          if (!readbackPending && targetTexture && targetWidth > 0 && targetHeight > 0) {
            readbackPending = true
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
                const bgra = new Uint8Array(mapped)
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
                } catch {}
                buffer.destroy()
                readbackPending = false
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
    }
  }, input)

  expect(result.ready).toBe(true)
  return result
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
  await expect
    .poll(
      async () => {
        lastResult = await inspectGpuReadback(page, input)
        return lastResult.ready && lastResult.hasGpu && predicate(lastResult)
      },
      { timeout: 15_000 },
    )
    .toBe(true)
  if (!lastResult) {
    throw new Error('expected readback result')
  }
  return lastResult
}

function premultiplyReadbackPoint(point: ReadbackPoint): ReadbackPoint {
  const alpha = point.a / 255
  return {
    r: Math.round(point.r * alpha),
    g: Math.round(point.g * alpha),
    b: Math.round(point.b * alpha),
    a: point.a,
  }
}

async function waitForReadbackSequence(page: Page, previousSequence: number): Promise<void> {
  await page.waitForFunction(
    (sequence) => {
      const inspector = (window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } })
        .__biligGpuReadbackInspector
      return (inspector?.getSequence() ?? 0) > sequence
    },
    previousSequence,
    { timeout: 15_000 },
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
