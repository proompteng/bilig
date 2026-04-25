import { useEffect, useMemo, useState, type ComponentProps } from 'react'
import { WorkbookPaneRendererV2, packGridScenePacketV2 } from '@bilig/grid'

const ROW_MARKER_WIDTH = 46
const HEADER_HEIGHT = 24
const COLUMN_WIDTH = 104
const ROW_HEIGHT = 22

const HEADER_FILL = gpuColor(243, 242, 238)
const GRID_LINE = gpuColor(236, 233, 225)
const SURFACE = gpuColor(255, 255, 255)
const ACCENT_SOFT = gpuColor(33, 86, 58, 0.12)
const ACCENT = gpuColor(33, 86, 58)
const HOVER = gpuColor(82, 96, 109, 0.08)
const VALUE_FILL = gpuColor(249, 239, 215, 0.85)
const TEXT_PRIMARY = '#1f2933'
const TEXT_MUTED = '#52606d'
const TEXT_ACCENT = '#163f29'

type RendererPane = ComponentProps<typeof WorkbookPaneRendererV2>['panes'][number]
type RendererPanes = ComponentProps<typeof WorkbookPaneRendererV2>['panes']
type RendererGpuScene = RendererPane['gpuScene']
type RendererGpuRect = RendererGpuScene['fillRects'][number]
type RendererTextScene = RendererPane['textScene']
type RendererTextItem = RendererTextScene['items'][number]
type RendererPaneId = 'corner' | 'top-body' | 'left-body' | 'body'

interface HostSize {
  readonly width: number
  readonly height: number
}

export function IsolatedWorkbookPaneRendererRoute() {
  const [host, setHost] = useState<HTMLDivElement | null>(null)
  const [hostSize, setHostSize] = useState<HostSize>({ width: 0, height: 0 })

  useEffect(() => {
    if (!host) {
      setHostSize({ width: 0, height: 0 })
      return
    }

    const update = () => {
      setHostSize({
        width: Math.max(0, Math.floor(host.clientWidth)),
        height: Math.max(0, Math.floor(host.clientHeight)),
      })
    }

    update()

    if (typeof ResizeObserver === 'undefined') {
      const frame = window.requestAnimationFrame(update)
      return () => {
        window.cancelAnimationFrame(frame)
      }
    }

    const observer = new ResizeObserver(update)
    observer.observe(host)
    return () => {
      observer.disconnect()
    }
  }, [host])

  const panes = useMemo(() => buildIsolatedRendererPanes(hostSize), [hostSize])

  return (
    <div className="h-dvh w-screen overflow-hidden bg-(--wb-surface)">
      <div className="relative h-full w-full" data-testid="isolated-pane-renderer-route" ref={setHost}>
        {host ? <WorkbookPaneRendererV2 active geometry={null} host={host} panes={panes} /> : null}
      </div>
    </div>
  )
}

function buildIsolatedRendererPanes(hostSize: HostSize): RendererPanes {
  const bodyWidth = Math.max(0, hostSize.width - ROW_MARKER_WIDTH)
  const bodyHeight = Math.max(0, hostSize.height - HEADER_HEIGHT)

  return [
    createRendererPane({
      generation: 1,
      paneId: 'corner',
      frame: { x: 0, y: 0, width: ROW_MARKER_WIDTH, height: HEADER_HEIGHT },
      surfaceSize: { width: ROW_MARKER_WIDTH, height: HEADER_HEIGHT },
      contentOffset: { x: 0, y: 0 },
      scrollAxes: { x: false, y: false },
      viewport: { colStart: 0, colEnd: 0, rowStart: 0, rowEnd: 0 },
      gpuScene: {
        fillRects: [{ x: 0, y: 0, width: ROW_MARKER_WIDTH, height: HEADER_HEIGHT, color: HEADER_FILL }],
        borderRects: [
          { x: ROW_MARKER_WIDTH - 1, y: 0, width: 1, height: HEADER_HEIGHT, color: GRID_LINE },
          { x: 0, y: HEADER_HEIGHT - 1, width: ROW_MARKER_WIDTH, height: 1, color: GRID_LINE },
        ],
      },
      textScene: { items: [] },
    }),
    createRendererPane({
      generation: 1,
      paneId: 'top-body',
      frame: { x: ROW_MARKER_WIDTH, y: 0, width: bodyWidth, height: HEADER_HEIGHT },
      surfaceSize: { width: bodyWidth, height: HEADER_HEIGHT },
      contentOffset: { x: 0, y: 0 },
      scrollAxes: { x: true, y: false },
      viewport: { colStart: 0, colEnd: Math.max(0, Math.ceil(bodyWidth / COLUMN_WIDTH) - 1), rowStart: 0, rowEnd: 0 },
      gpuScene: buildColumnHeaderGpuScene(bodyWidth),
      textScene: buildColumnHeaderTextScene(bodyWidth),
    }),
    createRendererPane({
      generation: 1,
      paneId: 'left-body',
      frame: { x: 0, y: HEADER_HEIGHT, width: ROW_MARKER_WIDTH, height: bodyHeight },
      surfaceSize: { width: ROW_MARKER_WIDTH, height: bodyHeight },
      contentOffset: { x: 0, y: 0 },
      scrollAxes: { x: false, y: true },
      viewport: { colStart: 0, colEnd: 0, rowStart: 0, rowEnd: Math.max(0, Math.ceil(bodyHeight / ROW_HEIGHT) - 1) },
      gpuScene: buildRowHeaderGpuScene(bodyHeight),
      textScene: buildRowHeaderTextScene(bodyHeight),
    }),
    createRendererPane({
      generation: 1,
      paneId: 'body',
      frame: { x: ROW_MARKER_WIDTH, y: HEADER_HEIGHT, width: bodyWidth, height: bodyHeight },
      surfaceSize: { width: bodyWidth, height: bodyHeight },
      contentOffset: { x: 0, y: 0 },
      scrollAxes: { x: true, y: true },
      viewport: {
        colStart: 0,
        colEnd: Math.max(0, Math.ceil(bodyWidth / COLUMN_WIDTH) - 1),
        rowStart: 0,
        rowEnd: Math.max(0, Math.ceil(bodyHeight / ROW_HEIGHT) - 1),
      },
      gpuScene: buildBodyGpuScene(bodyWidth, bodyHeight),
      textScene: buildBodyTextScene(bodyWidth, bodyHeight),
    }),
  ]
}

function createRendererPane(input: Omit<RendererPane, 'packedScene'> & { readonly paneId: RendererPaneId }): RendererPane {
  return {
    ...input,
    packedScene: packGridScenePacketV2({
      generation: input.generation,
      gpuScene: input.gpuScene,
      paneId: input.paneId,
      sheetName: 'Sheet1',
      surfaceSize: input.surfaceSize,
      textScene: input.textScene,
      viewport: input.viewport ?? { colStart: 0, colEnd: 0, rowStart: 0, rowEnd: 0 },
    }),
  }
}

function buildColumnHeaderGpuScene(width: number): RendererGpuScene {
  const columnCount = Math.max(1, Math.ceil(width / COLUMN_WIDTH))
  const fillRects: RendererGpuRect[] = []
  const borderRects: RendererGpuRect[] = []

  for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
    const left = columnIndex * COLUMN_WIDTH
    const cellWidth = Math.max(0, Math.min(COLUMN_WIDTH, width - left))
    if (cellWidth <= 0) {
      continue
    }
    fillRects.push({
      x: left,
      y: 0,
      width: cellWidth,
      height: HEADER_HEIGHT,
      color: columnIndex === 1 ? ACCENT_SOFT : HEADER_FILL,
    })
    borderRects.push({
      x: left + cellWidth - 1,
      y: 0,
      width: 1,
      height: HEADER_HEIGHT,
      color: GRID_LINE,
    })
  }

  borderRects.push({ x: 0, y: HEADER_HEIGHT - 1, width, height: 1, color: GRID_LINE })

  return { fillRects, borderRects }
}

function buildColumnHeaderTextScene(width: number): RendererTextScene {
  const columnCount = Math.max(1, Math.ceil(width / COLUMN_WIDTH))
  const items: RendererTextItem[] = []

  for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
    const left = columnIndex * COLUMN_WIDTH
    const cellWidth = Math.max(0, Math.min(COLUMN_WIDTH, width - left))
    if (cellWidth <= 0) {
      continue
    }
    items.push(
      createTextItem({
        align: 'center',
        color: columnIndex === 1 ? 'var(--wb-accent)' : 'var(--wb-text-muted)',
        height: HEADER_HEIGHT,
        text: indexToColumnLabel(columnIndex),
        width: cellWidth,
        x: left,
        y: 0,
      }),
    )
  }

  return { items }
}

function buildRowHeaderGpuScene(height: number): RendererGpuScene {
  const rowCount = Math.max(1, Math.ceil(height / ROW_HEIGHT))
  const fillRects: RendererGpuRect[] = []
  const borderRects: RendererGpuRect[] = []

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const top = rowIndex * ROW_HEIGHT
    const cellHeight = Math.max(0, Math.min(ROW_HEIGHT, height - top))
    if (cellHeight <= 0) {
      continue
    }
    fillRects.push({
      x: 0,
      y: top,
      width: ROW_MARKER_WIDTH,
      height: cellHeight,
      color: rowIndex === 2 ? ACCENT_SOFT : HEADER_FILL,
    })
    borderRects.push({
      x: 0,
      y: top + cellHeight - 1,
      width: ROW_MARKER_WIDTH,
      height: 1,
      color: GRID_LINE,
    })
  }

  borderRects.push({ x: ROW_MARKER_WIDTH - 1, y: 0, width: 1, height, color: GRID_LINE })

  return { fillRects, borderRects }
}

function buildRowHeaderTextScene(height: number): RendererTextScene {
  const rowCount = Math.max(1, Math.ceil(height / ROW_HEIGHT))
  const items: RendererTextItem[] = []

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const top = rowIndex * ROW_HEIGHT
    const cellHeight = Math.max(0, Math.min(ROW_HEIGHT, height - top))
    if (cellHeight <= 0) {
      continue
    }
    items.push(
      createTextItem({
        align: 'right',
        color: rowIndex === 2 ? 'var(--wb-accent)' : 'var(--wb-text-muted)',
        height: cellHeight,
        text: String(rowIndex + 1),
        width: ROW_MARKER_WIDTH,
        x: 0,
        y: top,
      }),
    )
  }

  return { items }
}

function buildBodyGpuScene(width: number, height: number): RendererGpuScene {
  const columnCount = Math.max(1, Math.ceil(width / COLUMN_WIDTH))
  const rowCount = Math.max(1, Math.ceil(height / ROW_HEIGHT))
  const fillRects: RendererGpuRect[] = [
    { x: 0, y: 0, width, height, color: SURFACE },
    {
      x: COLUMN_WIDTH + 1,
      y: ROW_HEIGHT * 2 + 1,
      width: Math.max(COLUMN_WIDTH * 2 - 2, 0),
      height: Math.max(ROW_HEIGHT * 3 - 2, 0),
      color: ACCENT_SOFT,
    },
    {
      x: COLUMN_WIDTH * 4 + 1,
      y: ROW_HEIGHT * 5 + 1,
      width: Math.max(COLUMN_WIDTH - 2, 0),
      height: Math.max(ROW_HEIGHT - 2, 0),
      color: VALUE_FILL,
    },
    {
      x: COLUMN_WIDTH * 6 + 1,
      y: ROW_HEIGHT * 8 + 1,
      width: Math.max(COLUMN_WIDTH - 2, 0),
      height: Math.max(ROW_HEIGHT - 2, 0),
      color: HOVER,
    },
  ]
  const borderRects: RendererGpuRect[] = []

  for (let columnIndex = 0; columnIndex <= columnCount; columnIndex += 1) {
    const x = Math.min(columnIndex * COLUMN_WIDTH, width - 1)
    borderRects.push({ x, y: 0, width: 1, height, color: GRID_LINE })
  }

  for (let rowIndex = 0; rowIndex <= rowCount; rowIndex += 1) {
    const y = Math.min(rowIndex * ROW_HEIGHT, height - 1)
    borderRects.push({ x: 0, y, width, height: 1, color: GRID_LINE })
  }

  borderRects.push(
    { x: COLUMN_WIDTH, y: ROW_HEIGHT * 2, width: COLUMN_WIDTH * 2, height: 1, color: ACCENT },
    { x: COLUMN_WIDTH, y: ROW_HEIGHT * 5 - 1, width: COLUMN_WIDTH * 2, height: 1, color: ACCENT },
    { x: COLUMN_WIDTH, y: ROW_HEIGHT * 2, width: 1, height: ROW_HEIGHT * 3, color: ACCENT },
    { x: COLUMN_WIDTH * 3 - 1, y: ROW_HEIGHT * 2, width: 1, height: ROW_HEIGHT * 3, color: ACCENT },
  )

  return { fillRects, borderRects }
}

function buildBodyTextScene(width: number, height: number): RendererTextScene {
  const columnCount = Math.max(1, Math.ceil(width / COLUMN_WIDTH))
  const rowCount = Math.max(1, Math.ceil(height / ROW_HEIGHT))
  const items: RendererTextItem[] = []

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const left = columnIndex * COLUMN_WIDTH
      const top = rowIndex * ROW_HEIGHT
      const cellWidth = Math.max(0, Math.min(COLUMN_WIDTH, width - left))
      const cellHeight = Math.max(0, Math.min(ROW_HEIGHT, height - top))
      if (cellWidth <= 0 || cellHeight <= 0) {
        continue
      }

      const text = resolveBodyCellText(columnIndex, rowIndex)
      if (!text) {
        continue
      }

      items.push(
        createTextItem({
          align: columnIndex >= 4 ? 'right' : 'left',
          color: TEXT_PRIMARY,
          height: cellHeight,
          text,
          width: cellWidth,
          x: left,
          y: top,
        }),
      )
    }
  }

  appendBodyTextVariants(items, width, height)

  return { items }
}

function appendBodyTextVariants(items: RendererTextItem[], width: number, height: number): void {
  const variantColumnLeft = COLUMN_WIDTH * 7
  const variantWidth = Math.max(0, Math.min(COLUMN_WIDTH * 2, width - variantColumnLeft))
  if (variantWidth <= 0 || height < ROW_HEIGHT * 9) {
    return
  }

  items.push(
    createTextItem({
      align: 'left',
      clipInsetLeft: 16,
      color: TEXT_ACCENT,
      font: 'italic 700 12px var(--font-sans)',
      fontSize: 12,
      height: ROW_HEIGHT,
      text: 'Clipped italic sample',
      width: variantWidth,
      x: variantColumnLeft,
      y: ROW_HEIGHT,
    }),
  )

  items.push(
    createTextItem({
      align: 'left',
      color: TEXT_PRIMARY,
      font: '400 12px var(--font-sans)',
      fontSize: 12,
      height: ROW_HEIGHT * 2,
      text: 'Wrap sample stays legible across multiple lines.',
      width: variantWidth,
      wrap: true,
      x: variantColumnLeft,
      y: ROW_HEIGHT * 2,
    }),
  )

  items.push(
    createTextItem({
      align: 'center',
      color: TEXT_PRIMARY,
      height: ROW_HEIGHT,
      text: 'Underline',
      underline: true,
      width: variantWidth,
      x: variantColumnLeft,
      y: ROW_HEIGHT * 5,
    }),
  )

  items.push(
    createTextItem({
      align: 'right',
      color: TEXT_MUTED,
      height: ROW_HEIGHT,
      strike: true,
      text: 'Strike',
      width: variantWidth,
      x: variantColumnLeft,
      y: ROW_HEIGHT * 8,
    }),
  )
}

function resolveBodyCellText(columnIndex: number, rowIndex: number): string {
  if (columnIndex === 7 && (rowIndex === 1 || rowIndex === 2 || rowIndex === 5 || rowIndex === 8)) {
    return ''
  }
  if (columnIndex === 8 && rowIndex === 2) {
    return ''
  }
  if (rowIndex === 0) {
    return ['Region', 'Owner', 'Status', 'ETA', 'Units', 'Delta', 'Risk'][columnIndex] ?? ''
  }
  if (columnIndex === 0) {
    return ['North', 'South', 'West', 'Central', 'Forecast', 'Backlog'][rowIndex - 1] ?? ''
  }
  if (columnIndex === 1) {
    return ['Avery', 'Kai', 'Mina', 'Jules', 'Ops', 'Queue'][rowIndex - 1] ?? ''
  }
  if (columnIndex === 2) {
    return ['Live', 'Review', 'Live', 'Blocked', 'Pending', 'Ready'][rowIndex - 1] ?? ''
  }
  if (columnIndex === 3) {
    return ['2d', '4d', '1d', '6d', '5d', '1d'][rowIndex - 1] ?? ''
  }
  if (rowIndex === 5 && columnIndex === 4) {
    return '4,280'
  }
  if (rowIndex === 8 && columnIndex === 6) {
    return 'Watch'
  }
  if (columnIndex >= 4) {
    return String((rowIndex + 1) * (columnIndex + 3) * 12)
  }
  return ''
}

function createTextItem(input: {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly text: string
  readonly color: string
  readonly align: RendererTextItem['align']
  readonly clipInsetTop?: number
  readonly clipInsetRight?: number
  readonly clipInsetBottom?: number
  readonly clipInsetLeft?: number
  readonly wrap?: boolean
  readonly font?: string
  readonly fontSize?: number
  readonly underline?: boolean
  readonly strike?: boolean
}): RendererTextItem {
  return {
    x: input.x,
    y: input.y,
    width: input.width,
    height: input.height,
    clipInsetTop: input.clipInsetTop ?? 0,
    clipInsetRight: input.clipInsetRight ?? 0,
    clipInsetBottom: input.clipInsetBottom ?? 0,
    clipInsetLeft: input.clipInsetLeft ?? 0,
    text: input.text,
    align: input.align,
    wrap: input.wrap ?? false,
    color: input.color,
    font: input.font ?? '500 12px var(--font-sans)',
    fontSize: input.fontSize ?? 12,
    underline: input.underline ?? false,
    strike: input.strike ?? false,
  }
}

function indexToColumnLabel(index: number): string {
  let value = index
  let label = ''
  do {
    label = String.fromCharCode(65 + (value % 26)) + label
    value = Math.floor(value / 26) - 1
  } while (value >= 0)
  return label
}

function gpuColor(red: number, green: number, blue: number, alpha = 1): RendererGpuRect['color'] {
  return {
    r: red / 255,
    g: green / 255,
    b: blue / 255,
    a: alpha,
  }
}
