import type { Viewport } from '@bilig/protocol'
import { resolveRowOffset, type GridMetrics } from '../gridMetrics.js'
import { resolveColumnOffset } from '../workbookGridViewport.js'
import type { WorkbookRenderPaneState } from '../renderer-v2/pane-scene-types.js'
import { getPaneFrame, resolvePaneLayout } from '../renderer-v2/pane-layout.js'
import {
  createGridTileKeyV2,
  GRID_SCENE_PACKET_V2_MAGIC,
  GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_VERSION,
  type GridScenePacketV2,
  type GridSceneTextRun,
} from '../renderer-v2/scene-packet-v2.js'
import type { GridRenderTile } from './render-tile-source.js'

interface AxisPlacementInput {
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly gridMetrics: GridMetrics
}

export function buildFixedRenderTileDataPaneStates(input: {
  readonly tiles: readonly GridRenderTile[]
  readonly sheetName: string
  readonly residentViewport: Viewport
  readonly visibleViewport: Viewport
  readonly freezeRows: number
  readonly freezeCols: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly hostWidth: number
  readonly hostHeight: number
  readonly gridMetrics: GridMetrics
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly columnWidths: Readonly<Record<number, number>>
  readonly rowHeights: Readonly<Record<number, number>>
}): readonly WorkbookRenderPaneState[] {
  const layout = resolvePaneLayout({
    frozenColumnWidth: input.frozenColumnWidth,
    frozenRowHeight: input.frozenRowHeight,
    headerHeight: input.gridMetrics.headerHeight,
    hostHeight: input.hostHeight,
    hostWidth: input.hostWidth,
    rowMarkerWidth: input.gridMetrics.rowMarkerWidth,
  })
  const bodyFrame = getPaneFrame(layout, 'body')
  if (bodyFrame.width <= 0 || bodyFrame.height <= 0) {
    return []
  }

  const packets = input.tiles.map((tile) =>
    createGridScenePacketV2FromRenderTile({
      columnWidths: input.columnWidths,
      gridMetrics: input.gridMetrics,
      rowHeights: input.rowHeights,
      sheetName: input.sheetName,
      sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
      sortedRowHeightOverrides: input.sortedRowHeightOverrides,
      tile,
    }),
  )
  const bodyTiles = packets.filter((packet) => intersects(packet.viewport, input.residentViewport))
  if (bodyTiles.length === 0) {
    return []
  }

  const bodyReference = bodyTiles.toSorted(
    (left, right) => left.viewport.rowStart - right.viewport.rowStart || left.viewport.colStart - right.viewport.colStart,
  )[0]!
  const residentSurfaceSize = resolveViewportSurfaceSize(input.residentViewport, input)
  const panes: WorkbookRenderPaneState[] = []

  bodyTiles.forEach((packet, index) => {
    panes.push(
      buildPlacementPane({
        frame: bodyFrame,
        id: index === 0 ? 'body' : `body:${packet.key.rowTile}:${packet.key.colTile}`,
        packet,
        reference: bodyReference.viewport,
        scrollAxes: { x: true, y: true },
        surfaceSize: index === 0 ? residentSurfaceSize : packet.surfaceSize,
        ...input,
      }),
    )
  })

  if (input.freezeRows > 0 && input.frozenRowHeight > 0) {
    const topViewport = {
      rowStart: 0,
      rowEnd: Math.max(0, input.freezeRows - 1),
      colStart: input.residentViewport.colStart,
      colEnd: input.residentViewport.colEnd,
    }
    const frame = getPaneFrame(layout, 'top')
    packets
      .filter((packet) => intersects(packet.viewport, topViewport))
      .forEach((packet) => {
        panes.push(
          buildPlacementPane({
            frame,
            id: `top:${packet.key.rowTile}:${packet.key.colTile}`,
            packet,
            reference: bodyReference.viewport,
            scrollAxes: { x: true, y: false },
            ...input,
          }),
        )
      })
  }

  if (input.freezeCols > 0 && input.frozenColumnWidth > 0) {
    const leftViewport = {
      rowStart: input.residentViewport.rowStart,
      rowEnd: input.residentViewport.rowEnd,
      colStart: 0,
      colEnd: Math.max(0, input.freezeCols - 1),
    }
    const frame = getPaneFrame(layout, 'left')
    packets
      .filter((packet) => intersects(packet.viewport, leftViewport))
      .forEach((packet) => {
        panes.push(
          buildPlacementPane({
            frame,
            id: `left:${packet.key.rowTile}:${packet.key.colTile}`,
            packet,
            reference: bodyReference.viewport,
            scrollAxes: { x: false, y: true },
            ...input,
          }),
        )
      })
  }

  if (input.freezeRows > 0 && input.freezeCols > 0 && input.frozenColumnWidth > 0 && input.frozenRowHeight > 0) {
    const cornerViewport = {
      rowStart: 0,
      rowEnd: Math.max(0, input.freezeRows - 1),
      colStart: 0,
      colEnd: Math.max(0, input.freezeCols - 1),
    }
    const frame = getPaneFrame(layout, 'corner')
    packets
      .filter((packet) => intersects(packet.viewport, cornerViewport))
      .forEach((packet) => {
        panes.push(
          buildPlacementPane({
            frame,
            id: `corner:${packet.key.rowTile}:${packet.key.colTile}`,
            packet,
            reference: bodyReference.viewport,
            scrollAxes: { x: false, y: false },
            ...input,
          }),
        )
      })
  }

  return panes
}

function buildPlacementPane(
  input: AxisPlacementInput & {
    readonly frame: WorkbookRenderPaneState['frame']
    readonly id: string
    readonly packet: GridScenePacketV2
    readonly reference: Viewport
    readonly scrollAxes: WorkbookRenderPaneState['scrollAxes']
    readonly surfaceSize?: WorkbookRenderPaneState['surfaceSize'] | undefined
  },
): WorkbookRenderPaneState {
  const packetX = resolveColumnOffset(input.packet.viewport.colStart, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth)
  const packetY = resolveRowOffset(input.packet.viewport.rowStart, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight)
  const referenceX = resolveColumnOffset(input.reference.colStart, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth)
  const referenceY = resolveRowOffset(input.reference.rowStart, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight)
  return {
    contentOffset: {
      x: input.scrollAxes.x ? packetX - referenceX : packetX,
      y: input.scrollAxes.y ? packetY - referenceY : packetY,
    },
    frame: input.frame,
    generation: input.packet.generation,
    packedScene: input.packet,
    paneId: input.id,
    scrollAxes: input.scrollAxes,
    surfaceSize: input.surfaceSize ?? input.packet.surfaceSize,
    viewport: input.packet.viewport,
  }
}

function createGridScenePacketV2FromRenderTile(
  input: AxisPlacementInput & {
    readonly sheetName: string
    readonly tile: GridRenderTile
    readonly columnWidths: Readonly<Record<number, number>>
    readonly rowHeights: Readonly<Record<number, number>>
  },
): GridScenePacketV2 {
  const surfaceSize = resolveViewportSurfaceSize(input.tile.bounds, input)
  return {
    borderRectCount: 0,
    cameraSeq: input.tile.lastCameraSeq,
    fillRectCount: input.tile.rectCount,
    generatedAt: input.tile.lastBatchId,
    generation: input.tile.lastBatchId,
    key: createGridTileKeyV2({
      axisVersionX: input.tile.version.axisX,
      axisVersionY: input.tile.version.axisY,
      colTile: input.tile.coord.colTile,
      dprBucket: input.tile.coord.dprBucket,
      freezeVersion: input.tile.version.freeze,
      paneId: 'body',
      rowTile: input.tile.coord.rowTile,
      selectionIndependentVersion: 0,
      sheetName: input.sheetName,
      styleVersion: input.tile.version.styles,
      textEpoch: input.tile.version.text,
      valueVersion: input.tile.version.values,
      viewport: input.tile.bounds,
    }),
    magic: GRID_SCENE_PACKET_V2_MAGIC,
    paneId: 'body',
    rectCount: input.tile.rectCount,
    rectInstances: input.tile.rectInstances,
    rects: buildRectSummary(input.tile),
    rectSignature: `render-tile:${input.tile.tileId}:rect:${input.tile.version.values}:${input.tile.version.styles}:${input.tile.rectCount}`,
    requestSeq: input.tile.lastBatchId,
    sheetName: input.sheetName,
    surfaceSize,
    textCount: input.tile.textCount,
    textMetrics: input.tile.textMetrics,
    textRuns: input.tile.textRuns.map(mapTextRun),
    textSignature: `render-tile:${input.tile.tileId}:text:${input.tile.version.text}:${input.tile.textCount}`,
    version: GRID_SCENE_PACKET_V2_VERSION,
    viewport: input.tile.bounds,
  }
}

function resolveViewportSurfaceSize(
  viewport: Viewport,
  input: AxisPlacementInput & {
    readonly columnWidths: Readonly<Record<number, number>>
    readonly rowHeights: Readonly<Record<number, number>>
  },
): { readonly width: number; readonly height: number } {
  return {
    width:
      resolveColumnOffset(viewport.colEnd + 1, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth) -
      resolveColumnOffset(viewport.colStart, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth),
    height:
      resolveRowOffset(viewport.rowEnd + 1, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight) -
      resolveRowOffset(viewport.rowStart, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight),
  }
}

function buildRectSummary(tile: GridRenderTile): Float32Array {
  const rects = new Float32Array(Math.max(1, tile.rectCount) * GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT)
  for (let index = 0; index < tile.rectCount; index += 1) {
    const rectOffset = index * GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT
    const instanceOffset = index * 20
    rects[rectOffset + 0] = tile.rectInstances[instanceOffset + 0] ?? 0
    rects[rectOffset + 1] = tile.rectInstances[instanceOffset + 1] ?? 0
    rects[rectOffset + 2] = tile.rectInstances[instanceOffset + 2] ?? 0
    rects[rectOffset + 3] = tile.rectInstances[instanceOffset + 3] ?? 0
    rects[rectOffset + 4] = tile.rectInstances[instanceOffset + 4] ?? 0
    rects[rectOffset + 5] = tile.rectInstances[instanceOffset + 5] ?? 0
    rects[rectOffset + 6] = tile.rectInstances[instanceOffset + 6] ?? 0
    rects[rectOffset + 7] = tile.rectInstances[instanceOffset + 7] ?? 0
  }
  return rects
}

function mapTextRun(run: GridRenderTile['textRuns'][number]): GridSceneTextRun {
  return {
    align: 'left',
    clipHeight: run.clipHeight,
    clipWidth: run.clipWidth,
    clipX: run.clipX,
    clipY: run.clipY,
    color: run.color,
    font: run.font,
    fontSize: run.fontSize,
    height: run.height,
    strike: run.strike,
    text: run.text,
    underline: run.underline,
    width: run.width,
    wrap: false,
    x: run.x,
    y: run.y,
  }
}

function intersects(left: Viewport, right: Viewport): boolean {
  return left.rowStart <= right.rowEnd && left.rowEnd >= right.rowStart && left.colStart <= right.colEnd && left.colEnd >= right.colStart
}
