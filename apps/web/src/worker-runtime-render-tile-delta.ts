import type { GridEngineLike } from '@bilig/grid'
import type { RecalcMetrics, WorkbookAxisEntrySnapshot } from '@bilig/protocol'
import type { RenderTileDeltaBatch, RenderTileDeltaSubscription, RenderTileReplaceMutation } from '@bilig/worker-transport'
import { getGridMetrics } from '../../../packages/grid/src/gridMetrics.js'
import { materializeGridRenderTileV3 } from '../../../packages/grid/src/renderer-v3/grid-tile-materializer.js'
import type { GridRenderTile } from '../../../packages/grid/src/renderer-v3/render-tile-source.js'
import { buildFreezeVersion, buildRenderedAxisState } from './worker-runtime-render-axis.js'
import { listViewportTileBounds } from './worker-viewport-tile-store.js'

interface RenderTileDeltaEngineLike extends GridEngineLike {
  getColumnAxisEntries(sheetName: string): readonly WorkbookAxisEntrySnapshot[]
  getRowAxisEntries(sheetName: string): readonly WorkbookAxisEntrySnapshot[]
  getLastMetrics(): Pick<RecalcMetrics, 'batchId'>
}

const FULL_SPAN_START = 0

export function buildWorkerRenderTileDeltaBatch(input: {
  engine: RenderTileDeltaEngineLike
  subscription: RenderTileDeltaSubscription
  generation: number
}): RenderTileDeltaBatch {
  const { engine, subscription, generation } = input
  const metrics = engine.getLastMetrics()
  const batchId = Number.isInteger(metrics.batchId) && metrics.batchId >= 0 ? metrics.batchId : 0
  const gridMetrics = getGridMetrics()
  const columnAxis = buildRenderedAxisState(engine.getColumnAxisEntries(subscription.sheetName), gridMetrics.columnWidth)
  const rowAxis = buildRenderedAxisState(engine.getRowAxisEntries(subscription.sheetName), gridMetrics.rowHeight)
  const freezeVersion = buildFreezeVersion(0, 0)

  return {
    magic: 'bilig.render.tile.delta',
    version: 1,
    sheetId: subscription.sheetId,
    batchId,
    cameraSeq: subscription.cameraSeq ?? 0,
    mutations: listViewportTileBounds(subscription).map((viewport) =>
      buildRenderTileReplaceMutation(
        materializeGridRenderTileV3({
          axisSeqX: columnAxis.version,
          axisSeqY: rowAxis.version,
          cameraSeq: subscription.cameraSeq ?? batchId,
          columnWidths: columnAxis.sizes,
          dprBucket: subscription.dprBucket ?? 1,
          engine,
          freezeSeq: freezeVersion,
          glyphAtlasSeq: 0,
          gridMetrics,
          materializedAtSeq: generation,
          packetSeq: batchId,
          rectSeq: batchId,
          rowHeights: rowAxis.sizes,
          sheetId: subscription.sheetId,
          sheetName: subscription.sheetName,
          sortedColumnWidthOverrides: columnAxis.sortedOverrides,
          sortedRowHeightOverrides: rowAxis.sortedOverrides,
          styleSeq: batchId,
          textSeq: batchId,
          valueSeq: batchId,
          viewport,
        }),
      ),
    ),
  }
}

export function buildRenderTileReplaceMutation(tile: GridRenderTile): RenderTileReplaceMutation {
  return {
    kind: 'tileReplace',
    tileId: tile.tileId,
    coord: tile.coord,
    version: tile.version,
    bounds: tile.bounds,
    rectInstances: tile.rectInstances,
    rectCount: tile.rectCount,
    textMetrics: tile.textMetrics,
    glyphRefs: new Uint32Array(0),
    textRuns: tile.textRuns.map((run) => ({
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
      x: run.x,
      y: run.y,
    })),
    textCount: tile.textCount,
    dirty: {
      rectSpans: tile.rectCount > 0 ? [{ offset: FULL_SPAN_START, length: tile.rectCount }] : [],
      textSpans: tile.textCount > 0 ? [{ offset: FULL_SPAN_START, length: tile.textCount }] : [],
      glyphSpans: [],
    },
  }
}
