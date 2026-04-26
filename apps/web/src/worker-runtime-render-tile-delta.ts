import type { GridEngineLike } from '@bilig/grid'
import type { RecalcMetrics, WorkbookAxisEntrySnapshot, WorkbookFreezePaneSnapshot } from '@bilig/protocol'
import type {
  RenderTileDeltaBatch,
  RenderTileDeltaSubscription,
  RenderTilePaneKind,
  RenderTileReplaceMutation,
  RenderTileTextRun,
} from '@bilig/worker-transport'
import type { GridScenePacketV2, GridSceneTextRun } from '../../../packages/grid/src/renderer-v2/scene-packet-v2.js'
import type { WorkbookPaneScenePacket, WorkbookPaneSceneRequest } from './resident-pane-scene-types.js'
import { buildWorkerResidentPaneScenes } from './worker-runtime-render-scene.js'

interface RenderTileDeltaEngineLike extends GridEngineLike {
  getColumnAxisEntries(sheetName: string): readonly WorkbookAxisEntrySnapshot[]
  getRowAxisEntries(sheetName: string): readonly WorkbookAxisEntrySnapshot[]
  getLastMetrics(): Pick<RecalcMetrics, 'batchId'>
  getFreezePane?(sheetName: string): WorkbookFreezePaneSnapshot | undefined
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
  const scenes = buildWorkerResidentPaneScenes({
    engine,
    generation,
    request: buildResidentPaneSceneRequest(engine, subscription, batchId),
  })
  return buildRenderTileDeltaBatchFromResidentPaneScenes({
    batchId,
    cameraSeq: subscription.cameraSeq ?? 0,
    scenes,
    sheetId: subscription.sheetId,
  })
}

export function buildRenderTileDeltaBatchFromResidentPaneScenes(input: {
  sheetId: number
  batchId: number
  cameraSeq: number
  scenes: readonly WorkbookPaneScenePacket[]
}): RenderTileDeltaBatch {
  return {
    magic: 'bilig.render.tile.delta',
    version: 1,
    sheetId: input.sheetId,
    batchId: input.batchId,
    cameraSeq: input.cameraSeq,
    mutations: input.scenes.map((scene) => buildRenderTileReplaceMutation(input.sheetId, scene.packedScene)),
  }
}

export function buildRenderTileReplaceMutation(sheetId: number, packet: GridScenePacketV2): RenderTileReplaceMutation {
  return {
    kind: 'tileReplace',
    tileId: buildRenderTileId(sheetId, packet.key.paneKind, packet.key.rowTile, packet.key.colTile, packet.key.dprBucket),
    coord: {
      sheetId,
      paneKind: packet.key.paneKind,
      rowTile: packet.key.rowTile,
      colTile: packet.key.colTile,
      dprBucket: packet.key.dprBucket,
    },
    version: {
      axisX: packet.key.axisVersionX,
      axisY: packet.key.axisVersionY,
      values: packet.key.valueVersion,
      styles: packet.key.styleVersion,
      text: packet.key.textEpoch,
      freeze: packet.key.freezeVersion,
    },
    bounds: packet.viewport,
    rectInstances: packet.rectInstances,
    rectCount: packet.rectCount,
    textMetrics: packet.textMetrics,
    glyphRefs: new Uint32Array(0),
    textRuns: packet.textRuns.map(mapTextRun),
    textCount: packet.textCount,
    dirty: {
      rectSpans: packet.rectCount > 0 ? [{ offset: FULL_SPAN_START, length: packet.rectCount }] : [],
      textSpans: packet.textCount > 0 ? [{ offset: FULL_SPAN_START, length: packet.textCount }] : [],
      glyphSpans: [],
    },
  }
}

function buildResidentPaneSceneRequest(
  engine: RenderTileDeltaEngineLike,
  subscription: RenderTileDeltaSubscription,
  batchId: number,
): WorkbookPaneSceneRequest {
  const freezePane = engine.getFreezePane?.(subscription.sheetName)
  return {
    sheetName: subscription.sheetName,
    residentViewport: {
      rowStart: subscription.rowStart,
      rowEnd: subscription.rowEnd,
      colStart: subscription.colStart,
      colEnd: subscription.colEnd,
    },
    freezeRows: freezePane?.rows ?? 0,
    freezeCols: freezePane?.cols ?? 0,
    dprBucket: subscription.dprBucket ?? 1,
    requestSeq: batchId,
    cameraSeq: subscription.cameraSeq ?? batchId,
    priority: 0,
    reason: 'visible',
  }
}

function mapTextRun(run: GridSceneTextRun): RenderTileTextRun {
  return {
    text: run.text,
    x: run.x,
    y: run.y,
    width: run.width,
    height: run.height,
    clipX: run.clipX,
    clipY: run.clipY,
    clipWidth: run.clipWidth,
    clipHeight: run.clipHeight,
    font: run.font,
    fontSize: run.fontSize,
    color: run.color,
    underline: run.underline,
    strike: run.strike,
  }
}

function buildRenderTileId(sheetId: number, paneKind: RenderTilePaneKind, rowTile: number, colTile: number, dprBucket: number): number {
  let hash = 2_166_136_261
  hash = mixHash(hash, sheetId)
  hash = mixHash(hash, paneKindToId(paneKind))
  hash = mixHash(hash, rowTile)
  hash = mixHash(hash, colTile)
  hash = mixHash(hash, dprBucket)
  return hash >>> 0
}

function mixHash(hash: number, value: number): number {
  return Math.imul((hash ^ value) >>> 0, 16_777_619) >>> 0
}

function paneKindToId(paneKind: RenderTilePaneKind): number {
  switch (paneKind) {
    case 'body':
      return 0
    case 'frozenTop':
      return 1
    case 'frozenLeft':
      return 2
    case 'frozenCorner':
      return 3
    case 'columnHeaderBody':
      return 4
    case 'columnHeaderFrozen':
      return 5
    case 'rowHeaderBody':
      return 6
    case 'rowHeaderFrozen':
      return 7
    case 'dynamicOverlay':
      return 8
  }
}
