import type { RecalcMetrics, WorkbookAxisEntrySnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '@bilig/grid'
import { buildResidentDataPaneScenes } from '../../../packages/grid/src/gridResidentDataLayer.js'
import { getGridMetrics } from '../../../packages/grid/src/gridMetrics.js'
import { resolveFrozenColumnWidth, resolveFrozenRowHeight } from '../../../packages/grid/src/workbookGridViewport.js'
import { createGridTileKeyV2 } from '../../../packages/grid/src/renderer-v2/scene-packet-v2.js'
import { validateGridScenePacketV2 } from '../../../packages/grid/src/renderer-v2/scene-packet-validator.js'
import type { WorkbookPaneScenePacket, WorkbookPaneSceneRequest } from './resident-pane-scene-types.js'
import { buildFreezeVersion, buildRenderedAxisState, resolveRevision } from './worker-runtime-render-axis.js'

interface ResidentPaneSceneEngineLike extends GridEngineLike {
  getColumnAxisEntries(sheetName: string): readonly WorkbookAxisEntrySnapshot[]
  getRowAxisEntries(sheetName: string): readonly WorkbookAxisEntrySnapshot[]
  getLastMetrics(): Pick<RecalcMetrics, 'batchId'>
}

export function buildWorkerResidentPaneScenes(input: {
  engine: ResidentPaneSceneEngineLike
  request: WorkbookPaneSceneRequest
  generation: number
}): WorkbookPaneScenePacket[] {
  const { engine, request, generation } = input
  const generatedAt = Date.now()
  const gridMetrics = getGridMetrics()
  const columnAxisEntries = engine.getColumnAxisEntries(request.sheetName)
  const rowAxisEntries = engine.getRowAxisEntries(request.sheetName)
  const columnAxis = buildRenderedAxisState(columnAxisEntries, gridMetrics.columnWidth)
  const rowAxis = buildRenderedAxisState(rowAxisEntries, gridMetrics.rowHeight)
  const batchVersion = resolveRevision(engine.getLastMetrics().batchId)
  const freezeRows = request.freezeRows
  const freezeCols = request.freezeCols
  const freezeVersion = buildFreezeVersion(freezeRows, freezeCols)
  const sceneVersion = batchVersion + resolveRevision(request.sceneRevision)

  const scenes = buildResidentDataPaneScenes({
    residentViewport: request.residentViewport,
    engine,
    sheetName: request.sheetName,
    columnWidths: columnAxis.sizes,
    rowHeights: rowAxis.sizes,
    freezeRows,
    freezeCols,
    frozenColumnWidth: resolveFrozenColumnWidth({
      freezeCols,
      columnWidths: columnAxis.sizes,
      gridMetrics,
    }),
    frozenRowHeight: resolveFrozenRowHeight({
      freezeRows,
      rowHeights: rowAxis.sizes,
      gridMetrics,
    }),
    gridMetrics,
    sortedColumnWidthOverrides: columnAxis.sortedOverrides,
    sortedRowHeightOverrides: rowAxis.sortedOverrides,
    packet: {
      generation,
      cameraSeq: request.cameraSeq ?? request.requestSeq ?? 0,
      generatedAt,
      requestSeq: request.requestSeq ?? 0,
      createKey: (paneId, viewport) =>
        createGridTileKeyV2({
          axisVersionX: columnAxis.version,
          axisVersionY: rowAxis.version,
          dprBucket: request.dprBucket ?? 1,
          freezeVersion,
          paneId,
          selectionIndependentVersion: sceneVersion,
          sheetName: request.sheetName,
          styleVersion: sceneVersion,
          valueVersion: sceneVersion,
          viewport,
        }),
    },
  })
  for (const scene of scenes) {
    const validation = validateGridScenePacketV2(scene.packedScene)
    if (!validation.ok) {
      throw new Error(`Invalid worker grid scene packet: ${validation.reason}`)
    }
  }
  return scenes
}

export function buildResidentPaneSceneCacheKey(request: WorkbookPaneSceneRequest): string {
  return [
    request.sheetName,
    request.residentViewport.rowStart,
    request.residentViewport.rowEnd,
    request.residentViewport.colStart,
    request.residentViewport.colEnd,
    request.freezeRows,
    request.freezeCols,
    request.dprBucket ?? 1,
    request.sceneRevision ?? 0,
  ].join(':')
}
