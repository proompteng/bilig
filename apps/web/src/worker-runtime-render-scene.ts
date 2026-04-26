import type { RecalcMetrics, WorkbookAxisEntrySnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '@bilig/grid'
import { buildResidentDataPaneScenes } from '../../../packages/grid/src/gridResidentDataLayer.js'
import { getGridMetrics } from '../../../packages/grid/src/gridMetrics.js'
import { resolveFrozenColumnWidth, resolveFrozenRowHeight } from '../../../packages/grid/src/workbookGridViewport.js'
import { createGridTileKeyV2 } from '../../../packages/grid/src/renderer-v2/scene-packet-v2.js'
import { validateGridScenePacketV2 } from '../../../packages/grid/src/renderer-v2/scene-packet-validator.js'
import type { WorkbookPaneScenePacket, WorkbookPaneSceneRequest } from './resident-pane-scene-types.js'

interface ResidentPaneSceneEngineLike extends GridEngineLike {
  getColumnAxisEntries(sheetName: string): readonly WorkbookAxisEntrySnapshot[]
  getRowAxisEntries(sheetName: string): readonly WorkbookAxisEntrySnapshot[]
  getLastMetrics(): Pick<RecalcMetrics, 'batchId'>
}

function buildRenderedAxisState(
  entries: readonly WorkbookAxisEntrySnapshot[],
  defaultSize: number,
): {
  sizes: Record<number, number>
  sortedOverrides: Array<readonly [number, number]>
} {
  const sizes: Record<number, number> = {}
  const sortedOverrides: Array<readonly [number, number]> = []
  for (const entry of entries) {
    const renderedSize = entry.hidden ? 0 : (entry.size ?? defaultSize)
    if (renderedSize === defaultSize) {
      continue
    }
    sizes[entry.index] = renderedSize
    sortedOverrides.push([entry.index, renderedSize] as const)
  }
  sortedOverrides.sort((left, right) => left[0] - right[0])
  return {
    sizes,
    sortedOverrides,
  }
}

function resolveRevision(value: number | undefined): number {
  return Number.isInteger(value) && value !== undefined && value >= 0 ? value : 0
}

function hashAxisEntries(entries: readonly WorkbookAxisEntrySnapshot[]): number {
  if (entries.length === 0) {
    return 0
  }
  let hash = 2_166_136_261
  for (const entry of [...entries].toSorted((left, right) => left.index - right.index)) {
    hash = mixRevisionInteger(hash, entry.index)
    hash = mixRevisionInteger(hash, Math.round((entry.size ?? -1) * 1_000))
    hash = mixRevisionInteger(hash, entry.hidden ? 1 : 0)
  }
  return hash >>> 0
}

function mixRevisionInteger(hash: number, value: number): number {
  return Math.imul((hash ^ value) >>> 0, 16_777_619) >>> 0
}

function buildFreezeVersion(freezeRows: number, freezeCols: number): number {
  return mixRevisionInteger(mixRevisionInteger(2_166_136_261, freezeRows), freezeCols)
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
  const columnAxisVersion = hashAxisEntries(columnAxisEntries)
  const rowAxisVersion = hashAxisEntries(rowAxisEntries)
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
          axisVersionX: columnAxisVersion,
          axisVersionY: rowAxisVersion,
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
