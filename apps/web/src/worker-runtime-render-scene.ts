import { formatAddress } from '@bilig/formula'
import type { WorkbookAxisEntrySnapshot } from '@bilig/protocol'
import { buildResidentDataPaneScenes } from '../../../packages/grid/src/gridResidentDataLayer.js'
import { getGridMetrics } from '../../../packages/grid/src/gridMetrics.js'
import { createGridSelection } from '../../../packages/grid/src/gridSelection.js'
import { resolveFrozenColumnWidth, resolveFrozenRowHeight } from '../../../packages/grid/src/workbookGridViewport.js'
import { validateGridScenePacketV2 } from '../../../packages/grid/src/renderer-v2/scene-packet-validator.js'
import type { WorkbookPaneScenePacket, WorkbookPaneSceneRequest } from './resident-pane-scene-types.js'
import { packWorkerGridScenePacket } from './worker-runtime-render-packet.js'
import type { WorkerEngine } from './worker-runtime-support.js'

interface ResidentPaneSceneEngineLike {
  workbook: WorkerEngine['workbook']
  getCell: WorkerEngine['getCell']
  getCellStyle: WorkerEngine['getCellStyle']
  getColumnAxisEntries: WorkerEngine['getColumnAxisEntries']
  getRowAxisEntries: WorkerEngine['getRowAxisEntries']
  subscribeCells: (sheetName: string, addresses: readonly string[], listener: () => void) => () => void
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

function normalizeSelectionRange(request: WorkbookPaneSceneRequest) {
  return (
    request.selectionRange ?? {
      x: request.selectedCell.col,
      y: request.selectedCell.row,
      width: 1,
      height: 1,
    }
  )
}

export function buildWorkerResidentPaneScenes(input: {
  engine: ResidentPaneSceneEngineLike
  request: WorkbookPaneSceneRequest
  generation: number
}): WorkbookPaneScenePacket[] {
  const { engine, request, generation } = input
  const gridMetrics = getGridMetrics()
  const selectedAddress = formatAddress(request.selectedCell.row, request.selectedCell.col)
  const selectedCellSnapshot = engine.getCell(request.sheetName, selectedAddress)
  const columnAxis = buildRenderedAxisState(engine.getColumnAxisEntries(request.sheetName), gridMetrics.columnWidth)
  const rowAxis = buildRenderedAxisState(engine.getRowAxisEntries(request.sheetName), gridMetrics.rowHeight)
  const freezeRows = request.freezeRows
  const freezeCols = request.freezeCols

  return buildResidentDataPaneScenes({
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
    gridSelection: createGridSelection(request.selectedCell.col, request.selectedCell.row),
    selectedCell: [request.selectedCell.col, request.selectedCell.row],
    selectedCellSnapshot,
    selectionRange: normalizeSelectionRange(request),
    editingCell: request.editingCell ? [request.editingCell.col, request.editingCell.row] : null,
  }).map((scene): WorkbookPaneScenePacket => {
    const packedScene = packWorkerGridScenePacket({
      generation,
      gpuScene: scene.gpuScene,
      paneId: scene.paneId,
      sheetName: request.sheetName,
      surfaceSize: scene.surfaceSize,
      textScene: scene.textScene,
      viewport: scene.viewport,
    })
    const validation = validateGridScenePacketV2(packedScene)
    if (!validation.ok) {
      throw new Error(`Invalid worker grid scene packet: ${validation.reason}`)
    }
    return {
      generation,
      paneId: scene.paneId,
      packedScene,
      viewport: scene.viewport,
      surfaceSize: scene.surfaceSize,
      gpuScene: scene.gpuScene,
      textScene: scene.textScene,
    }
  })
}

export function buildResidentPaneSceneCacheKey(request: WorkbookPaneSceneRequest): string {
  const range = request.selectionRange
  return [
    request.sheetName,
    request.residentViewport.rowStart,
    request.residentViewport.rowEnd,
    request.residentViewport.colStart,
    request.residentViewport.colEnd,
    request.freezeRows,
    request.freezeCols,
    request.selectedCell.col,
    request.selectedCell.row,
    range?.x ?? -1,
    range?.y ?? -1,
    range?.width ?? -1,
    range?.height ?? -1,
    request.editingCell?.col ?? -1,
    request.editingCell?.row ?? -1,
  ].join(':')
}
