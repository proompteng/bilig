import type { WorkbookPaneSceneRequest } from './resident-pane-scene-types.js'
import type { ViewportPatch } from '@bilig/worker-transport'

export interface ResidentScenePatchDamage {
  readonly damage: readonly { readonly cell: readonly [number, number] }[]
  readonly axisChanged: boolean
  readonly freezeChanged: boolean
}

function intersects(
  left: Pick<WorkbookPaneSceneRequest['residentViewport'], 'rowStart' | 'rowEnd' | 'colStart' | 'colEnd'>,
  right: Pick<ViewportPatch['viewport'], 'rowStart' | 'rowEnd' | 'colStart' | 'colEnd'>,
): boolean {
  return !(left.rowEnd < right.rowStart || right.rowEnd < left.rowStart || left.colEnd < right.colStart || right.colEnd < left.colStart)
}

function containsDamageCell(request: WorkbookPaneSceneRequest, damage: ResidentScenePatchDamage['damage']): boolean {
  return damage.some(
    ({ cell }) =>
      cell[1] >= request.residentViewport.rowStart &&
      cell[1] <= request.residentViewport.rowEnd &&
      cell[0] >= request.residentViewport.colStart &&
      cell[0] <= request.residentViewport.colEnd,
  )
}

export function residentPaneSceneRequestNeedsRefresh(
  request: WorkbookPaneSceneRequest,
  patch: ViewportPatch,
  applied?: ResidentScenePatchDamage,
): boolean {
  if (request.sheetName !== patch.viewport.sheetName) {
    return false
  }
  if (applied) {
    if (applied.freezeChanged) {
      return true
    }
    if (applied.axisChanged) {
      return intersects(request.residentViewport, patch.viewport)
    }
    return containsDamageCell(request, applied.damage)
  }
  if (patch.freezeRows !== undefined || patch.freezeCols !== undefined) {
    return true
  }
  if (patch.full) {
    return intersects(request.residentViewport, patch.viewport)
  }
  if (patch.columns.length > 0 || patch.rows.length > 0) {
    return intersects(request.residentViewport, patch.viewport)
  }
  if (patch.styles.length > 0) {
    return intersects(request.residentViewport, patch.viewport)
  }
  if (patch.cells.length > 0) {
    return intersects(request.residentViewport, patch.viewport)
  }
  return false
}
