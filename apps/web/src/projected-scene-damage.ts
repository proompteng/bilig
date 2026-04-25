import type { WorkbookPaneSceneRequest } from './resident-pane-scene-types.js'
import type { ViewportPatch } from '@bilig/worker-transport'

function intersects(
  left: Pick<WorkbookPaneSceneRequest['residentViewport'], 'rowStart' | 'rowEnd' | 'colStart' | 'colEnd'>,
  right: Pick<ViewportPatch['viewport'], 'rowStart' | 'rowEnd' | 'colStart' | 'colEnd'>,
): boolean {
  return !(left.rowEnd < right.rowStart || right.rowEnd < left.rowStart || left.colEnd < right.colStart || right.colEnd < left.colStart)
}

export function residentPaneSceneRequestNeedsRefresh(request: WorkbookPaneSceneRequest, patch: ViewportPatch): boolean {
  if (request.sheetName !== patch.viewport.sheetName) {
    return false
  }
  if (patch.full || patch.freezeRows !== undefined || patch.freezeCols !== undefined) {
    return true
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
