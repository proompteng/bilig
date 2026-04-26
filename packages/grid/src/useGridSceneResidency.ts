import type { WorkbookPaneScenePacket } from './renderer-v2/pane-scene-types.js'
import { noteTypeGpuScenePacketApplied } from './renderer-v2/grid-render-counters.js'

type WorkbookResidentSceneIdentity = Pick<WorkbookPaneScenePacket, 'paneId' | 'viewport'>

export function canUseWorkerResidentPaneScenes(input: {
  readonly workerResidentPaneScenes: readonly WorkbookResidentSceneIdentity[]
  readonly requiresLiveViewportState: boolean
  readonly hasHoverState: boolean
  readonly hasActiveHeaderDrag: boolean
}): boolean {
  return input.workerResidentPaneScenes.length > 0 && !input.requiresLiveViewportState && !input.hasActiveHeaderDrag
}

export function noteWorkerResidentPaneScenesApplied(workerResidentPaneScenes: readonly WorkbookResidentSceneIdentity[]): void {
  workerResidentPaneScenes.forEach((scene) => {
    noteTypeGpuScenePacketApplied(
      `${scene.paneId}:${scene.viewport.rowStart}:${scene.viewport.rowEnd}:${scene.viewport.colStart}:${scene.viewport.colEnd}`,
    )
  })
}
