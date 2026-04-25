import type { WorkbookPaneScenePacket } from './renderer/pane-scene-types.js'
import { noteTypeGpuScenePacketApplied } from './renderer/grid-render-counters.js'

export function canUseWorkerResidentPaneScenes(input: {
  readonly workerResidentPaneScenes: readonly WorkbookPaneScenePacket[]
  readonly requiresLiveViewportState: boolean
  readonly hasHoverState: boolean
  readonly hasActiveHeaderDrag: boolean
}): boolean {
  return input.workerResidentPaneScenes.length > 0 && !input.requiresLiveViewportState && !input.hasActiveHeaderDrag
}

export function noteWorkerResidentPaneScenesApplied(workerResidentPaneScenes: readonly WorkbookPaneScenePacket[]): void {
  workerResidentPaneScenes.forEach((scene) => {
    noteTypeGpuScenePacketApplied(
      `${scene.paneId}:${scene.viewport.rowStart}:${scene.viewport.rowEnd}:${scene.viewport.colStart}:${scene.viewport.colEnd}`,
    )
  })
}
