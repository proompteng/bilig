import { attachRuntimeImage, readRuntimeImage } from '@bilig/core/headless-runtime'
import type { WorkbookSnapshot } from '@bilig/protocol'

export function cloneWorkPaperSnapshotWithRuntimeImage(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  return attachWorkPaperRuntimeImage(snapshot, structuredClone(snapshot))
}

export function attachWorkPaperRuntimeImage(source: WorkbookSnapshot, target: WorkbookSnapshot): WorkbookSnapshot {
  const runtimeImage = readRuntimeImage(source)
  return runtimeImage ? attachRuntimeImage(target, structuredClone(runtimeImage)) : target
}
