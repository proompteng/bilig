import type { WorkbookSnapshot } from '@bilig/protocol'
import { projectWorkbookSemanticSnapshot, type WorkbookSemanticSnapshot } from './workbook-semantic-projection.js'

export function stableStringifyWorkbookSemanticSnapshot(snapshot: WorkbookSemanticSnapshot): string {
  return JSON.stringify(snapshot)
}

export function workbookSemanticSnapshotsEqual(left: WorkbookSnapshot, right: WorkbookSnapshot): boolean {
  return (
    stableStringifyWorkbookSemanticSnapshot(projectWorkbookSemanticSnapshot(left)) ===
    stableStringifyWorkbookSemanticSnapshot(projectWorkbookSemanticSnapshot(right))
  )
}
