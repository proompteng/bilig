import type { SheetMetadataSnapshot } from '@bilig/protocol'
import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import { readLargeSimpleDrawingRelationshipId, type LargeSimpleWorksheetScannedMetadata } from './xlsx-large-simple-worksheet-metadata.js'

export function releaseProjectedCellScanStorage(
  cellScan: ImportedWorksheetCellScan,
  options: {
    readonly releaseArenaAfterMaterialization: boolean | undefined
    readonly useLazyCells: boolean
  },
): void {
  if (options.releaseArenaAfterMaterialization !== true) {
    return
  }
  if (options.useLazyCells) {
    cellScan.arena.releaseMaterializationScratch()
  } else {
    cellScan.arena.release()
  }
  cellScan.styleIndexes.release()
}

export function sheetPivotArtifactsWithStreamedDefinitions(
  artifacts: SheetMetadataSnapshot['pivotArtifacts'],
  pivotTableDefinitionsXml: string | undefined,
): SheetMetadataSnapshot['pivotArtifacts'] {
  if (!pivotTableDefinitionsXml) {
    return artifacts
  }
  return {
    relationships: artifacts?.relationships ?? [],
    pivotTableDefinitionsXml,
  }
}

export function drawingRelationshipIdForScannedWorksheet(scanned: {
  readonly metadataScan: LargeSimpleWorksheetScannedMetadata | undefined
  readonly worksheetXml: string | undefined
}): string | undefined {
  return (
    scanned.metadataScan?.drawingRelationshipId ??
    (scanned.worksheetXml ? readLargeSimpleDrawingRelationshipId(scanned.worksheetXml) : undefined)
  )
}
