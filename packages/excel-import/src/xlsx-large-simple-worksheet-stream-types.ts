import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import type { LargeSimpleSharedStrings } from './xlsx-large-simple-shared-strings.js'
import type { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'
import type { LargeSimpleWorksheetScannedMetadata } from './xlsx-large-simple-worksheet-metadata.js'

export interface LargeSimpleWorksheetStreamScan {
  readonly cellScan: ImportedWorksheetCellScan
  readonly metadataXml: string | undefined
  readonly metadata: LargeSimpleWorksheetScannedMetadata | undefined
}

export interface LargeSimpleWorksheetStreamScanOptions {
  readonly hasSharedStrings: boolean
  readonly retainCells?: boolean
  readonly sharedStrings?: LargeSimpleSharedStrings
  readonly deferSharedStrings?: boolean
  readonly retainMetadataXml?: boolean
  readonly sheetName?: string
  readonly stringPool?: ImportedWorkbookStringPool
  readonly onRetainedBufferLength?: (length: number) => void
}
