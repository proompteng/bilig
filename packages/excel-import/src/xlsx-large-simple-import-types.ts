import type { SheetMetadataSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import type { ImportedWorkbookSheetPreview } from './workbook-import-helpers.js'
import type { ImportedWorkbookPreview } from './workbook-import-preview.js'
import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import type { LargeSimpleXlsxImportPhaseTelemetry } from './xlsx-large-simple-import-telemetry.js'
import type { LargeSimpleSharedStringIndexSet } from './xlsx-large-simple-shared-string-indexes.js'
import type { LargeSimpleSharedStrings } from './xlsx-large-simple-shared-strings.js'
import type { LargeSimpleWorksheetScannedMetadata } from './xlsx-large-simple-worksheet-metadata.js'

export interface LargeSimpleXlsxImportResult {
  snapshot: WorkbookSnapshot
  workbookName: string
  sheetNames: string[]
  warnings: string[]
  preview: ImportedWorkbookPreview
  stats: LargeSimpleXlsxImportStats
}

export interface LargeSimpleXlsxImportOptions {
  minByteLength?: number
  materializeCells?: boolean
  materializeMetadata?: boolean
  releaseArenaAfterMaterialization?: boolean
  releaseZipSource?: boolean
  allowUnsupportedFormulaText?: boolean
  allowUnsupportedCellMetadata?: boolean
  maxMaterializedLazyPackageArtifactBytes?: number
  releaseOwnedSourceBytes?: () => LargeSimpleXlsxOwnedSourceReleaseEvidence | undefined
}

export interface LargeSimpleXlsxImportSource {
  readonly byteLength: number
}

export interface LargeSimpleXlsxOwnedSourceReleaseEvidence {
  readonly ownedSourceBytesBeforeRelease?: number
  readonly ownedSourceBytesAfterRelease?: number
}

export interface LargeSimpleXlsxImportStats {
  readonly sheetCount: number
  readonly cellCount: number
  readonly formulaCellCount: number
  readonly valueCellCount: number
  readonly definedNameCount: number
  readonly tableCount: number
  readonly mergeCount: number
  readonly conditionalFormatCount: number
  readonly dataValidationCount: number
  readonly warningCount: number
  readonly dimensions: readonly LargeSimpleXlsxSheetDimension[]
  readonly phaseTelemetry: readonly LargeSimpleXlsxImportPhaseTelemetry[]
}

export interface LargeSimpleXlsxSheetDimension {
  readonly sheetName: string
  readonly rowCount: number
  readonly columnCount: number
  readonly nonEmptyCellCount: number
  readonly usedRange: ImportedWorksheetCellScan['usedRange']
}

export interface ParsedWorksheet {
  readonly sheet: WorkbookSnapshot['sheets'][number]
  readonly preview: ImportedWorkbookSheetPreview
  readonly stats: {
    readonly cellCount: number
    readonly formulaCellCount: number
    readonly valueCellCount: number
    readonly tableCount: number
    readonly mergeCount: number
    readonly conditionalFormatCount: number
    readonly dataValidationCount: number
    readonly dimension: LargeSimpleXlsxSheetDimension
  }
}

export interface ScannedWorksheet {
  readonly name: string
  readonly order: number
  readonly cellScan: ImportedWorksheetCellScan
  readonly worksheetXml: string | undefined
  readonly metadataScan: LargeSimpleWorksheetScannedMetadata | undefined
  readonly metadataInput: LargeSimpleSheetMetadataInput
  readonly sharedStringIndexes: LargeSimpleSharedStringIndexSet
  readonly sharedStrings?: LargeSimpleSharedStrings
  readonly hasUnresolvedSharedStringReferences?: boolean
}

export type LargeSimpleSheetMetadataInput = Pick<
  SheetMetadataSnapshot,
  | 'conditionalFormatArtifacts'
  | 'conditionalFormats'
  | 'controlArtifacts'
  | 'validations'
  | 'drawingArtifacts'
  | 'filters'
  | 'hyperlinks'
  | 'legacyCommentVml'
  | 'pivotArtifacts'
  | 'printerSettings'
  | 'printPageSetup'
  | 'sheetProtection'
>
