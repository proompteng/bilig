import type { SheetMetadataSnapshot } from '@bilig/protocol'

interface ImportedSheetMetadataInput {
  readonly rows?: SheetMetadataSnapshot['rows']
  readonly columns?: SheetMetadataSnapshot['columns']
  readonly rowMetadata?: SheetMetadataSnapshot['rowMetadata']
  readonly columnMetadata?: SheetMetadataSnapshot['columnMetadata']
  readonly sheetFormatPr?: SheetMetadataSnapshot['sheetFormatPr']
  readonly styleRanges?: SheetMetadataSnapshot['styleRanges']
  readonly freezePane?: SheetMetadataSnapshot['freezePane']
  readonly tabColor?: SheetMetadataSnapshot['tabColor']
  readonly sheetPr?: SheetMetadataSnapshot['sheetPr']
  readonly visibility?: SheetMetadataSnapshot['visibility']
  readonly merges?: SheetMetadataSnapshot['merges']
  readonly sheetProtection?: SheetMetadataSnapshot['sheetProtection']
  readonly protectedRanges?: SheetMetadataSnapshot['protectedRanges']
  readonly sorts?: SheetMetadataSnapshot['sorts']
  readonly filters?: SheetMetadataSnapshot['filters']
  readonly validations?: SheetMetadataSnapshot['validations']
  readonly conditionalFormats?: SheetMetadataSnapshot['conditionalFormats']
  readonly commentThreads?: SheetMetadataSnapshot['commentThreads']
  readonly legacyCommentVml?: SheetMetadataSnapshot['legacyCommentVml']
  readonly hyperlinks?: SheetMetadataSnapshot['hyperlinks']
  readonly printerSettings?: SheetMetadataSnapshot['printerSettings']
  readonly ignoredErrors?: SheetMetadataSnapshot['ignoredErrors']
  readonly cellMetadataRefs?: SheetMetadataSnapshot['cellMetadataRefs']
}

export function buildImportedSheetMetadata(input: ImportedSheetMetadataInput): SheetMetadataSnapshot | undefined {
  const metadata: SheetMetadataSnapshot = {
    ...(input.rows ? { rows: input.rows } : {}),
    ...(input.columns ? { columns: input.columns } : {}),
    ...(input.rowMetadata ? { rowMetadata: input.rowMetadata } : {}),
    ...(input.columnMetadata ? { columnMetadata: input.columnMetadata } : {}),
    ...(input.sheetFormatPr ? { sheetFormatPr: input.sheetFormatPr } : {}),
    ...(input.styleRanges ? { styleRanges: input.styleRanges } : {}),
    ...(input.freezePane ? { freezePane: input.freezePane } : {}),
    ...(input.tabColor ? { tabColor: input.tabColor } : {}),
    ...(input.sheetPr ? { sheetPr: input.sheetPr } : {}),
    ...(input.visibility ? { visibility: input.visibility } : {}),
    ...(input.merges ? { merges: input.merges } : {}),
    ...(input.sheetProtection ? { sheetProtection: input.sheetProtection } : {}),
    ...(input.protectedRanges ? { protectedRanges: input.protectedRanges } : {}),
    ...(input.sorts ? { sorts: input.sorts } : {}),
    ...(input.filters ? { filters: input.filters } : {}),
    ...(input.validations ? { validations: input.validations } : {}),
    ...(input.conditionalFormats ? { conditionalFormats: input.conditionalFormats } : {}),
    ...(input.commentThreads ? { commentThreads: input.commentThreads } : {}),
    ...(input.legacyCommentVml ? { legacyCommentVml: input.legacyCommentVml } : {}),
    ...(input.hyperlinks ? { hyperlinks: input.hyperlinks } : {}),
    ...(input.printerSettings ? { printerSettings: input.printerSettings } : {}),
    ...(input.ignoredErrors ? { ignoredErrors: input.ignoredErrors } : {}),
    ...(input.cellMetadataRefs ? { cellMetadataRefs: input.cellMetadataRefs } : {}),
  }
  return Object.keys(metadata).length > 0 ? metadata : undefined
}
