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
  readonly conditionalFormatArtifacts?: SheetMetadataSnapshot['conditionalFormatArtifacts']
  readonly commentThreads?: SheetMetadataSnapshot['commentThreads']
  readonly drawingArtifacts?: SheetMetadataSnapshot['drawingArtifacts']
  readonly controlArtifacts?: SheetMetadataSnapshot['controlArtifacts']
  readonly arrayFormulas?: SheetMetadataSnapshot['arrayFormulas']
  readonly dataTableFormulas?: SheetMetadataSnapshot['dataTableFormulas']
  readonly legacyCommentVml?: SheetMetadataSnapshot['legacyCommentVml']
  readonly hyperlinks?: SheetMetadataSnapshot['hyperlinks']
  readonly printerSettings?: SheetMetadataSnapshot['printerSettings']
  readonly ignoredErrors?: SheetMetadataSnapshot['ignoredErrors']
  readonly sparklines?: SheetMetadataSnapshot['sparklines']
  readonly styleArtifacts?: SheetMetadataSnapshot['styleArtifacts']
  readonly pivotArtifacts?: SheetMetadataSnapshot['pivotArtifacts']
  readonly cellMetadataRefs?: SheetMetadataSnapshot['cellMetadataRefs']
  readonly richTextArtifacts?: SheetMetadataSnapshot['richTextArtifacts']
  readonly threadedCommentArtifacts?: SheetMetadataSnapshot['threadedCommentArtifacts']
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
    ...(input.conditionalFormatArtifacts ? { conditionalFormatArtifacts: input.conditionalFormatArtifacts } : {}),
    ...(input.commentThreads ? { commentThreads: input.commentThreads } : {}),
    ...(input.drawingArtifacts ? { drawingArtifacts: input.drawingArtifacts } : {}),
    ...(input.controlArtifacts ? { controlArtifacts: input.controlArtifacts } : {}),
    ...(input.arrayFormulas ? { arrayFormulas: input.arrayFormulas } : {}),
    ...(input.dataTableFormulas ? { dataTableFormulas: input.dataTableFormulas } : {}),
    ...(input.legacyCommentVml ? { legacyCommentVml: input.legacyCommentVml } : {}),
    ...(input.hyperlinks ? { hyperlinks: input.hyperlinks } : {}),
    ...(input.printerSettings ? { printerSettings: input.printerSettings } : {}),
    ...(input.ignoredErrors ? { ignoredErrors: input.ignoredErrors } : {}),
    ...(input.sparklines ? { sparklines: input.sparklines } : {}),
    ...(input.styleArtifacts ? { styleArtifacts: input.styleArtifacts } : {}),
    ...(input.pivotArtifacts ? { pivotArtifacts: input.pivotArtifacts } : {}),
    ...(input.cellMetadataRefs ? { cellMetadataRefs: input.cellMetadataRefs } : {}),
    ...(input.richTextArtifacts ? { richTextArtifacts: input.richTextArtifacts } : {}),
    ...(input.threadedCommentArtifacts ? { threadedCommentArtifacts: input.threadedCommentArtifacts } : {}),
  }
  return Object.keys(metadata).length > 0 ? metadata : undefined
}
