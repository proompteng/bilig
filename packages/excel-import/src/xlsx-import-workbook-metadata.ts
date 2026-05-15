import type { WorkbookMetadataSnapshot } from '@bilig/protocol'

interface ImportedWorkbookMetadataInput {
  readonly properties?: WorkbookMetadataSnapshot['properties']
  readonly documentPropertyArtifacts?: WorkbookMetadataSnapshot['documentPropertyArtifacts']
  readonly workbookProtection?: WorkbookMetadataSnapshot['workbookProtection']
  readonly calculationSettings?: WorkbookMetadataSnapshot['calculationSettings']
  readonly macroPayloads?: WorkbookMetadataSnapshot['macroPayloads']
  readonly styles?: WorkbookMetadataSnapshot['styles']
  readonly definedNames?: WorkbookMetadataSnapshot['definedNames']
  readonly tables?: WorkbookMetadataSnapshot['tables']
  readonly spills?: WorkbookMetadataSnapshot['spills']
  readonly pivots?: WorkbookMetadataSnapshot['pivots']
  readonly unsupportedFormulaDependencies?: WorkbookMetadataSnapshot['unsupportedFormulaDependencies']
  readonly unsupportedPivots?: WorkbookMetadataSnapshot['unsupportedPivots']
  readonly pivotArtifacts?: WorkbookMetadataSnapshot['pivotArtifacts']
  readonly drawingArtifacts?: WorkbookMetadataSnapshot['drawingArtifacts']
  readonly chartArtifacts?: WorkbookMetadataSnapshot['chartArtifacts']
  readonly chartSheetArtifacts?: WorkbookMetadataSnapshot['chartSheetArtifacts']
  readonly controlArtifacts?: WorkbookMetadataSnapshot['controlArtifacts']
  readonly dataModelArtifacts?: WorkbookMetadataSnapshot['dataModelArtifacts']
  readonly externalLinkArtifacts?: WorkbookMetadataSnapshot['externalLinkArtifacts']
  readonly slicerConnectionArtifacts?: WorkbookMetadataSnapshot['slicerConnectionArtifacts']
  readonly threadedCommentArtifacts?: WorkbookMetadataSnapshot['threadedCommentArtifacts']
  readonly viewState?: WorkbookMetadataSnapshot['viewState']
  readonly charts?: WorkbookMetadataSnapshot['charts']
  readonly styleArtifacts?: WorkbookMetadataSnapshot['styleArtifacts']
  readonly cellMetadata?: WorkbookMetadataSnapshot['cellMetadata']
}

export function buildImportedWorkbookMetadata(input: ImportedWorkbookMetadataInput): WorkbookMetadataSnapshot | undefined {
  const metadata: WorkbookMetadataSnapshot = {
    ...(input.properties ? { properties: input.properties } : {}),
    ...(input.documentPropertyArtifacts ? { documentPropertyArtifacts: input.documentPropertyArtifacts } : {}),
    ...(input.workbookProtection ? { workbookProtection: input.workbookProtection } : {}),
    ...(input.calculationSettings ? { calculationSettings: input.calculationSettings } : {}),
    ...(input.macroPayloads ? { macroPayloads: input.macroPayloads } : {}),
    ...(input.styles ? { styles: input.styles } : {}),
    ...(input.definedNames ? { definedNames: input.definedNames } : {}),
    ...(input.tables ? { tables: input.tables } : {}),
    ...(input.spills ? { spills: input.spills } : {}),
    ...(input.pivots ? { pivots: input.pivots } : {}),
    ...(input.unsupportedFormulaDependencies ? { unsupportedFormulaDependencies: input.unsupportedFormulaDependencies } : {}),
    ...(input.unsupportedPivots ? { unsupportedPivots: input.unsupportedPivots } : {}),
    ...(input.pivotArtifacts ? { pivotArtifacts: input.pivotArtifacts } : {}),
    ...(input.drawingArtifacts ? { drawingArtifacts: input.drawingArtifacts } : {}),
    ...(input.chartArtifacts ? { chartArtifacts: input.chartArtifacts } : {}),
    ...(input.chartSheetArtifacts ? { chartSheetArtifacts: input.chartSheetArtifacts } : {}),
    ...(input.controlArtifacts ? { controlArtifacts: input.controlArtifacts } : {}),
    ...(input.dataModelArtifacts ? { dataModelArtifacts: input.dataModelArtifacts } : {}),
    ...(input.externalLinkArtifacts ? { externalLinkArtifacts: input.externalLinkArtifacts } : {}),
    ...(input.slicerConnectionArtifacts ? { slicerConnectionArtifacts: input.slicerConnectionArtifacts } : {}),
    ...(input.threadedCommentArtifacts ? { threadedCommentArtifacts: input.threadedCommentArtifacts } : {}),
    ...(input.viewState ? { viewState: input.viewState } : {}),
    ...(input.charts ? { charts: input.charts } : {}),
    ...(input.styleArtifacts ? { styleArtifacts: input.styleArtifacts } : {}),
    ...(input.cellMetadata ? { cellMetadata: input.cellMetadata } : {}),
  }
  return Object.keys(metadata).length > 0 ? metadata : undefined
}
