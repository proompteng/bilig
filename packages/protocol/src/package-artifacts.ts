export interface WorkbookPackageRelationshipSnapshot {
  id: string
  type: string
  target: string
  targetMode?: string
}

export interface WorkbookThemeArtifactSnapshot {
  path: string
  xml: string
  relationship: WorkbookPackageRelationshipSnapshot
  contentType?: string
}

export interface WorkbookStyleArtifactsSnapshot {
  stylesXml: string
  theme?: WorkbookThemeArtifactSnapshot
}

export interface WorkbookDocumentPropertyPartSnapshot {
  path: string
  xml: string
  relationship: WorkbookPackageRelationshipSnapshot
  contentType?: string
}

export interface WorkbookDocumentPropertiesArtifactsSnapshot {
  core?: WorkbookDocumentPropertyPartSnapshot
  app?: WorkbookDocumentPropertyPartSnapshot
  custom?: WorkbookDocumentPropertyPartSnapshot
}

export interface WorkbookPreservedPackagePartSnapshot {
  path: string
  storage: 'base64'
  dataBase64: string
  byteLength: number
}

export interface WorkbookContentTypeDefaultSnapshot {
  extension: string
  contentType: string
}

export interface WorkbookContentTypeOverrideSnapshot {
  partName: string
  contentType: string
}

export interface WorkbookDrawingArtifactsSnapshot {
  parts: WorkbookPreservedPackagePartSnapshot[]
  contentTypeDefaults?: WorkbookContentTypeDefaultSnapshot[]
  contentTypeOverrides?: WorkbookContentTypeOverrideSnapshot[]
}

export interface WorkbookSheetDrawingArtifactsSnapshot {
  relationshipTarget: string
}

export interface WorkbookControlArtifactsSnapshot {
  parts: WorkbookPreservedPackagePartSnapshot[]
  contentTypeDefaults?: WorkbookContentTypeDefaultSnapshot[]
  contentTypeOverrides?: WorkbookContentTypeOverrideSnapshot[]
}

export interface WorkbookSheetControlArtifactsSnapshot {
  controlsXml: string
  worksheetRootOpenTag: string
  relationships: WorkbookPackageRelationshipSnapshot[]
}

export interface WorkbookDataModelArtifactsSnapshot {
  parts: WorkbookPreservedPackagePartSnapshot[]
  workbookRelationships: WorkbookPackageRelationshipSnapshot[]
  contentTypeDefaults?: WorkbookContentTypeDefaultSnapshot[]
  contentTypeOverrides?: WorkbookContentTypeOverrideSnapshot[]
}

export interface WorkbookThreadedCommentArtifactsSnapshot {
  parts: WorkbookPreservedPackagePartSnapshot[]
  workbookRelationships?: WorkbookPackageRelationshipSnapshot[]
  contentTypeDefaults?: WorkbookContentTypeDefaultSnapshot[]
  contentTypeOverrides?: WorkbookContentTypeOverrideSnapshot[]
}

export interface WorkbookSheetThreadedCommentArtifactsSnapshot {
  relationships: WorkbookPackageRelationshipSnapshot[]
}

export interface WorkbookSheetArrayFormulaSnapshot {
  address: string
  formulaXml: string
}

export interface WorkbookSheetArrayFormulasSnapshot {
  formulas: WorkbookSheetArrayFormulaSnapshot[]
}

export interface WorkbookSheetDataTableFormulaSnapshot {
  address: string
  formulaXml: string
}

export interface WorkbookSheetDataTableFormulasSnapshot {
  formulas: WorkbookSheetDataTableFormulaSnapshot[]
}
