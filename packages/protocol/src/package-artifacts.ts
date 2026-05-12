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
