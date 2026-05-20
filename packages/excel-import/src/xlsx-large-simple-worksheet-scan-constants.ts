export const unsupportedWorksheetTagNames = new Set(['controls', 'picture'])

export const metadataWorksheetTagNames = new Set([
  'autoFilter',
  'colBreaks',
  'cols',
  'conditionalFormatting',
  'dataValidations',
  'drawing',
  'extLst',
  'headerFooter',
  'hyperlinks',
  'legacyDrawing',
  'mergeCells',
  'oleObjects',
  'pageMargins',
  'pageSetup',
  'pivotTableDefinition',
  'printOptions',
  'rowBreaks',
  'sheetFormatPr',
  'tableParts',
])
export const richTextRunPattern = /<(?:[A-Za-z_][\w.-]*:)?r\b/u
