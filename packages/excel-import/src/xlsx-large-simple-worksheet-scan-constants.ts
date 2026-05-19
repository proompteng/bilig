export const unsupportedWorksheetTagNames = new Set(['dataValidations', 'legacyDrawing', 'oleObjects', 'picture', 'sheetProtection'])

export const metadataWorksheetTagNames = new Set([
  'autoFilter',
  'colBreaks',
  'cols',
  'conditionalFormatting',
  'drawing',
  'headerFooter',
  'hyperlinks',
  'mergeCells',
  'pageMargins',
  'pageSetup',
  'printOptions',
  'rowBreaks',
  'sheetFormatPr',
  'tableParts',
])
export const richTextRunPattern = /<(?:[A-Za-z_][\w.-]*:)?r\b/u
