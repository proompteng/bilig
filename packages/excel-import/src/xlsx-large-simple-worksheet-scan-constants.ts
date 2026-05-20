export const unsupportedWorksheetTagNames = new Set(['oleObjects', 'picture'])

export const metadataWorksheetTagNames = new Set([
  'autoFilter',
  'colBreaks',
  'cols',
  'conditionalFormatting',
  'drawing',
  'extLst',
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
