import { strFromU8, strToU8, zipSync, type Unzipped } from 'fflate'

const xlsxWorksheetXmlPathPattern = /^xl\/worksheets\/[^/]+\.xml$/u
const xmlAttributePattern = /\s([A-Za-z_:][\w:.-]*)=("[^"]*"|'[^']*')/gu
const commonStyleOnlyBlankCellPattern = /^<c r=(?:"[^"]*"|'[^']*') s=(?:"[^"]*"|'[^']*')(?: t=(?:"z"|'z'))?(?:\/>|>)$/u
const commonSelfClosingStyleOnlyBlankCellPattern = /<c r=(?:"[^"]*"|'[^']*') s=(?:"[^"]*"|'[^']*')(?: t=(?:"z"|'z'))?\s*\/>/gu
const commonExpandedStyleOnlyBlankCellPattern = /<c r=(?:"[^"]*"|'[^']*') s=(?:"[^"]*"|'[^']*')(?: t=(?:"z"|'z'))?>\s*<\/c>/gu
const commonNoOpEmptyRowPattern = /^<row r=(?:"[^"]*"|'[^']*')(?: spans=(?:"[^"]*"|'[^']*'))?(?:\/>|>)$/u
const commonSelfClosingNoOpEmptyRowPattern = /<row r=(?:"[^"]*"|'[^']*')(?: spans=(?:"[^"]*"|'[^']*'))?\s*\/>/gu
const commonExpandedNoOpEmptyRowPattern = /<row r=(?:"[^"]*"|'[^']*')(?: spans=(?:"[^"]*"|'[^']*'))?>\s*<\/row>/gu
const possibleStyleOnlyBlankCellPattern = /<c\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<c\b(?:[^>"']|"[^"]*"|'[^']*')*>\s*<\/c>/u
const possibleNoOpEmptyRowPattern = /<row\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<row\b(?:[^>"']|"[^"]*"|'[^']*')*>\s*<\/row>/u

function forEachXmlAttribute(tag: string, visit: (name: string, value: string) => boolean): boolean {
  xmlAttributePattern.lastIndex = 0
  let match: RegExpExecArray | null
  while ((match = xmlAttributePattern.exec(tag)) !== null) {
    const name = match[1]
    const quotedValue = match[2]
    if (!name || !quotedValue || !visit(name, quotedValue.slice(1, -1))) {
      xmlAttributePattern.lastIndex = 0
      return false
    }
  }
  xmlAttributePattern.lastIndex = 0
  return true
}

function xmlAttributeValue(tag: string, attributeName: string): string | undefined {
  let value: string | undefined
  forEachXmlAttribute(tag, (name, nextValue) => {
    if (name === attributeName) {
      value = nextValue
      return false
    }
    return true
  })
  return value
}

function readSheetDefaultRowHeight(sheetXml: string): string | undefined {
  const sheetFormatTag = /<sheetFormatPr\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(sheetXml)?.[0]
  if (!sheetFormatTag) {
    return undefined
  }
  return xmlAttributeValue(sheetFormatTag, 'defaultRowHeight')
}

function isStyleOnlyBlankCellXml(cellXml: string): boolean {
  const openingTag = /^<c\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(cellXml)?.[0]
  if (!openingTag) {
    return false
  }
  if (commonStyleOnlyBlankCellPattern.test(openingTag)) {
    return true
  }
  let hasAddress = false
  let hasStyle = false
  let type: string | undefined
  let attributeCount = 0
  return (
    forEachXmlAttribute(openingTag, (name, value) => {
      attributeCount += 1
      switch (name) {
        case 'r':
          hasAddress = true
          return true
        case 's':
          hasStyle = true
          return true
        case 't':
          type = value
          return true
        default:
          return false
      }
    }) &&
    hasAddress &&
    hasStyle &&
    (attributeCount === 2 || (attributeCount === 3 && type === 'z'))
  )
}

function isNoOpEmptyRowXml(rowXml: string, defaultRowHeight: string | undefined): boolean {
  const openingTag = /^<row\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(rowXml)?.[0]
  if (!openingTag) {
    return false
  }
  if (commonNoOpEmptyRowPattern.test(openingTag)) {
    return true
  }
  if (/<c\b/u.test(rowXml)) {
    return false
  }
  let customHeight: string | undefined
  let height: string | undefined
  if (
    !forEachXmlAttribute(openingTag, (name, value) => {
      switch (name) {
        case 'r':
        case 'spans':
        case 'x14ac:dyDescent':
          return true
        case 'customHeight':
          customHeight = value
          return true
        case 'ht':
          height = value
          return true
        default:
          return false
      }
    })
  ) {
    return false
  }
  const hasDefaultHeight = height !== undefined && height === defaultRowHeight
  return (height === undefined || hasDefaultHeight) && (customHeight === undefined || (customHeight === '1' && hasDefaultHeight))
}

function stripNoOpEmptyRows(sheetXml: string): string {
  const withoutCommonRows = sheetXml.replace(commonSelfClosingNoOpEmptyRowPattern, '').replace(commonExpandedNoOpEmptyRowPattern, '')
  if (!possibleNoOpEmptyRowPattern.test(withoutCommonRows)) {
    return withoutCommonRows
  }
  const defaultRowHeight = readSheetDefaultRowHeight(sheetXml)
  return withoutCommonRows.replace(/<row\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<row\b(?:[^>"']|"[^"]*"|'[^']*')*>\s*<\/row>/gu, (rowXml) =>
    isNoOpEmptyRowXml(rowXml, defaultRowHeight) ? '' : rowXml,
  )
}

function stripStyleOnlyBlankCells(sheetXml: string): string {
  const withoutCommonBlankCells = sheetXml
    .replace(commonSelfClosingStyleOnlyBlankCellPattern, '')
    .replace(commonExpandedStyleOnlyBlankCellPattern, '')
  const withoutBlankCells = possibleStyleOnlyBlankCellPattern.test(withoutCommonBlankCells)
    ? withoutCommonBlankCells.replace(/<c\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<c\b(?:[^>"']|"[^"]*"|'[^']*')*>\s*<\/c>/gu, (cellXml) =>
        isStyleOnlyBlankCellXml(cellXml) ? '' : cellXml,
      )
    : withoutCommonBlankCells
  return stripNoOpEmptyRows(withoutBlankCells)
}

function stripWorkbookWorksheetXml(data: Uint8Array, zip: Unzipped, stripWorksheet: (sheetXml: string) => string): Uint8Array {
  let outputZip: Unzipped | null = null
  for (const path of Object.keys(zip)) {
    if (!xlsxWorksheetXmlPathPattern.test(path)) {
      continue
    }
    const worksheetBytes = zip[path]
    if (!worksheetBytes) {
      continue
    }
    const worksheetXml = strFromU8(worksheetBytes)
    const strippedWorksheetXml = stripWorksheet(worksheetXml)
    if (strippedWorksheetXml === worksheetXml) {
      continue
    }
    outputZip ??= { ...zip }
    outputZip[path] = strToU8(strippedWorksheetXml)
  }
  return outputZip ? zipSync(outputZip) : data
}

export function stripNoOpEmptyRowsFromXlsx(data: Uint8Array, zip: Unzipped): Uint8Array {
  return stripWorkbookWorksheetXml(data, zip, stripNoOpEmptyRows)
}

export function stripStyleOnlyBlankCellsForSheetJs(data: Uint8Array, zip: Unzipped): Uint8Array {
  return stripWorkbookWorksheetXml(data, zip, stripStyleOnlyBlankCells)
}
