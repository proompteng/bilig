import { strFromU8, strToU8, zipSync, type Unzipped } from 'fflate'

const xlsxWorksheetXmlPathPattern = /^xl\/worksheets\/[^/]+\.xml$/u

function cellAttributeValues(tag: string): Map<string, string> {
  const attributes = new Map<string, string>()
  for (const match of tag.matchAll(/\s([A-Za-z_:][\w:.-]*)=("[^"]*"|'[^']*')/gu)) {
    const name = match[1]
    const quotedValue = match[2]
    if (!name || !quotedValue) {
      continue
    }
    attributes.set(name, quotedValue.slice(1, -1))
  }
  return attributes
}

function isStyleOnlyBlankCellXml(cellXml: string): boolean {
  const openingTag = /^<c\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(cellXml)?.[0]
  if (!openingTag) {
    return false
  }
  const attributes = cellAttributeValues(openingTag)
  if (!attributes.has('r') || !attributes.has('s')) {
    return false
  }
  if (attributes.size === 2) {
    return true
  }
  return attributes.size === 3 && attributes.get('t') === 'z'
}

function stripStyleOnlyBlankCells(sheetXml: string): string {
  return sheetXml.replace(/<c\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<c\b(?:[^>"']|"[^"]*"|'[^']*')*>\s*<\/c>/gu, (cellXml) =>
    isStyleOnlyBlankCellXml(cellXml) ? '' : cellXml,
  )
}

export function stripStyleOnlyBlankCellsForSheetJs(data: Uint8Array, zip: Unzipped): Uint8Array {
  let changed = false
  for (const path of Object.keys(zip)) {
    if (!xlsxWorksheetXmlPathPattern.test(path)) {
      continue
    }
    const worksheetBytes = zip[path]
    if (!worksheetBytes) {
      continue
    }
    const worksheetXml = strFromU8(worksheetBytes)
    const strippedWorksheetXml = stripStyleOnlyBlankCells(worksheetXml)
    if (strippedWorksheetXml === worksheetXml) {
      continue
    }
    zip[path] = strToU8(strippedWorksheetXml)
    changed = true
  }
  return changed ? zipSync(zip) : data
}
