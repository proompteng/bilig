export function readKnownXmlLocalName(bytes: Uint8Array, startIndex: number, endIndex: number): string | null {
  switch (endIndex - startIndex) {
    case 1:
      if (bytes[startIndex] === 99) {
        return 'c'
      }
      if (bytes[startIndex] === 102) {
        return 'f'
      }
      if (bytes[startIndex] === 114) {
        return 'r'
      }
      if (bytes[startIndex] === 116) {
        return 't'
      }
      return bytes[startIndex] === 118 ? 'v' : null
    case 2:
      return matchesAscii(bytes, startIndex, endIndex, 'is') ? 'is' : null
    case 3:
      if (matchesAscii(bytes, startIndex, endIndex, 'col')) {
        return 'col'
      }
      return matchesAscii(bytes, startIndex, endIndex, 'row') ? 'row' : null
    case 4:
      return matchesAscii(bytes, startIndex, endIndex, 'cols') ? 'cols' : null
    case 7:
      return matchesAscii(bytes, startIndex, endIndex, 'drawing') ? 'drawing' : null
    case 9:
      if (matchesAscii(bytes, startIndex, endIndex, 'dimension')) {
        return 'dimension'
      }
      if (matchesAscii(bytes, startIndex, endIndex, 'colBreaks')) {
        return 'colBreaks'
      }
      if (matchesAscii(bytes, startIndex, endIndex, 'mergeCell')) {
        return 'mergeCell'
      }
      if (matchesAscii(bytes, startIndex, endIndex, 'pageSetup')) {
        return 'pageSetup'
      }
      if (matchesAscii(bytes, startIndex, endIndex, 'rowBreaks')) {
        return 'rowBreaks'
      }
      if (matchesAscii(bytes, startIndex, endIndex, 'sheetData')) {
        return 'sheetData'
      }
      if (matchesAscii(bytes, startIndex, endIndex, 'tablePart')) {
        return 'tablePart'
      }
      return matchesAscii(bytes, startIndex, endIndex, 'worksheet') ? 'worksheet' : null
    case 10:
      if (matchesAscii(bytes, startIndex, endIndex, 'autoFilter')) {
        return 'autoFilter'
      }
      if (matchesAscii(bytes, startIndex, endIndex, 'hyperlinks')) {
        return 'hyperlinks'
      }
      if (matchesAscii(bytes, startIndex, endIndex, 'mergeCells')) {
        return 'mergeCells'
      }
      if (matchesAscii(bytes, startIndex, endIndex, 'tableParts')) {
        return 'tableParts'
      }
      return null
    case 11:
      return matchesAscii(bytes, startIndex, endIndex, 'pageMargins') ? 'pageMargins' : null
    case 12:
      if (matchesAscii(bytes, startIndex, endIndex, 'headerFooter')) {
        return 'headerFooter'
      }
      return matchesAscii(bytes, startIndex, endIndex, 'printOptions') ? 'printOptions' : null
    case 13:
      return matchesAscii(bytes, startIndex, endIndex, 'sheetFormatPr') ? 'sheetFormatPr' : null
    case 15:
      if (matchesAscii(bytes, startIndex, endIndex, 'dataValidations')) {
        return 'dataValidations'
      }
      return matchesAscii(bytes, startIndex, endIndex, 'sheetProtection') ? 'sheetProtection' : null
    case 21:
      return matchesAscii(bytes, startIndex, endIndex, 'conditionalFormatting') ? 'conditionalFormatting' : null
    default:
      return null
  }
}

function matchesAscii(bytes: Uint8Array, startIndex: number, endIndex: number, value: string): boolean {
  if (endIndex - startIndex !== value.length) {
    return false
  }
  for (let index = 0; index < value.length; index += 1) {
    if (bytes[startIndex + index] !== value.charCodeAt(index)) {
      return false
    }
  }
  return true
}
