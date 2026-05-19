import { toLiteralInput } from './workbook-import-helpers.js'
import { decodeExcelEscapedText } from './xlsx-escaped-text.js'

export function decodeXmlText(value: string): string {
  if (!value.includes('&')) {
    return value
  }
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/gu, (_match, entity: string) => {
    if (entity.startsWith('#x')) {
      const codePoint = Number.parseInt(entity.slice(2), 16)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    if (entity.startsWith('#')) {
      const codePoint = Number.parseInt(entity.slice(1), 10)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    switch (entity) {
      case 'amp':
        return '&'
      case 'lt':
        return '<'
      case 'gt':
        return '>'
      case 'quot':
        return '"'
      case 'apos':
        return "'"
      default:
        return ''
    }
  })
}

export function stringItemText(xml: string): string {
  return normalizeWorksheetText(
    [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/gu)]
      .map((match) => decodeXmlText(match[1] ?? ''))
      .join(''),
  )
}

export function normalizeWorksheetText(value: string): string {
  const literal = toLiteralInput(decodeExcelEscapedText(value))
  return typeof literal === 'string' ? literal : value
}

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}
