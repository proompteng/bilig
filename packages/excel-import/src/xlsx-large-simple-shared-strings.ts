import type { WorkbookRichTextCellSnapshot } from '@bilig/protocol'
import { decodeExcelEscapedText } from './xlsx-escaped-text.js'

export interface LargeSimpleSharedStringEntry {
  readonly text: string
  readonly xml?: string
  readonly rich: boolean
}

const sharedStringElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?si\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?si)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/gu
const richTextRunPattern = /<(?:[A-Za-z_][\w.-]*:)?r\b/u

export function readLargeSimpleSharedStrings(sharedStringsXml: string): LargeSimpleSharedStringEntry[] {
  return [...sharedStringsXml.matchAll(sharedStringElementPattern)].map((match) => {
    const xml = match[0]
    const rich = richTextRunPattern.test(xml)
    const entry: LargeSimpleSharedStringEntry = { text: stringItemText(xml), rich }
    if (rich) {
      Object.assign(entry, { xml })
    }
    return entry
  })
}

export function readLargeSimpleRichTextCellArtifact(
  address: string,
  openingTag: string,
  cellXml: string,
  sharedStrings: readonly LargeSimpleSharedStringEntry[],
): WorkbookRichTextCellSnapshot | undefined {
  const type = readXmlAttribute(openingTag, 't')
  if (type === 's') {
    const entry = sharedStrings[readSharedStringIndex(cellXml) ?? -1]
    return entry?.rich
      ? {
          address,
          text: entry.text,
          storage: 'sharedString',
          xml: entry.xml ?? '',
        }
      : undefined
  }
  if (type === 'inlineStr') {
    const inlineStringXml = readStringElement(cellXml, 'is')
    if (inlineStringXml && richTextRunPattern.test(inlineStringXml)) {
      return {
        address,
        text: stringItemText(inlineStringXml),
        storage: 'inlineString',
        xml: inlineStringXml,
      }
    }
  }
  return undefined
}

function readSharedStringIndex(cellXml: string): number | null {
  const rawValue = readElementText(cellXml, 'v')?.trim()
  if (!rawValue) {
    return null
  }
  const index = Number(decodeXmlText(rawValue))
  return Number.isSafeInteger(index) && index >= 0 ? index : null
}

function readStringElement(xml: string, elementName: 'is'): string | null {
  return new RegExp(`<((?:[A-Za-z_][\\w.-]*:)?${elementName})\\b[^>]*(?:/>|>[\\s\\S]*?</\\1>)`, 'u').exec(xml)?.[0] ?? null
}

function readElementText(xml: string, elementName: 'v'): string | null {
  return (
    new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${elementName}\\b[^>]*>([\\s\\S]*?)</(?:[A-Za-z_][\\w.-]*:)?${elementName}>`, 'u').exec(xml)?.[1] ??
    null
  )
}

function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function stringItemText(xml: string): string {
  return [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/gu)]
    .map((match) => decodeExcelEscapedText(decodeXmlText(match[1] ?? '')))
    .join('')
}

function decodeXmlText(value: string): string {
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

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}
