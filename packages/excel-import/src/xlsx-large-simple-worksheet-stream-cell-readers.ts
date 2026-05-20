import type { WorkbookRichTextCellSnapshot } from '@bilig/protocol'
import { parseLargeSimpleSharedFormulaIndex, readLargeSimpleFormulaTypeCode } from './xlsx-large-simple-formula-records.js'
import type { LargeSimpleSharedStrings } from './xlsx-large-simple-shared-strings.js'
import { stringItemText } from './xlsx-large-simple-worksheet-stream-text.js'
import { decodeBytes, encodeCellAddress } from './xlsx-large-simple-xml-byte-utils.js'
import { richTextRunPattern } from './xlsx-large-simple-worksheet-scan-constants.js'
import {
  findClosingTag,
  findNextOpeningTag,
  findTagEnd,
  isSelfClosingTag,
  readElementXml,
  readXmlAttributeFromTag,
  readXmlAttributeRangeFromTag,
} from './xlsx-large-simple-worksheet-stream-xml.js'

const extensionElementPattern = /<(?:[A-Za-z_][\w.-]*:)?ext\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?ext>)/gu
const slicerListElementPattern = /<(?:[A-Za-z_][\w.-]*:)?slicerList\b/u

export function readSlicerListExtensionXml(xml: string): string | undefined {
  extensionElementPattern.lastIndex = 0
  return [...xml.matchAll(extensionElementPattern)].find((match) => slicerListElementPattern.test(match[0]))?.[0]
}

export function readElementAttribute(xml: string, name: string): string | null {
  return new RegExp(`\\s${name}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

export function readPositiveIntegerAttributeFromTag(
  bytes: Uint8Array,
  nameEnd: number,
  tagEnd: number,
  attributeName: string,
): number | null {
  const range = readXmlAttributeRangeFromTag(bytes, nameEnd, tagEnd, attributeName)
  if (!range || range.start === range.end) {
    return null
  }
  let value = 0
  for (let index = range.start; index < range.end; index += 1) {
    const byte = bytes[index] ?? 0
    if (byte < 48 || byte > 57) {
      return null
    }
    value = value * 10 + byte - 48
  }
  return value > 0 && Number.isSafeInteger(value) ? value : null
}

export function readInlineStringCellValue(bytes: Uint8Array, contentStart: number, contentEnd: number): string | undefined {
  const inlineStringXml = readElementXml(bytes, contentStart, contentEnd, 'is')
  return inlineStringXml ? stringItemText(inlineStringXml) : undefined
}

export function readFormulaSpec(
  bytes: Uint8Array,
  contentStart: number,
  contentEnd: number,
  allowUnsupportedFormulaText: boolean,
): { readonly typeCode: number; readonly sharedIndex: number | null; readonly rawFormula: string } | null | undefined {
  const tag = findNextOpeningTag(bytes, contentStart, 'f', contentEnd)
  if (!tag) {
    return undefined
  }
  const tagEnd = findTagEnd(bytes, tag.nameEnd, contentEnd)
  if (tagEnd === null) {
    return null
  }
  const type = readXmlAttributeFromTag(bytes, tag.nameEnd, tagEnd, 't')
  if (!allowUnsupportedFormulaText && (type === 'array' || type === 'dataTable')) {
    return null
  }
  const selfClosing = isSelfClosingTag(bytes, tagEnd)
  const closing = selfClosing ? { start: tagEnd + 1, end: tagEnd + 1 } : findClosingTag(bytes, tagEnd + 1, 'f', contentEnd)
  if (!closing) {
    return null
  }
  return {
    typeCode: readLargeSimpleFormulaTypeCode(type),
    sharedIndex: parseLargeSimpleSharedFormulaIndex(readXmlAttributeFromTag(bytes, tag.nameEnd, tagEnd, 'si')),
    rawFormula: selfClosing ? '' : decodeBytes(bytes, tagEnd + 1, closing.start).trim(),
  }
}

export function readRichTextCellArtifact(
  bytes: Uint8Array,
  contentStart: number,
  contentEnd: number,
  row: number,
  column: number,
  type: string | null,
  sharedStringIndex: number | null,
  sharedStrings: LargeSimpleSharedStrings,
): WorkbookRichTextCellSnapshot | undefined {
  if (type === 's') {
    const entry = sharedStringIndex === null ? undefined : sharedStrings[sharedStringIndex]
    return entry?.rich
      ? {
          address: encodeCellAddress(row, column),
          text: entry.text,
          storage: 'sharedString',
          xml: entry.xml ?? '',
        }
      : undefined
  }
  if (type !== 'inlineStr') {
    return undefined
  }
  const inlineStringXml = readElementXml(bytes, contentStart, contentEnd, 'is')
  if (!inlineStringXml || !richTextRunPattern.test(inlineStringXml)) {
    return undefined
  }
  return {
    address: encodeCellAddress(row, column),
    text: stringItemText(inlineStringXml),
    storage: 'inlineString',
    xml: inlineStringXml,
  }
}
