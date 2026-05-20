import type { CellStyleFontSnapshot, CellStyleRecord } from '@bilig/protocol'

type ImportedCellStyle = Omit<CellStyleRecord, 'id'>

const elementTextCache = new Map<string, RegExp>()

export interface LargeSimpleWorkbookStylesScanOptions {
  readonly onRetainedBufferLength?: (length: number) => void
}

export function readLargeSimpleWorkbookStyles(
  stylesXml: string | null,
  requiredStyleIndexes: ReadonlySet<number>,
): Map<number, ImportedCellStyle> | null {
  if (!stylesXml || requiredStyleIndexes.size === 0) {
    return new Map()
  }
  const fills = readFillStyles(stylesXml)
  const fonts = readFontStyles(stylesXml)
  const cellXfsXml = extractElementXml(stylesXml, 'cellXfs')
  if (!fills || !fonts || !cellXfsXml) {
    return null
  }
  const cellXfs = [
    ...cellXfsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?xf\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?xf>)/gu),
  ]
  const styles = new Map<number, ImportedCellStyle>()
  for (const styleIndex of requiredStyleIndexes) {
    const xfXml = cellXfs[styleIndex]?.[0]
    if (!xfXml) {
      return null
    }
    const openingTag = /<(?:[A-Za-z_][\w.-]*:)?xf\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(xfXml)?.[0]
    if (!openingTag) {
      return null
    }
    const numFmtId = readNonNegativeIntegerAttribute(openingTag, 'numFmtId')
    if (numFmtId !== null && numFmtId !== 0) {
      return null
    }
    const fillId = readNonNegativeIntegerAttribute(openingTag, 'fillId')
    const fontId = readNonNegativeIntegerAttribute(openingTag, 'fontId')
    const borderId = readNonNegativeIntegerAttribute(openingTag, 'borderId')
    if (isApplied(openingTag, 'applyBorder', borderId) || hasChildElement(xfXml, 'alignment') || hasChildElement(xfXml, 'protection')) {
      return null
    }
    const fill = isApplied(openingTag, 'applyFill', fillId) ? fills[fillId ?? -1] : undefined
    const font = isApplied(openingTag, 'applyFont', fontId) ? fonts[fontId ?? -1] : undefined
    if (fill === null || font === null) {
      return null
    }
    const style: ImportedCellStyle = {
      ...(fill ? { fill } : {}),
      ...(font ? { font } : {}),
    }
    if (Object.keys(style).length > 0) {
      styles.set(styleIndex, style)
    }
  }
  return styles
}

export function readLargeSimpleWorkbookStylesFromChunks(
  readChunks: (onChunk: (chunk: Uint8Array) => void) => boolean,
  requiredStyleIndexes: ReadonlySet<number>,
  options: LargeSimpleWorkbookStylesScanOptions = {},
): Map<number, ImportedCellStyle> | null {
  if (requiredStyleIndexes.size === 0) {
    return new Map()
  }
  const cellXfs = collectIndexedXmlElementsFromChunks(readChunks, 'cellXfs', 'xf', requiredStyleIndexes, options)
  if (!cellXfs) {
    return null
  }
  const requiredFillIndexes = new Set<number>()
  const requiredFontIndexes = new Set<number>()
  const styleRefs = new Map<number, StyleComponentRefs>()
  for (const styleIndex of requiredStyleIndexes) {
    const xfXml = cellXfs.get(styleIndex)
    if (!xfXml) {
      return null
    }
    const refs = readStyleComponentRefs(xfXml)
    if (!refs) {
      return null
    }
    styleRefs.set(styleIndex, refs)
    if (refs.fillApplied && refs.fillId !== null) {
      requiredFillIndexes.add(refs.fillId)
    }
    if (refs.fontApplied && refs.fontId !== null) {
      requiredFontIndexes.add(refs.fontId)
    }
  }
  const fills = collectIndexedXmlElementsFromChunks(readChunks, 'fills', 'fill', requiredFillIndexes, options)
  const fonts = collectIndexedXmlElementsFromChunks(readChunks, 'fonts', 'font', requiredFontIndexes, options)
  if (!fills || !fonts) {
    return null
  }
  const styles = new Map<number, ImportedCellStyle>()
  for (const [styleIndex, refs] of styleRefs) {
    const fill = refs.fillApplied && refs.fillId !== null ? readFillStyle(fills.get(refs.fillId) ?? '') : undefined
    const font = refs.fontApplied && refs.fontId !== null ? readFontStyle(fonts.get(refs.fontId) ?? '') : undefined
    if (fill === null || font === null) {
      return null
    }
    const style: ImportedCellStyle = {
      ...(fill ? { fill } : {}),
      ...(font ? { font } : {}),
    }
    if (Object.keys(style).length > 0) {
      styles.set(styleIndex, style)
    }
  }
  return styles
}

interface StyleComponentRefs {
  readonly fillId: number | null
  readonly fontId: number | null
  readonly fillApplied: boolean
  readonly fontApplied: boolean
}

function readStyleComponentRefs(xfXml: string): StyleComponentRefs | null {
  const openingTag = /<(?:[A-Za-z_][\w.-]*:)?xf\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(xfXml)?.[0]
  if (!openingTag) {
    return null
  }
  const numFmtId = readNonNegativeIntegerAttribute(openingTag, 'numFmtId')
  if (numFmtId !== null && numFmtId !== 0) {
    return null
  }
  const fillId = readNonNegativeIntegerAttribute(openingTag, 'fillId')
  const fontId = readNonNegativeIntegerAttribute(openingTag, 'fontId')
  const borderId = readNonNegativeIntegerAttribute(openingTag, 'borderId')
  if (isApplied(openingTag, 'applyBorder', borderId) || hasChildElement(xfXml, 'alignment') || hasChildElement(xfXml, 'protection')) {
    return null
  }
  return {
    fillId,
    fontId,
    fillApplied: isApplied(openingTag, 'applyFill', fillId),
    fontApplied: isApplied(openingTag, 'applyFont', fontId),
  }
}

function readFillStyles(stylesXml: string): Array<ImportedCellStyle['fill'] | null> | null {
  const fillsXml = extractElementXml(stylesXml, 'fills')
  if (!fillsXml) {
    return []
  }
  return [...fillsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?fill\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?fill>/gu)].map((match) =>
    readFillStyle(match[0] ?? ''),
  )
}

function readFillStyle(fillXml: string): ImportedCellStyle['fill'] | null | undefined {
  const patternFill = extractElementXml(fillXml, 'patternFill')
  if (!patternFill) {
    return undefined
  }
  const openingTag = /<(?:[A-Za-z_][\w.-]*:)?patternFill\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(patternFill)?.[0]
  const patternType = openingTag ? readAttribute(openingTag, 'patternType') : undefined
  if (!patternType || patternType === 'none' || patternType === 'gray125') {
    return undefined
  }
  if (patternType !== 'solid') {
    return null
  }
  const color = readColor(patternFill, 'fgColor') ?? readColor(patternFill, 'bgColor')
  return color ? { backgroundColor: color } : undefined
}

function readFontStyles(stylesXml: string): Array<CellStyleFontSnapshot | null | undefined> | null {
  const fontsXml = extractElementXml(stylesXml, 'fonts')
  if (!fontsXml) {
    return []
  }
  return [...fontsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?font\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?font>/gu)].map((match) =>
    readFontStyle(match[0] ?? ''),
  )
}

function readFontStyle(fontXml: string): CellStyleFontSnapshot | null | undefined {
  const family = readElementValue(fontXml, 'name')
  const size = readElementNumberValue(fontXml, 'sz')
  const color = readColor(fontXml, 'color')
  const font: CellStyleFontSnapshot = {
    ...(family ? { family } : {}),
    ...(size ? { size } : {}),
    ...(hasBooleanElement(fontXml, 'b') ? { bold: true } : {}),
    ...(hasBooleanElement(fontXml, 'i') ? { italic: true } : {}),
    ...(hasBooleanElement(fontXml, 'u') ? { underline: true } : {}),
    ...(color ? { color } : {}),
  }
  return Object.keys(font).length > 0 ? font : undefined
}

function extractElementXml(xml: string, elementName: string): string | null {
  let pattern = elementTextCache.get(elementName)
  if (!pattern) {
    const qualifiedName = `(?:[A-Za-z_][\\w.-]*:)?${elementName}`
    pattern = new RegExp(`<${qualifiedName}\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/${qualifiedName}>)`, 'u')
    elementTextCache.set(elementName, pattern)
  }
  return pattern.exec(xml)?.[0] ?? null
}

function hasChildElement(xml: string, elementName: string): boolean {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${elementName}\\b`, 'u').test(xml.replace(/^<[^>]*>/u, ''))
}

function readElementValue(xml: string, elementName: string): string | undefined {
  const elementXml = extractElementXml(xml, elementName)
  if (!elementXml) {
    return undefined
  }
  return readAttribute(elementXml, 'val')
}

function readElementNumberValue(xml: string, elementName: string): number | undefined {
  const value = readElementValue(xml, elementName)
  if (value === undefined) {
    return undefined
  }
  const number = Number(value)
  return Number.isFinite(number) && number > 0 ? number : undefined
}

function hasBooleanElement(xml: string, elementName: string): boolean {
  const elementXml = extractElementXml(xml, elementName)
  if (!elementXml) {
    return false
  }
  const value = readAttribute(elementXml, 'val')
  return value === undefined || value === '1' || value.toLocaleLowerCase('en-US') === 'true'
}

function readColor(xml: string, elementName: string): string | undefined {
  const elementXml = extractElementXml(xml, elementName)
  const rgb = elementXml ? readAttribute(elementXml, 'rgb') : undefined
  if (!rgb) {
    return undefined
  }
  const normalized = rgb.trim().replace(/^#/, '')
  if (/^[0-9a-fA-F]{8}$/u.test(normalized)) {
    return `#${normalized.slice(2).toLocaleLowerCase('en-US')}`
  }
  if (/^[0-9a-fA-F]{6}$/u.test(normalized)) {
    return `#${normalized.toLocaleLowerCase('en-US')}`
  }
  return undefined
}

function collectIndexedXmlElementsFromChunks(
  readChunks: (onChunk: (chunk: Uint8Array) => void) => boolean,
  parentName: string,
  childName: string,
  requiredIndexes: ReadonlySet<number>,
  options: LargeSimpleWorkbookStylesScanOptions,
): Map<number, string> | null {
  if (requiredIndexes.size === 0) {
    return new Map()
  }
  const collector = new IndexedXmlElementCollector(parentName, childName, requiredIndexes, options)
  if (!readChunks((chunk) => collector.push(chunk))) {
    return null
  }
  return collector.finish()
}

class IndexedXmlElementCollector {
  private readonly decoder = new TextDecoder()
  private buffer = ''
  private index = 0
  private inParent = false
  private childIndex = 0
  private failed = false
  private skippingUnrequiredChild = false
  private readonly elements = new Map<number, string>()

  constructor(
    private readonly parentName: string,
    private readonly childName: string,
    private readonly requiredIndexes: ReadonlySet<number>,
    private readonly options: LargeSimpleWorkbookStylesScanOptions,
  ) {}

  push(chunk: Uint8Array): void {
    if (this.failed || chunk.byteLength === 0 || this.isComplete()) {
      return
    }
    this.buffer += this.decoder.decode(chunk, { stream: true })
    this.process(false)
    this.releaseBufferIfComplete()
    if (!this.isComplete()) {
      this.compact()
    }
    this.reportRetainedBufferLength()
  }

  finish(): Map<number, string> | null {
    if (this.failed) {
      return null
    }
    if (!this.isComplete()) {
      this.buffer += this.decoder.decode()
      this.process(true)
    }
    this.releaseBufferIfComplete()
    if (!this.isComplete()) {
      this.compact()
    }
    this.reportRetainedBufferLength()
    if (this.failed || this.elements.size !== this.requiredIndexes.size) {
      return null
    }
    return this.elements
  }

  private process(final: boolean): void {
    while (!this.failed && this.elements.size < this.requiredIndexes.size) {
      if (this.skippingUnrequiredChild) {
        if (!this.finishSkippingUnrequiredChild(final)) {
          return
        }
        continue
      }
      if (!this.inParent) {
        const parent = findNextOpeningTag(this.buffer, this.index, this.parentName)
        if (!parent) {
          this.index = Math.max(0, this.buffer.length - this.parentName.length - 4)
          return
        }
        const tagEnd = findStringTagEnd(this.buffer, parent.nameEnd)
        if (tagEnd === null) {
          if (final) {
            this.failed = true
          }
          this.index = parent.start
          return
        }
        if (isSelfClosingStringTag(this.buffer, tagEnd)) {
          this.index = tagEnd + 1
          continue
        }
        this.inParent = true
        this.childIndex = 0
        this.index = tagEnd + 1
        continue
      }
      const next = findNextParentBoundaryOrChild(this.buffer, this.index, this.parentName, this.childName)
      if (!next) {
        this.index = Math.max(0, this.buffer.length - Math.max(this.parentName.length, this.childName.length) - 4)
        return
      }
      if (next.kind === 'parent-close') {
        const tagEnd = findStringTagEnd(this.buffer, next.nameEnd)
        if (tagEnd === null) {
          if (final) {
            this.failed = true
          }
          this.index = next.start
          return
        }
        this.inParent = false
        this.index = tagEnd + 1
        continue
      }
      const tagEnd = findStringTagEnd(this.buffer, next.nameEnd)
      if (tagEnd === null) {
        if (final) {
          this.failed = true
        }
        this.index = next.start
        return
      }
      const childStart = next.start
      const childEnd = isSelfClosingStringTag(this.buffer, tagEnd)
        ? tagEnd + 1
        : findClosingStringElementEnd(this.buffer, tagEnd + 1, this.childName)
      if (childEnd === null) {
        if (final) {
          this.failed = true
        }
        if (this.requiredIndexes.has(this.childIndex)) {
          this.index = childStart
        } else {
          this.skippingUnrequiredChild = true
          this.index = tagEnd + 1
        }
        return
      }
      if (this.requiredIndexes.has(this.childIndex)) {
        this.elements.set(this.childIndex, this.buffer.slice(childStart, childEnd))
      }
      this.childIndex += 1
      this.index = childEnd
    }
  }

  private isComplete(): boolean {
    return this.elements.size === this.requiredIndexes.size
  }

  private releaseBufferIfComplete(): void {
    if (!this.isComplete()) {
      return
    }
    this.buffer = ''
    this.index = 0
    this.skippingUnrequiredChild = false
  }

  private finishSkippingUnrequiredChild(final: boolean): boolean {
    const childEnd = findClosingStringElementEnd(this.buffer, this.index, this.childName)
    if (childEnd === null) {
      if (final) {
        this.failed = true
      }
      this.index = this.buffer.length
      return false
    }
    this.childIndex += 1
    this.index = childEnd
    this.skippingUnrequiredChild = false
    return true
  }

  private compact(): void {
    if (this.skippingUnrequiredChild) {
      const retainLength = closingStringTagRetainLength(this.childName)
      this.buffer = this.buffer.slice(Math.max(0, this.buffer.length - retainLength))
      this.index = 0
      return
    }
    if (this.index === 0) {
      return
    }
    if (this.index >= this.buffer.length) {
      this.buffer = ''
      this.index = 0
      return
    }
    this.buffer = this.buffer.slice(this.index)
    this.index = 0
  }

  private reportRetainedBufferLength(): void {
    this.options.onRetainedBufferLength?.(this.buffer.length)
  }
}

function closingStringTagRetainLength(elementName: string): number {
  return Math.max(256, elementName.length + 4)
}

function findNextParentBoundaryOrChild(
  xml: string,
  startIndex: number,
  parentName: string,
  childName: string,
): { readonly kind: 'child' | 'parent-close'; readonly start: number; readonly nameEnd: number } | null {
  let index = startIndex
  while (index < xml.length) {
    const tagStart = xml.indexOf('<', index)
    if (tagStart < 0) {
      return null
    }
    if (xml.charCodeAt(tagStart + 1) === 47) {
      const tag = readStringTagName(xml, tagStart + 2)
      if (tag?.localName === parentName) {
        return { kind: 'parent-close', start: tagStart, nameEnd: tag.endIndex }
      }
      index = tagStart + 1
      continue
    }
    const tag = readStringTagName(xml, tagStart + 1)
    if (tag?.localName === childName) {
      return { kind: 'child', start: tagStart, nameEnd: tag.endIndex }
    }
    index = tagStart + 1
  }
  return null
}

function findNextOpeningTag(
  xml: string,
  startIndex: number,
  localName: string,
): { readonly start: number; readonly nameEnd: number } | null {
  let index = startIndex
  while (index < xml.length) {
    const tagStart = xml.indexOf('<', index)
    if (tagStart < 0) {
      return null
    }
    const tag = readStringTagName(xml, tagStart + 1)
    if (tag?.localName === localName) {
      return { start: tagStart, nameEnd: tag.endIndex }
    }
    index = tagStart + 1
  }
  return null
}

function findClosingStringElementEnd(xml: string, startIndex: number, localName: string): number | null {
  let index = startIndex
  while (index < xml.length) {
    const tagStart = xml.indexOf('</', index)
    if (tagStart < 0) {
      return null
    }
    const tag = readStringTagName(xml, tagStart + 2)
    if (tag?.localName === localName) {
      const tagEnd = findStringTagEnd(xml, tag.endIndex)
      return tagEnd === null ? null : tagEnd + 1
    }
    index = tagStart + 2
  }
  return null
}

function findStringTagEnd(xml: string, startIndex: number): number | null {
  let quote: string | null = null
  for (let index = startIndex; index < xml.length; index += 1) {
    const char = xml[index]
    if (quote !== null) {
      if (char === quote) {
        quote = null
      }
      continue
    }
    if (char === '"' || char === "'") {
      quote = char
      continue
    }
    if (char === '>') {
      return index
    }
  }
  return null
}

function isSelfClosingStringTag(xml: string, tagEnd: number): boolean {
  let index = tagEnd - 1
  while (index >= 0 && /\s/u.test(xml[index] ?? '')) {
    index -= 1
  }
  return xml[index] === '/'
}

function readStringTagName(xml: string, startIndex: number): { readonly localName: string; readonly endIndex: number } | null {
  const first = xml.charCodeAt(startIndex)
  if (!Number.isFinite(first) || first === 33 || first === 47 || first === 63) {
    return null
  }
  let index = startIndex
  let localNameStart = startIndex
  while (index < xml.length && isXmlNameChar(xml[index] ?? '')) {
    if (xml[index] === ':') {
      localNameStart = index + 1
    }
    index += 1
  }
  return index === localNameStart ? null : { localName: xml.slice(localNameStart, index), endIndex: index }
}

function isXmlNameChar(char: string): boolean {
  return /[A-Za-z0-9_.:-]/u.test(char)
}

function isApplied(tag: string, attributeName: string, componentId: number | null): boolean {
  const value = readAttribute(tag, attributeName)
  if (value === '1' || value?.toLocaleLowerCase('en-US') === 'true') {
    return true
  }
  if (value === '0' || value?.toLocaleLowerCase('en-US') === 'false') {
    return false
  }
  return componentId !== null && componentId > 0
}

function readNonNegativeIntegerAttribute(tag: string, attributeName: string): number | null {
  const value = readAttribute(tag, attributeName)
  if (!value || !/^[0-9]+$/u.test(value)) {
    return null
  }
  const number = Number(value)
  return Number.isSafeInteger(number) ? number : null
}

function readAttribute(tag: string, attributeName: string): string | undefined {
  const pattern = new RegExp(`\\s${escapeRegExp(attributeName)}=(?:"([^"]*)"|'([^']*)')`, 'u')
  const match = pattern.exec(tag)
  return match?.[1] ?? match?.[2]
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}
