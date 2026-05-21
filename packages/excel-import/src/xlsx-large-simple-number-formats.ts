import * as XLSX from 'xlsx'

import type { WorkbookSnapshot } from '@bilig/protocol'
import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import { encodeCellAddress, packArenaCellAddress } from './xlsx-large-simple-arena-helpers.js'
import type { LargeSimpleWorkbookStylesScanOptions } from './xlsx-large-simple-styles.js'
import { decodeCellAddress } from './xlsx-large-simple-xml-byte-utils.js'

export function readLargeSimpleWorkbookNumberFormatsFromChunks(
  readChunks: (onChunk: (chunk: Uint8Array) => void) => boolean,
  requiredStyleIndexes: ReadonlySet<number>,
  options: LargeSimpleWorkbookStylesScanOptions = {},
): Map<number, string> | null {
  if (requiredStyleIndexes.size === 0) {
    return new Map()
  }
  const cellXfs = collectIndexedXmlElementsFromChunks(readChunks, 'cellXfs', 'xf', requiredStyleIndexes, options)
  if (!cellXfs) {
    return null
  }
  const styleFormatIds = new Map<number, number>()
  const requiredCustomFormatIds = new Set<number>()
  for (const styleIndex of requiredStyleIndexes) {
    const xfXml = cellXfs.get(styleIndex)
    if (!xfXml) {
      return null
    }
    const openingTag = /<(?:[A-Za-z_][\w.-]*:)?xf\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(xfXml)?.[0]
    if (!openingTag) {
      return null
    }
    const formatId = readNonNegativeIntegerAttribute(openingTag, 'numFmtId')
    if (formatId === null || formatId === 0) {
      continue
    }
    styleFormatIds.set(styleIndex, formatId)
    if (!builtinNumberFormatCode(formatId)) {
      requiredCustomFormatIds.add(formatId)
    }
  }
  if (styleFormatIds.size === 0) {
    return new Map()
  }
  const customFormats = collectNumberFormatsFromChunks(readChunks, requiredCustomFormatIds, options)
  if (!customFormats) {
    return null
  }
  const formatsByStyleIndex = new Map<number, string>()
  for (const [styleIndex, formatId] of styleFormatIds) {
    const format = builtinNumberFormatCode(formatId) ?? customFormats.get(formatId)
    if (format) {
      formatsByStyleIndex.set(styleIndex, format)
    }
  }
  return formatsByStyleIndex
}

export function applyLargeSimpleNumberFormatsToCells(
  cells: WorkbookSnapshot['sheets'][number]['cells'],
  cellScan: ImportedWorksheetCellScan,
  numberFormatsByStyleIndex: ReadonlyMap<number, string>,
): void {
  const cellsByPackedAddress = new Map<number, WorkbookSnapshot['sheets'][number]['cells'][number]>()
  for (const cell of cells) {
    const decoded = decodeCellAddress(cell.address)
    if (!decoded) {
      continue
    }
    cellsByPackedAddress.set(packArenaCellAddress(decoded.row, decoded.column), cell)
  }
  let addedCellCount = 0
  cellScan.styleIndexes.forEach((row, column, styleIndex) => {
    const format = numberFormatsByStyleIndex.get(styleIndex)
    if (!format) {
      return
    }
    const packedAddress = packArenaCellAddress(row, column)
    const existingCell = cellsByPackedAddress.get(packedAddress)
    if (existingCell) {
      existingCell.format ??= format
      return
    }
    const cell: WorkbookSnapshot['sheets'][number]['cells'][number] = {
      address: encodeCellAddress(row, column),
      format,
    }
    cellsByPackedAddress.set(packedAddress, cell)
    cells.push(cell)
    addedCellCount += 1
  })
  if (addedCellCount > 0) {
    cells.sort(compareLargeSimpleSnapshotCells)
  }
}

function compareLargeSimpleSnapshotCells(
  left: WorkbookSnapshot['sheets'][number]['cells'][number],
  right: WorkbookSnapshot['sheets'][number]['cells'][number],
): number {
  const leftDecoded = decodeCellAddress(left.address)
  const rightDecoded = decodeCellAddress(right.address)
  if (!leftDecoded || !rightDecoded) {
    return left.address.localeCompare(right.address)
  }
  return leftDecoded.row - rightDecoded.row || leftDecoded.column - rightDecoded.column
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

function collectNumberFormatsFromChunks(
  readChunks: (onChunk: (chunk: Uint8Array) => void) => boolean,
  requiredFormatIds: ReadonlySet<number>,
  options: LargeSimpleWorkbookStylesScanOptions,
): Map<number, string> | null {
  if (requiredFormatIds.size === 0) {
    return new Map()
  }
  const collector = new NumberFormatCollector(requiredFormatIds, options)
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

class NumberFormatCollector {
  private readonly decoder = new TextDecoder()
  private buffer = ''
  private index = 0
  private inParent = false
  private failed = false
  private skippingElement = false
  private readonly formats = new Map<number, string>()

  constructor(
    private readonly requiredFormatIds: ReadonlySet<number>,
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
    return this.failed || this.formats.size !== this.requiredFormatIds.size ? null : this.formats
  }

  private process(final: boolean): void {
    while (!this.failed && this.formats.size < this.requiredFormatIds.size) {
      if (this.skippingElement) {
        if (!this.finishSkippingElement(final)) {
          return
        }
        continue
      }
      if (!this.inParent) {
        const parent = findNextOpeningTag(this.buffer, this.index, 'numFmts')
        if (!parent) {
          this.index = Math.max(0, this.buffer.length - 'numFmts'.length - 4)
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
          return
        }
        this.inParent = true
        this.index = tagEnd + 1
        continue
      }
      const next = findNextParentBoundaryOrChild(this.buffer, this.index, 'numFmts', 'numFmt')
      if (!next) {
        this.index = Math.max(0, this.buffer.length - 'numFmts'.length - 4)
        return
      }
      const tagEnd = findStringTagEnd(this.buffer, next.nameEnd)
      if (tagEnd === null) {
        if (final) {
          this.failed = true
        }
        this.index = next.start
        return
      }
      if (next.kind === 'parent-close') {
        this.inParent = false
        this.index = tagEnd + 1
        return
      }
      const childStart = next.start
      const childEnd = isSelfClosingStringTag(this.buffer, tagEnd)
        ? tagEnd + 1
        : findClosingStringElementEnd(this.buffer, tagEnd + 1, 'numFmt')
      if (childEnd === null) {
        if (final) {
          this.failed = true
        }
        this.skippingElement = true
        this.index = tagEnd + 1
        return
      }
      this.collectFormat(this.buffer.slice(childStart, childEnd))
      this.index = childEnd
    }
  }

  private collectFormat(elementXml: string): void {
    const tagEnd = findStringTagEnd(elementXml, 0)
    const openingTag = tagEnd === null ? elementXml : elementXml.slice(0, tagEnd + 1)
    const id = readNonNegativeIntegerAttribute(openingTag, 'numFmtId')
    if (id === null || !this.requiredFormatIds.has(id)) {
      return
    }
    const format = normalizeNumberFormatCode(decodeXmlAttribute(readAttribute(openingTag, 'formatCode') ?? ''))
    if (format) {
      this.formats.set(id, format)
    }
  }

  private isComplete(): boolean {
    return this.formats.size === this.requiredFormatIds.size
  }

  private releaseBufferIfComplete(): void {
    if (!this.isComplete()) {
      return
    }
    this.buffer = ''
    this.index = 0
    this.skippingElement = false
  }

  private finishSkippingElement(final: boolean): boolean {
    const childEnd = findClosingStringElementEnd(this.buffer, this.index, 'numFmt')
    if (childEnd === null) {
      if (final) {
        this.failed = true
      }
      this.index = this.buffer.length
      return false
    }
    this.collectFormat(this.buffer.slice(0, childEnd))
    this.index = childEnd
    this.skippingElement = false
    return true
  }

  private compact(): void {
    if (this.skippingElement) {
      const retainLength = closingStringTagRetainLength('numFmt')
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

function builtinNumberFormatCode(formatId: number): string | undefined {
  const ssf: unknown = XLSX.SSF
  if (!isRecord(ssf) || !isRecord(ssf['_table'])) {
    return undefined
  }
  return normalizeNumberFormatCode(ssf['_table'][String(formatId)])
}

function normalizeNumberFormatCode(value: unknown): string | undefined {
  if (typeof value !== 'string') {
    return undefined
  }
  const trimmed = value.trim()
  return trimmed.length > 0 && trimmed !== 'General' ? trimmed : undefined
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

function decodeXmlAttribute(value: string): string {
  return value
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&amp;/gu, '&')
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}
