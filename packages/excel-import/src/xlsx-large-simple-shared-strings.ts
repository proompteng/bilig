import type { WorkbookRichTextCellSnapshot } from '@bilig/protocol'
import { decodeExcelEscapedText } from './xlsx-escaped-text.js'
import type { ImportedWorkbookArenaDedupeMode } from './xlsx-large-simple-arena-types.js'
import type { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'

export interface LargeSimpleSharedStringEntry {
  readonly text: string
  readonly xml?: string
  readonly rich: boolean
}

export interface LargeSimpleSharedStringTable {
  readonly length: number
  readonly [index: number]: LargeSimpleSharedStringEntry | undefined
}

export type LargeSimpleSharedStrings = readonly LargeSimpleSharedStringEntry[] | LargeSimpleSharedStringTable

const sharedStringElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?si\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?si)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/gu
const richTextRunPattern = /<(?:[A-Za-z_][\w.-]*:)?r\b/u
const partialSharedStringTagRetainLength = 256
const sparseReferencedSharedStringDensityDivisor = 64

export interface LargeSimpleReferencedSharedStringScanOptions {
  readonly onRetainedBufferLength?: (length: number) => void
  readonly stringPool?: ImportedWorkbookStringPool
  readonly deduplicateText?: ImportedWorkbookArenaDedupeMode
  readonly dedupeMaxEntries?: number
}

interface LargeSimpleSharedStringReadOptions extends LargeSimpleReferencedSharedStringScanOptions {
  readonly plainEntryPool?: LargeSimpleSharedStringEntryPool
}

interface LargeSimpleSharedStringEntryPool {
  readonly entries: Map<string, LargeSimpleSharedStringEntry>
  readonly keys: string[]
  evictionIndex: number
}

export function readLargeSimpleSharedStrings(
  sharedStringsXml: string,
  options: LargeSimpleReferencedSharedStringScanOptions = {},
): LargeSimpleSharedStringEntry[] {
  const readOptions = sharedStringReadOptions(options)
  return [...sharedStringsXml.matchAll(sharedStringElementPattern)].map((match) => {
    const xml = match[0]
    return readLargeSimpleSharedStringEntry(xml, readOptions)
  })
}

export function readLargeSimpleReferencedSharedStringsFromChunks(
  readChunks: (onChunk: (chunk: Uint8Array) => void) => boolean,
  referencedIndexes: ReadonlySet<number>,
  options: LargeSimpleReferencedSharedStringScanOptions = {},
): LargeSimpleSharedStrings | null {
  if (referencedIndexes.size === 0) {
    return []
  }
  const scanner = new LargeSimpleSharedStringChunkScanner(referencedIndexes, options)
  if (!readChunks((chunk) => scanner.push(chunk))) {
    return null
  }
  return scanner.finish()
}

export function createLargeSimpleSharedStringSubset(
  sharedStrings: LargeSimpleSharedStrings,
  referencedIndexes: ReadonlySet<number>,
): LargeSimpleSharedStrings | null {
  if (referencedIndexes.size === 0) {
    return []
  }
  const entries = new Map<number, LargeSimpleSharedStringEntry>()
  let maxReferencedIndex = 0
  for (const index of referencedIndexes) {
    const entry = sharedStrings[index]
    if (!entry) {
      return null
    }
    entries.set(index, entry)
    maxReferencedIndex = Math.max(maxReferencedIndex, index)
  }
  return createReferencedSharedStringTable(entries, maxReferencedIndex, { preferSparse: true })
}

export function hasReferencedLargeSimpleRichSharedStrings(
  sharedStrings: LargeSimpleSharedStrings,
  referencedIndexes: ReadonlySet<number>,
): boolean {
  for (const index of referencedIndexes) {
    if (sharedStrings[index]?.rich) {
      return true
    }
  }
  return false
}

export function collectReferencedLargeSimpleRichSharedStringIndexes(
  sharedStrings: LargeSimpleSharedStrings,
  referencedIndexes: ReadonlySet<number>,
): Set<number> | null {
  const richIndexes = new Set<number>()
  for (const index of referencedIndexes) {
    const entry = sharedStrings[index]
    if (!entry) {
      return null
    }
    if (entry.rich) {
      richIndexes.add(index)
    }
  }
  return richIndexes
}

export function readLargeSimpleRichTextCellArtifact(
  address: string,
  openingTag: string,
  cellXml: string,
  sharedStrings: LargeSimpleSharedStrings,
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

class LargeSimpleSharedStringChunkScanner {
  private readonly decoder = new TextDecoder()
  private readonly denseEntries: LargeSimpleSharedStringEntry[] | null
  private readonly sparseEntries: Map<number, LargeSimpleSharedStringEntry> | null
  private buffer = ''
  private index = 0
  private sharedStringIndex = 0
  private failed = false
  private skippingUnreferencedElementName: string | null = null
  private readonly maxReferencedIndex: number
  private readonly readOptions: LargeSimpleSharedStringReadOptions

  constructor(
    private readonly referencedIndexes: ReadonlySet<number>,
    private readonly options: LargeSimpleReferencedSharedStringScanOptions,
  ) {
    this.readOptions = sharedStringReadOptions(options)
    let maxIndex = 0
    for (const index of referencedIndexes) {
      maxIndex = Math.max(maxIndex, index)
    }
    this.maxReferencedIndex = maxIndex
    if (maxIndex + 1 <= referencedIndexes.size * sparseReferencedSharedStringDensityDivisor) {
      this.denseEntries = []
      this.denseEntries.length = maxIndex + 1
      this.sparseEntries = null
    } else {
      this.denseEntries = null
      this.sparseEntries = new Map()
    }
  }

  push(chunk: Uint8Array): void {
    if (this.failed || chunk.byteLength === 0) {
      return
    }
    this.buffer += this.decoder.decode(chunk, { stream: true })
    this.process(false)
    this.compact()
    this.reportRetainedBufferLength()
  }

  finish(): LargeSimpleSharedStrings | null {
    if (this.failed) {
      return null
    }
    this.buffer += this.decoder.decode()
    this.process(true)
    this.compact()
    this.reportRetainedBufferLength()
    if (this.failed) {
      return null
    }
    for (const index of this.referencedIndexes) {
      if (!this.hasEntry(index)) {
        return null
      }
    }
    return this.denseEntries ?? createReferencedSharedStringTable(this.sparseEntries ?? new Map(), this.maxReferencedIndex)
  }

  private process(final: boolean): void {
    while (!this.failed && this.index < this.buffer.length) {
      if (this.skippingUnreferencedElementName !== null) {
        if (!this.finishSkippingUnreferencedElement(final)) {
          return
        }
        continue
      }
      const opening = findNextElementOpening(this.buffer, 'si', this.index)
      if (!opening) {
        this.index = final ? this.buffer.length : Math.max(this.index, this.buffer.length - partialSharedStringTagRetainLength)
        return
      }
      const tagEnd = findTagEnd(this.buffer, opening.nameEnd)
      if (tagEnd === null) {
        if (final) {
          this.failed = true
        }
        return
      }
      const xmlEnd = isSelfClosingTag(this.buffer, tagEnd) ? tagEnd + 1 : findClosingElementEnd(this.buffer, opening.name, tagEnd + 1)
      if (xmlEnd === null) {
        if (final) {
          this.failed = true
        }
        if (!this.referencedIndexes.has(this.sharedStringIndex)) {
          this.skippingUnreferencedElementName = opening.name
          this.index = tagEnd + 1
        }
        return
      }
      if (this.referencedIndexes.has(this.sharedStringIndex)) {
        this.setEntry(this.sharedStringIndex, readLargeSimpleSharedStringEntry(this.buffer.slice(opening.start, xmlEnd), this.readOptions))
      }
      this.sharedStringIndex += 1
      this.index = xmlEnd
    }
  }

  private finishSkippingUnreferencedElement(final: boolean): boolean {
    const elementName = this.skippingUnreferencedElementName
    if (elementName === null) {
      return true
    }
    const xmlEnd = findClosingElementEnd(this.buffer, elementName, this.index)
    if (xmlEnd === null) {
      if (final) {
        this.failed = true
      }
      this.index = this.buffer.length
      return false
    }
    this.sharedStringIndex += 1
    this.index = xmlEnd
    this.skippingUnreferencedElementName = null
    return true
  }

  private compact(): void {
    if (this.skippingUnreferencedElementName !== null) {
      const retainLength = closingTagRetainLength(this.skippingUnreferencedElementName)
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

  private hasEntry(index: number): boolean {
    return this.denseEntries ? this.denseEntries[index] !== undefined : (this.sparseEntries?.has(index) ?? false)
  }

  private setEntry(index: number, entry: LargeSimpleSharedStringEntry): void {
    if (this.denseEntries) {
      this.denseEntries[index] = entry
      return
    }
    this.sparseEntries?.set(index, entry)
  }
}

function createReferencedSharedStringTable(
  entries: ReadonlyMap<number, LargeSimpleSharedStringEntry>,
  maxReferencedIndex: number,
  options: { readonly preferSparse?: boolean } = {},
): LargeSimpleSharedStrings {
  const length = maxReferencedIndex + 1
  if (options.preferSparse !== true && length <= entries.size * sparseReferencedSharedStringDensityDivisor) {
    const output: LargeSimpleSharedStringEntry[] = []
    output.length = length
    for (const [index, entry] of entries) {
      output[index] = entry
    }
    return output
  }
  return new Proxy(
    { length },
    {
      get(target, property) {
        if (property === 'length') {
          return target.length
        }
        return typeof property === 'string' && isArrayIndexProperty(property) ? entries.get(Number(property)) : undefined
      },
      has(_target, property) {
        return property === 'length' || (typeof property === 'string' && isArrayIndexProperty(property) && entries.has(Number(property)))
      },
    },
  ) as LargeSimpleSharedStringTable
}

function isArrayIndexProperty(property: string): boolean {
  if (property.length === 0 || !/^(?:0|[1-9][0-9]*)$/u.test(property)) {
    return false
  }
  const index = Number(property)
  return Number.isSafeInteger(index) && index >= 0 && index < 2 ** 32 - 1
}

function closingTagRetainLength(elementName: string): number {
  return Math.max(partialSharedStringTagRetainLength, elementName.length + 4)
}

function readLargeSimpleSharedStringEntry(xml: string, options: LargeSimpleSharedStringReadOptions): LargeSimpleSharedStringEntry {
  const rich = richTextRunPattern.test(xml)
  if (!rich) {
    return internPlainSharedStringEntry(internSharedStringText(stringItemText(xml), options), options)
  }
  return lazyRichSharedStringEntry(xml, options)
}

function lazyRichSharedStringEntry(xml: string, options: LargeSimpleSharedStringReadOptions): LargeSimpleSharedStringEntry {
  let text: string | undefined
  return {
    rich: true,
    xml,
    get text() {
      text ??= internSharedStringText(stringItemText(xml), options)
      return text
    },
  }
}

function internSharedStringText(value: string, options: LargeSimpleReferencedSharedStringScanOptions): string {
  const mode = options.deduplicateText ?? 'bounded'
  if (mode === false || !options.stringPool) {
    return value
  }
  if (mode === 'bounded') {
    return options.stringPool.internBounded(value, options.dedupeMaxEntries ?? 8192)
  }
  return options.stringPool.intern(value)
}

function sharedStringReadOptions(options: LargeSimpleReferencedSharedStringScanOptions): LargeSimpleSharedStringReadOptions {
  if ((options.deduplicateText ?? 'bounded') === false) {
    return options
  }
  return {
    ...options,
    plainEntryPool: {
      entries: new Map(),
      keys: [],
      evictionIndex: 0,
    },
  }
}

function internPlainSharedStringEntry(text: string, options: LargeSimpleSharedStringReadOptions): LargeSimpleSharedStringEntry {
  const pool = options.plainEntryPool
  if (!pool) {
    return { text, rich: false }
  }
  const existing = pool.entries.get(text)
  if (existing) {
    return existing
  }
  const entry = { text, rich: false } satisfies LargeSimpleSharedStringEntry
  pool.entries.set(text, entry)
  pool.keys.push(text)
  if ((options.deduplicateText ?? 'bounded') === 'bounded') {
    evictPlainSharedStringEntries(pool, options.dedupeMaxEntries ?? 8192)
  }
  return entry
}

function evictPlainSharedStringEntries(pool: LargeSimpleSharedStringEntryPool, maxEntries: number): void {
  const limit = Math.max(0, Math.trunc(maxEntries))
  while (pool.keys.length - pool.evictionIndex > limit) {
    const key = pool.keys[pool.evictionIndex]
    pool.evictionIndex += 1
    if (key !== undefined) {
      pool.entries.delete(key)
    }
  }
  if (pool.evictionIndex > limit && pool.evictionIndex * 2 > pool.keys.length) {
    pool.keys.splice(0, pool.evictionIndex)
    pool.evictionIndex = 0
  }
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

function findNextElementOpening(
  xml: string,
  localName: string,
  startIndex: number,
): { readonly start: number; readonly name: string; readonly nameEnd: number } | null {
  let index = startIndex
  while (index < xml.length) {
    const openingStart = xml.indexOf('<', index)
    if (openingStart === -1) {
      return null
    }
    const name = readElementName(xml, openingStart + 1)
    if (!name) {
      index = openingStart + 1
      continue
    }
    if (readLocalName(name.name) === localName) {
      return {
        start: openingStart,
        name: name.name,
        nameEnd: name.end,
      }
    }
    index = name.end
  }
  return null
}

function readElementName(xml: string, startIndex: number): { readonly name: string; readonly end: number } | null {
  const first = xml[startIndex]
  if (!first || first === '/' || first === '?' || first === '!') {
    return null
  }
  let index = startIndex
  while (index < xml.length && isXmlNameCharacter(xml.charCodeAt(index))) {
    index += 1
  }
  return index > startIndex ? { name: xml.slice(startIndex, index), end: index } : null
}

function readLocalName(name: string): string {
  return name.includes(':') ? name.slice(name.lastIndexOf(':') + 1) : name
}

function findTagEnd(xml: string, startIndex: number): number | null {
  let quote: string | null = null
  for (let index = startIndex; index < xml.length; index += 1) {
    const character = xml[index]
    if (quote) {
      if (character === quote) {
        quote = null
      }
      continue
    }
    if (character === '"' || character === "'") {
      quote = character
      continue
    }
    if (character === '>') {
      return index
    }
  }
  return null
}

function isSelfClosingTag(xml: string, tagEnd: number): boolean {
  let index = tagEnd - 1
  while (index >= 0 && isAsciiWhitespace(xml.charCodeAt(index))) {
    index -= 1
  }
  return xml[index] === '/'
}

function findClosingElementEnd(xml: string, name: string, startIndex: number): number | null {
  const closingStart = xml.indexOf(`</${name}`, startIndex)
  if (closingStart === -1) {
    return null
  }
  const closingNameEnd = closingStart + name.length + 2
  const next = xml[closingNameEnd]
  if (next !== '>' && !isAsciiWhitespace(next?.charCodeAt(0) ?? 0)) {
    return findClosingElementEnd(xml, name, closingNameEnd)
  }
  const tagEnd = findTagEnd(xml, closingNameEnd)
  return tagEnd === null ? null : tagEnd + 1
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

function isAsciiWhitespace(value: number): boolean {
  return value === 9 || value === 10 || value === 12 || value === 13 || value === 32
}

function isXmlNameCharacter(value: number): boolean {
  return (
    (value >= 65 && value <= 90) ||
    (value >= 97 && value <= 122) ||
    (value >= 48 && value <= 57) ||
    value === 45 ||
    value === 46 ||
    value === 58 ||
    value === 95
  )
}
