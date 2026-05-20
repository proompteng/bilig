import { inflateSync, strFromU8, unzipSync, type Unzipped } from 'fflate'
import { Inflate } from 'fflate-stream'

export type XlsxZipEntries = Unzipped
export type XlsxZipSource = Uint8Array | XlsxZipEntries

export interface XlsxZipByteSource {
  readonly byteLength: number
  readRange(start: number, end: number): Uint8Array
  release?(): void
}

export interface XlsxZipEntryMetadata {
  readonly path: string
  readonly compressedSize: number
  readonly compressionMethod: number
}

export function readXlsxZipEntries(source: XlsxZipSource): XlsxZipEntries {
  if (!(source instanceof Uint8Array)) {
    return source
  }
  const zip = unzipSync(source)
  return shouldUseCentralDirectoryZipFallback(source, zip) ? (unzipFromCentralDirectory(byteSourceFromUint8Array(source)) ?? zip) : zip
}

export function readXlsxZipEntriesLazy(source: XlsxZipSource): XlsxZipEntries {
  if (!(source instanceof Uint8Array)) {
    return source
  }
  return readXlsxZipEntriesLazyFromByteSource(byteSourceFromUint8Array(source)) ?? readXlsxZipEntries(source)
}

export function readXlsxZipEntriesLazyFromByteSource(source: XlsxZipByteSource): XlsxZipEntries | null {
  const entries = readCentralDirectoryEntries(source)
  if (!entries) {
    return null
  }
  const output: Unzipped = {}
  const metadata: XlsxZipCentralDirectorySource = {
    source,
    entriesByPath: new Map(entries.map((entry) => [entry.path, entry])),
  }
  for (const entry of entries) {
    defineLazyZipEntry(output, metadata, entry)
  }
  defineLazyZipCentralDirectorySource(output, metadata)
  return output
}

export function readXlsxZipEntryMetadata(source: XlsxZipByteSource): readonly XlsxZipEntryMetadata[] | null {
  return (
    readCentralDirectoryEntries(source)?.map((entry) => ({
      path: entry.path,
      compressedSize: entry.compressedSize,
      compressionMethod: entry.compressionMethod,
    })) ?? null
  )
}

export function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

export function getZipText(zip: XlsxZipEntries, path: string): string | null {
  const file = zip[normalizeZipPath(path)]
  return file ? strFromU8(file) : null
}

export function forEachInflatedXlsxZipEntryChunk(
  zip: XlsxZipEntries,
  path: string,
  onChunk: (chunk: Uint8Array) => void,
  options: { readonly chunkSize?: number } = {},
): boolean {
  const metadata = (zip as XlsxZipEntriesWithCentralDirectorySource)[xlsxZipCentralDirectorySourceSymbol]
  const source = metadata?.source
  const entry = metadata?.entriesByPath.get(normalizeZipPath(path))
  if (!metadata || !source || !entry) {
    const inflated = zip[normalizeZipPath(path)]
    if (!inflated) {
      return false
    }
    const chunkSize = options.chunkSize ?? defaultZipEntryChunkSize
    for (let offset = 0; offset < inflated.byteLength; offset += chunkSize) {
      onChunk(inflated.subarray(offset, Math.min(inflated.byteLength, offset + chunkSize)))
    }
    return true
  }
  inflateCentralDirectoryEntryChunks(source, entry.localHeaderOffset, entry.compressedSize, entry.compressionMethod, onChunk, {
    chunkSize: options.chunkSize ?? defaultZipEntryChunkSize,
  })
  return true
}

export function releaseLazyXlsxZipSource(zip: XlsxZipEntries): boolean {
  const metadata = (zip as XlsxZipEntriesWithCentralDirectorySource)[xlsxZipCentralDirectorySourceSymbol]
  if (!metadata?.source) {
    return false
  }
  metadata.source.release?.()
  metadata.source = null
  return true
}

export function readLazyXlsxZipSource(zip: XlsxZipEntries): Uint8Array | undefined {
  const metadata = (zip as XlsxZipEntriesWithCentralDirectorySource)[xlsxZipCentralDirectorySourceSymbol]
  const source = metadata?.source
  return source && isUint8ArrayZipByteSource(source) ? source.bytes : undefined
}

export function readLazyXlsxZipSourceByteLength(zip: XlsxZipEntries): number | undefined {
  const metadata = (zip as XlsxZipEntriesWithCentralDirectorySource)[xlsxZipCentralDirectorySourceSymbol]
  return metadata ? (metadata.source?.byteLength ?? 0) : undefined
}

const localFileHeaderSignature = 0x04034b50
const centralDirectoryFileHeaderSignature = 0x02014b50
const endOfCentralDirectorySignature = 0x06054b50
const storedCompressionMethod = 0
const deflatedCompressionMethod = 8
const zip64Sentinel = 0xffffffff
const maxEndOfCentralDirectorySearch = 65_557
const defaultZipEntryChunkSize = 64 * 1024
const xlsxZipCentralDirectorySourceSymbol: unique symbol = Symbol('bilig.xlsxZipCentralDirectorySource')
const textDecoder = new TextDecoder()

interface CentralDirectoryEntry {
  readonly path: string
  readonly localHeaderOffset: number
  readonly compressedSize: number
  readonly compressionMethod: number
}

interface XlsxZipCentralDirectorySource {
  source: XlsxZipByteSource | null
  readonly entriesByPath: ReadonlyMap<string, CentralDirectoryEntry>
}

interface Uint8ArrayXlsxZipByteSource extends XlsxZipByteSource {
  readonly bytes: Uint8Array
}

type XlsxZipEntriesWithCentralDirectorySource = XlsxZipEntries & {
  readonly [xlsxZipCentralDirectorySourceSymbol]?: XlsxZipCentralDirectorySource
}

function shouldUseCentralDirectoryZipFallback(source: Uint8Array, zip: XlsxZipEntries): boolean {
  const entries = Object.values(zip)
  return source.byteLength > 0 && entries.length > 0 && entries.every((entry) => entry.byteLength === 0)
}

function unzipFromCentralDirectory(source: XlsxZipByteSource): XlsxZipEntries | null {
  const entries = readCentralDirectoryEntries(source)
  if (!entries) {
    return null
  }
  const output: Unzipped = {}
  for (const entry of entries) {
    output[entry.path] = inflateCentralDirectoryEntry(source, entry.localHeaderOffset, entry.compressedSize, entry.compressionMethod)
  }
  return output
}

function byteSourceFromUint8Array(source: Uint8Array): Uint8ArrayXlsxZipByteSource {
  return {
    byteLength: source.byteLength,
    bytes: source,
    readRange(start, end) {
      return source.subarray(start, end)
    },
  }
}

function isUint8ArrayZipByteSource(source: XlsxZipByteSource): source is Uint8ArrayXlsxZipByteSource {
  return 'bytes' in source && source.bytes instanceof Uint8Array
}

function readCentralDirectoryEntries(source: XlsxZipByteSource): CentralDirectoryEntry[] | null {
  const endOfCentralDirectoryOffset = findEndOfCentralDirectoryOffset(source)
  if (endOfCentralDirectoryOffset === null) {
    return null
  }
  const endOfCentralDirectory = source.readRange(endOfCentralDirectoryOffset, endOfCentralDirectoryOffset + 22)
  const centralDirectoryOffset = readUint32(endOfCentralDirectory, 16)
  const centralDirectorySize = readUint32(endOfCentralDirectory, 12)
  if (
    centralDirectoryOffset === zip64Sentinel ||
    centralDirectorySize === zip64Sentinel ||
    centralDirectoryOffset + centralDirectorySize > source.byteLength
  ) {
    return null
  }
  const centralDirectory = source.readRange(centralDirectoryOffset, centralDirectoryOffset + centralDirectorySize)

  const entries: CentralDirectoryEntry[] = []
  let offset = 0
  const endOffset = centralDirectory.byteLength
  while (offset + 46 <= endOffset && readUint32(centralDirectory, offset) === centralDirectoryFileHeaderSignature) {
    const compressionMethod = readUint16(centralDirectory, offset + 10)
    const compressedSize = readUint32(centralDirectory, offset + 20)
    const fileNameLength = readUint16(centralDirectory, offset + 28)
    const extraFieldLength = readUint16(centralDirectory, offset + 30)
    const fileCommentLength = readUint16(centralDirectory, offset + 32)
    const localHeaderOffset = readUint32(centralDirectory, offset + 42)
    const fileNameStart = offset + 46
    const fileNameEnd = fileNameStart + fileNameLength
    const nextOffset = fileNameEnd + extraFieldLength + fileCommentLength
    if (
      fileNameEnd > endOffset ||
      nextOffset > endOffset ||
      localHeaderOffset === zip64Sentinel ||
      localHeaderOffset + 30 > source.byteLength
    ) {
      return null
    }
    const path = normalizeZipPath(textDecoder.decode(centralDirectory.subarray(fileNameStart, fileNameEnd)))
    entries.push({ path, localHeaderOffset, compressedSize, compressionMethod })
    offset = nextOffset
  }
  return offset === endOffset ? entries : null
}

function defineLazyZipEntry(output: Unzipped, metadata: XlsxZipCentralDirectorySource, entry: CentralDirectoryEntry): void {
  Object.defineProperty(output, entry.path, {
    configurable: true,
    enumerable: true,
    get() {
      const source = metadata.source
      if (!source) {
        throw new Error('XLSX ZIP source has been released')
      }
      const bytes = inflateCentralDirectoryEntry(source, entry.localHeaderOffset, entry.compressedSize, entry.compressionMethod)
      Object.defineProperty(output, entry.path, {
        configurable: true,
        enumerable: true,
        value: bytes,
        writable: true,
      })
      return bytes
    },
  })
}

function defineLazyZipCentralDirectorySource(output: Unzipped, metadata: XlsxZipCentralDirectorySource): void {
  Object.defineProperty(output, xlsxZipCentralDirectorySourceSymbol, {
    configurable: false,
    enumerable: false,
    value: metadata,
  })
}

function inflateCentralDirectoryEntry(
  source: XlsxZipByteSource,
  localHeaderOffset: number,
  compressedSize: number,
  compressionMethod: number,
): Uint8Array {
  const localHeader = source.readRange(localHeaderOffset, localHeaderOffset + 30)
  if (readUint32(localHeader, 0) !== localFileHeaderSignature) {
    throw new Error('Invalid XLSX local file header')
  }
  const fileNameLength = readUint16(localHeader, 26)
  const extraFieldLength = readUint16(localHeader, 28)
  const dataStart = localHeaderOffset + 30 + fileNameLength + extraFieldLength
  const dataEnd = dataStart + compressedSize
  if (dataEnd > source.byteLength) {
    throw new Error('Invalid XLSX compressed data range')
  }
  const compressed = source.readRange(dataStart, dataEnd)
  if (compressionMethod === storedCompressionMethod) {
    return new Uint8Array(compressed)
  }
  if (compressionMethod === deflatedCompressionMethod) {
    return inflateSync(compressed)
  }
  throw new Error(`Unsupported XLSX compression method: ${String(compressionMethod)}`)
}

function inflateCentralDirectoryEntryChunks(
  source: XlsxZipByteSource,
  localHeaderOffset: number,
  compressedSize: number,
  compressionMethod: number,
  onChunk: (chunk: Uint8Array) => void,
  options: { readonly chunkSize: number },
): void {
  const localHeader = source.readRange(localHeaderOffset, localHeaderOffset + 30)
  if (readUint32(localHeader, 0) !== localFileHeaderSignature) {
    throw new Error('Invalid XLSX local file header')
  }
  const fileNameLength = readUint16(localHeader, 26)
  const extraFieldLength = readUint16(localHeader, 28)
  const dataStart = localHeaderOffset + 30 + fileNameLength + extraFieldLength
  const dataEnd = dataStart + compressedSize
  if (dataEnd > source.byteLength) {
    throw new Error('Invalid XLSX compressed data range')
  }
  if (compressionMethod === storedCompressionMethod) {
    forEachSourceChunk(source, dataStart, dataEnd, options.chunkSize, onChunk)
    return
  }
  if (compressionMethod === deflatedCompressionMethod) {
    const inflate = new Inflate((chunk) => {
      onChunk(chunk)
    })
    if (compressedSize === 0) {
      inflate.push(source.readRange(dataStart, dataEnd), true)
      return
    }
    for (let offset = 0; offset < compressedSize; offset += options.chunkSize) {
      const end = Math.min(compressedSize, offset + options.chunkSize)
      inflate.push(source.readRange(dataStart + offset, dataStart + end), end === compressedSize)
    }
    return
  }
  throw new Error(`Unsupported XLSX compression method: ${String(compressionMethod)}`)
}

function forEachSourceChunk(
  source: XlsxZipByteSource,
  start: number,
  end: number,
  chunkSize: number,
  onChunk: (chunk: Uint8Array) => void,
): void {
  const normalizedChunkSize = Math.max(1, Math.trunc(chunkSize))
  if (start === end) {
    onChunk(source.readRange(start, end))
    return
  }
  for (let offset = start; offset < end; offset += normalizedChunkSize) {
    onChunk(source.readRange(offset, Math.min(end, offset + normalizedChunkSize)))
  }
}

function findEndOfCentralDirectoryOffset(source: XlsxZipByteSource): number | null {
  const tailStart = Math.max(0, source.byteLength - maxEndOfCentralDirectorySearch)
  const tail = source.readRange(tailStart, source.byteLength)
  for (let offset = tail.byteLength - 22; offset >= 0; offset -= 1) {
    if (readUint32(tail, offset) === endOfCentralDirectorySignature) {
      return tailStart + offset
    }
  }
  return null
}

function readUint16(source: Uint8Array, offset: number): number {
  return source[offset]! | (source[offset + 1]! << 8)
}

function readUint32(source: Uint8Array, offset: number): number {
  return (source[offset]! | (source[offset + 1]! << 8) | (source[offset + 2]! << 16) | (source[offset + 3]! << 24)) >>> 0
}
