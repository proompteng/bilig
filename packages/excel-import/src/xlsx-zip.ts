import { inflateSync, strFromU8, unzipSync, type Unzipped } from 'fflate'
import { Inflate } from 'fflate-stream'

export type XlsxZipEntries = Unzipped
export type XlsxZipSource = Uint8Array | XlsxZipEntries

export function readXlsxZipEntries(source: XlsxZipSource): XlsxZipEntries {
  if (!(source instanceof Uint8Array)) {
    return source
  }
  const zip = unzipSync(source)
  return shouldUseCentralDirectoryZipFallback(source, zip) ? unzipFromCentralDirectory(source) : zip
}

export function readXlsxZipEntriesLazy(source: XlsxZipSource): XlsxZipEntries {
  if (!(source instanceof Uint8Array)) {
    return source
  }
  const entries = readCentralDirectoryEntries(source)
  if (!entries) {
    return readXlsxZipEntries(source)
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
    return false
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
  metadata.source = null
  return true
}

export function readLazyXlsxZipSource(zip: XlsxZipEntries): Uint8Array | undefined {
  const metadata = (zip as XlsxZipEntriesWithCentralDirectorySource)[xlsxZipCentralDirectorySourceSymbol]
  return metadata?.source ?? undefined
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
  source: Uint8Array | null
  readonly entriesByPath: ReadonlyMap<string, CentralDirectoryEntry>
}

type XlsxZipEntriesWithCentralDirectorySource = XlsxZipEntries & {
  readonly [xlsxZipCentralDirectorySourceSymbol]?: XlsxZipCentralDirectorySource
}

function shouldUseCentralDirectoryZipFallback(source: Uint8Array, zip: XlsxZipEntries): boolean {
  const entries = Object.values(zip)
  return source.byteLength > 0 && entries.length > 0 && entries.every((entry) => entry.byteLength === 0)
}

function unzipFromCentralDirectory(source: Uint8Array): XlsxZipEntries {
  const entries = readCentralDirectoryEntries(source)
  if (!entries) {
    return unzipSync(source)
  }
  const output: Unzipped = {}
  for (const entry of entries) {
    output[entry.path] = inflateCentralDirectoryEntry(source, entry.localHeaderOffset, entry.compressedSize, entry.compressionMethod)
  }
  return output
}

function readCentralDirectoryEntries(source: Uint8Array): CentralDirectoryEntry[] | null {
  const endOfCentralDirectoryOffset = findEndOfCentralDirectoryOffset(source)
  if (endOfCentralDirectoryOffset === null) {
    return null
  }
  const centralDirectoryOffset = readUint32(source, endOfCentralDirectoryOffset + 16)
  const centralDirectorySize = readUint32(source, endOfCentralDirectoryOffset + 12)
  if (
    centralDirectoryOffset === zip64Sentinel ||
    centralDirectorySize === zip64Sentinel ||
    centralDirectoryOffset + centralDirectorySize > source.byteLength
  ) {
    return null
  }

  const entries: CentralDirectoryEntry[] = []
  let offset = centralDirectoryOffset
  const endOffset = centralDirectoryOffset + centralDirectorySize
  while (offset + 46 <= endOffset && readUint32(source, offset) === centralDirectoryFileHeaderSignature) {
    const compressionMethod = readUint16(source, offset + 10)
    const compressedSize = readUint32(source, offset + 20)
    const fileNameLength = readUint16(source, offset + 28)
    const extraFieldLength = readUint16(source, offset + 30)
    const fileCommentLength = readUint16(source, offset + 32)
    const localHeaderOffset = readUint32(source, offset + 42)
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
    const path = normalizeZipPath(textDecoder.decode(source.subarray(fileNameStart, fileNameEnd)))
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
  source: Uint8Array,
  localHeaderOffset: number,
  compressedSize: number,
  compressionMethod: number,
): Uint8Array {
  if (readUint32(source, localHeaderOffset) !== localFileHeaderSignature) {
    throw new Error('Invalid XLSX local file header')
  }
  const fileNameLength = readUint16(source, localHeaderOffset + 26)
  const extraFieldLength = readUint16(source, localHeaderOffset + 28)
  const dataStart = localHeaderOffset + 30 + fileNameLength + extraFieldLength
  const dataEnd = dataStart + compressedSize
  if (dataEnd > source.byteLength) {
    throw new Error('Invalid XLSX compressed data range')
  }
  const compressed = source.subarray(dataStart, dataEnd)
  if (compressionMethod === storedCompressionMethod) {
    return new Uint8Array(compressed)
  }
  if (compressionMethod === deflatedCompressionMethod) {
    return inflateSync(compressed)
  }
  throw new Error(`Unsupported XLSX compression method: ${String(compressionMethod)}`)
}

function inflateCentralDirectoryEntryChunks(
  source: Uint8Array,
  localHeaderOffset: number,
  compressedSize: number,
  compressionMethod: number,
  onChunk: (chunk: Uint8Array) => void,
  options: { readonly chunkSize: number },
): void {
  if (readUint32(source, localHeaderOffset) !== localFileHeaderSignature) {
    throw new Error('Invalid XLSX local file header')
  }
  const fileNameLength = readUint16(source, localHeaderOffset + 26)
  const extraFieldLength = readUint16(source, localHeaderOffset + 28)
  const dataStart = localHeaderOffset + 30 + fileNameLength + extraFieldLength
  const dataEnd = dataStart + compressedSize
  if (dataEnd > source.byteLength) {
    throw new Error('Invalid XLSX compressed data range')
  }
  const compressed = source.subarray(dataStart, dataEnd)
  if (compressionMethod === storedCompressionMethod) {
    forEachChunk(compressed, options.chunkSize, onChunk)
    return
  }
  if (compressionMethod === deflatedCompressionMethod) {
    const inflate = new Inflate((chunk) => {
      onChunk(chunk)
    })
    if (compressed.byteLength === 0) {
      inflate.push(compressed, true)
      return
    }
    for (let offset = 0; offset < compressed.byteLength; offset += options.chunkSize) {
      const end = Math.min(compressed.byteLength, offset + options.chunkSize)
      inflate.push(compressed.subarray(offset, end), end === compressed.byteLength)
    }
    return
  }
  throw new Error(`Unsupported XLSX compression method: ${String(compressionMethod)}`)
}

function forEachChunk(bytes: Uint8Array, chunkSize: number, onChunk: (chunk: Uint8Array) => void): void {
  const normalizedChunkSize = Math.max(1, Math.trunc(chunkSize))
  if (bytes.byteLength === 0) {
    onChunk(bytes)
    return
  }
  for (let offset = 0; offset < bytes.byteLength; offset += normalizedChunkSize) {
    onChunk(bytes.subarray(offset, Math.min(bytes.byteLength, offset + normalizedChunkSize)))
  }
}

function findEndOfCentralDirectoryOffset(source: Uint8Array): number | null {
  const minOffset = Math.max(0, source.byteLength - maxEndOfCentralDirectorySearch)
  for (let offset = source.byteLength - 22; offset >= minOffset; offset -= 1) {
    if (readUint32(source, offset) === endOfCentralDirectorySignature) {
      return offset
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
