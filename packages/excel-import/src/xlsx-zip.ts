import { inflateSync, strFromU8, unzipSync, type Unzipped } from 'fflate'

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
  for (const entry of entries) {
    defineLazyZipEntry(output, source, entry)
  }
  return output
}

export function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

export function getZipText(zip: XlsxZipEntries, path: string): string | null {
  const file = zip[normalizeZipPath(path)]
  return file ? strFromU8(file) : null
}

const localFileHeaderSignature = 0x04034b50
const centralDirectoryFileHeaderSignature = 0x02014b50
const endOfCentralDirectorySignature = 0x06054b50
const storedCompressionMethod = 0
const deflatedCompressionMethod = 8
const zip64Sentinel = 0xffffffff
const maxEndOfCentralDirectorySearch = 65_557
const textDecoder = new TextDecoder()

interface CentralDirectoryEntry {
  readonly path: string
  readonly localHeaderOffset: number
  readonly compressedSize: number
  readonly compressionMethod: number
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

function defineLazyZipEntry(output: Unzipped, source: Uint8Array, entry: CentralDirectoryEntry): void {
  Object.defineProperty(output, entry.path, {
    configurable: true,
    enumerable: true,
    get() {
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
