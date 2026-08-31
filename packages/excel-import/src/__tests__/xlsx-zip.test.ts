import { deflateRawSync, inflateRawSync } from 'node:zlib'
import { createHash } from 'node:crypto'

import { describe, expect, it } from 'vitest'

import {
  forEachInflatedXlsxZipEntryChunk,
  forEachInflatedXlsxZipEntryChunkAsync,
  getZipText,
  readLazyXlsxZipSource,
  readLazyXlsxZipSourceByteLength,
  readXlsxZipEntriesLazyFromByteSource,
  readXlsxZipEntries,
  readXlsxZipEntriesLazy,
  releaseInflatedLazyXlsxZipEntries,
  releaseLazyXlsxZipSource,
  type XlsxZipByteSource,
} from '../xlsx-zip.js'

describe('XLSX ZIP reader', () => {
  it('inflates streamed ZIP entries from the central directory', () => {
    const zip = readXlsxZipEntries(buildStreamedZip('xl/workbook.xml', '<workbook><sheets/></workbook>'))

    expect(getZipText(zip, 'xl/workbook.xml')).toBe('<workbook><sheets/></workbook>')
  })

  it('enumerates lazy central-directory entries before inflating them', () => {
    const zip = readXlsxZipEntriesLazy(buildStreamedZip('xl/workbook.xml', '<workbook><sheets/></workbook>'))

    expect(Object.keys(zip)).toEqual(['xl/workbook.xml'])
    const descriptorBefore = Object.getOwnPropertyDescriptor(zip, 'xl/workbook.xml')
    expect(descriptorBefore && 'get' in descriptorBefore && typeof descriptorBefore.get).toBe('function')
    expect(descriptorBefore && 'value' in descriptorBefore).toBe(false)

    expect(getZipText(zip, 'xl/workbook.xml')).toBe('<workbook><sheets/></workbook>')
    const descriptorAfter = Object.getOwnPropertyDescriptor(zip, 'xl/workbook.xml')
    expect(descriptorAfter?.value).toBeInstanceOf(Uint8Array)
  })

  it('releases inflated lazy ZIP entries while keeping them readable from the source', () => {
    const zip = readXlsxZipEntriesLazy(buildStreamedZip('xl/workbook.xml', '<workbook><sheets/></workbook>'))

    expect(getZipText(zip, 'xl/workbook.xml')).toBe('<workbook><sheets/></workbook>')
    const inflatedDescriptor = Object.getOwnPropertyDescriptor(zip, 'xl/workbook.xml')
    expect(inflatedDescriptor?.value).toBeInstanceOf(Uint8Array)

    expect(releaseInflatedLazyXlsxZipEntries(zip)).toBeGreaterThan(0)
    const releasedDescriptor = Object.getOwnPropertyDescriptor(zip, 'xl/workbook.xml')
    expect(releasedDescriptor && 'get' in releasedDescriptor && typeof releasedDescriptor.get).toBe('function')

    expect(getZipText(zip, 'xl/workbook.xml')).toBe('<workbook><sheets/></workbook>')
  })

  it('releases lazy central-directory source bytes after streamed consumers finish', () => {
    const zip = readXlsxZipEntriesLazy(buildStreamedZip('xl/workbook.xml', '<workbook><sheets/></workbook>'))
    const chunks: Uint8Array[] = []
    expect(readLazyXlsxZipSourceByteLength(zip)).toBeGreaterThan(0)
    expect(readLazyXlsxZipSource(zip)).toBeInstanceOf(Uint8Array)

    expect(forEachInflatedXlsxZipEntryChunk(zip, 'xl/workbook.xml', (chunk) => chunks.push(chunk))).toBe(true)
    expect(Buffer.concat(chunks).toString()).toBe('<workbook><sheets/></workbook>')
    expect(releaseLazyXlsxZipSource(zip)).toBe(true)
    expect(readLazyXlsxZipSourceByteLength(zip)).toBe(0)
    expect(readLazyXlsxZipSource(zip)).toBeUndefined()
    expect(releaseLazyXlsxZipSource(zip)).toBe(false)
    expect(forEachInflatedXlsxZipEntryChunk(zip, 'xl/workbook.xml', () => undefined)).toBe(false)
    expect(() => getZipText(zip, 'xl/workbook.xml')).toThrow(/released/u)
  })

  it('fails fast for corrupt small streamed ZIP entries', () => {
    const path = 'xl/workbook.xml'
    const zip = readXlsxZipEntriesLazy(corruptFirstCompressedByte(buildStreamedZip(path, '<workbook><sheets/></workbook>')))

    expect(() => forEachInflatedXlsxZipEntryChunk(zip, path, () => undefined)).toThrow()
  })

  it('streams highly compressed large entries without inflating the whole entry first', () => {
    const path = 'xl/worksheets/sheet1.xml'
    const payload = semiCompressibleBytes(2 * 1024 * 1024)
    const source = new CountingZipByteSource(buildStreamedZip(path, payload))
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    const chunks: Uint8Array[] = []
    source.maxReadLength = 0

    expect(zip).not.toBeNull()
    expect(
      forEachInflatedXlsxZipEntryChunk(zip!, path, (chunk) => chunks.push(chunk), {
        chunkSize: 512,
      }),
    ).toBe(true)

    expect(Buffer.concat(chunks)).toEqual(Buffer.from(payload))
    expect(chunks.length).toBeGreaterThan(1)
    expect(source.readIntoCount).toBeGreaterThan(0)
    expect(source.maxReadLength).toBeLessThan(16 * 1024)
  }, 30_000)

  it('can force streaming inflation for small deflated entries', () => {
    const path = 'xl/worksheets/sheet1.xml'
    const payload = semiCompressibleBytes(128 * 1024)
    const source = new CountingZipByteSource(buildStreamedZip(path, payload))
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    const hash = createHash('sha256')
    let streamedByteLength = 0

    expect(zip).not.toBeNull()
    source.resetCounts()
    expect(
      forEachInflatedXlsxZipEntryChunk(
        zip!,
        path,
        (chunk) => {
          streamedByteLength += chunk.byteLength
          hash.update(chunk)
        },
        {
          chunkSize: 4096,
          forceStreamingInflate: true,
        },
      ),
    ).toBe(true)

    expect(streamedByteLength).toBe(payload.byteLength)
    expect(hash.digest('hex')).toBe(createHash('sha256').update(payload).digest('hex'))
    expect(source.readIntoCount).toBeGreaterThan(0)
    expect(source.maxReadLength).toBeLessThanOrEqual(4096)
  })

  it('stops streaming compressed entries when the consumer returns false', () => {
    const path = 'xl/worksheets/sheet1.xml'
    const payload = semiCompressibleBytes(2 * 1024 * 1024)
    const source = new CountingZipByteSource(buildStreamedZip(path, payload))
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    let emittedChunks = 0

    expect(zip).not.toBeNull()
    source.resetCounts()
    expect(
      forEachInflatedXlsxZipEntryChunk(
        zip!,
        path,
        () => {
          emittedChunks += 1
          return false
        },
        {
          chunkSize: 4096,
          forceStreamingInflate: true,
        },
      ),
    ).toBe(true)

    expect(emittedChunks).toBe(1)
    expect(source.readIntoCount).toBeGreaterThan(0)
    expect(source.readIntoCount).toBeLessThan(8)
    expect(source.maxReadLength).toBeLessThanOrEqual(4096)
  })

  it('async-streams compressed entries with native backpressure', async () => {
    const path = 'xl/worksheets/sheet1.xml'
    const payload = semiCompressibleBytes(512 * 1024)
    const source = new CountingZipByteSource(buildStreamedZip(path, payload))
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    let emittedByteLength = 0
    let emittedChunks = 0

    expect(zip).not.toBeNull()
    source.resetCounts()
    expect(
      await forEachInflatedXlsxZipEntryChunkAsync(
        zip!,
        path,
        (chunk) => {
          emittedChunks += 1
          emittedByteLength += chunk.byteLength
          return emittedChunks < 2
        },
        {
          chunkSize: 4096,
          forceStreamingInflate: true,
        },
      ),
    ).toBe(true)

    expect(emittedChunks).toBe(2)
    expect(emittedByteLength).toBeLessThan(payload.byteLength)
    expect(source.maxReadLength).toBeLessThanOrEqual(4096)
  })

  it('prefers byte-source native async inflation hooks when available', async () => {
    const path = 'xl/worksheets/sheet1.xml'
    const payload = semiCompressibleBytes(512 * 1024)
    const source = new NativeInflateZipByteSource(buildStreamedZip(path, payload))
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    let emittedChunks = 0
    let emittedByteLength = 0

    expect(zip).not.toBeNull()
    expect(
      await forEachInflatedXlsxZipEntryChunkAsync(
        zip!,
        path,
        (chunk) => {
          emittedChunks += 1
          emittedByteLength += chunk.byteLength
          return emittedChunks < 3
        },
        {
          chunkSize: 4096,
          forceStreamingInflate: true,
        },
      ),
    ).toBe(true)

    expect(source.nativeInflateCount).toBe(1)
    expect(emittedChunks).toBe(3)
    expect(emittedByteLength).toBeLessThan(payload.byteLength)
  })

  it('streams stored entries through reusable readRangeInto scratch buffers', () => {
    const path = 'xl/worksheets/sheet1.xml'
    const payload = semiCompressibleBytes(512 * 1024)
    const source = new CountingZipByteSource(buildStoredZip(path, payload))
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    const hash = createHash('sha256')
    let streamedByteLength = 0

    expect(zip).not.toBeNull()
    source.resetCounts()
    expect(
      forEachInflatedXlsxZipEntryChunk(
        zip!,
        path,
        (chunk) => {
          streamedByteLength += chunk.byteLength
          hash.update(chunk)
        },
        {
          chunkSize: 4096,
        },
      ),
    ).toBe(true)

    expect(streamedByteLength).toBe(payload.byteLength)
    expect(hash.digest('hex')).toBe(createHash('sha256').update(payload).digest('hex'))
    expect(source.readIntoCount).toBeGreaterThan(0)
    expect(source.rangeCount).toBe(0)
    expect(source.maxReadLength).toBeLessThanOrEqual(4096)
  })
})

function buildStreamedZip(path: string, text: string | Uint8Array): Uint8Array {
  const fileName = Buffer.from(path)
  const payload = typeof text === 'string' ? Buffer.from(text) : Buffer.from(text)
  const compressed = deflateRawSync(payload)
  const localHeader = Buffer.alloc(30)
  localHeader.writeUInt32LE(0x04034b50, 0)
  localHeader.writeUInt16LE(20, 4)
  localHeader.writeUInt16LE(0x08, 6)
  localHeader.writeUInt16LE(8, 8)
  localHeader.writeUInt16LE(fileName.length, 26)

  const dataDescriptor = Buffer.alloc(16)
  dataDescriptor.writeUInt32LE(0x08074b50, 0)
  dataDescriptor.writeUInt32LE(compressed.length, 8)
  dataDescriptor.writeUInt32LE(payload.length, 12)

  const centralDirectoryOffset = localHeader.length + fileName.length + compressed.length + dataDescriptor.length
  const centralDirectory = Buffer.alloc(46)
  centralDirectory.writeUInt32LE(0x02014b50, 0)
  centralDirectory.writeUInt16LE(20, 4)
  centralDirectory.writeUInt16LE(20, 6)
  centralDirectory.writeUInt16LE(0x08, 8)
  centralDirectory.writeUInt16LE(8, 10)
  centralDirectory.writeUInt32LE(compressed.length, 20)
  centralDirectory.writeUInt32LE(payload.length, 24)
  centralDirectory.writeUInt16LE(fileName.length, 28)

  const endOfCentralDirectory = Buffer.alloc(22)
  endOfCentralDirectory.writeUInt32LE(0x06054b50, 0)
  endOfCentralDirectory.writeUInt16LE(1, 8)
  endOfCentralDirectory.writeUInt16LE(1, 10)
  endOfCentralDirectory.writeUInt32LE(centralDirectory.length + fileName.length, 12)
  endOfCentralDirectory.writeUInt32LE(centralDirectoryOffset, 16)

  return Buffer.concat([localHeader, fileName, compressed, dataDescriptor, centralDirectory, fileName, endOfCentralDirectory])
}

function buildStoredZip(path: string, text: Uint8Array): Uint8Array {
  const fileName = Buffer.from(path)
  const payload = Buffer.from(text)
  const localHeader = Buffer.alloc(30)
  localHeader.writeUInt32LE(0x04034b50, 0)
  localHeader.writeUInt16LE(20, 4)
  localHeader.writeUInt16LE(0, 8)
  localHeader.writeUInt32LE(payload.length, 18)
  localHeader.writeUInt32LE(payload.length, 22)
  localHeader.writeUInt16LE(fileName.length, 26)

  const centralDirectoryOffset = localHeader.length + fileName.length + payload.length
  const centralDirectory = Buffer.alloc(46)
  centralDirectory.writeUInt32LE(0x02014b50, 0)
  centralDirectory.writeUInt16LE(20, 4)
  centralDirectory.writeUInt16LE(20, 6)
  centralDirectory.writeUInt16LE(0, 10)
  centralDirectory.writeUInt32LE(payload.length, 20)
  centralDirectory.writeUInt32LE(payload.length, 24)
  centralDirectory.writeUInt16LE(fileName.length, 28)

  const endOfCentralDirectory = Buffer.alloc(22)
  endOfCentralDirectory.writeUInt32LE(0x06054b50, 0)
  endOfCentralDirectory.writeUInt16LE(1, 8)
  endOfCentralDirectory.writeUInt16LE(1, 10)
  endOfCentralDirectory.writeUInt32LE(centralDirectory.length + fileName.length, 12)
  endOfCentralDirectory.writeUInt32LE(centralDirectoryOffset, 16)

  return Buffer.concat([localHeader, fileName, payload, centralDirectory, fileName, endOfCentralDirectory])
}

class CountingZipByteSource implements XlsxZipByteSource {
  readonly byteLength: number
  maxReadLength = 0
  rangeCount = 0
  readIntoCount = 0

  constructor(private readonly bytes: Uint8Array) {
    this.byteLength = bytes.byteLength
  }

  readRange(start: number, end: number): Uint8Array {
    this.rangeCount += 1
    this.maxReadLength = Math.max(this.maxReadLength, end - start)
    return this.bytes.subarray(start, end)
  }

  readRangeInto(start: number, end: number, target: Uint8Array): Uint8Array {
    const length = end - start
    this.readIntoCount += 1
    this.maxReadLength = Math.max(this.maxReadLength, length)
    target.set(this.bytes.subarray(start, end), 0)
    return target.subarray(0, length)
  }

  resetCounts(): void {
    this.maxReadLength = 0
    this.rangeCount = 0
    this.readIntoCount = 0
  }
}

class NativeInflateZipByteSource extends CountingZipByteSource {
  nativeInflateCount = 0

  async inflateRawRangeChunksAsync(
    start: number,
    end: number,
    onChunk: (chunk: Uint8Array) => boolean | void,
    options: { readonly chunkSize: number },
  ): Promise<boolean> {
    this.nativeInflateCount += 1
    const inflated = inflateRawSync(this.readRange(start, end))
    const chunkSize = Math.max(1, Math.trunc(options.chunkSize))
    for (let offset = 0; offset < inflated.byteLength; offset += chunkSize) {
      if (onChunk(inflated.subarray(offset, Math.min(inflated.byteLength, offset + chunkSize))) === false) {
        return true
      }
    }
    return true
  }
}

function semiCompressibleBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length)
  let state = 0x12345678
  for (let index = 0; index < bytes.length; index += 1) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    bytes[index] = index % 10 === 0 ? state & 0xff : 65 + (index % 26)
  }
  return bytes
}

function corruptFirstCompressedByte(bytes: Uint8Array): Uint8Array {
  const nameLength = bytes[26] | (bytes[27] << 8)
  const extraLength = bytes[28] | (bytes[29] << 8)
  const compressedDataStart = 30 + nameLength + extraLength
  const corrupted = new Uint8Array(bytes)
  corrupted[compressedDataStart] = corrupted[compressedDataStart] ^ 0xff
  return corrupted
}
