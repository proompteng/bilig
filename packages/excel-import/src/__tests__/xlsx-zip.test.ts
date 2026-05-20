import { deflateRawSync } from 'node:zlib'

import { describe, expect, it } from 'vitest'

import {
  forEachInflatedXlsxZipEntryChunk,
  getZipText,
  readLazyXlsxZipSource,
  readLazyXlsxZipSourceByteLength,
  readXlsxZipEntriesLazyFromByteSource,
  readXlsxZipEntries,
  readXlsxZipEntriesLazy,
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
    expect(source.maxReadLength).toBeLessThan(16 * 1024)
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

class CountingZipByteSource implements XlsxZipByteSource {
  readonly byteLength: number
  maxReadLength = 0

  constructor(private readonly bytes: Uint8Array) {
    this.byteLength = bytes.byteLength
  }

  readRange(start: number, end: number): Uint8Array {
    this.maxReadLength = Math.max(this.maxReadLength, end - start)
    return this.bytes.subarray(start, end)
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
