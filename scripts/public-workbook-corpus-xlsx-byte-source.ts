import { closeSync, fstatSync, openSync, readSync } from 'node:fs'

import type { XlsxZipByteSource } from '../packages/excel-import/src/xlsx-zip.js'

export class FileBackedXlsxZipByteSource implements XlsxZipByteSource {
  readonly byteLength: number
  private fd: number | null

  constructor(path: string) {
    this.fd = openSync(path, 'r')
    this.byteLength = fstatSync(this.fd).size
  }

  readRange(start: number, end: number): Uint8Array {
    if (this.fd === null) {
      throw new Error('XLSX ZIP file source has been released')
    }
    if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > this.byteLength) {
      throw new Error('Invalid XLSX ZIP file byte range')
    }
    const output = Buffer.allocUnsafe(end - start)
    let offset = 0
    while (offset < output.byteLength) {
      const bytesRead = readSync(this.fd, output, offset, output.byteLength - offset, start + offset)
      if (bytesRead === 0) {
        throw new Error('Unexpected end of XLSX ZIP file source')
      }
      offset += bytesRead
    }
    return output
  }

  release(): void {
    if (this.fd === null) {
      return
    }
    closeSync(this.fd)
    this.fd = null
  }
}

export function isZipWorkbookSource(source: XlsxZipByteSource): boolean {
  const bytes = source.readRange(0, Math.min(4, source.byteLength))
  return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04
}
