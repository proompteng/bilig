import type { WorkbookPreservedPackagePartSnapshot } from '@bilig/protocol'

const binaryChunkSize = 0x8000

function encodeBinaryString(bytes: Uint8Array): string {
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += binaryChunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + binaryChunkSize))
  }
  return binary
}

function decodeBinaryString(binary: string): Uint8Array {
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index)
  }
  return bytes
}

function encodeBase64(bytes: Uint8Array): string {
  const btoa = globalThis.btoa
  if (typeof btoa === 'function') {
    return btoa(encodeBinaryString(bytes))
  }
  return Buffer.from(bytes).toString('base64')
}

function decodeBase64(dataBase64: string): Uint8Array {
  const atob = globalThis.atob
  if (typeof atob === 'function') {
    return decodeBinaryString(atob(dataBase64))
  }
  return new Uint8Array(Buffer.from(dataBase64, 'base64'))
}

class LazyEncodedPartSnapshot implements WorkbookPreservedPackagePartSnapshot {
  readonly storage = 'base64' as const
  declare readonly dataBase64: string
  private dataBase64Cache: string | undefined

  constructor(
    readonly path: string,
    readonly byteLength: number,
    private readonly readBytes: () => Uint8Array | undefined,
  ) {
    Object.defineProperty(this, 'dataBase64', {
      configurable: true,
      enumerable: true,
      get: () => this.getDataBase64(),
    })
  }

  private getDataBase64(): string {
    this.dataBase64Cache ??= encodeBase64(this.readBytes() ?? new Uint8Array())
    return this.dataBase64Cache
  }
}

export function encodedPartSnapshot(path: string, bytes: Uint8Array): WorkbookPreservedPackagePartSnapshot {
  return {
    path,
    storage: 'base64',
    dataBase64: encodeBase64(bytes),
    byteLength: bytes.byteLength,
  }
}

export function lazyEncodedPartSnapshot(
  path: string,
  byteLength: number,
  readBytes: () => Uint8Array | undefined,
): WorkbookPreservedPackagePartSnapshot {
  return new LazyEncodedPartSnapshot(path, byteLength, readBytes)
}

export function materializePreservedPackageParts(parts: readonly WorkbookPreservedPackagePartSnapshot[]): {
  readonly byteLength: number
  readonly partCount: number
} {
  let byteLength = 0
  let partCount = 0
  for (const part of parts) {
    if (part.storage !== 'base64') {
      continue
    }
    void part.dataBase64
    byteLength += part.byteLength
    partCount += 1
  }
  return { byteLength, partCount }
}

export function decodedPartBytes(part: WorkbookPreservedPackagePartSnapshot): Uint8Array | undefined {
  if (part.storage !== 'base64') {
    return undefined
  }
  const bytes = decodeBase64(part.dataBase64)
  return bytes.byteLength === part.byteLength ? bytes : undefined
}
