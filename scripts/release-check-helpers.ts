import { gzipSync } from 'node:zlib'

export const releaseCheckGzipOptions = { level: 9 } as const

type GzipSync = typeof gzipSync

export function measureGzipBytes(bytes: Uint8Array, gzip: GzipSync = gzipSync): number {
  return gzip(bytes, releaseCheckGzipOptions).byteLength
}
