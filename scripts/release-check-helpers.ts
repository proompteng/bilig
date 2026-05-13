import { gzipSync } from 'node:zlib'

export const releaseCheckGzipOptions = { level: 9 } as const

type GzipMeasure = (bytes: Uint8Array, options: typeof releaseCheckGzipOptions) => { readonly byteLength: number }

export function measureGzipBytes(bytes: Uint8Array, gzip: GzipMeasure = gzipSync): number {
  return gzip(bytes, releaseCheckGzipOptions).byteLength
}
