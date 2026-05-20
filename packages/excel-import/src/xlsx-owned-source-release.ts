import type { LargeSimpleXlsxOwnedSourceReleaseEvidence } from './xlsx-large-simple-import.js'

export interface OwnedXlsxSourceBytes {
  bytes: Uint8Array
}

export function releaseOwnedXlsxSourceBytes(
  source: OwnedXlsxSourceBytes,
  afterRelease?: (releasedBytes: Uint8Array) => void,
): LargeSimpleXlsxOwnedSourceReleaseEvidence {
  const ownedSourceBytesBeforeRelease = source.bytes.byteLength
  source.bytes = new Uint8Array(0)
  afterRelease?.(source.bytes)
  return {
    ownedSourceBytesBeforeRelease,
    ownedSourceBytesAfterRelease: source.bytes.byteLength,
  }
}
