import { describe, expect, it } from 'vitest'
import type { Unzipped } from 'fflate'

import {
  hasFullImporterOnlyPackageMetadata,
  shouldBypassLargeSimpleByteThresholdForPackageArtifacts,
} from '../xlsx-large-simple-package-artifact-threshold.js'

describe('large simple XLSX package artifact threshold', () => {
  it('bypasses the byte threshold for data-model package artifacts', () => {
    expect(shouldBypassLargeSimpleByteThresholdForPackageArtifacts(zipWith(['xl/model/item.data']))).toBe(true)
    expect(shouldBypassLargeSimpleByteThresholdForPackageArtifacts(zipWith(['xl/customData/item1.xml']))).toBe(true)
    expect(shouldBypassLargeSimpleByteThresholdForPackageArtifacts(zipWith(['customXml/item1.xml']))).toBe(true)
  })

  it('keeps pivot-only packages on the normal fidelity fallback threshold', () => {
    expect(
      shouldBypassLargeSimpleByteThresholdForPackageArtifacts(
        zipWith(['xl/pivotTables/pivotTable1.xml', 'xl/pivotCache/pivotCacheDefinition1.xml']),
      ),
    ).toBe(false)
  })

  it('does not force SheetJS fallback for package metadata the streaming importer preserves', () => {
    expect(hasFullImporterOnlyPackageMetadata(zipWith(['xl/comments1.xml']))).toBe(true)
    expect(hasFullImporterOnlyPackageMetadata(zipWith(['xl/threadedComments/threadedComment1.xml']))).toBe(true)
    expect(hasFullImporterOnlyPackageMetadata(zipWith(['xl/printerSettings/printerSettings1.bin']))).toBe(false)
    expect(hasFullImporterOnlyPackageMetadata(zipWith(['xl/drawings/vmlDrawing1.vml']))).toBe(false)
  })
})

function zipWith(paths: readonly string[]): Unzipped {
  return Object.fromEntries(paths.map((path) => [path, new Uint8Array()]))
}
