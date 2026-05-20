import { describe, expect, it } from 'vitest'
import type { Unzipped } from 'fflate'

import { shouldBypassLargeSimpleByteThresholdForPackageArtifacts } from '../xlsx-large-simple-package-artifact-threshold.js'

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
})

function zipWith(paths: readonly string[]): Unzipped {
  return Object.fromEntries(paths.map((path) => [path, new Uint8Array()]))
}
