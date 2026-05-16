import { describe, expect, it } from 'vitest'

import { formatFetchCheckpointProgress } from '../public-workbook-corpus-fetch-progress.ts'
import type { PublicWorkbookCorpusFetchCheckpointProgress } from '../public-workbook-corpus-types.ts'

describe('public workbook corpus fetch progress formatting', () => {
  it('formats checkpoint counters without a failure sample section when there are no failures', () => {
    expect(formatFetchCheckpointProgress(progress({ failedSourceSamples: [] }))).toBe(
      'Cached 12 public workbook artifacts; exhausted 8 sources; +3 exhausted this batch; 5 committed; 0 failed; 2 duplicate hashes; 1 duplicate fingerprints',
    )
  })

  it('normalizes and truncates failure samples before printing them', () => {
    const longError = `first line\n\tsecond line ${'x'.repeat(260)}`
    const formatted = formatFetchCheckpointProgress(
      progress({
        failedSourceSamples: [
          {
            sourceId: 'ckan:finance',
            fileName: 'budget workbook.xlsx',
            error: longError,
          },
        ],
      }),
    )
    const expectedError = longError.replace(/\s+/gu, ' ').slice(0, 240)

    expect(formatted).toContain(`failure samples: ckan:finance budget workbook.xlsx: ${expectedError}`)
    expect(expectedError).toHaveLength(240)
    expect(formatted).not.toContain('\n')
    expect(formatted).not.toContain('\t')
  })
})

function progress(overrides: Partial<PublicWorkbookCorpusFetchCheckpointProgress>): PublicWorkbookCorpusFetchCheckpointProgress {
  return {
    artifactCount: 12,
    exhaustedSourceCount: 8,
    exhaustedSourceDelta: 3,
    committedArtifactCount: 5,
    failedSourceCount: 0,
    duplicateHashSourceCount: 2,
    duplicateFingerprintSourceCount: 1,
    failedSourceSamples: [],
    ...overrides,
  }
}
