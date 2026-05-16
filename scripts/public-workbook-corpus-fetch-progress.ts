import type { PublicWorkbookCorpusFetchCheckpointProgress } from './public-workbook-corpus-types.ts'

export function formatFetchCheckpointProgress(progress: PublicWorkbookCorpusFetchCheckpointProgress): string {
  const parts = [
    `Cached ${String(progress.artifactCount)} public workbook artifacts`,
    `exhausted ${String(progress.exhaustedSourceCount)} sources`,
    `+${String(progress.exhaustedSourceDelta)} exhausted this batch`,
    `${String(progress.committedArtifactCount)} committed`,
    `${String(progress.failedSourceCount)} failed`,
    `${String(progress.duplicateHashSourceCount)} duplicate hashes`,
    `${String(progress.duplicateFingerprintSourceCount)} duplicate fingerprints`,
  ]
  if (progress.failedSourceSamples.length > 0) {
    parts.push(`failure samples: ${progress.failedSourceSamples.map(formatFetchFailureSample).join(' | ')}`)
  }
  return parts.join('; ')
}

function formatFetchFailureSample(sample: PublicWorkbookCorpusFetchCheckpointProgress['failedSourceSamples'][number]): string {
  return `${sample.sourceId} ${sample.fileName}: ${sample.error.replace(/\s+/gu, ' ').slice(0, 240)}`
}
