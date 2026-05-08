import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookManifest } from './public-workbook-corpus-types.ts'
import { publicWorkbookCorpusCaseNeedsEvidenceRefresh } from './public-workbook-corpus-evidence.ts'

export function selectMissingPublicWorkbookArtifacts(args: {
  readonly manifest: PublicWorkbookManifest
  readonly cases: readonly PublicWorkbookCorpusCase[]
  readonly limit: number
}): PublicWorkbookArtifact[] {
  const normalizedLimit = Math.max(0, Math.trunc(args.limit))
  if (normalizedLimit === 0) {
    return []
  }
  return listMissingPublicWorkbookArtifacts(args).slice(0, normalizedLimit)
}

export function listMissingPublicWorkbookArtifacts(args: {
  readonly manifest: PublicWorkbookManifest
  readonly cases: readonly PublicWorkbookCorpusCase[]
}): PublicWorkbookArtifact[] {
  const casesById = new Map(args.cases.map((entry) => [entry.id, entry]))
  return args.manifest.artifacts
    .filter((artifact) => {
      const corpusCase = casesById.get(artifact.id)
      return !corpusCase || !publicWorkbookCorpusCaseMatchesArtifact(corpusCase, artifact)
    })
    .toSorted(compareVerificationArtifactPriority)
}

export function selectStalePublicWorkbookArtifacts(args: {
  readonly manifest: PublicWorkbookManifest
  readonly cases: readonly PublicWorkbookCorpusCase[]
  readonly limit: number
}): PublicWorkbookArtifact[] {
  const normalizedLimit = Math.max(0, Math.trunc(args.limit))
  if (normalizedLimit === 0) {
    return []
  }
  return listStalePublicWorkbookArtifacts(args).slice(0, normalizedLimit)
}

export function listStalePublicWorkbookArtifacts(args: {
  readonly manifest: PublicWorkbookManifest
  readonly cases: readonly PublicWorkbookCorpusCase[]
}): PublicWorkbookArtifact[] {
  const casesById = new Map(args.cases.map((entry) => [entry.id, entry]))
  return args.manifest.artifacts
    .filter((artifact) => {
      const corpusCase = casesById.get(artifact.id)
      return (
        corpusCase !== undefined &&
        publicWorkbookCorpusCaseMatchesArtifact(corpusCase, artifact) &&
        publicWorkbookCorpusCaseNeedsEvidenceRefresh(corpusCase)
      )
    })
    .toSorted(compareVerificationArtifactPriority)
}

export function indexPublicWorkbookCorpusCases(cases: readonly PublicWorkbookCorpusCase[]): Map<string, PublicWorkbookCorpusCase> {
  return new Map(cases.map((entry) => [entry.id, entry]))
}

export function publicWorkbookCorpusCaseMatchesArtifact(corpusCase: PublicWorkbookCorpusCase, artifact: PublicWorkbookArtifact): boolean {
  return (
    corpusCase.id === artifact.id &&
    corpusCase.sourceId === artifact.sourceId &&
    corpusCase.sourceUrl === artifact.sourceUrl &&
    corpusCase.fileName === artifact.fileName &&
    corpusCase.sha256 === artifact.sha256 &&
    corpusCase.byteSize === artifact.byteSize
  )
}

function compareVerificationArtifactPriority(left: PublicWorkbookArtifact, right: PublicWorkbookArtifact): number {
  return left.byteSize - right.byteSize || left.fileName.localeCompare(right.fileName) || left.id.localeCompare(right.id)
}
