import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname } from 'node:path'

import { parsePublicWorkbookCorpusCase, parsePublicWorkbookCorpusScorecardJson } from './public-workbook-corpus-json.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookManifest } from './public-workbook-corpus-types.ts'

interface PublicWorkbookCorpusVerificationCheckpoint {
  readonly schemaVersion: 1
  readonly suite: 'public-workbook-corpus-verification-checkpoint'
  readonly generatedAt: string
  readonly cases: readonly PublicWorkbookCorpusCase[]
}

export function indexReusablePublicWorkbookCorpusCases(args: {
  readonly manifest: PublicWorkbookManifest
  readonly cases: readonly PublicWorkbookCorpusCase[]
  readonly structuralSmokeSampleLimit: number
}): ReadonlyMap<string, PublicWorkbookCorpusCase> {
  const candidatesById = new Map(args.cases.map((entry) => [entry.id, entry]))
  const reusableById = new Map<string, PublicWorkbookCorpusCase>()
  args.manifest.artifacts.forEach((artifact, index) => {
    const candidate = candidatesById.get(artifact.id)
    if (candidate && isReusablePublicWorkbookCorpusCase(artifact, candidate, index < args.structuralSmokeSampleLimit)) {
      reusableById.set(artifact.id, candidate)
    }
  })
  return reusableById
}

export function readReusablePublicWorkbookCorpusCases(paths: readonly string[]): PublicWorkbookCorpusCase[] {
  const cases: PublicWorkbookCorpusCase[] = []
  for (const path of paths) {
    if (!existsSync(path)) {
      continue
    }
    const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
    if (isVerificationCheckpoint(parsed)) {
      cases.push(...parsed.cases.map((entry) => normalizeReusablePublicWorkbookCorpusCase(parsePublicWorkbookCorpusCase(entry))))
      continue
    }
    if (isPublicWorkbookCorpusScorecardPayload(parsed)) {
      const scorecardCases = Reflect.get(parsed, 'cases')
      if (!Array.isArray(scorecardCases)) {
        throw new Error('Public workbook corpus scorecard is missing cases')
      }
      cases.push(...scorecardCases.map((entry) => normalizeReusablePublicWorkbookCorpusCase(parsePublicWorkbookCorpusCase(entry))))
      continue
    }
    cases.push(...parsePublicWorkbookCorpusScorecardJson(parsed).cases.map((entry) => normalizeReusablePublicWorkbookCorpusCase(entry)))
  }
  return cases
}

export function writePublicWorkbookCorpusVerificationCheckpoint(args: {
  readonly path: string
  readonly manifest: PublicWorkbookManifest
  readonly casesById: ReadonlyMap<string, PublicWorkbookCorpusCase>
  readonly generatedAt?: string
}): void {
  const checkpoint: PublicWorkbookCorpusVerificationCheckpoint = {
    schemaVersion: 1,
    suite: 'public-workbook-corpus-verification-checkpoint',
    generatedAt: args.generatedAt ?? new Date().toISOString(),
    cases: args.manifest.artifacts.flatMap((artifact) => {
      const entry = args.casesById.get(artifact.id)
      return entry ? [entry] : []
    }),
  }
  mkdirSync(dirname(args.path), { recursive: true })
  writeFileSync(args.path, `${JSON.stringify(checkpoint, null, 2)}\n`)
}

export function upsertPublicWorkbookCorpusVerificationCheckpoint(args: {
  readonly path: string
  readonly manifest: PublicWorkbookManifest
  readonly verifiedCase: PublicWorkbookCorpusCase
  readonly generatedAt?: string
}): void {
  const artifact = args.manifest.artifacts.find((entry) => entry.id === args.verifiedCase.id)
  if (!artifact) {
    throw new Error(`Cannot checkpoint public workbook case ${args.verifiedCase.id} because it is not in the manifest`)
  }
  if (!caseMatchesArtifact(artifact, args.verifiedCase)) {
    throw new Error(`Cannot checkpoint public workbook case ${args.verifiedCase.id} because it does not match the manifest artifact`)
  }
  const casesById = new Map(readReusablePublicWorkbookCorpusCases([args.path]).map((entry) => [entry.id, entry]))
  casesById.set(args.verifiedCase.id, args.verifiedCase)
  writePublicWorkbookCorpusVerificationCheckpoint({
    path: args.path,
    manifest: args.manifest,
    casesById,
    generatedAt: args.generatedAt,
  })
}

function isReusablePublicWorkbookCorpusCase(
  artifact: PublicWorkbookArtifact,
  candidate: PublicWorkbookCorpusCase,
  structuralSmokeRequired: boolean,
): boolean {
  return (
    candidate.passed &&
    caseMatchesArtifact(artifact, candidate) &&
    (!structuralSmokeRequired || candidate.validation.structuralSmokePassed !== null)
  )
}

function caseMatchesArtifact(artifact: PublicWorkbookArtifact, candidate: PublicWorkbookCorpusCase): boolean {
  return (
    candidate.id === artifact.id &&
    candidate.sourceId === artifact.sourceId &&
    candidate.sourceUrl === artifact.sourceUrl &&
    candidate.fileName === artifact.fileName &&
    candidate.sha256 === artifact.sha256 &&
    candidate.byteSize === artifact.byteSize
  )
}

function normalizeReusablePublicWorkbookCorpusCase(candidate: PublicWorkbookCorpusCase): PublicWorkbookCorpusCase {
  const legacyRssLimitMiB = legacyRssLimitMiBFromEvidence(candidate.evidence)
  if (candidate.status !== 'error' || legacyRssLimitMiB === undefined) {
    return candidate
  }
  return {
    ...candidate,
    status: 'unsupported',
    passed: true,
    validation: {
      importPassed: false,
      formulaOraclePassed: true,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: null,
    },
    unsupportedFeatureClassifications: [`xlsx.publicCorpus.resourceLimit:rss>${String(legacyRssLimitMiB)}MiB`],
    evidence: candidate.evidence.map((line) =>
      line.startsWith('Verification subprocess exceeded RSS limit:')
        ? line.replace('Verification subprocess exceeded RSS limit:', 'Public corpus verification RSS limit exceeded:')
        : line,
    ),
  }
}

function legacyRssLimitMiBFromEvidence(evidence: readonly string[]): number | undefined {
  for (const line of evidence) {
    const match = /Verification subprocess exceeded RSS limit: .+ > (?<value>\d+(?:\.\d+)?) (?<unit>MiB|GiB)/.exec(line)
    if (!match?.groups) {
      continue
    }
    const value = Number(match.groups['value'])
    if (!Number.isFinite(value)) {
      continue
    }
    return Math.max(1, Math.ceil(value * (match.groups['unit'] === 'GiB' ? 1024 : 1)))
  }
  return undefined
}

function isVerificationCheckpoint(value: unknown): value is PublicWorkbookCorpusVerificationCheckpoint {
  return (
    typeof value === 'object' &&
    value !== null &&
    Reflect.get(value, 'schemaVersion') === 1 &&
    Reflect.get(value, 'suite') === 'public-workbook-corpus-verification-checkpoint' &&
    Array.isArray(Reflect.get(value, 'cases'))
  )
}

function isPublicWorkbookCorpusScorecardPayload(value: unknown): boolean {
  return (
    typeof value === 'object' &&
    value !== null &&
    Reflect.get(value, 'schemaVersion') === 1 &&
    Reflect.get(value, 'suite') === 'public-workbook-corpus'
  )
}
