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
      cases.push(...parsed.cases.map((entry) => parsePublicWorkbookCorpusCase(entry)))
      continue
    }
    cases.push(...parsePublicWorkbookCorpusScorecardJson(parsed).cases)
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

function isReusablePublicWorkbookCorpusCase(
  artifact: PublicWorkbookArtifact,
  candidate: PublicWorkbookCorpusCase,
  structuralSmokeRequired: boolean,
): boolean {
  return (
    candidate.passed &&
    candidate.id === artifact.id &&
    candidate.sourceId === artifact.sourceId &&
    candidate.sourceUrl === artifact.sourceUrl &&
    candidate.sha256 === artifact.sha256 &&
    candidate.byteSize === artifact.byteSize &&
    (!structuralSmokeRequired || candidate.validation.structuralSmokePassed !== null)
  )
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
