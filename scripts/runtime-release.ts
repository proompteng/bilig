import { incrementMajor, incrementMinor, incrementPatch } from './runtime-package-set.ts'

export const RUNTIME_RELEASE_TAG_PREFIX = 'libraries-v'

export const RUNTIME_AFFECTING_PATH_PATTERNS = [
  'packages/protocol/**',
  'packages/workbook-domain/**',
  'packages/wasm-kernel/**',
  'packages/formula/**',
  'packages/core/**',
  'packages/headless/**',
  'scripts/runtime-package-set.ts',
  'scripts/publish-runtime-package-set.ts',
  'scripts/check-package-publish.ts',
  'scripts/gen-formula-dominance-snapshot.ts',
  'scripts/gen-workpaper-hyperformula-audit.ts',
  'scripts/gen-workpaper-benchmark-baseline.ts',
  'scripts/gen-workpaper-vs-hyperformula-benchmark.ts',
  'scripts/workpaper-external-smoke.ts',
  '.github/workflows/headless-package.yml',
] as const

export type ReleaseType = 'none' | 'patch' | 'minor' | 'major'

export interface ConventionalCommit {
  type: string
  scope: string | null
  description: string
  breaking: boolean
}

export interface RuntimeReleaseCommit {
  sha: string
  shortSha: string
  subject: string
  body: string
  files: string[]
  runtimeAffecting: boolean
  conventional: ConventionalCommit | null
  releaseType: ReleaseType
}

const CONVENTIONAL_HEADER_PATTERN = /^(?<type>[a-z][a-z0-9-]*)(?:\((?<scope>[^()\r\n]+)\))?(?<breaking>!)?: (?<description>.+)$/u

const NO_RELEASE_TYPES = new Set(['build', 'chore', 'ci', 'docs', 'refactor', 'style', 'test'])
const PATCH_TYPES = new Set(['fix', 'perf', 'revert'])

export function parseConventionalCommit(input: { subject: string; body: string }): ConventionalCommit | null {
  const match = CONVENTIONAL_HEADER_PATTERN.exec(input.subject.trim())
  if (!match?.groups) {
    return null
  }
  const type = match.groups.type?.trim()
  const description = match.groups.description?.trim()
  if (!type || !description) {
    return null
  }
  const body = input.body.trim()
  return {
    type,
    scope: match.groups.scope?.trim() || null,
    description,
    breaking: Boolean(match.groups.breaking) || /(^|\n)BREAKING CHANGE:\s+/u.test(body),
  }
}

export function releaseTypeForConventionalCommit(commit: ConventionalCommit): ReleaseType {
  if (commit.breaking) {
    return 'major'
  }
  if (commit.type === 'feat') {
    return 'minor'
  }
  if (PATCH_TYPES.has(commit.type)) {
    return 'patch'
  }
  if (NO_RELEASE_TYPES.has(commit.type)) {
    return 'none'
  }
  return 'none'
}

export function compareReleaseTypes(left: ReleaseType, right: ReleaseType): number {
  return releaseWeight(left) - releaseWeight(right)
}

export function maxReleaseType(left: ReleaseType, right: ReleaseType): ReleaseType {
  return compareReleaseTypes(left, right) >= 0 ? left : right
}

export function bumpVersion(version: string, releaseType: Exclude<ReleaseType, 'none'>): string {
  switch (releaseType) {
    case 'patch':
      return incrementPatch(version)
    case 'minor':
      return incrementMinor(version)
    case 'major':
      return incrementMajor(version)
  }
}

export function extractVersionFromRuntimeTag(tagName: string): string | null {
  if (!tagName.startsWith(RUNTIME_RELEASE_TAG_PREFIX)) {
    return null
  }
  return tagName.slice(RUNTIME_RELEASE_TAG_PREFIX.length)
}

export function isRuntimeAffectingPath(path: string): boolean {
  return RUNTIME_AFFECTING_PATH_PATTERNS.some((pattern) => matchesPathPattern(path, pattern))
}

export function summarizeReleaseNotes(input: {
  targetVersion: string
  lastTag: string | null
  commits: RuntimeReleaseCommit[]
  releaseType: Exclude<ReleaseType, 'none'>
  manualOverride: boolean
}): string {
  const breaking = input.commits.filter((commit) => commit.releaseType === 'major')
  const features = input.commits.filter(
    (commit) => commit.runtimeAffecting && commit.releaseType === 'minor' && commit.releaseType !== 'major',
  )
  const fixes = input.commits.filter((commit) => commit.runtimeAffecting && commit.releaseType === 'patch')
  const internal = input.commits.filter(
    (commit) => commit.runtimeAffecting && commit.releaseType === 'none' && commit.conventional !== null,
  )

  const lines = [`# Libraries v${input.targetVersion}`, '']

  lines.push(`- Release type: ${input.releaseType}`)
  lines.push(`- Previous libraries tag: ${input.lastTag ?? 'none'}`)
  lines.push(`- Manual override: ${input.manualOverride ? 'yes' : 'no'}`)
  lines.push('')

  pushCommitSection(lines, 'Breaking changes', breaking)
  pushCommitSection(lines, 'Features', features)
  pushCommitSection(lines, 'Fixes', fixes)
  pushCommitSection(lines, 'Internal runtime changes', internal)

  return `${lines.join('\n').trim()}\n`
}

function pushCommitSection(lines: string[], title: string, commits: RuntimeReleaseCommit[]): void {
  if (commits.length === 0) {
    return
  }
  lines.push(`## ${title}`)
  for (const commit of commits) {
    lines.push(`- ${commit.subject} (${commit.shortSha})`)
  }
  lines.push('')
}

function releaseWeight(value: ReleaseType): number {
  switch (value) {
    case 'none':
      return 0
    case 'patch':
      return 1
    case 'minor':
      return 2
    case 'major':
      return 3
  }
}

function matchesPathPattern(path: string, pattern: string): boolean {
  if (pattern.endsWith('/**')) {
    return path.startsWith(pattern.slice(0, -2))
  }
  return path === pattern
}
