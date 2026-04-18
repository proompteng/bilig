#!/usr/bin/env bun

import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import {
  DEFAULT_COMPETITIVE_SAMPLE_COUNT,
  DEFAULT_COMPETITIVE_WARMUP_COUNT,
} from '../packages/benchmarks/src/benchmark-workpaper-vs-hyperformula.ts'
import {
  EXPANDED_COMPARATIVE_WORKLOADS,
  runWorkPaperVsHyperFormulaExpandedBenchmarkSuite,
  type ExpandedComparativeBenchmarkResult,
} from '../packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts'

interface ExpandedCompetitiveBenchmarkArtifact {
  schemaVersion: 1
  suite: 'workpaper-vs-hyperformula-expanded'
  generatedAt: string
  host: {
    arch: string
    nodeVersion: string
    platform: string
  }
  benchmark: {
    sampleCount: number
    warmupCount: number
  }
  engines: {
    hyperformula: {
      commit: string
      licenseKey: string
      metadataSource: 'fallback' | 'local-checkout'
      packageName: 'hyperformula'
      sourcePath: string
      version: string
    }
    workpaper: {
      packageName: '@bilig/headless'
      sourcePath: string
      version: string
    }
  }
  results: ExpandedComparativeBenchmarkResult[]
}

interface ArtifactShapeInput {
  schemaVersion: 1
  suite: 'workpaper-vs-hyperformula-expanded'
  results: Array<{
    category: string
    comparable: boolean
    comparison?: Record<string, unknown>
    engines: {
      hyperformula: Record<string, unknown>
      workpaper: Record<string, unknown>
    }
    fixture: Record<string, unknown>
    note?: string
    workload: string
  }>
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-hyperformula-expanded.json')
const localHyperFormulaRoot = '/Users/gregkonush/github.com/hyperformula'
const isCheckMode = process.argv.slice(2).includes('--check')

const sampleCount = DEFAULT_COMPETITIVE_SAMPLE_COUNT
const warmupCount = DEFAULT_COMPETITIVE_WARMUP_COUNT

if (isCheckMode) {
  if (!existsSync(outputPath)) {
    throw new Error(
      'WorkPaper expanded competitive benchmark artifact is missing. Run: bun scripts/gen-workpaper-vs-hyperformula-expanded-benchmark.ts',
    )
  }

  const existing = parseArtifactForShape(readFileSync(outputPath, 'utf8'))
  const actualShape = normalizeArtifactShape(existing)
  const actualWorkloads = actualShape.workloads.map((workload) => workload.workload)
  if (JSON.stringify(actualWorkloads) !== JSON.stringify([...EXPANDED_COMPARATIVE_WORKLOADS])) {
    throw new Error(
      'WorkPaper expanded competitive benchmark artifact workload coverage is out of date. Run: bun scripts/gen-workpaper-vs-hyperformula-expanded-benchmark.ts',
    )
  }
  const expectedShape = normalizeArtifactShape({
    schemaVersion: 1,
    suite: 'workpaper-vs-hyperformula-expanded',
    results: [...EXPANDED_COMPARATIVE_WORKLOADS].map((workload) =>
      workload === 'dynamic-array-filter' ? leadershipShape(workload, [], []) : comparableShape(workload, [], []),
    ),
  })

  if (JSON.stringify(actualShape) !== JSON.stringify(expectedShape)) {
    throw new Error(
      'WorkPaper expanded competitive benchmark artifact shape is out of date. Run: bun scripts/gen-workpaper-vs-hyperformula-expanded-benchmark.ts',
    )
  }

  console.log(
    JSON.stringify(
      {
        mode: 'check',
        outputPath,
        workloads: actualShape.workloads.map((workload) => workload.workload),
      },
      null,
      2,
    ),
  )
  process.exit(0)
}

const workpaperVersion = readPackageVersion(join(rootDir, 'packages', 'headless', 'package.json'))
const hyperformulaMetadata = readHyperFormulaMetadata(localHyperFormulaRoot)

const artifact: ExpandedCompetitiveBenchmarkArtifact = {
  schemaVersion: 1,
  suite: 'workpaper-vs-hyperformula-expanded',
  generatedAt: new Date().toISOString(),
  host: {
    arch: process.arch,
    nodeVersion: process.version,
    platform: process.platform,
  },
  benchmark: {
    sampleCount,
    warmupCount,
  },
  engines: {
    workpaper: {
      packageName: '@bilig/headless',
      sourcePath: join(rootDir, 'packages', 'headless'),
      version: workpaperVersion,
    },
    hyperformula: hyperformulaMetadata,
  },
  results: runWorkPaperVsHyperFormulaExpandedBenchmarkSuite({
    sampleCount,
    warmupCount,
  }),
}

mkdirSync(dirname(outputPath), { recursive: true })
writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(artifact, null, 2)}\n`))
console.log(
  JSON.stringify(
    {
      mode: 'write',
      outputPath,
      workloads: artifact.results.map((result) => result.workload),
    },
    null,
    2,
  ),
)

function comparableShape(workload: string, fixtureKeys: string[], verificationKeys: string[]): ArtifactShapeInput['results'][number] {
  return {
    workload,
    category: 'directly-comparable',
    comparable: true,
    fixture: Object.fromEntries(fixtureKeys.map((key) => [key, 'placeholder'])),
    comparison: {
      fasterEngine: 'workpaper',
      meanSpeedup: 1,
      verificationEquivalent: true,
    },
    engines: {
      workpaper: measuredEngineShape(verificationKeys),
      hyperformula: measuredEngineShape(verificationKeys),
    },
  }
}

function leadershipShape(workload: string, fixtureKeys: string[], verificationKeys: string[]): ArtifactShapeInput['results'][number] {
  return {
    workload,
    category: 'leadership',
    comparable: false,
    fixture: Object.fromEntries(fixtureKeys.map((key) => [key, 'placeholder'])),
    note: 'placeholder',
    engines: {
      workpaper: measuredEngineShape(verificationKeys),
      hyperformula: {
        evidence: [],
        reason: '',
        status: 'unsupported',
      },
    },
  }
}

function measuredEngineShape(verificationKeys: string[]): Record<string, unknown> {
  return {
    status: 'supported',
    elapsedMs: {
      max: 0,
      mean: 0,
      median: 0,
      min: 0,
      p95: 0,
      samples: [],
    },
    memoryDeltaBytes: {
      arrayBuffersBytes: { max: 0, mean: 0, median: 0, min: 0, p95: 0, samples: [] },
      externalBytes: { max: 0, mean: 0, median: 0, min: 0, p95: 0, samples: [] },
      heapTotalBytes: { max: 0, mean: 0, median: 0, min: 0, p95: 0, samples: [] },
      heapUsedBytes: { max: 0, mean: 0, median: 0, min: 0, p95: 0, samples: [] },
      rssBytes: { max: 0, mean: 0, median: 0, min: 0, p95: 0, samples: [] },
    },
    verification: Object.fromEntries(verificationKeys.map((key) => [key, 'placeholder'])),
  }
}

function normalizeArtifactShape(input: ArtifactShapeInput) {
  return {
    schemaVersion: input.schemaVersion,
    suite: input.suite,
    workloads: input.results.map((result) => ({
      workload: result.workload,
      category: result.category,
      comparable: result.comparable,
      hasComparison: result.comparison !== undefined,
      hasNote: result.note !== undefined,
      hyperformulaStatus: result.engines.hyperformula.status,
    })),
  }
}

function parseArtifactForShape(json: string): ArtifactShapeInput {
  const parsed = parseJsonRecord(json)
  if (!isArtifactShapeInput(parsed)) {
    throw new Error('Expanded competitive benchmark artifact has an unexpected format')
  }
  return parsed
}

function readPackageVersion(packagePath: string): string {
  const pkg = parseJsonRecord(readFileSync(packagePath, 'utf8'))
  if (typeof pkg.version !== 'string' || pkg.version.length === 0) {
    throw new Error(`Unable to read package version from ${packagePath}`)
  }
  return pkg.version
}

function readHyperFormulaMetadata(localRoot: string): ExpandedCompetitiveBenchmarkArtifact['engines']['hyperformula'] {
  const fallback = {
    packageName: 'hyperformula' as const,
    version: '3.2.0',
    sourcePath: localRoot,
    licenseKey: 'gpl-v3',
    metadataSource: 'fallback' as const,
    commit: 'unknown',
  }

  if (!existsSync(localRoot)) {
    return fallback
  }

  const packageJsonPath = join(localRoot, 'package.json')
  if (!existsSync(packageJsonPath)) {
    return fallback
  }

  const pkg = parseJsonRecord(readFileSync(packageJsonPath, 'utf8'))
  const version = typeof pkg.version === 'string' && pkg.version.length > 0 ? pkg.version : fallback.version
  const commit = readGitCommit(localRoot)
  return {
    ...fallback,
    version,
    commit,
    metadataSource: 'local-checkout',
  }
}

function readGitCommit(cwd: string): string {
  try {
    return execFileSync('git', ['rev-parse', 'HEAD'], {
      cwd,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
    }).trim()
  } catch {
    return 'unknown'
  }
}

function parseJsonRecord(json: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(json)
  if (!isRecord(parsed)) {
    throw new Error('Expected JSON object')
  }
  return parsed
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value)
}

function isArtifactShapeInput(value: Record<string, unknown>): value is ArtifactShapeInput {
  return value.schemaVersion === 1 && value.suite === 'workpaper-vs-hyperformula-expanded' && Array.isArray(value.results)
}

function formatJsonForRepo(content: string): string {
  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-expanded-benchmark-'))
  const tempFile = join(tempDir, 'artifact.json')
  writeFileSync(tempFile, content)
  try {
    execFileSync('pnpm', ['exec', 'oxfmt', '--write', tempFile], {
      cwd: rootDir,
      stdio: 'ignore',
    })
    return readFileSync(tempFile, 'utf8')
  } finally {
    rmSync(tempDir, { force: true, recursive: true })
  }
}
