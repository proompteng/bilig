#!/usr/bin/env bun

import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import {
  DEFAULT_COMPETITIVE_SAMPLE_COUNT,
  DEFAULT_COMPETITIVE_WARMUP_COUNT,
  runWorkPaperVsHyperFormulaBenchmarkSuite,
  type ComparativeBenchmarkResult,
} from '../packages/benchmarks/src/benchmark-workpaper-vs-hyperformula.ts'

interface CompetitiveBenchmarkArtifact {
  schemaVersion: 1
  suite: 'workpaper-vs-hyperformula'
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
  results: ComparativeBenchmarkResult[]
}

interface CompetitiveBenchmarkArtifactShapeInput {
  schemaVersion: 1
  suite: 'workpaper-vs-hyperformula'
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
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-hyperformula.json')
const localHyperFormulaRoot = '/Users/gregkonush/github.com/hyperformula'
const isCheckMode = process.argv.slice(2).includes('--check')

const sampleCount = DEFAULT_COMPETITIVE_SAMPLE_COUNT
const warmupCount = DEFAULT_COMPETITIVE_WARMUP_COUNT

if (isCheckMode) {
  if (!existsSync(outputPath)) {
    throw new Error(`WorkPaper competitive benchmark artifact is missing. Run: bun scripts/gen-workpaper-vs-hyperformula-benchmark.ts`)
  }

  const existing = parseArtifactForShape(readFileSync(outputPath, 'utf8'))
  const actualShape = normalizeArtifactShape(existing)
  const expectedShape = normalizeArtifactShape({
    schemaVersion: 1,
    suite: 'workpaper-vs-hyperformula',
    results: [
      comparableShape('build-from-sheets', ['cols', 'materializedCells', 'rows'], ['dimensions', 'terminalValue']),
      comparableShape('single-edit-recalc', ['downstreamCount'], ['changeCount', 'terminalFormula', 'terminalValue']),
      comparableShape('batch-edit-recalc', ['editCount'], ['changeCount', 'sampleFormulaValue']),
      comparableShape('range-read', ['cols', 'materializedCells', 'rows'], ['readCols', 'readRows', 'terminalValue', 'topLeftValue']),
      comparableShape('lookup-no-column-index', ['rowCount', 'useColumnIndex'], ['changeCount', 'formulaValue']),
      comparableShape('lookup-with-column-index', ['rowCount', 'useColumnIndex'], ['changeCount', 'formulaValue']),
      leadershipShape('dynamic-array-filter', ['formula', 'rowCount'], ['changeCount', 'spillHeight', 'spillIsArray', 'spillValue']),
    ],
  })

  if (JSON.stringify(actualShape) !== JSON.stringify(expectedShape)) {
    throw new Error(
      `WorkPaper competitive benchmark artifact shape is out of date. Run: bun scripts/gen-workpaper-vs-hyperformula-benchmark.ts`,
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

const artifact: CompetitiveBenchmarkArtifact = {
  schemaVersion: 1,
  suite: 'workpaper-vs-hyperformula',
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
  results: runWorkPaperVsHyperFormulaBenchmarkSuite({
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

function comparableShape(
  workload: string,
  fixtureKeys: string[],
  verificationKeys: string[],
): CompetitiveBenchmarkArtifactShapeInput['results'][number] {
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

function leadershipShape(
  workload: string,
  fixtureKeys: string[],
  verificationKeys: string[],
): CompetitiveBenchmarkArtifactShapeInput['results'][number] {
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
      arrayBuffersBytes: {
        max: 0,
        mean: 0,
        median: 0,
        min: 0,
        p95: 0,
        samples: [],
      },
      externalBytes: {
        max: 0,
        mean: 0,
        median: 0,
        min: 0,
        p95: 0,
        samples: [],
      },
      heapTotalBytes: {
        max: 0,
        mean: 0,
        median: 0,
        min: 0,
        p95: 0,
        samples: [],
      },
      heapUsedBytes: {
        max: 0,
        mean: 0,
        median: 0,
        min: 0,
        p95: 0,
        samples: [],
      },
      rssBytes: {
        max: 0,
        mean: 0,
        median: 0,
        min: 0,
        p95: 0,
        samples: [],
      },
    },
    verification: Object.fromEntries(verificationKeys.map((key) => [key, 'placeholder'])),
  }
}

function parseArtifactForShape(serialized: string): CompetitiveBenchmarkArtifactShapeInput {
  const candidate = toRecord(JSON.parse(serialized), `competitive benchmark artifact ${outputPath}`)
  if (candidate.schemaVersion !== 1 || candidate.suite !== 'workpaper-vs-hyperformula') {
    throw new Error(`Unexpected competitive benchmark artifact header: ${outputPath}`)
  }

  const results = candidate.results
  if (!Array.isArray(results)) {
    throw new Error(`Competitive benchmark artifact is missing results: ${outputPath}`)
  }

  return {
    schemaVersion: 1,
    suite: 'workpaper-vs-hyperformula',
    results: results.map((result, index) => {
      const record = toRecord(result, `competitive benchmark result ${index}`)
      return {
        workload: requireString(record.workload, `competitive benchmark result ${index} workload`),
        category: requireString(record.category, `competitive benchmark result ${index} category`),
        comparable: requireBoolean(record.comparable, `competitive benchmark result ${index} comparable`),
        fixture: toRecord(record.fixture, `competitive benchmark result ${index} fixture`),
        comparison:
          record.comparison === undefined ? undefined : toRecord(record.comparison, `competitive benchmark result ${index} comparison`),
        note: record.note === undefined ? undefined : requireString(record.note, `competitive benchmark result ${index} note`),
        engines: {
          workpaper: toRecord(
            toRecord(record.engines, `competitive benchmark result ${index} engines`).workpaper,
            `competitive benchmark result ${index} workpaper engine`,
          ),
          hyperformula: toRecord(
            toRecord(record.engines, `competitive benchmark result ${index} engines`).hyperformula,
            `competitive benchmark result ${index} hyperformula engine`,
          ),
        },
      }
    }),
  }
}

function normalizeArtifactShape(input: CompetitiveBenchmarkArtifactShapeInput): {
  schemaVersion: 1
  suite: 'workpaper-vs-hyperformula'
  workloads: Array<{
    category: string
    comparable: boolean
    comparisonKeys: string[]
    fixtureKeys: string[]
    hyperformulaStatus: string
    hyperformulaVerificationKeys: string[]
    notePresent: boolean
    workpaperStatus: string
    workpaperVerificationKeys: string[]
    workload: string
  }>
} {
  return {
    schemaVersion: 1,
    suite: 'workpaper-vs-hyperformula',
    workloads: input.results.map((result) => ({
      workload: result.workload,
      category: result.category,
      comparable: result.comparable,
      fixtureKeys: Object.keys(result.fixture).toSorted(),
      comparisonKeys: result.comparison ? Object.keys(result.comparison).toSorted() : [],
      workpaperStatus: requireString(result.engines.workpaper.status, `${result.workload} workpaper status`),
      workpaperVerificationKeys: extractVerificationKeys(result.engines.workpaper),
      hyperformulaStatus: requireString(result.engines.hyperformula.status, `${result.workload} hyperformula status`),
      hyperformulaVerificationKeys: extractVerificationKeys(result.engines.hyperformula),
      notePresent: result.note !== undefined,
    })),
  }
}

function extractVerificationKeys(engine: Record<string, unknown>): string[] {
  const verification = engine.verification
  if (verification === undefined) {
    return []
  }
  return Object.keys(toRecord(verification, 'engine verification')).toSorted()
}

function readHyperFormulaMetadata(root: string): CompetitiveBenchmarkArtifact['engines']['hyperformula'] {
  const fallback = {
    commit: '6de904b8876f920f287b63a95934c479acf78307',
    metadataSource: 'fallback' as const,
    packageName: 'hyperformula' as const,
    sourcePath: root,
    version: '3.2.0',
    licenseKey: 'gpl-v3',
  }
  if (!existsSync(root)) {
    return fallback
  }

  const packagePath = join(root, 'package.json')
  const version = existsSync(packagePath) ? readPackageVersion(packagePath) : fallback.version
  let commit = fallback.commit
  try {
    commit = execFileSync('git', ['-C', root, 'rev-parse', 'HEAD'], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
    }).trim()
  } catch {
    commit = fallback.commit
  }

  return {
    commit,
    licenseKey: 'gpl-v3',
    metadataSource: 'local-checkout',
    packageName: 'hyperformula',
    sourcePath: root,
    version,
  }
}

function readPackageVersion(packagePath: string): string {
  const candidate = toRecord(JSON.parse(readFileSync(packagePath, 'utf8')), packagePath)
  return requireString(candidate.version, `${packagePath} version`)
}

function requireString(value: unknown, context: string): string {
  if (typeof value !== 'string') {
    throw new Error(`Expected ${context} to be a string`)
  }
  return value
}

function requireBoolean(value: unknown, context: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`Expected ${context} to be a boolean`)
  }
  return value
}

function toRecord(value: unknown, context: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`Expected ${context} to be an object`)
  }
  const record: Record<string, unknown> = {}
  for (const [key, entryValue] of Object.entries(value)) {
    record[key] = entryValue
  }
  return record
}

function formatJsonForRepo(serializedJson: string): string {
  const tempDir = mkdtempSync(join(tmpdir(), 'workpaper-competitive-bench-'))
  const tempFilePath = join(tempDir, 'artifact.json')
  writeFileSync(tempFilePath, serializedJson)
  const oxfmtPath = join(rootDir, 'node_modules', '.bin', 'oxfmt')

  const formatResult = Bun.spawnSync([oxfmtPath, '--write', tempFilePath], {
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (formatResult.exitCode !== 0) {
    const stderr = new TextDecoder().decode(formatResult.stderr).trim()
    rmSync(tempDir, { recursive: true, force: true })
    throw new Error(`Unable to format generated competitive benchmark artifact: ${stderr}`)
  }

  const formattedJson = readFileSync(tempFilePath, 'utf8')
  rmSync(tempDir, { recursive: true, force: true })
  return formattedJson
}
