import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import {
  DEFAULT_COMPETITIVE_SAMPLE_COUNT,
  DEFAULT_COMPETITIVE_WARMUP_COUNT,
} from '../packages/benchmarks/src/benchmark-workpaper-vs-hyperformula.ts'
import {
  WORKPAPER_XLSX_CALC_WORKLOADS,
  buildWorkPaperVsXlsxCalcBenchmarkReport,
  runWorkPaperVsXlsxCalcBenchmarkSuite,
  type WorkPaperXlsxCalcBenchmarkResult,
  type WorkPaperXlsxCalcScorecard,
} from '../packages/benchmarks/src/benchmark-workpaper-vs-xlsx-calc.ts'
import { readJsonObject } from './json-scorecard-helpers.ts'
import { formatJsonForRepo } from './scorecard-format.ts'
import { deriveWorkPaperXlsxCalcScorecard, parseWorkPaperXlsxCalcArtifact } from './workpaper-vs-xlsx-calc-artifact.ts'

interface WorkPaperVsXlsxCalcBenchmarkArtifact {
  readonly schemaVersion: 1
  readonly suite: 'workpaper-vs-xlsx-calc'
  readonly generatedAt: string
  readonly host: {
    readonly arch: string
    readonly nodeVersion: string
    readonly platform: string
  }
  readonly benchmark: {
    readonly sampleCount: number
    readonly warmupCount: number
  }
  readonly engines: {
    readonly workpaper: {
      readonly packageName: '@bilig/headless'
      readonly sourcePath: string
      readonly version: string
    }
    readonly xlsxCalc: {
      readonly coverageTier: 'workbook-wide'
      readonly packageName: 'xlsx-calc'
      readonly sourcePath: string
      readonly version: string
    }
  }
  readonly scorecard: WorkPaperXlsxCalcScorecard
  readonly results: readonly WorkPaperXlsxCalcBenchmarkResult[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-xlsx-calc.json')
const isCheckMode = process.argv.slice(2).includes('--check')
const sampleCount = DEFAULT_COMPETITIVE_SAMPLE_COUNT
const warmupCount = DEFAULT_COMPETITIVE_WARMUP_COUNT

if (isCheckMode) {
  if (!existsSync(outputPath)) {
    throw new Error('WorkPaper vs xlsx-calc benchmark artifact is missing. Run: pnpm workpaper:bench:xlsx-calc:generate')
  }

  const artifact = parseWorkPaperXlsxCalcArtifact(readJsonObject(outputPath))
  const actualWorkloads = artifact.results.map((result) => result.workload)
  if (JSON.stringify(actualWorkloads) !== JSON.stringify([...WORKPAPER_XLSX_CALC_WORKLOADS])) {
    throw new Error('WorkPaper vs xlsx-calc benchmark workload coverage is out of date. Run: pnpm workpaper:bench:xlsx-calc:generate')
  }

  const derivedScorecard = deriveWorkPaperXlsxCalcScorecard(artifact.results, artifact.scorecard.coverageNote)
  if (JSON.stringify(artifact.scorecard) !== JSON.stringify(derivedScorecard)) {
    throw new Error('WorkPaper vs xlsx-calc scorecard does not match benchmark results. Run: pnpm workpaper:bench:xlsx-calc:generate')
  }

  console.log(
    JSON.stringify(
      {
        mode: 'check',
        outputPath,
        workloads: actualWorkloads,
      },
      null,
      2,
    ),
  )
  process.exit(0)
}

const report = buildWorkPaperVsXlsxCalcBenchmarkReport(
  runWorkPaperVsXlsxCalcBenchmarkSuite({
    sampleCount,
    warmupCount,
  }),
)
const artifact: WorkPaperVsXlsxCalcBenchmarkArtifact = {
  schemaVersion: 1,
  suite: 'workpaper-vs-xlsx-calc',
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
      version: readPackageVersion(join(rootDir, 'packages', 'headless', 'package.json')),
    },
    xlsxCalc: {
      coverageTier: 'workbook-wide',
      packageName: 'xlsx-calc',
      sourcePath: join(rootDir, 'packages', 'benchmarks', 'node_modules', 'xlsx-calc'),
      version: readPackageVersion(join(rootDir, 'packages', 'benchmarks', 'node_modules', 'xlsx-calc', 'package.json')),
    },
  },
  scorecard: report.scorecard,
  results: report.results,
}

mkdirSync(dirname(outputPath), { recursive: true })
writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(artifact, null, 2)}\n`))
console.log(
  JSON.stringify(
    {
      mode: 'write',
      outputPath,
      workloads: artifact.results.map((result) => result.workload),
      meanAndP95WinCount: artifact.scorecard.meanAndP95WinCount,
      comparableWorkloadCount: artifact.scorecard.comparableWorkloadCount,
    },
    null,
    2,
  ),
)

function readPackageVersion(packagePath: string): string {
  const parsed: unknown = JSON.parse(readFileSync(packagePath, 'utf8'))
  if (!isRecord(parsed) || typeof parsed.version !== 'string' || parsed.version.length === 0) {
    throw new Error(`Unable to read package version from ${packagePath}`)
  }
  return parsed.version
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value)
}
