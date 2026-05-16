import { existsSync, statSync } from 'node:fs'
import { extname, join, resolve } from 'node:path'

import { assertPublicCorpusRunNotStopped, publicCorpusStopMarkerOverrideFlag } from './public-workbook-corpus-cli.ts'
import type { WorkPaperXlsxCorpusOptions } from './check-workpaper-xlsx-corpus-types.ts'

export interface WorkPaperXlsxCorpusCliOptions extends WorkPaperXlsxCorpusOptions {
  readonly isolateFiles: boolean
  readonly paths: readonly string[]
  readonly jsonOut?: string
  readonly maxMismatches: number
  readonly minMatchRate: number
  readonly stopMarkerPath: string
}

export interface WorkPaperXlsxCorpusInternalCliOptions extends WorkPaperXlsxCorpusOptions {
  readonly paths: readonly string[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultCorpusRunStopMarkerPath = join(rootDir, '.agent-coordination', '20260507T074946Z-codex-stop-interactive-corpus-runs.md')
const checkedInFixtureCorpusDirectories = new Set([resolve(rootDir, 'packages/headless/fixtures/xlsx-corpus')])
const xlsxExtensions = new Set(['.xls', '.xlsm', '.xlsx'])

export function parseWorkPaperXlsxCorpusCliArgs(argv: readonly string[]): WorkPaperXlsxCorpusCliOptions {
  const paths: string[] = []
  let childProcessTimeoutMs: number | undefined
  let jsonOut: string | undefined
  let maxMismatches = 0
  let minMatchRate = 1
  let evaluationTimeoutMs: number | undefined
  let isolateFiles = true
  let maxFileBytes: number | undefined
  let mismatchSampleLimit: number | undefined
  let stopMarkerPath = defaultCorpusRunStopMarkerPath

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index]
    switch (arg) {
      case '--help':
      case '-h':
        throw new CliUsageError(usageText(), 0)
      case '--allow-mismatches':
        maxMismatches = Number.POSITIVE_INFINITY
        minMatchRate = 0
        break
      case publicCorpusStopMarkerOverrideFlag:
        break
      case '--child-timeout-ms':
        childProcessTimeoutMs = parseNonNegativeInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--json-out':
        jsonOut = requiredArgValue(argv, index, arg)
        index += 1
        break
      case '--max-mismatches':
        maxMismatches = parseNonNegativeInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--max-file-bytes':
        maxFileBytes = parsePositiveInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--min-match-rate':
        minMatchRate = parseMatchRate(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--mismatch-sample-limit':
        mismatchSampleLimit = parseNonNegativeInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--timeout-ms':
        evaluationTimeoutMs = parseNonNegativeInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--no-isolate':
        if (process.env['BILIG_ALLOW_UNISOLATED_XLSX_CORPUS'] !== '1') {
          throw new CliUsageError(
            '--no-isolate is disabled for corpus CLI runs because it can retain workbook memory across large corpora. Set BILIG_ALLOW_UNISOLATED_XLSX_CORPUS=1 only for a single-file debugger run.',
            2,
          )
        }
        isolateFiles = false
        break
      case '--corpus-run-stop-marker':
        stopMarkerPath = resolve(requiredArgValue(argv, index, arg))
        index += 1
        break
      default:
        if (arg.startsWith('-')) {
          throw new CliUsageError(`Unknown option: ${arg}\n\n${usageText()}`, 2)
        }
        paths.push(arg)
        break
    }
  }

  if (paths.length === 0) {
    throw new CliUsageError(`Missing XLSX file or directory path.\n\n${usageText()}`, 2)
  }

  return {
    childProcessTimeoutMs,
    isolateFiles,
    paths,
    jsonOut,
    maxFileBytes,
    maxMismatches,
    minMatchRate,
    stopMarkerPath,
    evaluationTimeoutMs,
    mismatchSampleLimit,
  }
}

export function parseWorkPaperXlsxCorpusInternalCliArgs(argv: readonly string[]): WorkPaperXlsxCorpusInternalCliOptions {
  const paths: string[] = []
  let evaluationTimeoutMs: number | undefined
  let maxFileBytes: number | undefined
  let mismatchSampleLimit: number | undefined

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index]
    switch (arg) {
      case '--internal-check-file-json':
        paths.push(requiredArgValue(argv, index, arg))
        index += 1
        break
      case '--mismatch-sample-limit':
        mismatchSampleLimit = parseNonNegativeInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--max-file-bytes':
        maxFileBytes = parsePositiveInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--timeout-ms':
        evaluationTimeoutMs = parseNonNegativeInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      default:
        throw new CliUsageError(`Unknown internal option: ${arg}`, 2)
    }
  }

  if (paths.length !== 1) {
    throw new CliUsageError('Internal XLSX corpus check expects exactly one file path.', 2)
  }

  return {
    paths,
    evaluationTimeoutMs,
    maxFileBytes,
    mismatchSampleLimit,
  }
}

export function assertUnisolatedCliDebuggerPath(paths: readonly string[]): void {
  if (paths.length !== 1) {
    throw new CliUsageError('--no-isolate only supports one explicit XLSX file for debugger runs.', 2)
  }

  const path = resolve(paths[0])
  if (!existsSync(path) || !statSync(path).isFile() || !xlsxExtensions.has(extname(path).toLowerCase())) {
    throw new CliUsageError('--no-isolate only supports one explicit XLSX file for debugger runs.', 2)
  }
}

export function assertBroadCorpusSweepNotStopped(paths: readonly string[], stopMarkerPath: string): void {
  if (!isBroadCorpusSweep(paths)) {
    return
  }
  try {
    assertPublicCorpusRunNotStopped({
      commandName: 'workpaper:xlsx-corpus directory sweep',
      stopMarkerPath,
    })
  } catch (error) {
    if (error instanceof Error) {
      throw new CliUsageError(error.message, 2)
    }
    throw error
  }
}

export class CliUsageError extends Error {
  readonly exitCode: number

  constructor(message: string, exitCode: number) {
    super(message)
    this.name = 'CliUsageError'
    this.exitCode = exitCode
  }
}

function isBroadCorpusSweep(paths: readonly string[]): boolean {
  return (
    paths.length > 1 ||
    paths.some((entry) => {
      const path = resolve(entry)
      return existsSync(path) && statSync(path).isDirectory() && !checkedInFixtureCorpusDirectories.has(path)
    })
  )
}

function requiredArgValue(argv: readonly string[], index: number, option: string): string {
  const value = argv[index + 1]
  if (value === undefined || value.startsWith('-')) {
    throw new CliUsageError(`Missing value for ${option}\n\n${usageText()}`, 2)
  }
  return value
}

function parseNonNegativeInteger(value: string, option: string): number {
  const parsed = Number.parseInt(value, 10)
  if (!Number.isFinite(parsed) || parsed < 0 || String(parsed) !== value) {
    throw new CliUsageError(`${option} expects a non-negative integer, got ${value}`, 2)
  }
  return parsed
}

function parsePositiveInteger(value: string, option: string): number {
  const parsed = Number.parseInt(value, 10)
  if (!Number.isFinite(parsed) || parsed <= 0 || String(parsed) !== value) {
    throw new CliUsageError(`${option} expects a positive integer, got ${value}`, 2)
  }
  return parsed
}

function parseMatchRate(value: string, option: string): number {
  const parsed = Number(value)
  if (!/^(?:0(?:\.\d+)?|1(?:\.0+)?)$/.test(value) || !Number.isFinite(parsed) || parsed < 0 || parsed > 1) {
    throw new CliUsageError(`${option} expects a number between 0 and 1, got ${value}`, 2)
  }
  return parsed
}

function usageText(): string {
  return [
    'Usage: bun scripts/check-workpaper-xlsx-corpus.ts [options] <xlsx-file-or-directory> [...]',
    '',
    'Options:',
    '  --timeout-ms <ms>              WorkPaper initial evaluation timeout per workbook. Default: 30000.',
    '  --child-timeout-ms <ms>        Child-process timeout per workbook. Default: timeout-ms + 1000.',
    '  --max-file-bytes <bytes>       Fail a workbook before loading it when the file is larger. Default: 52428800.',
    '  --no-isolate                  Debug-only: requires BILIG_ALLOW_UNISOLATED_XLSX_CORPUS=1.',
    '  --max-mismatches <count>       Maximum comparable cached-result mismatches before failing. Default: 0.',
    '  --min-match-rate <ratio>       Minimum comparable cached-result match rate before failing. Default: 1.',
    '  --mismatch-sample-limit <n>    Number of mismatch samples to keep in JSON output. Default: 25.',
    '  --json-out <path>              Also write the JSON report to a file.',
    '  --corpus-run-stop-marker <path> Fail closed for directory or multi-file corpus sweeps while marker exists.',
    '  --allow-mismatches             Report mismatches without failing the process.',
  ].join('\n')
}
