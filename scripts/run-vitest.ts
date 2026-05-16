#!/usr/bin/env bun

import { spawnSync } from 'node:child_process'
import { resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { assertLocalCiResourceGuardAllowsRun } from './ci-local-resource-guard.ts'
import { ensureWasmKernelArtifact } from './ensure-wasm-kernel.js'

const DEFAULT_CI_FILE_CHUNK_SIZE = 3
const DEFAULT_CI_BATCH_COOLDOWN_MS = 1_000
const BROAD_CORPUS_FILE_THRESHOLD = 4

export function buildVitestArgs(args: readonly string[], env: NodeJS.ProcessEnv = process.env): string[] {
  if (!env['BILIG_CI_PROFILE'] || hasArg(args, '--maxWorkers')) {
    return [...args]
  }
  return [...args, '--maxWorkers', env['BILIG_VITEST_MAX_WORKERS'] ?? '1']
}

export function buildVitestArgBatches(args: readonly string[], env: NodeJS.ProcessEnv = process.env): string[][] {
  return splitVitestRunArgsForCi(args, env).map((batchArgs) => buildVitestArgs(batchArgs, env))
}

export function readVitestBatchCooldownMs(env: NodeJS.ProcessEnv = process.env): number {
  if (!env['BILIG_CI_PROFILE']) {
    return 0
  }
  return readNonNegativeInt(env['BILIG_VITEST_BATCH_COOLDOWN_MS']) ?? DEFAULT_CI_BATCH_COOLDOWN_MS
}

export function isBroadCorpusVitestRun(args: readonly string[]): boolean {
  const runIndex = args.indexOf('--run')
  if (runIndex < 0) {
    return false
  }

  const runFiles = args.slice(runIndex + 1).filter((arg) => !arg.startsWith('-'))
  const publicCorpusTestFiles = runFiles.filter((arg) => arg.includes('public-workbook-corpus'))
  if (publicCorpusTestFiles.length >= BROAD_CORPUS_FILE_THRESHOLD) {
    return true
  }

  const excelCorpusTestFiles = runFiles.filter((arg) => arg.includes('packages/excel-import/src/__tests__/xlsx-'))
  return publicCorpusTestFiles.length > 0 && excelCorpusTestFiles.length > 0
}

function splitVitestRunArgsForCi(args: readonly string[], env: NodeJS.ProcessEnv): string[][] {
  if (!env['BILIG_CI_PROFILE']) {
    return [[...args]]
  }

  const runIndex = args.indexOf('--run')
  if (runIndex < 0) {
    return [[...args]]
  }

  const prefixArgs = args.slice(0, runIndex + 1)
  const runArgs = args.slice(runIndex + 1)
  if (runArgs.length === 0 || runArgs.some((arg) => arg.startsWith('-'))) {
    return [[...args]]
  }

  const chunkSize =
    readPositiveInt(env['BILIG_VITEST_FILE_CHUNK_SIZE']) ?? (isBroadCorpusVitestRun(args) ? runArgs.length : DEFAULT_CI_FILE_CHUNK_SIZE)
  if (runArgs.length <= chunkSize) {
    return [[...args]]
  }

  const batches: string[][] = []
  for (let start = 0; start < runArgs.length; start += chunkSize) {
    batches.push([...prefixArgs, ...runArgs.slice(start, start + chunkSize)])
  }
  return batches
}

function hasArg(args: readonly string[], flag: string): boolean {
  return args.some((arg) => arg === flag || arg.startsWith(`${flag}=`))
}

function readPositiveInt(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined
  }

  const parsed = parseDecimalInt(value)
  return Number.isInteger(parsed) && parsed > 0 ? parsed : undefined
}

function readNonNegativeInt(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined
  }

  const parsed = parseDecimalInt(value)
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : undefined
}

function parseDecimalInt(value: string): number {
  if (!/^(?:0|[1-9]\d*)$/u.test(value)) {
    return Number.NaN
  }
  const parsed = Number(value)
  return Number.isSafeInteger(parsed) ? parsed : Number.NaN
}

function sleepSync(ms: number): void {
  if (ms <= 0) {
    return
  }
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms)
}

function main(): never {
  const rootDir = fileURLToPath(new URL('..', import.meta.url))
  const requestedArgs = process.argv.slice(2)
  if (isBroadCorpusVitestRun(requestedArgs)) {
    assertLocalCiResourceGuardAllowsRun(rootDir, process.env, { runLabel: 'public workbook corpus Vitest lane' })
  }

  const vitestBin = process.platform === 'win32' ? 'node_modules\\.bin\\vitest.cmd' : 'node_modules/.bin/vitest'
  ensureWasmKernelArtifact()
  const batches = buildVitestArgBatches(requestedArgs)
  const batchCooldownMs = readVitestBatchCooldownMs()
  for (const [index, args] of batches.entries()) {
    if (index > 0) {
      sleepSync(batchCooldownMs)
    }
    const result = spawnSync(vitestBin, args, {
      cwd: process.cwd(),
      env: process.env,
      stdio: 'inherit',
    })

    if (result.error) {
      throw result.error
    }

    if (result.signal) {
      process.stderr.write(`vitest terminated by signal ${result.signal}\n`)
    }

    if (result.status !== 0) {
      process.exit(result.status ?? 1)
    }
  }

  process.exit(0)
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  try {
    main()
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`)
    process.exit(1)
  }
}
