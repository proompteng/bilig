import { existsSync } from 'node:fs'
import { isAbsolute, relative, resolve } from 'node:path'

import { parseStrictBooleanEnvFlag } from './strict-env.js'

export function readStringArg(name: string, fallback: string): string {
  let value: string | null = null
  let count = 0
  process.argv.forEach((arg, index) => {
    const parsed = readArgValueForName(name, arg, index)
    if (parsed !== null) {
      count += 1
      assertArgSpecifiedOnce(name, count)
      value = parsed
    }
  })
  return value ?? fallback
}

function assertArgSpecifiedOnce(name: string, count: number): void {
  if (count > 1) {
    throw new Error(`Expected ${name} to be specified once`)
  }
}

function readArgValueForName(name: string, arg: string, index: number): string | null {
  if (arg === name) {
    return readArgValueAt(name, index)
  }
  const inlinePrefix = `${name}=`
  if (arg.startsWith(inlinePrefix)) {
    const value = arg.slice(inlinePrefix.length)
    if (isBlankArgValue(value)) {
      throw new Error(`Expected ${name} to have a value`)
    }
    return value
  }
  return null
}

function readArgValueAt(name: string, index: number): string {
  const next = process.argv[index + 1]
  if (next === undefined || isBlankArgValue(next) || next.startsWith('--')) {
    throw new Error(`Expected ${name} to have a value`)
  }
  return next
}

function isBlankArgValue(value: string): boolean {
  return value.trim().length === 0
}

export function readNumberArg(name: string, fallback: number): number {
  const raw = readStringArg(name, String(fallback))
  return parsePositiveIntegerArg(raw, `Expected ${name} to be a positive integer`)
}

const bytesPerMiB = 1024 * 1024
const maxSafeMiB = Math.floor(Number.MAX_SAFE_INTEGER / bytesPerMiB)

export function readMegabytesArg(name: string, fallbackBytes: number): number {
  const raw = readStringArg(name, String(Math.ceil(fallbackBytes / bytesPerMiB)))
  const parsed = parsePositiveIntegerArg(raw, `Expected ${name} to be a positive integer number of MiB`)
  if (parsed > maxSafeMiB) {
    throw new Error(`Expected ${name} to be a positive integer number of MiB`)
  }
  return parsed * bytesPerMiB
}

function parsePositiveIntegerArg(raw: string, errorMessage: string): number {
  const parsed = Number(raw)
  if (!/^\d+$/.test(raw) || parsed <= 0 || !Number.isSafeInteger(parsed)) {
    throw new Error(errorMessage)
  }
  return parsed
}

export function readFlagArg(name: string): boolean {
  let value = false
  let count = 0
  process.argv.forEach((arg, index) => {
    if (arg === name) {
      count += 1
      assertArgSpecifiedOnce(name, count)
      const next = process.argv[index + 1]
      value = next === undefined || next.startsWith('--') ? true : readBooleanArgValue(name, next)
      return
    }

    const inlinePrefix = `${name}=`
    if (arg.startsWith(inlinePrefix)) {
      count += 1
      assertArgSpecifiedOnce(name, count)
      const raw = arg.slice(inlinePrefix.length)
      value = readBooleanArgValue(name, raw)
    }
  })
  return value
}

function readBooleanArgValue(name: string, raw: string): boolean {
  if (raw === 'true') {
    return true
  }
  if (raw === 'false') {
    return false
  }
  throw new Error(`Expected ${name} to be true or false`)
}

export function readDebugOnlyFlagArg(name: string, envVar: string, reason: string): boolean {
  const enabled = readFlagArg(name)
  if (enabled && !parseStrictBooleanEnvFlag(process.env[envVar], envVar, false)) {
    throw new Error(`${name} is disabled for public corpus CLI runs because ${reason}. Set ${envVar}=1 only for focused debugging.`)
  }
  return enabled
}

export const publicCorpusStopMarkerOverrideFlag = '--allow-active-stop-marker'
export const publicCorpusStopMarkerOverrideEnvVar = 'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE'
const rootDir = resolve(new URL('..', import.meta.url).pathname)

export function assertPublicCorpusRunNotStopped(args: { readonly commandName: string; readonly stopMarkerPath: string }): void {
  if (!existsSync(args.stopMarkerPath)) {
    return
  }
  if (readFlagArg(publicCorpusStopMarkerOverrideFlag) && process.env[publicCorpusStopMarkerOverrideEnvVar] === '1') {
    return
  }
  throw new Error(
    `${args.commandName} is disabled while the public corpus stop marker is active: ${formatPublicCorpusStopMarkerPathForMessage(
      args.stopMarkerPath,
    )}. The marker protects the interactive host from broad workbook corpus runs. Resume only after the user explicitly asks, then pass ${publicCorpusStopMarkerOverrideFlag} with ${publicCorpusStopMarkerOverrideEnvVar}=1.`,
  )
}

export function formatPublicCorpusStopMarkerPathForMessage(path: string, displayRootDir = rootDir): string {
  const relativePath = relative(displayRootDir, path)
  if (!relativePath || relativePath.startsWith('..') || isAbsolute(relativePath)) {
    return path
  }
  return relativePath
}

export function readVerifyConcurrencyArg(defaultVerifyConcurrency: number): number {
  const verifyConcurrency = readNumberArg('--verify-concurrency', defaultVerifyConcurrency)
  if (
    verifyConcurrency > 1 &&
    !parseStrictBooleanEnvFlag(process.env['BILIG_ALLOW_PARALLEL_PUBLIC_CORPUS_VERIFY'], 'BILIG_ALLOW_PARALLEL_PUBLIC_CORPUS_VERIFY', false)
  ) {
    throw new Error(
      `--verify-concurrency greater than 1 is disabled for public corpus CLI runs because each worker can consume substantial memory. Set BILIG_ALLOW_PARALLEL_PUBLIC_CORPUS_VERIFY=1 only on a host sized for parallel verification.`,
    )
  }
  return verifyConcurrency
}

export function readFetchConcurrencyArg(defaultFetchConcurrency: number): number {
  const fetchConcurrency = readNumberArg('--fetch-concurrency', defaultFetchConcurrency)
  if (
    fetchConcurrency > 1 &&
    !parseStrictBooleanEnvFlag(process.env['BILIG_ALLOW_PARALLEL_PUBLIC_CORPUS_FETCH'], 'BILIG_ALLOW_PARALLEL_PUBLIC_CORPUS_FETCH', false)
  ) {
    throw new Error(
      `--fetch-concurrency greater than 1 is disabled for public corpus CLI runs because each fetch can spawn workbook fingerprinting workers and retain downloaded workbook bytes. Set BILIG_ALLOW_PARALLEL_PUBLIC_CORPUS_FETCH=1 only on a host sized for parallel fetch/fingerprint runs.`,
    )
  }
  return fetchConcurrency
}

const maxInteractiveVerificationBatchLimit = 20
const largeVerifyMissingLimitEnvVar = 'BILIG_ALLOW_LARGE_PUBLIC_CORPUS_VERIFY_MISSING'

export function readVerifyMissingLimitArg(defaultLimit: number, dryRun: boolean): number {
  return readPublicCorpusVerificationBatchLimitArg(defaultLimit, dryRun, {
    commandName: 'verify-missing',
    envVar: largeVerifyMissingLimitEnvVar,
  })
}

export function readPublicCorpusVerificationBatchLimitArg(
  defaultLimit: number,
  dryRun: boolean,
  args: { readonly commandName: string; readonly envVar: string },
): number {
  const limit = readNumberArg('--limit', defaultLimit)
  if (!dryRun && limit > maxInteractiveVerificationBatchLimit && process.env[args.envVar] !== '1') {
    throw new Error(
      `--limit above ${String(
        maxInteractiveVerificationBatchLimit,
      )} is disabled for public corpus ${args.commandName} runs because it can start many workbook verification workers. Set ${args.envVar}=1 only when intentionally resuming a large corpus verification tranche.`,
    )
  }
  return limit
}

export function readFetchRunArgs(defaults: { readonly batchSize: number; readonly concurrency: number }): {
  readonly fetchBatchSize: number
  readonly fetchConcurrency: number
  readonly inProcessFingerprinting: boolean
} {
  return {
    fetchBatchSize: readNumberArg('--fetch-batch-size', defaults.batchSize),
    fetchConcurrency: readFetchConcurrencyArg(defaults.concurrency),
    inProcessFingerprinting: readDebugOnlyFlagArg(
      '--in-process-fingerprint',
      'BILIG_ALLOW_IN_PROCESS_PUBLIC_CORPUS_FINGERPRINT',
      'it can retain workbook fingerprinting memory across large fetch runs',
    ),
  }
}

export function readRepeatedStringArg(name: string): string[] {
  const values: string[] = []
  process.argv.forEach((arg, index) => {
    const parsed = readArgValueForName(name, arg, index)
    if (parsed !== null) {
      values.push(parsed)
    }
  })
  return values
}
