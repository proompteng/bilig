import { existsSync } from 'node:fs'

export function readStringArg(name: string, fallback: string): string {
  const index = process.argv.indexOf(name)
  return index >= 0 ? (process.argv[index + 1] ?? fallback) : fallback
}

export function readNumberArg(name: string, fallback: number): number {
  const raw = readStringArg(name, String(fallback))
  const parsed = Number(raw)
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`Expected ${name} to be a positive number`)
  }
  return Math.trunc(parsed)
}

export function readMegabytesArg(name: string, fallbackBytes: number): number {
  const raw = readStringArg(name, String(Math.ceil(fallbackBytes / 1024 / 1024)))
  const parsed = Number(raw)
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`Expected ${name} to be a positive number of MiB`)
  }
  return Math.trunc(parsed * 1024 * 1024)
}

export function readFlagArg(name: string): boolean {
  return process.argv.includes(name)
}

export function readDebugOnlyFlagArg(name: string, envVar: string, reason: string): boolean {
  const enabled = readFlagArg(name)
  if (enabled && process.env[envVar] !== '1') {
    throw new Error(`${name} is disabled for public corpus CLI runs because ${reason}. Set ${envVar}=1 only for focused debugging.`)
  }
  return enabled
}

export const publicCorpusStopMarkerOverrideFlag = '--allow-active-stop-marker'
export const publicCorpusStopMarkerOverrideEnvVar = 'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE'

export function assertPublicCorpusRunNotStopped(args: { readonly commandName: string; readonly stopMarkerPath: string }): void {
  if (!existsSync(args.stopMarkerPath)) {
    return
  }
  if (readFlagArg(publicCorpusStopMarkerOverrideFlag) && process.env[publicCorpusStopMarkerOverrideEnvVar] === '1') {
    return
  }
  throw new Error(
    `${args.commandName} is disabled while the public corpus stop marker is active: ${args.stopMarkerPath}. The marker protects the interactive host from broad workbook corpus runs. Resume only after the user explicitly asks, then pass ${publicCorpusStopMarkerOverrideFlag} with ${publicCorpusStopMarkerOverrideEnvVar}=1.`,
  )
}

export function readVerifyConcurrencyArg(defaultVerifyConcurrency: number): number {
  const verifyConcurrency = readNumberArg('--verify-concurrency', defaultVerifyConcurrency)
  if (verifyConcurrency > 1 && process.env['BILIG_ALLOW_PARALLEL_PUBLIC_CORPUS_VERIFY'] !== '1') {
    throw new Error(
      `--verify-concurrency greater than 1 is disabled for public corpus CLI runs because each worker can consume substantial memory. Set BILIG_ALLOW_PARALLEL_PUBLIC_CORPUS_VERIFY=1 only on a host sized for parallel verification.`,
    )
  }
  return verifyConcurrency
}

export function readFetchConcurrencyArg(defaultFetchConcurrency: number): number {
  const fetchConcurrency = readNumberArg('--fetch-concurrency', defaultFetchConcurrency)
  if (fetchConcurrency > 1 && process.env['BILIG_ALLOW_PARALLEL_PUBLIC_CORPUS_FETCH'] !== '1') {
    throw new Error(
      `--fetch-concurrency greater than 1 is disabled for public corpus CLI runs because each fetch can spawn workbook fingerprinting workers and retain downloaded workbook bytes. Set BILIG_ALLOW_PARALLEL_PUBLIC_CORPUS_FETCH=1 only on a host sized for parallel fetch/fingerprint runs.`,
    )
  }
  return fetchConcurrency
}

const maxInteractiveVerifyMissingLimit = 20
const largeVerifyMissingLimitEnvVar = 'BILIG_ALLOW_LARGE_PUBLIC_CORPUS_VERIFY_MISSING'

export function readVerifyMissingLimitArg(defaultLimit: number, dryRun: boolean): number {
  const limit = readNumberArg('--limit', defaultLimit)
  if (!dryRun && limit > maxInteractiveVerifyMissingLimit && process.env[largeVerifyMissingLimitEnvVar] !== '1') {
    throw new Error(
      `--limit above ${String(
        maxInteractiveVerifyMissingLimit,
      )} is disabled for public corpus verify-missing runs because it can start many workbook verification workers. Set ${largeVerifyMissingLimitEnvVar}=1 only when intentionally resuming a large corpus verification tranche.`,
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
    if (arg === name && process.argv[index + 1]) {
      values.push(process.argv[index + 1])
    }
  })
  return values
}
