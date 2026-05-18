import { spawn } from 'node:child_process'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { parsePublicWorkbookCorpusCase } from './public-workbook-corpus-json.ts'
import { compactRepoLocalPaths } from './public-workbook-corpus-output.ts'
import { formatByteSize, startChildRssWatchdog, terminateChildProcess } from './public-workbook-corpus-process.ts'
import { unsupportedRssLimitCase } from './public-workbook-corpus-resource-limits.ts'
import {
  startVerificationRuntimeMetrics,
  withPeakRssBytes,
  withVerificationRuntimeMetrics,
} from './public-workbook-corpus-verification-metrics.ts'
import { artifactBaseEvidence, failedCase } from './public-workbook-corpus-verify-cases.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase } from './public-workbook-corpus-types.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const publicWorkbookCorpusVerifyWorkerScriptPath = fileURLToPath(new URL('./public-workbook-corpus-verify-worker.ts', import.meta.url))
const noop = (): void => undefined

export const verificationWorkerPhasePrefix = 'bilig-public-workbook-verify-phase='

export function verifyCachedWorkbookArtifactIsolated(args: {
  readonly artifact: PublicWorkbookArtifact
  readonly cacheDir: string
  readonly manifestPath: string
  readonly runStructuralSmoke: boolean
  readonly timeoutMs: number
  readonly maxRssBytes: number
  readonly maxCellCount: number
  readonly rssCheckIntervalMs?: number
}): Promise<PublicWorkbookCorpusCase> {
  const baseEvidence = artifactBaseEvidence(args.artifact)
  const runtimeMetrics = startVerificationRuntimeMetrics()
  return new Promise<PublicWorkbookCorpusCase>((resolvePromise) => {
    const childArgs = [
      publicWorkbookCorpusVerifyWorkerScriptPath,
      'verify-artifact-worker',
      '--manifest',
      args.manifestPath,
      '--cache-dir',
      args.cacheDir,
      '--artifact-id',
      args.artifact.id,
      '--artifact-json-base64',
      Buffer.from(JSON.stringify(args.artifact), 'utf8').toString('base64'),
      '--verify-max-rss-mb',
      String(Math.ceil(args.maxRssBytes / 1024 / 1024)),
      '--verify-max-cells',
      String(args.maxCellCount),
      ...(args.runStructuralSmoke ? ['--structural-smoke'] : []),
    ]
    const child = spawn(process.execPath, childArgs, {
      detached: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    let stdout = ''
    let stderr = ''
    let stderrRemainder = ''
    let latestWorkerPhase = 'startup'
    let peakRssBytes = 0
    let timer: ReturnType<typeof setTimeout>
    let stopRssWatchdog = noop
    const finish = createOneShotResolver(resolvePromise, () => {
      clearTimeout(timer)
      stopRssWatchdog()
    })
    const terminateChild = (signal: 'SIGTERM' | 'SIGKILL'): void => {
      terminateChildProcess(child, signal, { processGroup: true })
    }
    stopRssWatchdog = startChildRssWatchdog(child, {
      maxRssBytes: args.maxRssBytes,
      intervalMs: args.rssCheckIntervalMs,
      onSample: (rssBytes) => {
        peakRssBytes = Math.max(peakRssBytes, rssBytes)
      },
      onLimitExceeded: (rssBytes) => {
        terminateChild('SIGTERM')
        const forceKillTimer = setTimeout(() => terminateChild('SIGKILL'), 5_000)
        forceKillTimer.unref()
        finish(
          withVerificationRuntimeMetrics(
            unsupportedRssLimitCase(args.artifact, baseEvidence, rssBytes, args.maxRssBytes, [
              `rss-limit-phase=${latestWorkerPhase}`,
              `peak-rss=${formatByteSize(Math.max(peakRssBytes, rssBytes))}`,
              'The workbook was isolated in a subprocess so the corpus verification run could continue.',
            ]),
            runtimeMetrics,
            Math.max(peakRssBytes, rssBytes),
          ),
        )
      },
    })
    timer = setTimeout(() => {
      terminateChild('SIGTERM')
      const forceKillTimer = setTimeout(() => terminateChild('SIGKILL'), 5_000)
      forceKillTimer.unref()
      finish(
        withVerificationRuntimeMetrics(
          failedCase(args.artifact, 'error', baseEvidence, [
            `Verification timed out after ${String(args.timeoutMs)}ms`,
            'The workbook was isolated in a subprocess so the corpus verification run could continue.',
          ]),
          runtimeMetrics,
          peakRssBytes,
        ),
      )
    }, args.timeoutMs)
    child.stdout.setEncoding('utf8')
    child.stderr.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => {
      stdout += chunk
    })
    child.stderr.on('data', (chunk: string) => {
      stderr += chunk
      const lines = `${stderrRemainder}${chunk}`.split(/\r?\n/u)
      stderrRemainder = lines.pop() ?? ''
      for (const line of lines) {
        if (line.startsWith(verificationWorkerPhasePrefix)) {
          latestWorkerPhase = line.slice(verificationWorkerPhasePrefix.length)
        }
      }
    })
    child.on('error', (error) => {
      finish(
        withVerificationRuntimeMetrics(
          failedCase(args.artifact, 'error', baseEvidence, [`Verification subprocess failed to start: ${error.message}`]),
          runtimeMetrics,
          peakRssBytes,
        ),
      )
    })
    child.on('close', (code, signal) => {
      if (code !== 0) {
        const failureDetails = compactVerificationWorkerOutput(stderr || stdout)
        finish(
          withVerificationRuntimeMetrics(
            failedCase(args.artifact, 'error', baseEvidence, [
              `Verification subprocess exited with ${signal ? `signal ${signal}` : `code ${String(code)}`}`,
              ...(failureDetails ? [failureDetails] : []),
            ]),
            runtimeMetrics,
            peakRssBytes,
          ),
        )
        return
      }
      try {
        const parsed: unknown = JSON.parse(stdout)
        finish(withPeakRssBytes(parsePublicWorkbookCorpusCase(parsed), peakRssBytes))
      } catch (error) {
        const details = compactVerificationWorkerOutput(stderr || stdout)
        finish(
          withVerificationRuntimeMetrics(
            failedCase(args.artifact, 'error', baseEvidence, [
              `Verification subprocess returned invalid JSON: ${error instanceof Error ? error.message : String(error)}`,
              ...(details ? [details] : []),
            ]),
            runtimeMetrics,
            peakRssBytes,
          ),
        )
      }
    })
  })
}

export function compactVerificationWorkerOutput(value: string): string | null {
  const withoutPhaseMarkers = value
    .split(/\r?\n/u)
    .filter((line) => !line.startsWith(verificationWorkerPhasePrefix))
    .join('\n')
  const compacted = compactRepoLocalPaths(withoutPhaseMarkers, rootDir).replace(/\s+/gu, ' ').trim()
  return compacted.length > 0 ? compacted.slice(0, 1_000) : null
}

function createOneShotResolver<T>(resolveValue: (value: T) => void, cleanup: () => void): (value: T) => void {
  let settled = false
  return (value) => {
    if (settled) {
      return
    }
    settled = true
    cleanup()
    resolveValue(value)
  }
}
