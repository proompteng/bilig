import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

import { fingerprintWorkbookFileIsolated } from './public-workbook-corpus-fetch.ts'
import { startSelfRssGuard } from './public-workbook-corpus-process.ts'
import { fingerprintWorkbookBytes, inspectWorkbookFootprint } from './public-workbook-corpus-workbook.ts'

export async function writeFingerprintArtifactResult(args: {
  readonly filePath: string
  readonly fileName: string
  readonly fingerprintTimeoutMs: number
  readonly fingerprintMaxRssBytes: number
}): Promise<void> {
  if (!args.filePath) {
    throw new Error('Expected --file for fingerprint-artifact')
  }
  const workbookFingerprint = await fingerprintWorkbookFileIsolated(resolve(args.filePath), args.fileName, args.fingerprintTimeoutMs, {
    maxRssBytes: args.fingerprintMaxRssBytes,
    rssCheckIntervalMs: 250,
  })
  process.stdout.write(`${JSON.stringify({ workbookFingerprint })}\n`)
}

export function writeFingerprintArtifactWorkerResult(args: {
  readonly filePath: string
  readonly fileName: string
  readonly fingerprintMaxRssBytes: number
}): void {
  const stopSelfRssGuard = startSelfRssGuard(args.fingerprintMaxRssBytes, 'Workbook fingerprinting worker')
  try {
    if (!args.filePath) {
      throw new Error('Expected --file for fingerprint-artifact-worker')
    }
    const workbookFingerprint = fingerprintWorkbookBytes(readFileSync(resolve(args.filePath)), args.fileName)
    process.stdout.write(`${JSON.stringify({ workbookFingerprint })}\n`)
  } catch (error) {
    process.stderr.write(`${formatWorkerError(error)}\n`)
    process.exitCode = 1
  } finally {
    stopSelfRssGuard()
  }
}

function formatWorkerError(error: unknown): string {
  if (error instanceof Error) {
    return `${error.name}: ${error.message}`
  }
  return String(error)
}

export function writeFootprintWorkerResult(args: { readonly fileName: string; readonly verifyMaxRssBytes: number }): void {
  const stopSelfRssGuard = startSelfRssGuard(args.verifyMaxRssBytes, 'Workbook footprint worker')
  try {
    const bytes = readFileSync(0)
    const footprint = inspectWorkbookFootprint(bytes, args.fileName)
    process.stdout.write(`${JSON.stringify({ footprint })}\n`)
  } finally {
    stopSelfRssGuard()
  }
}
