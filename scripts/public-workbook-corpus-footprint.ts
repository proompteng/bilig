import { spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

import { asRecord } from './public-workbook-corpus-json.ts'
import { startChildRssWatchdog, terminateChildProcess } from './public-workbook-corpus-process.ts'
import type { WorkbookFootprint } from './public-workbook-corpus-workbook.ts'
import type { WorkbookExternalWorkbookReferenceSnapshot } from '../packages/protocol/src/types.js'

const noop = (): void => undefined
const tsxExecutablePath = fileURLToPath(new URL('../node_modules/.bin/tsx', import.meta.url))

export interface PublicWorkbookCorpusWorkerOptions {
  readonly timeoutMs: number
  readonly maxRssBytes: number
  readonly rssCheckIntervalMs: number
  readonly onPhase?: (phase: string) => void
}

export function inspectWorkbookFootprintIsolated(args: {
  readonly bytes: Uint8Array
  readonly filePath?: string
  readonly fileName: string
  readonly scriptPath: string
  readonly options: PublicWorkbookCorpusWorkerOptions
}): Promise<WorkbookFootprint | null> {
  return new Promise<WorkbookFootprint | null>((resolvePromise) => {
    const childArgs = [
      args.scriptPath,
      'footprint-worker',
      '--file-name',
      args.fileName,
      '--verify-max-rss-mb',
      String(Math.ceil(args.options.maxRssBytes / 1024 / 1024)),
      ...(args.filePath ? ['--file', args.filePath] : []),
    ]
    const child = spawn(existsSync(tsxExecutablePath) ? tsxExecutablePath : process.execPath, childArgs, {
      stdio: [args.filePath ? 'ignore' : 'pipe', 'pipe', 'pipe'],
    })
    let stdout = ''
    let stopRssWatchdog = noop
    let timer: ReturnType<typeof setTimeout>
    const finish = createOneShotResolver(resolvePromise, () => {
      clearTimeout(timer)
      stopRssWatchdog()
    })
    const terminateChild = (signal: 'SIGTERM' | 'SIGKILL'): void => {
      terminateChildProcess(child, signal)
    }
    stopRssWatchdog = startChildRssWatchdog(child, {
      maxRssBytes: args.options.maxRssBytes,
      intervalMs: args.options.rssCheckIntervalMs,
      onLimitExceeded: () => {
        terminateChild('SIGTERM')
        const forceKillTimer = setTimeout(() => terminateChild('SIGKILL'), 5_000)
        forceKillTimer.unref()
        finish(null)
      },
    })
    timer = setTimeout(() => {
      terminateChild('SIGTERM')
      const forceKillTimer = setTimeout(() => terminateChild('SIGKILL'), 5_000)
      forceKillTimer.unref()
      finish(null)
    }, args.options.timeoutMs)
    child.stdout.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => {
      stdout += chunk
    })
    if (!args.filePath && child.stdin) {
      child.stdin.on('error', () => undefined)
      child.stdin.end(Buffer.from(args.bytes))
    }
    child.on('error', () => {
      finish(null)
    })
    child.on('close', (code) => {
      if (code !== 0) {
        finish(null)
        return
      }
      try {
        const parsed: unknown = JSON.parse(stdout)
        finish(readFootprintWorkerResult(parsed))
      } catch {
        finish(null)
      }
    })
  })
}

export function readFootprintWorkerResult(value: unknown): WorkbookFootprint | null {
  const record = asRecord(value)
  const footprint = asRecord(record['footprint'])
  const featureCounts = asRecord(footprint['featureCounts'])
  const largeSimpleXlsxImport = readLargeSimpleXlsxImport(footprint['largeSimpleXlsxImport'])
  return {
    featureCounts: {
      sheetCount: readInteger(featureCounts, 'sheetCount'),
      cellCount: readInteger(featureCounts, 'cellCount'),
      formulaCellCount: readInteger(featureCounts, 'formulaCellCount'),
      valueCellCount: readInteger(featureCounts, 'valueCellCount'),
      definedNameCount: readInteger(featureCounts, 'definedNameCount'),
      tableCount: readInteger(featureCounts, 'tableCount'),
      chartCount: readInteger(featureCounts, 'chartCount'),
      pivotCount: readInteger(featureCounts, 'pivotCount'),
      mergeCount: readInteger(featureCounts, 'mergeCount'),
      styleRangeCount: readInteger(featureCounts, 'styleRangeCount'),
      conditionalFormatCount: readInteger(featureCounts, 'conditionalFormatCount'),
      dataValidationCount: readInteger(featureCounts, 'dataValidationCount'),
      macroPayloadCount: readInteger(featureCounts, 'macroPayloadCount'),
      warningCount: readInteger(featureCounts, 'warningCount'),
    },
    workbookMetadata: readWorkbookMetadata(asRecord(footprint['workbookMetadata'])),
    externalWorkbookReferences: readExternalWorkbookReferences(footprint['externalWorkbookReferences']),
    ...(largeSimpleXlsxImport ? { largeSimpleXlsxImport } : {}),
  }
}

function readLargeSimpleXlsxImport(value: unknown): WorkbookFootprint['largeSimpleXlsxImport'] | undefined {
  if (value === undefined) {
    return undefined
  }
  const record = asRecord(value)
  const eligible = record['eligible']
  const blockers = record['blockers']
  if (typeof eligible !== 'boolean' || !Array.isArray(blockers)) {
    throw new Error('Expected largeSimpleXlsxImport eligibility record')
  }
  return {
    eligible,
    blockers: blockers.map((entry) => {
      if (typeof entry !== 'string') {
        throw new Error('Expected string largeSimpleXlsxImport blocker')
      }
      return entry
    }),
  }
}

function readWorkbookMetadata(record: Record<string, unknown>): WorkbookFootprint['workbookMetadata'] {
  return {
    workbookName: readRequiredString(record, 'workbookName'),
    sheetNames: readRequiredStringArray(record, 'sheetNames'),
    dimensions: readRequiredArray(record, 'dimensions').map((entry) => {
      const dimension = asRecord(entry)
      const usedRange = readOptionalUsedRange(dimension['usedRange'])
      const parsedDimension: WorkbookFootprint['workbookMetadata']['dimensions'][number] = {
        sheetName: readRequiredString(dimension, 'sheetName'),
        rowCount: readInteger(dimension, 'rowCount'),
        columnCount: readInteger(dimension, 'columnCount'),
        nonEmptyCellCount: readInteger(dimension, 'nonEmptyCellCount'),
      }
      if (usedRange !== undefined) {
        Object.assign(parsedDimension, { usedRange })
      }
      return parsedDimension
    }),
  }
}

function readExternalWorkbookReferences(value: unknown): readonly WorkbookExternalWorkbookReferenceSnapshot[] {
  if (!Array.isArray(value)) {
    return []
  }
  return value.map((entry) => {
    const record = asRecord(entry)
    const parsed: WorkbookExternalWorkbookReferenceSnapshot = {
      bookIndex: readInteger(record, 'bookIndex'),
    }
    const packagePath = readOptionalString(record, 'packagePath')
    const target = readOptionalString(record, 'target')
    const targetMode = readOptionalString(record, 'targetMode')
    const workbookName = readOptionalString(record, 'workbookName')
    const sheetNames = readOptionalStringArray(record, 'sheetNames')
    if (packagePath) {
      Object.assign(parsed, { packagePath })
    }
    if (target) {
      Object.assign(parsed, { target })
    }
    if (targetMode) {
      Object.assign(parsed, { targetMode })
    }
    if (workbookName) {
      Object.assign(parsed, { workbookName })
    }
    if (sheetNames.length > 0) {
      Object.assign(parsed, { sheetNames })
    }
    return parsed
  })
}

function readOptionalUsedRange(value: unknown): WorkbookFootprint['workbookMetadata']['dimensions'][number]['usedRange'] | undefined {
  if (value === undefined) {
    return undefined
  }
  if (value === null) {
    return null
  }
  const record = asRecord(value)
  return {
    startRow: readInteger(record, 'startRow'),
    startColumn: readInteger(record, 'startColumn'),
    endRow: readInteger(record, 'endRow'),
    endColumn: readInteger(record, 'endColumn'),
  }
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

function readInteger(record: Record<string, unknown>, key: string): number {
  const value = record[key]
  if (!Number.isInteger(value) || value < 0) {
    throw new Error(`Expected non-negative integer ${key}`)
  }
  return value
}

function readRequiredString(record: Record<string, unknown>, key: string): string {
  const value = record[key]
  if (typeof value !== 'string' || value.length === 0) {
    throw new Error(`Expected non-empty string ${key}`)
  }
  return value
}

function readOptionalString(record: Record<string, unknown>, key: string): string | undefined {
  const value = record[key]
  return typeof value === 'string' && value.length > 0 ? value : undefined
}

function readRequiredStringArray(record: Record<string, unknown>, key: string): string[] {
  return readRequiredArray(record, key).map((entry) => {
    if (typeof entry !== 'string') {
      throw new Error(`Expected string entry in ${key}`)
    }
    return entry
  })
}

function readOptionalStringArray(record: Record<string, unknown>, key: string): string[] {
  const value = record[key]
  if (!Array.isArray(value)) {
    return []
  }
  return value.flatMap((entry) => (typeof entry === 'string' && entry.length > 0 ? [entry] : []))
}

function readRequiredArray(record: Record<string, unknown>, key: string): unknown[] {
  const value = record[key]
  if (!Array.isArray(value)) {
    throw new Error(`Expected array ${key}`)
  }
  return value
}
