import { spawn } from 'node:child_process'

import { asRecord } from './public-workbook-corpus-json.ts'
import { startChildRssWatchdog, terminateChildProcess } from './public-workbook-corpus-process.ts'
import type { WorkbookFootprint } from './public-workbook-corpus-workbook.ts'

const noop = (): void => undefined

export interface PublicWorkbookCorpusWorkerOptions {
  readonly timeoutMs: number
  readonly maxRssBytes: number
  readonly rssCheckIntervalMs: number
  readonly onPhase?: (phase: string) => void
}

export function inspectWorkbookFootprintIsolated(args: {
  readonly bytes: Uint8Array
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
    ]
    const child = spawn(process.execPath, childArgs, {
      stdio: ['pipe', 'pipe', 'pipe'],
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
    child.stdin.on('error', () => undefined)
    child.stdin.end(Buffer.from(args.bytes))
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

function readRequiredStringArray(record: Record<string, unknown>, key: string): string[] {
  return readRequiredArray(record, key).map((entry) => {
    if (typeof entry !== 'string') {
      throw new Error(`Expected string entry in ${key}`)
    }
    return entry
  })
}

function readRequiredArray(record: Record<string, unknown>, key: string): unknown[] {
  const value = record[key]
  if (!Array.isArray(value)) {
    throw new Error(`Expected array ${key}`)
  }
  return value
}
