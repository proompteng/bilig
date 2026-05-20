import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

import { tryInspectLargeSimpleXlsxHeadless } from '../packages/excel-import/src/xlsx-large-simple-headless-inspect.js'
import type { LargeSimpleXlsxImportStats } from '../packages/excel-import/src/xlsx-large-simple-import.js'
import { readXlsxZipEntriesLazyFromByteSource } from '../packages/excel-import/src/xlsx-zip.js'
import { startSelfRssGuard } from './public-workbook-corpus-process.ts'
import {
  fingerprintLargeSimpleDataOnlyWorkbookSource,
  fingerprintFormulaFreeWorkbookFootprint,
  fingerprintWorkbookBytes,
  inspectWorkbookFootprintForWorker,
  type WorkbookFootprint,
} from './public-workbook-corpus-workbook.ts'
import { FileBackedXlsxZipByteSource, isZipWorkbookSource } from './public-workbook-corpus-xlsx-byte-source.ts'
import { inspectXlsxWorkbookFootprintLowMemoryFromByteSource } from './public-workbook-corpus-xlsx-footprint.ts'

export async function writeFingerprintArtifactResult(args: {
  readonly filePath: string
  readonly fileName: string
  readonly fingerprintTimeoutMs: number
  readonly fingerprintMaxRssBytes: number
}): Promise<void> {
  if (!args.filePath) {
    throw new Error('Expected --file for fingerprint-artifact')
  }
  const { fingerprintWorkbookFileIsolated } = await import('./public-workbook-corpus-fetch.ts')
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
    const filePath = resolve(args.filePath)
    const workbookFingerprint =
      tryFingerprintWorkbookFromFile(filePath, args.fileName) ?? fingerprintWorkbookBytes(readFileSync(filePath), args.fileName)
    process.stdout.write(`${JSON.stringify({ workbookFingerprint })}\n`)
  } catch (error) {
    process.stderr.write(`${formatWorkerError(error)}\n`)
    process.exitCode = 1
  } finally {
    stopSelfRssGuard()
  }
}

function tryFingerprintWorkbookFromFile(filePath: string, fileName: string): string | null {
  return tryFingerprintLargeSimpleWorkbookFromFile(filePath, fileName) ?? tryFingerprintFormulaFreeWorkbookFromFile(filePath, fileName)
}

function tryFingerprintLargeSimpleWorkbookFromFile(filePath: string, fileName: string): string | null {
  const source = new FileBackedXlsxZipByteSource(filePath)
  try {
    return isZipWorkbookSource(source) ? fingerprintLargeSimpleDataOnlyWorkbookSource(source, fileName) : null
  } catch {
    return null
  } finally {
    source.release()
  }
}

function tryFingerprintFormulaFreeWorkbookFromFile(filePath: string, fileName: string): string | null {
  const source = new FileBackedXlsxZipByteSource(filePath)
  try {
    if (!isZipWorkbookSource(source)) {
      return null
    }
    const footprint = inspectXlsxWorkbookFootprintLowMemoryFromByteSource(source, fileName)
    return footprint ? fingerprintFormulaFreeWorkbookFootprint(footprint) : null
  } catch {
    return null
  } finally {
    source.release()
  }
}

function formatWorkerError(error: unknown): string {
  if (error instanceof Error) {
    return `${error.name}: ${error.message}`
  }
  return String(error)
}

export async function writeFootprintWorkerResult(args: {
  readonly filePath: string
  readonly fileName: string
  readonly verifyMaxRssBytes: number
}): Promise<void> {
  const stopSelfRssGuard = startSelfRssGuard(args.verifyMaxRssBytes, 'Workbook footprint worker')
  try {
    const footprintFromFile = args.filePath ? tryInspectLargeSimpleWorkbookFootprintFromFile(resolve(args.filePath), args.fileName) : null
    if (footprintFromFile) {
      process.stdout.write(`${JSON.stringify({ footprint: footprintFromFile })}\n`)
      return
    }
    const bytes = args.filePath ? readFileSync(resolve(args.filePath)) : readFileSync(0)
    const footprint = await inspectWorkbookFootprintForWorker(bytes, args.fileName)
    process.stdout.write(`${JSON.stringify({ footprint })}\n`)
  } finally {
    stopSelfRssGuard()
  }
}

function tryInspectLargeSimpleWorkbookFootprintFromFile(filePath: string, fileName: string): WorkbookFootprint | null {
  const source = new FileBackedXlsxZipByteSource(filePath)
  try {
    if (!isZipWorkbookSource(source)) {
      return null
    }
    const sourceFootprint = inspectXlsxWorkbookFootprintLowMemoryFromByteSource(source, fileName)
    if (sourceFootprint) {
      return sourceFootprint
    }
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    if (!zip) {
      return null
    }
    const inspected = tryInspectLargeSimpleXlsxHeadless({ byteLength: source.byteLength }, fileName, zip, {
      minByteLength: 0,
      releaseZipSource: true,
    })
    return inspected ? footprintFromLargeSimpleInspect(inspected) : null
  } finally {
    source.release()
  }
}

function footprintFromLargeSimpleInspect(inspected: NonNullable<ReturnType<typeof tryInspectLargeSimpleXlsxHeadless>>): WorkbookFootprint {
  const featureCounts = featureCountsFromLargeSimpleStats(inspected.stats)
  return {
    featureCounts,
    workbookMetadata: {
      workbookName: inspected.workbookName,
      sheetNames: inspected.sheetNames,
      dimensions: inspected.stats.dimensions,
    },
    externalWorkbookReferences: [],
    largeSimpleXlsxImport: { eligible: true, blockers: [] },
  }
}

function featureCountsFromLargeSimpleStats(stats: LargeSimpleXlsxImportStats): WorkbookFootprint['featureCounts'] {
  return {
    sheetCount: stats.sheetCount,
    cellCount: stats.cellCount,
    formulaCellCount: stats.formulaCellCount,
    valueCellCount: stats.valueCellCount,
    definedNameCount: stats.definedNameCount,
    tableCount: stats.tableCount,
    chartCount: 0,
    pivotCount: 0,
    mergeCount: stats.mergeCount,
    styleRangeCount: 0,
    conditionalFormatCount: stats.conditionalFormatCount,
    dataValidationCount: stats.dataValidationCount ?? 0,
    macroPayloadCount: 0,
    warningCount: stats.warningCount,
  }
}
