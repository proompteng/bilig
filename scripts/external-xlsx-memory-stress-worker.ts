#!/usr/bin/env bun

import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { basename, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { importXlsxFromZipByteSource } from '../packages/excel-import/src/xlsx-byte-source-import.js'
import { tryInspectLargeSimpleXlsxHeadless } from '../packages/excel-import/src/xlsx-large-simple-headless-inspect.js'
import { readXlsxZipEntriesLazyFromByteSource } from '../packages/excel-import/src/xlsx-zip.js'
import type { ImportedWorkbook } from '../packages/excel-import/src/workbook-import-result.js'
import { readFlagArg, readStringArg } from './public-workbook-corpus-cli.ts'
import { FileBackedXlsxZipByteSource } from './public-workbook-corpus-xlsx-byte-source.ts'

interface PublicXlsxImportModule {
  readonly importXlsx: (bytes: Uint8Array, fileName: string) => ImportedWorkbook
}

export interface ExternalXlsxStressWorkerSummary {
  readonly importMode: 'headless-inspect' | 'public-snapshot'
  readonly sheets: number
  readonly cells: number
  readonly formulas: number
  readonly warnings: number
  readonly workbookMetadataKeys: readonly string[]
  readonly sheetMetadataKeys: readonly string[]
}

const requireModule = createRequire(import.meta.url)
const headlessInspectCellThreshold = 1_000_000

export function runExternalXlsxStressWorker(): void {
  const filePath = resolve(readStringArg('--file', ''))
  const fileName = readStringArg('--file-name', basename(filePath))
  const usePublicImport = readFlagArg('--public-import')
  const summary = usePublicImport
    ? summarizeExternalXlsxImportedWorkbook(importPublicXlsx(readFileSync(filePath), fileName))
    : summarizeFileBackedXlsx(filePath, fileName)
  collectGarbage()
  process.stdout.write(`${JSON.stringify(summary)}\n`)
}

export function summarizeExternalXlsxImportedWorkbook(imported: ImportedWorkbook): ExternalXlsxStressWorkerSummary {
  let cells = imported.stats?.cellCount ?? 0
  let formulas = imported.stats?.formulaCellCount ?? 0
  const sheetMetadataKeys = new Set<string>()
  for (const sheet of imported.snapshot.sheets) {
    if (!imported.stats) {
      cells += sheet.cells.length
      for (const cell of sheet.cells) {
        if ('formula' in cell) {
          formulas += 1
        }
      }
    }
    for (const key of Object.keys(sheet.metadata ?? {})) {
      sheetMetadataKeys.add(key)
    }
  }
  return {
    importMode: 'public-snapshot',
    sheets: imported.snapshot.sheets.length,
    cells,
    formulas,
    warnings: imported.warnings.length,
    workbookMetadataKeys: Object.keys(imported.snapshot.workbook.metadata ?? {}).toSorted(),
    sheetMetadataKeys: [...sheetMetadataKeys].toSorted(),
  }
}

function summarizeFileBackedXlsx(filePath: string, fileName: string): ExternalXlsxStressWorkerSummary {
  const headless = inspectFileBackedXlsxHeadless(filePath, fileName)
  if (headless && headless.stats.cellCount >= headlessInspectCellThreshold) {
    return {
      importMode: 'headless-inspect',
      sheets: headless.sheetNames.length,
      cells: headless.stats.cellCount,
      formulas: headless.stats.formulaCellCount,
      warnings: headless.warnings.length,
      workbookMetadataKeys: [],
      sheetMetadataKeys: [],
    }
  }
  return summarizeExternalXlsxImportedWorkbook(importFileBackedXlsx(filePath, fileName))
}

function inspectFileBackedXlsxHeadless(filePath: string, fileName: string): ReturnType<typeof tryInspectLargeSimpleXlsxHeadless> {
  const source = new FileBackedXlsxZipByteSource(filePath)
  try {
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    return zip
      ? tryInspectLargeSimpleXlsxHeadless({ byteLength: source.byteLength }, fileName, zip, {
          allowUnsupportedWorksheetFeaturesForMetrics: true,
          minByteLength: 0,
          releaseZipSource: true,
        })
      : null
  } finally {
    source.release()
  }
}

function importPublicXlsx(bytes: Uint8Array, fileName: string): ImportedWorkbook {
  const importer = readPublicXlsxImportModule(requireModule('../packages/excel-import/src/index.js'))
  return importer.importXlsx(bytes, fileName)
}

function importFileBackedXlsx(filePath: string, fileName: string): ImportedWorkbook {
  const source = new FileBackedXlsxZipByteSource(filePath)
  try {
    return importXlsxFromZipByteSource(source, fileName, { attachSourceReaderForUntouchedExport: false })
  } finally {
    source.release()
  }
}

function readPublicXlsxImportModule(value: unknown): PublicXlsxImportModule {
  if (isRecord(value) && typeof value['importXlsx'] === 'function') {
    return { importXlsx: value['importXlsx'] }
  }
  throw new Error('Public XLSX importer module is missing required exports')
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function collectGarbage(): void {
  if (typeof Bun !== 'undefined' && typeof Bun.gc === 'function') {
    Bun.gc(true)
    return
  }
  const gc = Reflect.get(globalThis, 'gc')
  if (typeof gc === 'function') {
    gc()
  }
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  try {
    runExternalXlsxStressWorker()
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`)
    process.exit(1)
  }
}
