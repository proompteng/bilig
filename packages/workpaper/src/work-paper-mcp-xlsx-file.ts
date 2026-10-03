import { existsSync, mkdirSync, renameSync, statSync, writeFileSync } from 'node:fs'
import { basename, dirname, resolve } from 'node:path'

import { exportWorkPaperDocument, serializeWorkPaperDocument } from './persistence.js'
import { WorkPaper } from './work-paper.js'
import { createFileBackedWorkPaperMcpToolServer, createFileBackedWorkPaperMcpToolServerFromFile } from './work-paper-mcp-file-server.js'
import type { WorkPaperMcpToolServer } from './work-paper-mcp-server.js'
import { importXlsxFile } from './xlsx.js'

const workPaperMcpFromXlsxBytesLimit = 1_000_000

interface FileBackedWorkPaperMcpFromXlsxOptions {
  readonly fromXlsxPath: string
  readonly overwriteWorkPaper?: boolean
  readonly workpaperPath: string
  readonly writable?: boolean
}

function createFileBackedWorkPaperMcpToolServerFromXlsxFile(input: FileBackedWorkPaperMcpFromXlsxOptions): WorkPaperMcpToolServer {
  const workpaperPath = resolve(input.workpaperPath)
  if (existsSync(workpaperPath) && !input.overwriteWorkPaper) {
    throw new Error(`WorkPaper JSON already exists at ${workpaperPath}; pass --overwrite-workpaper to replace it`)
  }

  const xlsxPath = resolve(input.fromXlsxPath)
  assertWorkPaperMcpXlsxImportWithinSmallWorkbookLimit(xlsxPath)
  const imported = importXlsxFile(xlsxPath, basename(xlsxPath), { preferNativeSimpleImport: true })
  const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, { useColumnIndex: true })
  mkdirSync(dirname(workpaperPath), { recursive: true })
  writeFileAtomically(workpaperPath, serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true })))

  return createFileBackedWorkPaperMcpToolServerFromFile({
    workpaperPath,
    writable: input.writable ?? false,
  })
}

function createWorkPaperMcpToolServerFromXlsxFile(input: { readonly fromXlsxPath: string }): WorkPaperMcpToolServer {
  const xlsxPath = resolve(input.fromXlsxPath)
  assertWorkPaperMcpXlsxImportWithinSmallWorkbookLimit(xlsxPath)
  const imported = importXlsxFile(xlsxPath, basename(xlsxPath), { preferNativeSimpleImport: true })
  const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, { useColumnIndex: true })

  return createFileBackedWorkPaperMcpToolServer({
    workbook,
    sourcePath: xlsxPath,
    writable: false,
  })
}

function writeFileAtomically(path: string, contents: string): void {
  const tempPath = resolve(dirname(path), `.${basename(path)}.${process.pid.toString()}.tmp`)
  writeFileSync(tempPath, contents)
  renameSync(tempPath, path)
}

function assertWorkPaperMcpXlsxImportWithinSmallWorkbookLimit(xlsxPath: string): void {
  const byteLength = statSync(xlsxPath).size
  if (byteLength <= workPaperMcpFromXlsxBytesLimit) {
    return
  }
  throw new Error(
    `bilig-workpaper-mcp --from-xlsx is a small-workbook WorkPaper materialization path ` +
      `(${byteLength.toLocaleString('en-US')} bytes > ${workPaperMcpFromXlsxBytesLimit.toLocaleString('en-US')} bytes). ` +
      'Use @bilig/xlsx file-backed compatibility, cache inspection, or streaming-native recalc APIs for large XLSX files.',
  )
}

export {
  createFileBackedWorkPaperMcpToolServerFromXlsxFile,
  createWorkPaperMcpToolServerFromXlsxFile,
  type FileBackedWorkPaperMcpFromXlsxOptions,
}
