import type { WorkbookSnapshot } from '@bilig/protocol'
import {
  createFileImportedXlsxSourceReader as createFileImportedXlsxSourceReaderImpl,
  createTempFileImportedXlsxSourceReader as createTempFileImportedXlsxSourceReaderImpl,
  exportXlsx as exportXlsxImpl,
  exportXlsxToFile as exportXlsxToFileImpl,
  exportXlsxSourceLiteralPatches as exportXlsxSourceLiteralPatchesImpl,
  exportXlsxSourceLiteralPatchesToFile as exportXlsxSourceLiteralPatchesToFileImpl,
  exportXlsxSourceLiteralPatchesToFileAsync as exportXlsxSourceLiteralPatchesToFileAsyncImpl,
  importXlsx as importXlsxImpl,
  importXlsxFromZipByteSource as importXlsxFromZipByteSourceImpl,
  type ImportedWorkbook,
  type XlsxByteSourceImportOptions,
  type XlsxImportOptions,
} from '@bilig/excel-import'
import { writeFile } from 'node:fs/promises'
import { basename } from 'node:path'

export interface XlsxSourceLiteralPatch {
  readonly sheetName: string
  readonly address: string
  readonly value: string | number | boolean | null
  readonly preserveFormula?: boolean
}

export interface XlsxSourceLiteralPatchExportInput {
  readonly source: Uint8Array | XlsxSourceReader
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly sheetNames?: readonly string[]
  readonly workbookName?: string
}

export interface XlsxSourceLiteralPatchFileExportInput extends XlsxSourceLiteralPatchExportInput {
  readonly outputPath: string
}

export interface XlsxSourceLiteralPatchFileExportResult {
  readonly bytesWritten: number
}

export interface XlsxSourceReader {
  readonly byteLength: number
  readBytes(): Uint8Array
  readRange?(start: number, end: number): Uint8Array
  readRangeInto?(start: number, end: number, target: Uint8Array): Uint8Array
  release?(): void
}

export interface XlsxZipSourceReader extends XlsxSourceReader {
  readRange(start: number, end: number): Uint8Array
}

export interface WorkPaperXlsxExportSource {
  exportSnapshot(): WorkbookSnapshot
  exportSourcePreservingXlsxSnapshot?(): WorkbookSnapshot | null
}

const importedXlsxSourceBytes = Symbol.for('bilig.importedXlsxSourceBytes')
const importedXlsxSourceCellPatches = Symbol.for('bilig.importedXlsxSourceCellPatches')

type SnapshotWithImportedXlsxSource = WorkbookSnapshot & {
  readonly [importedXlsxSourceBytes]?: Uint8Array | XlsxSourceReader
  readonly [importedXlsxSourceCellPatches]?: readonly unknown[]
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isXlsxSourceReader(value: unknown): value is XlsxSourceReader {
  return isRecord(value) && typeof value['byteLength'] === 'number' && typeof value['readBytes'] === 'function'
}

function isXlsxSourceReference(value: unknown): value is Uint8Array | XlsxSourceReader {
  return value instanceof Uint8Array || isXlsxSourceReader(value)
}

function isScalarXlsxPatchValue(value: unknown): value is XlsxSourceLiteralPatch['value'] {
  return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean'
}

function isXlsxSourceLiteralPatch(value: unknown): value is XlsxSourceLiteralPatch {
  return (
    isRecord(value) &&
    (value['kind'] === undefined || value['kind'] === 'literal') &&
    typeof value['sheetName'] === 'string' &&
    typeof value['address'] === 'string' &&
    isScalarXlsxPatchValue(value['value'])
  )
}

function sourcePreservingPatchInputFromSnapshot(snapshot: WorkbookSnapshot): XlsxSourceLiteralPatchExportInput | null {
  const sourceSnapshot = snapshot as SnapshotWithImportedXlsxSource
  const source = sourceSnapshot[importedXlsxSourceBytes]
  const patches = sourceSnapshot[importedXlsxSourceCellPatches]
  if (!isXlsxSourceReference(source) || patches === undefined || patches.length === 0) {
    return null
  }
  const literalPatches = patches.filter(isXlsxSourceLiteralPatch).map((patch): XlsxSourceLiteralPatch => {
    if (patch.preserveFormula === true) {
      return {
        sheetName: patch.sheetName,
        address: patch.address,
        value: patch.value,
        preserveFormula: true,
      }
    }
    return {
      sheetName: patch.sheetName,
      address: patch.address,
      value: patch.value,
    }
  })
  return literalPatches.length > 0
    ? {
        source,
        patches: literalPatches,
        sheetNames: snapshot.sheets.map((sheet) => sheet.name),
        workbookName: snapshot.workbook.name,
      }
    : null
}

function isWorkbookSnapshot(value: unknown): value is WorkbookSnapshot {
  return isRecord(value) && value['version'] === 1 && isRecord(value['workbook']) && Array.isArray(value['sheets'])
}

export function exportXlsx(snapshot: WorkbookSnapshot): Uint8Array
export function exportXlsx(workbook: WorkPaperXlsxExportSource): Uint8Array
export function exportXlsx(input: WorkbookSnapshot | WorkPaperXlsxExportSource): Uint8Array {
  if (!isWorkbookSnapshot(input)) {
    return exportWorkPaperXlsx(input)
  }
  return exportWorkbookSnapshotXlsx(input)
}

function exportWorkbookSnapshotXlsx(snapshot: WorkbookSnapshot): Uint8Array {
  return exportXlsxImpl(snapshot)
}

export function exportXlsxToFile(snapshot: WorkbookSnapshot, outputPath: string): XlsxSourceLiteralPatchFileExportResult {
  return exportXlsxToFileImpl(snapshot, outputPath)
}

export function createFileImportedXlsxSourceReader(path: string, byteLength?: number): XlsxZipSourceReader {
  return createFileImportedXlsxSourceReaderImpl(path, byteLength)
}

export function createTempFileImportedXlsxSourceReader(bytes: Uint8Array): XlsxZipSourceReader {
  return createTempFileImportedXlsxSourceReaderImpl(bytes)
}

export function exportXlsxSourceLiteralPatches(input: XlsxSourceLiteralPatchExportInput): Uint8Array {
  return exportXlsxSourceLiteralPatchesImpl(input)
}

export function exportXlsxSourceLiteralPatchesToFile(input: XlsxSourceLiteralPatchFileExportInput): XlsxSourceLiteralPatchFileExportResult {
  return exportXlsxSourceLiteralPatchesToFileImpl(input)
}

export function exportXlsxSourceLiteralPatchesToFileAsync(
  input: XlsxSourceLiteralPatchFileExportInput,
): Promise<XlsxSourceLiteralPatchFileExportResult> {
  return exportXlsxSourceLiteralPatchesToFileAsyncImpl(input)
}

export function exportWorkPaperXlsx(workbook: WorkPaperXlsxExportSource): Uint8Array {
  const sourcePreservingSnapshot = workbook.exportSourcePreservingXlsxSnapshot?.()
  const sourcePreservingInput = sourcePreservingSnapshot ? sourcePreservingPatchInputFromSnapshot(sourcePreservingSnapshot) : null
  if (sourcePreservingInput) {
    return exportXlsxSourceLiteralPatches(sourcePreservingInput)
  }
  return exportWorkbookSnapshotXlsx(sourcePreservingSnapshot ?? workbook.exportSnapshot())
}

export async function exportWorkPaperXlsxToFileAsync(
  workbook: WorkPaperXlsxExportSource,
  outputPath: string,
): Promise<XlsxSourceLiteralPatchFileExportResult> {
  const sourcePreservingSnapshot = workbook.exportSourcePreservingXlsxSnapshot?.()
  const sourcePreservingInput = sourcePreservingSnapshot ? sourcePreservingPatchInputFromSnapshot(sourcePreservingSnapshot) : null
  if (sourcePreservingInput) {
    return exportXlsxSourceLiteralPatchesToFileAsync({
      ...sourcePreservingInput,
      outputPath,
    })
  }

  const exported = exportWorkbookSnapshotXlsx(workbook.exportSnapshot())
  await writeFile(outputPath, exported)
  return { bytesWritten: exported.byteLength }
}

export function importXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string, options?: XlsxImportOptions): ImportedWorkbook {
  return importXlsxImpl(bytes, fileName, options)
}

export function importXlsxFile(path: string, fileName = basename(path), options?: XlsxByteSourceImportOptions): ImportedWorkbook {
  const source = createFileImportedXlsxSourceReader(path)
  try {
    return importXlsxFromZipByteSourceImpl(source, fileName, options)
  } catch (error) {
    source.release?.()
    throw error
  }
}

export type { XlsxImportOptions } from '@bilig/xlsx'
