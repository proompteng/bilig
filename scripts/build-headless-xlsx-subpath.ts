#!/usr/bin/env bun

import { cpSync, existsSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const excelImportDistDir = join(rootDir, 'packages', 'excel-import', 'dist')
const headlessDistDir = join(rootDir, 'packages', 'workpaper', 'dist')
const bundledXlsxDistDir = join(headlessDistDir, 'xlsx-internal')

if (!existsSync(join(excelImportDistDir, 'index.js')) || !existsSync(join(excelImportDistDir, 'index.d.ts'))) {
  throw new Error('Build @bilig/excel-import before building the @bilig/workpaper XLSX subpath')
}

mkdirSync(headlessDistDir, { recursive: true })
rmSync(bundledXlsxDistDir, { recursive: true, force: true })
cpSync(excelImportDistDir, bundledXlsxDistDir, { recursive: true })

writeFileSync(
  join(headlessDistDir, 'xlsx-runtime.js'),
  `import { writeFile } from 'node:fs/promises'
import { basename } from 'node:path'
import {
  createFileImportedXlsxSourceReader,
  exportXlsx as exportWorkbookSnapshotXlsx,
  exportXlsxSourceLiteralPatches,
  exportXlsxSourceLiteralPatchesToFileAsync,
  importXlsxFromZipByteSource,
} from './xlsx-internal/index.js'

export * from './xlsx-internal/index.js'

const importedXlsxSourceBytes = Symbol.for('bilig.importedXlsxSourceBytes')
const importedXlsxSourceCellPatches = Symbol.for('bilig.importedXlsxSourceCellPatches')

function isSourceReference(value) {
  return value instanceof Uint8Array || (value && typeof value.byteLength === 'number' && typeof value.readBytes === 'function')
}

function sourcePreservingPatchInputFromSnapshot(snapshot) {
  const source = snapshot[importedXlsxSourceBytes]
  const patches = snapshot[importedXlsxSourceCellPatches]
  if (!isSourceReference(source) || !Array.isArray(patches) || patches.length === 0) {
    return null
  }
  const literalPatches = patches
    .filter((patch) => {
      const value = patch?.value
      return (
        patch &&
        (patch.kind === undefined || patch.kind === 'literal') &&
        typeof patch.sheetName === 'string' &&
        typeof patch.address === 'string' &&
        (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean')
      )
    })
    .map((patch) => {
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

function isWorkbookSnapshot(value) {
  return value && value.version === 1 && typeof value.workbook === 'object' && Array.isArray(value.sheets)
}

export function exportXlsx(input) {
  return isWorkbookSnapshot(input) ? exportWorkbookSnapshotXlsx(input) : exportWorkPaperXlsx(input)
}

export function exportWorkPaperXlsx(workbook) {
  const snapshot =
    typeof workbook.exportSourcePreservingXlsxSnapshot === 'function'
      ? workbook.exportSourcePreservingXlsxSnapshot()
      : null
  const sourcePreservingInput = snapshot ? sourcePreservingPatchInputFromSnapshot(snapshot) : null
  return sourcePreservingInput
    ? exportXlsxSourceLiteralPatches(sourcePreservingInput)
    : exportWorkbookSnapshotXlsx(snapshot ?? workbook.exportSnapshot())
}

export async function exportWorkPaperXlsxToFileAsync(workbook, outputPath) {
  const snapshot =
    typeof workbook.exportSourcePreservingXlsxSnapshot === 'function'
      ? workbook.exportSourcePreservingXlsxSnapshot()
      : null
  const sourcePreservingInput = snapshot ? sourcePreservingPatchInputFromSnapshot(snapshot) : null
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

export function importXlsxFile(path, fileName = basename(path), options) {
  const source = createFileImportedXlsxSourceReader(path)
  try {
    return importXlsxFromZipByteSource(source, fileName, options)
  } catch (error) {
    source.release?.()
    throw error
  }
}
`,
)
writeFileSync(
  join(headlessDistDir, 'xlsx-runtime.d.ts'),
  `import type { WorkbookSnapshot } from '@bilig/protocol'
import type { ImportedWorkbook, XlsxByteSourceImportOptions, XlsxSourceLiteralPatchFileExportResult } from './xlsx-internal/index.js'

export * from './xlsx-internal/index.js'

export interface WorkPaperXlsxExportSource {
  exportSnapshot(): WorkbookSnapshot
  exportSourcePreservingXlsxSnapshot?(): WorkbookSnapshot | null
}

export declare function exportXlsx(snapshot: WorkbookSnapshot): Uint8Array
export declare function exportXlsx(workbook: WorkPaperXlsxExportSource): Uint8Array
export declare function exportWorkPaperXlsx(workbook: WorkPaperXlsxExportSource): Uint8Array
export declare function exportWorkPaperXlsxToFileAsync(
  workbook: WorkPaperXlsxExportSource,
  outputPath: string,
): Promise<XlsxSourceLiteralPatchFileExportResult>
export declare function importXlsxFile(path: string, fileName?: string, options?: XlsxByteSourceImportOptions): ImportedWorkbook
`,
)
writeFileSync(
  join(headlessDistDir, 'formula-clinic-bin.js'),
  `#!/usr/bin/env node
import { importXlsx } from './xlsx.js'
import { runFormulaClinicCli } from './formula-clinic-cli.js'

process.exitCode = runFormulaClinicCli({
  argv: process.argv.slice(2),
  importXlsx,
  writeStderr: (text) => process.stderr.write(text),
  writeStdout: (text) => process.stdout.write(text),
})
`,
)
