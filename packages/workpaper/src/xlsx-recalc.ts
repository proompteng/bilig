import { writeFileSync } from 'node:fs'
import { WorkPaper, type RawCellContent, type WorkPaperChange } from './work-paper.js'
import { exportXlsx, exportXlsxToFile, importXlsx } from './xlsx-runtime.js'
import {
  type CellValue,
  ErrorCode,
  formatErrorCode,
  type LiteralInput,
  ValueTag,
  type CompatibilityMode,
  type WorkbookCalculationMode,
  type WorkbookDateSystem,
} from '@bilig/protocol'
import {
  patchXlsxTextParts,
  type ImportedWorkbookDiagnostics,
  recalculateXlsxFileToFileStreamingNative,
  type XlsxExternalWorkbookInput,
  type XlsxFormulaRecalcNativeDiagnostics,
  type XlsxImportOptions,
  type XlsxTextPartPatch,
} from '@bilig/xlsx'
import { parseA1CellReference, parseQualifiedCellTarget } from './xlsx-recalc-targets.js'

export { WorkPaper } from './work-paper.js'
export { exportXlsx, exportXlsxToFile, importXlsx } from './xlsx-runtime.js'
export { parseQualifiedA1, parseQualifiedCellTarget } from './xlsx-recalc-targets.js'
export type { XlsxExternalWorkbookInput, XlsxExternalWorkbookHydrationDiagnostics } from '@bilig/xlsx'
export {
  type StreamingNativeFormulaCounts,
  type XlsxFormulaRecalcNativeDiagnostics,
  type XlsxFormulaRecalcPhaseRss,
  StreamingNativeXlsxRecalcError,
} from '@bilig/xlsx'
export {
  type StreamingNativeXlsxCacheFormulaInspection,
  type StreamingNativeXlsxCacheInspectionLimit,
  type StreamingNativeXlsxCacheInspectionResult,
  type StreamingNativeXlsxCacheLiteral,
  type StreamingNativeXlsxCacheStatus,
  type StreamingNativeXlsxCacheStatusSummary,
  inspectXlsxCacheFileStreamingNative,
} from '@bilig/xlsx'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSnapshot>
type WorkbookSnapshot = Parameters<typeof WorkPaper.buildFromSnapshot>[0]
type WorkbookCalculationSettings = NonNullable<NonNullable<WorkbookSnapshot['workbook']['metadata']>['calculationSettings']>
interface FormulaErrorCache {
  readonly sheetName: string
  readonly address: string
  readonly value: string
}

const sourcePreservingOutputCalculationSettings = Symbol.for('bilig.sourcePreservingXlsxOutputCalculationSettings')
const importedXlsxSourceBytes = Symbol.for('bilig.importedXlsxSourceBytes')
const importedXlsxSourceCellPatches = Symbol.for('bilig.importedXlsxSourceCellPatches')
const legacyWorkPaperBytesApiLimit = 1_000_000

interface ImportedXlsxSourceCellPatch {
  readonly kind: 'literal'
  readonly sheetName: string
  readonly address: string
  readonly value: string | number | boolean | null
  readonly preserveFormula?: boolean
}

type SnapshotWithImportedXlsxSource = WorkbookSnapshot & {
  readonly [importedXlsxSourceBytes]?: unknown
  readonly [importedXlsxSourceCellPatches]?: readonly ImportedXlsxSourceCellPatch[]
}

interface PreparedXlsxFormulaRecalcOutput {
  readonly outputSnapshot: WorkbookSnapshot
  readonly errorCaches: readonly FormulaErrorCache[]
  readonly warnings: readonly string[]
  readonly sheetNames: readonly string[]
  readonly reads: Readonly<Record<string, XlsxFormulaRecalcCellValue>>
  readonly changes: readonly WorkPaperChange[]
  readonly diagnostics?: ImportedWorkbookDiagnostics
}

export type XlsxFormulaRecalcCellValue = CellValue
export type XlsxFormulaRecalcChange = WorkPaperChange
export type XlsxFormulaRecalcDiagnostics = ImportedWorkbookDiagnostics & Partial<XlsxFormulaRecalcNativeDiagnostics>

export interface XlsxFormulaRecalcEdit {
  readonly target: string
  readonly value: LiteralInput
}

export type XlsxFormulaRecalcWorkPaperEngine = 'auto' | 'workpaper'
export type XlsxFormulaRecalcWorkPaperFallbackPolicy = 'error' | 'workpaper'

export interface XlsxFormulaRecalcWorkPaperConfig {
  readonly calculationSettings?:
    | {
        readonly mode?: WorkbookCalculationMode
        readonly compatibilityMode?: CompatibilityMode
        readonly dateSystem?: WorkbookDateSystem
        readonly iterate?: boolean | null
        readonly iterateCount?: number | null
        readonly iterateDelta?: string | null
        readonly fullPrecision?: boolean | null
        readonly fullCalcOnLoad?: boolean | null
        readonly calcOnSave?: boolean | null
        readonly calcCompleted?: boolean | null
        readonly concurrentCalc?: boolean | null
      }
    | undefined
  readonly evaluationTimeoutMs?: number
  readonly maxRows?: number
  readonly maxColumns?: number
  readonly useColumnIndex?: boolean
  readonly [key: string]: unknown
}

export interface XlsxFormulaRecalcOptions {
  readonly fileName?: string
  readonly externalWorkbooks?: readonly XlsxExternalWorkbookInput[]
  readonly edits?: readonly XlsxFormulaRecalcEdit[]
  readonly reads?: readonly string[]
  readonly config?: XlsxFormulaRecalcWorkPaperConfig
  readonly engine?: XlsxFormulaRecalcWorkPaperEngine
  readonly maxRssBytes?: number
  readonly fallbackPolicy?: XlsxFormulaRecalcWorkPaperFallbackPolicy
}

export interface XlsxFormulaRecalcFileOptions extends XlsxFormulaRecalcOptions {
  readonly outputPath: string
}

export interface XlsxFormulaRecalcResult {
  readonly xlsx: Uint8Array
  readonly warnings: readonly string[]
  readonly sheetNames: readonly string[]
  readonly reads: Readonly<Record<string, XlsxFormulaRecalcCellValue>>
  readonly changes: readonly XlsxFormulaRecalcChange[]
  readonly diagnostics?: XlsxFormulaRecalcDiagnostics
}

export interface XlsxFormulaRecalcFileResult {
  readonly bytesWritten: number
  readonly warnings: readonly string[]
  readonly sheetNames: readonly string[]
  readonly reads: Readonly<Record<string, XlsxFormulaRecalcCellValue>>
  readonly changes: readonly XlsxFormulaRecalcChange[]
  readonly diagnostics?: XlsxFormulaRecalcDiagnostics
}

export type XlsxFormulaRecalcFileToFileEngine = 'auto' | 'streaming-native'
export type XlsxFormulaRecalcFileToFileFallbackPolicy = 'error'

export interface XlsxFormulaRecalcFileToFileOptions {
  readonly outputPath: string
  readonly externalWorkbooks?: readonly XlsxExternalWorkbookInput[]
  readonly edits?: readonly XlsxFormulaRecalcEdit[]
  readonly reads?: readonly string[]
  readonly engine?: XlsxFormulaRecalcFileToFileEngine
  readonly maxRssBytes?: number
  readonly fallbackPolicy?: XlsxFormulaRecalcFileToFileFallbackPolicy
}

export function recalculateXlsx(input: Uint8Array | ArrayBuffer | Buffer, options: XlsxFormulaRecalcOptions = {}): XlsxFormulaRecalcResult {
  assertBytesApiEngine((options as { readonly engine?: string }).engine)
  const bytes = toUint8Array(input)
  assertLegacyWorkPaperBytesApiWithinLimit(bytes, 'recalculateXlsx')
  return withPreparedRecalculatedXlsxOutput(bytes, options, (prepared) => {
    const exportedXlsx = toUint8Array(exportXlsx(prepared.outputSnapshot))
    return {
      xlsx: addFormulaErrorCachesToXlsxBytes(exportedXlsx, prepared.outputSnapshot, prepared.errorCaches),
      warnings: prepared.warnings,
      sheetNames: prepared.sheetNames,
      reads: prepared.reads,
      changes: prepared.changes,
      ...(prepared.diagnostics ? { diagnostics: prepared.diagnostics } : {}),
    }
  })
}

export async function recalculateXlsxFileToFile(
  inputPath: string,
  options: XlsxFormulaRecalcFileToFileOptions,
): Promise<XlsxFormulaRecalcFileResult> {
  const legacyOptions = options as {
    readonly config?: unknown
    readonly engine?: string
    readonly fallbackPolicy?: string
  }
  if (legacyOptions.engine === 'workpaper') {
    throw legacyWorkPaperFileToFileError('engine')
  }
  if (legacyOptions.fallbackPolicy === 'workpaper') {
    throw legacyWorkPaperFileToFileError('fallbackPolicy')
  }
  if (legacyOptions.config !== undefined) {
    throw new Error('streaming-native file-to-file recalc does not support WorkPaper config options')
  }
  return await recalculateXlsxFileToFileStreamingNative(inputPath, {
    outputPath: options.outputPath,
    ...(options.edits === undefined ? {} : { edits: nativeLiteralEdits(options.edits) }),
    ...(options.reads === undefined ? {} : { reads: options.reads }),
    ...(options.externalWorkbooks === undefined ? {} : { externalWorkbooks: options.externalWorkbooks }),
    ...(options.maxRssBytes === undefined ? {} : { maxRssBytes: options.maxRssBytes }),
  })
}

export function recalculateXlsxToFile(
  input: Uint8Array | ArrayBuffer | Buffer,
  options: XlsxFormulaRecalcFileOptions,
): XlsxFormulaRecalcFileResult {
  assertBytesApiEngine(options.engine)
  const bytes = toUint8Array(input)
  assertLegacyWorkPaperBytesApiWithinLimit(bytes, 'recalculateXlsxToFile')
  return withPreparedRecalculatedXlsxOutput(bytes, options, (prepared) => {
    if (prepared.errorCaches.length === 0) {
      const exported = exportXlsxToFile(prepared.outputSnapshot, options.outputPath)
      return {
        bytesWritten: exported.bytesWritten,
        warnings: prepared.warnings,
        sheetNames: prepared.sheetNames,
        reads: prepared.reads,
        changes: prepared.changes,
        ...(prepared.diagnostics ? { diagnostics: prepared.diagnostics } : {}),
      }
    }
    const exportedXlsx = addFormulaErrorCachesToXlsxBytes(
      toUint8Array(exportXlsx(prepared.outputSnapshot)),
      prepared.outputSnapshot,
      prepared.errorCaches,
    )
    writeFileSync(options.outputPath, exportedXlsx)
    return {
      bytesWritten: exportedXlsx.byteLength,
      warnings: prepared.warnings,
      sheetNames: prepared.sheetNames,
      reads: prepared.reads,
      changes: prepared.changes,
      ...(prepared.diagnostics ? { diagnostics: prepared.diagnostics } : {}),
    }
  })
}

function assertBytesApiEngine(engine: string | undefined): void {
  if (engine === 'streaming-native') {
    throw new Error('streaming-native engine requires recalculateXlsxFileToFile() with file-backed input and output paths')
  }
}

function assertLegacyWorkPaperBytesApiWithinLimit(input: Uint8Array, apiName: string): void {
  if (input.byteLength <= legacyWorkPaperBytesApiLimit) {
    return
  }
  throw new Error(
    `${apiName} legacy bytes API is small-workbook only (${input.byteLength.toLocaleString('en-US')} bytes > ` +
      `${legacyWorkPaperBytesApiLimit.toLocaleString('en-US')} bytes). Use recalculateXlsxFileToFile() for file-backed streaming-native XLSX jobs.`,
  )
}

function nativeLiteralEdits(edits: readonly XlsxFormulaRecalcEdit[]): readonly { readonly target: string; readonly value: LiteralInput }[] {
  return edits.map((edit) => {
    const value = edit.value
    if (value === null || typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') {
      return { target: edit.target, value }
    }
    throw new Error(`streaming-native file-to-file recalc supports literal edits only: ${edit.target}`)
  })
}

function legacyWorkPaperFileToFileError(optionName: 'engine' | 'fallbackPolicy'): Error {
  return new Error(
    `bilig-workpaper/xlsx file-to-file recalc uses the streaming-native engine for large XLSX jobs; ${optionName}=workpaper is only available through the legacy bytes APIs.`,
  )
}

function withPreparedRecalculatedXlsxOutput<Result>(
  input: Uint8Array | ArrayBuffer | Buffer,
  options: XlsxFormulaRecalcOptions,
  consume: (prepared: PreparedXlsxFormulaRecalcOutput) => Result,
): Result {
  const imported = importXlsx(
    toUint8Array(input),
    options.fileName ?? 'workbook.xlsx',
    xlsxFormulaRecalcImportOptions(options.externalWorkbooks),
  )
  const originalCalculationSettings = imported.snapshot.workbook.metadata?.calculationSettings
  const workbook = WorkPaper.buildFromSnapshot(snapshotForFreshFormulaRecalculation(imported.snapshot), {
    evaluationTimeoutMs: 30_000,
    useColumnIndex: true,
    ...options.config,
  })

  try {
    const changes: WorkPaperChange[] = []
    for (const edit of options.edits ?? []) {
      appendChanges(changes, workbook.setCellContents(parseQualifiedCellTarget(workbook, edit.target), edit.value))
    }
    appendChanges(changes, workbook.rebuildAndRecalculate())

    const reads: Record<string, XlsxFormulaRecalcCellValue> = {}
    for (const target of options.reads ?? []) {
      reads[target] = workbook.getCellValue(parseQualifiedCellTarget(workbook, target))
    }

    const sourcePreservingOutputSnapshot = sourcePreservingSnapshotForRecalculationExport(
      workbook,
      imported.snapshot,
      options.config?.calculationSettings === undefined ? originalCalculationSettings : undefined,
      options.edits?.length ?? 0,
    )
    if (sourcePreservingOutputSnapshot) {
      return consume({
        outputSnapshot: sourcePreservingOutputSnapshot,
        errorCaches: [],
        warnings: imported.warnings,
        sheetNames: imported.sheetNames,
        reads,
        changes,
        ...(imported.diagnostics ? { diagnostics: imported.diagnostics } : {}),
      })
    }

    const outputFormulaCaches = snapshotWithFormulaCachedValues(workbook, workbook.exportSnapshot())
    return consume({
      outputSnapshot: restoreOutputCalculationSettings(
        outputFormulaCaches.snapshot,
        options.config?.calculationSettings === undefined ? originalCalculationSettings : undefined,
      ),
      errorCaches: outputFormulaCaches.errorCaches,
      warnings: imported.warnings,
      sheetNames: imported.sheetNames,
      reads,
      changes,
      ...(imported.diagnostics ? { diagnostics: imported.diagnostics } : {}),
    })
  } finally {
    workbook.dispose()
  }
}

function sourcePreservingSnapshotForRecalculationExport(
  workbook: WorkPaperInstance,
  importedSnapshot: WorkbookSnapshot,
  originalCalculationSettings: WorkbookCalculationSettings | undefined,
  editCount: number,
): WorkbookSnapshot | null {
  const sourcePreservingSnapshot = workbook.exportSourcePreservingXlsxSnapshot?.()
  if (!sourcePreservingSnapshot && editCount > 0) {
    return null
  }
  const source = readImportedXlsxSourceReference(sourcePreservingSnapshot ?? importedSnapshot)
  if (source === undefined) {
    return null
  }
  const formulaCachePatches = formulaCachePatchesForImportedSnapshot(workbook, importedSnapshot)
  if (formulaCachePatches === null) {
    return null
  }
  const outputSnapshot = attachSourcePreservingRecalculationPatches(
    restoreOutputCalculationSettings(
      sourcePreservingSnapshot ?? minimalSourcePreservingSnapshot(importedSnapshot),
      originalCalculationSettings,
    ),
    source,
    mergeImportedXlsxSourceCellPatches(readImportedXlsxSourceCellPatches(sourcePreservingSnapshot), formulaCachePatches),
  )
  if (originalCalculationSettings !== undefined) {
    Object.defineProperty(outputSnapshot, sourcePreservingOutputCalculationSettings, {
      configurable: true,
      enumerable: false,
      value: calculationSettingsAfterExplicitRecalculation(originalCalculationSettings),
    })
  }
  return outputSnapshot
}

function formulaCachePatchesForImportedSnapshot(
  workbook: WorkPaperInstance,
  importedSnapshot: WorkbookSnapshot,
): readonly ImportedXlsxSourceCellPatch[] | null {
  const patches: ImportedXlsxSourceCellPatch[] = []
  for (const sheet of importedSnapshot.sheets) {
    const sheetId = workbook.getSheetId(sheet.name)
    if (sheetId === undefined) {
      continue
    }
    for (const cell of sheet.cells) {
      if (typeof cell.formula !== 'string' || cell.formula.trim().length === 0) {
        continue
      }
      const address = cellA1Address(cell)
      if (address === null) {
        return null
      }
      const coordinates = cellCoordinates(cell)
      if (coordinates === null) {
        return null
      }
      const cellValue = workbook.getCellValue({ sheet: sheetId, row: coordinates.row, col: coordinates.col })
      const cachedValue = literalInputForFormulaCache(cellValue)
      if (cachedValue !== undefined) {
        patches.push({ kind: 'literal', sheetName: sheet.name, address, value: cachedValue, preserveFormula: true })
        continue
      }
      if (errorInputForFormulaCache(cellValue) !== undefined) {
        return null
      }
    }
  }
  return patches
}

function readImportedXlsxSourceReference(snapshot: WorkbookSnapshot | null | undefined): unknown {
  return snapshot === null || snapshot === undefined ? undefined : (snapshot as SnapshotWithImportedXlsxSource)[importedXlsxSourceBytes]
}

function readImportedXlsxSourceCellPatches(snapshot: WorkbookSnapshot | null | undefined): readonly ImportedXlsxSourceCellPatch[] {
  return snapshot === null || snapshot === undefined
    ? []
    : ((snapshot as SnapshotWithImportedXlsxSource)[importedXlsxSourceCellPatches] ?? [])
}

function attachSourcePreservingRecalculationPatches(
  snapshot: WorkbookSnapshot,
  source: unknown,
  patches: readonly ImportedXlsxSourceCellPatch[],
): WorkbookSnapshot {
  Object.defineProperty(snapshot, importedXlsxSourceBytes, {
    configurable: true,
    enumerable: false,
    value: source,
  })
  Object.defineProperty(snapshot, importedXlsxSourceCellPatches, {
    configurable: true,
    enumerable: false,
    value: patches,
  })
  return snapshot
}

function mergeImportedXlsxSourceCellPatches(
  basePatches: readonly ImportedXlsxSourceCellPatch[],
  formulaCachePatches: readonly ImportedXlsxSourceCellPatch[],
): readonly ImportedXlsxSourceCellPatch[] {
  const merged = new Map<string, ImportedXlsxSourceCellPatch>()
  for (const patch of basePatches) {
    merged.set(`${patch.sheetName}!${patch.address}`, patch)
  }
  for (const patch of formulaCachePatches) {
    merged.set(`${patch.sheetName}!${patch.address}`, patch)
  }
  return [...merged.values()]
}

function minimalSourcePreservingSnapshot(importedSnapshot: WorkbookSnapshot): WorkbookSnapshot {
  return {
    version: importedSnapshot.version,
    workbook: { name: importedSnapshot.workbook.name },
    sheets: importedSnapshot.sheets.map((sheet) => ({
      ...(sheet.id === undefined ? {} : { id: sheet.id }),
      name: sheet.name,
      order: sheet.order,
      cells: [],
    })),
  }
}

function cellA1Address(cell: WorkbookSnapshot['sheets'][number]['cells'][number]): string | null {
  if (typeof cell.address === 'string' && cell.address.length > 0) {
    return cell.address
  }
  const coordinates = cellCoordinates(cell)
  return coordinates ? encodeA1CellReference(coordinates.row, coordinates.col) : null
}

function cellCoordinates(cell: WorkbookSnapshot['sheets'][number]['cells'][number]): { readonly row: number; readonly col: number } | null {
  const row = cell.row
  const col = cell.col
  if (Number.isInteger(row) && Number.isInteger(col) && row !== undefined && col !== undefined) {
    return { row, col }
  }
  if (typeof cell.address !== 'string') {
    return null
  }
  try {
    return parseA1CellReference(cell.address)
  } catch {
    return null
  }
}

function encodeA1CellReference(row: number, col: number): string {
  let column = ''
  for (let value = col + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    column = String.fromCharCode(((value - 1) % 26) + 65) + column
  }
  return `${column}${String(row + 1)}`
}

export const recalculateSheetjsWorkbook = recalculateXlsx

function appendChanges(target: WorkPaperChange[], changes: readonly WorkPaperChange[]): void {
  for (const change of changes) {
    target.push(change)
  }
}

export const xlsxCacheDoctorSchemaVersion = 'xlsx-cache-doctor.v1'

export type XlsxCacheInspectionLimit = number | 'all'
export type XlsxCacheStatus = 'fresh' | 'stale' | 'missing-cache' | 'unsupported-recalculation'

export interface XlsxCacheInspectionOptions {
  readonly fileName?: string
  readonly externalWorkbooks?: readonly XlsxExternalWorkbookInput[]
  readonly edits?: readonly XlsxFormulaRecalcEdit[]
  readonly inspectLimit?: XlsxCacheInspectionLimit
  readonly config?: XlsxFormulaRecalcWorkPaperConfig
}

export interface XlsxCacheStatusSummary {
  readonly inspected: number
  readonly stale: number
  readonly fresh: number
  readonly missingCache: number
  readonly unsupportedRecalculation: number
}

export interface XlsxCacheFormulaInspection {
  readonly target: string
  readonly formula: string
  readonly cachedValue?: RawCellContent
  readonly recalculatedValue: XlsxFormulaRecalcCellValue | undefined
  readonly literalRecalculatedValue?: RawCellContent | string
  readonly cacheStatus: XlsxCacheStatus
  readonly staleCachedValue: boolean | null
}

export interface XlsxCacheInspectionResult {
  readonly schemaVersion: typeof xlsxCacheDoctorSchemaVersion
  readonly sheetNames: readonly string[]
  readonly formulaCellCount: number
  readonly inspectedFormulaCellCount: number
  readonly uninspectedFormulaCellCount: number
  readonly inspectionLimit: XlsxCacheInspectionLimit
  readonly staleCachedFormulaCount: number
  readonly cacheStatusSummary: XlsxCacheStatusSummary
  readonly suggestedReads: readonly string[]
  readonly formulas: readonly XlsxCacheFormulaInspection[]
  readonly warnings: readonly string[]
  readonly diagnostics?: ImportedWorkbookDiagnostics
  readonly inspectionCompleted: true
  readonly recalculationCompleted: true
  readonly excelParity: 'not_proven'
}

interface XlsxCacheFormulaCell {
  readonly target: string
  readonly formula: string
  readonly cachedValue?: RawCellContent
}

export function inspectXlsxCache(
  input: Uint8Array | ArrayBuffer | Buffer,
  options: XlsxCacheInspectionOptions = {},
): XlsxCacheInspectionResult {
  const bytes = toUint8Array(input)
  assertLegacyWorkPaperBytesApiWithinLimit(bytes, 'inspectXlsxCache')
  const imported = importXlsx(bytes, options.fileName ?? 'workbook.xlsx', xlsxFormulaRecalcImportOptions(options.externalWorkbooks))
  const formulaCells = collectXlsxCacheFormulaCells(imported.snapshot)
  const inspectionLimit = normalizeXlsxCacheInspectionLimit(options.inspectLimit ?? 'all')
  const inspectedFormulaCells = inspectionLimit === 'all' ? formulaCells : formulaCells.slice(0, inspectionLimit)
  const uninspectedFormulaCellCount = formulaCells.length - inspectedFormulaCells.length
  const suggestedReads = inspectedFormulaCells.map((cell) => cell.target)
  const recalculated = recalculateXlsx(bytes, {
    ...(options.externalWorkbooks ? { externalWorkbooks: options.externalWorkbooks } : {}),
    ...(options.fileName ? { fileName: options.fileName } : {}),
    ...(options.edits ? { edits: options.edits } : {}),
    reads: suggestedReads,
    ...(options.config ? { config: options.config } : {}),
  })
  const formulas = inspectedFormulaCells.map((cell) => {
    const recalculatedValue = recalculated.reads[cell.target]
    const literalRecalculatedValue = literalValueForXlsxCacheInspection(recalculatedValue)
    const cacheStatus = xlsxCacheStatusForInspection(cell.cachedValue, literalRecalculatedValue)
    return {
      target: cell.target,
      formula: cell.formula,
      ...(cell.cachedValue !== undefined ? { cachedValue: cell.cachedValue } : {}),
      recalculatedValue,
      ...(literalRecalculatedValue !== undefined ? { literalRecalculatedValue } : {}),
      cacheStatus,
      staleCachedValue: staleCachedValueForXlsxCacheInspection(cacheStatus),
    }
  })

  return {
    schemaVersion: xlsxCacheDoctorSchemaVersion,
    sheetNames: imported.sheetNames,
    formulaCellCount: formulaCells.length,
    inspectedFormulaCellCount: inspectedFormulaCells.length,
    uninspectedFormulaCellCount,
    inspectionLimit,
    staleCachedFormulaCount: formulas.filter((formula) => formula.staleCachedValue === true).length,
    cacheStatusSummary: buildXlsxCacheStatusSummary(formulas),
    suggestedReads,
    formulas,
    warnings: recalculated.warnings,
    ...(recalculated.diagnostics ? { diagnostics: recalculated.diagnostics } : {}),
    inspectionCompleted: true,
    recalculationCompleted: true,
    excelParity: 'not_proven',
  }
}

function xlsxFormulaRecalcImportOptions(externalWorkbooks: readonly XlsxExternalWorkbookInput[] | undefined): XlsxImportOptions {
  return externalWorkbooks && externalWorkbooks.length > 0
    ? { externalWorkbooks, externalLinkCacheArtifactMode: 'replace-refreshed' }
    : { preferNativeSimpleImport: true }
}

function collectXlsxCacheFormulaCells(snapshot: WorkbookSnapshot): XlsxCacheFormulaCell[] {
  const cells: XlsxCacheFormulaCell[] = []
  for (const sheet of snapshot.sheets.toSorted((left, right) => left.order - right.order)) {
    for (const cell of sheet.cells.toSorted((left, right) => compareA1Addresses(left.address, right.address))) {
      if (typeof cell.formula !== 'string' || cell.formula.trim().length === 0) {
        continue
      }
      cells.push({
        target: formatQualifiedTarget(sheet.name, cell.address),
        formula: cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`,
        ...(cell.value !== undefined ? { cachedValue: cell.value } : {}),
      })
    }
  }
  return cells
}

function normalizeXlsxCacheInspectionLimit(limit: XlsxCacheInspectionLimit): XlsxCacheInspectionLimit {
  if (limit === 'all') {
    return limit
  }
  if (Number.isInteger(limit) && limit > 0) {
    return limit
  }
  throw new Error(`Expected inspectLimit to be "all" or a positive integer, received: ${String(limit)}`)
}

function compareA1Addresses(left: string, right: string): number {
  const leftParts = parseA1AddressForSort(left)
  const rightParts = parseA1AddressForSort(right)
  return leftParts.row - rightParts.row || leftParts.col - rightParts.col || left.localeCompare(right)
}

function parseA1AddressForSort(address: string): { readonly row: number; readonly col: number } {
  const match = /^([A-Z]+)(\d+)$/iu.exec(address)
  if (!match) {
    return { row: Number.MAX_SAFE_INTEGER, col: Number.MAX_SAFE_INTEGER }
  }
  const [, letters = '', rowText = ''] = match
  let col = 0
  for (const letter of letters.toUpperCase()) {
    col = col * 26 + letter.charCodeAt(0) - 64
  }
  return { row: Number(rowText), col }
}

function formatQualifiedTarget(sheetName: string, address: string): string {
  return `${quoteSheetNameForTarget(sheetName)}!${address}`
}

function quoteSheetNameForTarget(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/u.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

function literalValueForXlsxCacheInspection(value: XlsxFormulaRecalcCellValue | undefined): RawCellContent | string | undefined {
  if (value === undefined || typeof value !== 'object' || value === null || !('tag' in value)) {
    return undefined
  }
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
      return 'value' in value && typeof value.value === 'number' && Number.isFinite(value.value) ? value.value : undefined
    case ValueTag.Boolean:
      return 'value' in value && typeof value.value === 'boolean' ? value.value : undefined
    case ValueTag.String:
      return 'value' in value && typeof value.value === 'string' ? value.value : undefined
    case ValueTag.Error:
      return 'code' in value && typeof value.code === 'number' ? formatErrorCode(value.code) : undefined
  }
}

function xlsxCacheStatusForInspection(
  cachedValue: RawCellContent | undefined,
  literalRecalculatedValue: RawCellContent | string | undefined,
): XlsxCacheStatus {
  if (cachedValue === undefined) {
    return 'missing-cache'
  }
  if (literalRecalculatedValue === undefined) {
    return 'unsupported-recalculation'
  }
  return literalValuesEqual(cachedValue, literalRecalculatedValue) ? 'fresh' : 'stale'
}

function staleCachedValueForXlsxCacheInspection(cacheStatus: XlsxCacheStatus): XlsxCacheFormulaInspection['staleCachedValue'] {
  switch (cacheStatus) {
    case 'stale':
      return true
    case 'fresh':
      return false
    case 'missing-cache':
    case 'unsupported-recalculation':
      return null
  }
}

function buildXlsxCacheStatusSummary(formulas: readonly XlsxCacheFormulaInspection[]): XlsxCacheStatusSummary {
  return {
    inspected: formulas.length,
    stale: formulas.filter((formula) => formula.cacheStatus === 'stale').length,
    fresh: formulas.filter((formula) => formula.cacheStatus === 'fresh').length,
    missingCache: formulas.filter((formula) => formula.cacheStatus === 'missing-cache').length,
    unsupportedRecalculation: formulas.filter((formula) => formula.cacheStatus === 'unsupported-recalculation').length,
  }
}

function literalValuesEqual(left: RawCellContent, right: RawCellContent | string): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function snapshotWithFormulaCachedValues(
  workbook: WorkPaperInstance,
  snapshot: WorkbookSnapshot,
): { readonly snapshot: WorkbookSnapshot; readonly errorCaches: readonly FormulaErrorCache[] } {
  const next: WorkbookSnapshot = {
    ...snapshot,
    workbook: {
      ...snapshot.workbook,
      ...(snapshot.workbook.metadata === undefined ? {} : { metadata: structuredClone(snapshot.workbook.metadata) }),
    },
    sheets: snapshot.sheets.map((sheet) => ({
      ...sheet,
      ...(sheet.metadata === undefined ? {} : { metadata: structuredClone(sheet.metadata) }),
      cells: sheet.cells.map((cell) => ({ ...cell })),
    })),
  }
  const errorCaches: FormulaErrorCache[] = []
  for (const sheet of next.sheets) {
    const sheetId = workbook.getSheetId(sheet.name)
    if (sheetId === undefined) {
      continue
    }
    for (const cell of sheet.cells) {
      if (typeof cell.formula !== 'string' || cell.formula.trim().length === 0) {
        continue
      }
      const address = parseA1CellReference(cell.address)
      const cellValue = workbook.getCellValue({ sheet: sheetId, row: address.row, col: address.col })
      const cachedValue = literalInputForFormulaCache(cellValue)
      if (cachedValue !== undefined) {
        cell.value = cachedValue
        continue
      }
      const cachedError = errorInputForFormulaCache(cellValue)
      if (cachedError !== undefined) {
        delete cell.value
        errorCaches.push({ sheetName: sheet.name, address: cell.address, value: cachedError })
      }
    }
  }
  return { snapshot: next, errorCaches }
}

function literalInputForFormulaCache(value: XlsxFormulaRecalcCellValue): string | number | boolean | null | undefined {
  if (typeof value !== 'object' || value === null || !('tag' in value)) {
    return undefined
  }
  if (value.tag === ValueTag.Empty) {
    return null
  }
  if (value.tag === ValueTag.Number && 'value' in value && typeof value.value === 'number' && Number.isFinite(value.value)) {
    return value.value
  }
  if (value.tag === ValueTag.Boolean && 'value' in value && typeof value.value === 'boolean') {
    return value.value
  }
  if (value.tag === ValueTag.String && 'value' in value && typeof value.value === 'string') {
    return value.value
  }
  return undefined
}

function errorInputForFormulaCache(value: XlsxFormulaRecalcCellValue): string | undefined {
  if (typeof value !== 'object' || value === null || !('tag' in value) || value.tag !== ValueTag.Error || !('code' in value)) {
    return undefined
  }
  if (typeof value.code !== 'number') {
    return undefined
  }
  // Desktop Excel displays #FIELD! for linked-data field failures but saves #VALUE! in worksheet formula caches.
  return value.code === ErrorCode.Field ? '#VALUE!' : formatErrorCode(value.code)
}

function addFormulaErrorCachesToXlsxBytes(
  bytes: Uint8Array,
  snapshot: WorkbookSnapshot,
  errorCaches: readonly FormulaErrorCache[],
): Uint8Array {
  if (errorCaches.length === 0) {
    return bytes
  }
  const cachesBySheetName = new Map<string, FormulaErrorCache[]>()
  for (const cache of errorCaches) {
    const sheetCaches = cachesBySheetName.get(cache.sheetName) ?? []
    sheetCaches.push(cache)
    cachesBySheetName.set(cache.sheetName, sheetCaches)
  }
  const orderedSheets = snapshot.sheets.toSorted((left, right) => left.order - right.order)
  const patches: XlsxTextPartPatch[] = []
  for (let index = 0; index < orderedSheets.length; index += 1) {
    const sheet = orderedSheets[index]!
    const sheetCaches = cachesBySheetName.get(sheet.name)
    if (!sheetCaches || sheetCaches.length === 0) {
      continue
    }
    const sheetPath = `xl/worksheets/sheet${String(index + 1)}.xml`
    patches.push({
      path: sheetPath,
      patchText: (sheetXml) =>
        sheetCaches.reduce((nextXml, cache) => replaceFormulaCellErrorCache(nextXml, cache.address, cache.value), sheetXml),
    })
  }
  return patchXlsxTextParts(bytes, patches)
}

function replaceFormulaCellErrorCache(sheetXml: string, address: string, value: string): string {
  const pattern = new RegExp(`<c\\b(?=[^>]*\\br="${escapeRegExp(address)}")[\\s\\S]*?<\\/c>`, 'u')
  if (!pattern.test(sheetXml)) {
    return sheetXml
  }
  return sheetXml.replace(pattern, (cellXml) => writeCellErrorCache(cellXml, value))
}

function writeCellErrorCache(cellXml: string, value: string): string {
  const escapedValue = escapeXmlText(value)
  const typedCellXml = cellXml.replace(/^<c\b([^>]*)>/u, (_match: string, attributes: string) => {
    const nextAttributes = attributes.replace(/\s+t="[^"]*"/u, '')
    return `<c${nextAttributes} t="e">`
  })
  if (/<v>[\s\S]*?<\/v>/u.test(typedCellXml)) {
    return typedCellXml.replace(/<v>[\s\S]*?<\/v>/u, `<v>${escapedValue}</v>`)
  }
  if (/<\/f>/u.test(typedCellXml)) {
    return typedCellXml.replace(/<\/f>/u, `</f><v>${escapedValue}</v>`)
  }
  return typedCellXml.replace(/<\/c>$/u, `<v>${escapedValue}</v></c>`)
}

function escapeXmlText(value: string): string {
  return value.replace(/&/gu, '&amp;').replace(/</gu, '&lt;').replace(/>/gu, '&gt;')
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function snapshotForFreshFormulaRecalculation(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  const calculationSettings = snapshot.workbook.metadata?.calculationSettings
  if (calculationSettings === undefined) {
    return snapshot
  }
  if (calculationSettings.mode !== 'manual' && calculationSettings.fullCalcOnLoad !== false) {
    return snapshot
  }

  return {
    ...snapshot,
    workbook: {
      ...snapshot.workbook,
      metadata: {
        ...snapshot.workbook.metadata,
        calculationSettings: {
          ...calculationSettings,
          mode: 'automatic',
          fullCalcOnLoad: true,
        },
      },
    },
  }
}

function restoreOutputCalculationSettings(
  snapshot: WorkbookSnapshot,
  originalCalculationSettings: WorkbookCalculationSettings | undefined,
): WorkbookSnapshot {
  if (originalCalculationSettings === undefined) {
    return snapshot
  }
  return {
    ...snapshot,
    workbook: {
      ...snapshot.workbook,
      metadata: {
        ...snapshot.workbook.metadata,
        calculationSettings: calculationSettingsAfterExplicitRecalculation(originalCalculationSettings),
      },
    },
  }
}

function calculationSettingsAfterExplicitRecalculation(settings: WorkbookCalculationSettings): WorkbookCalculationSettings {
  const recalculated = structuredClone(settings)
  delete recalculated.calcCompleted
  delete recalculated.calcOnSave
  delete recalculated.forceFullCalc
  delete recalculated.fullCalcOnLoad
  return recalculated
}

function toUint8Array(input: Uint8Array | ArrayBuffer | Buffer): Uint8Array {
  if (input instanceof Uint8Array) {
    return input
  }
  return new Uint8Array(input)
}
