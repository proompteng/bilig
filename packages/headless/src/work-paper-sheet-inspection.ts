import { MAX_COLS, MAX_ROWS, ValueTag, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
import { FORMULA_SPILL_PRODUCING_FUNCTION_NAMES, compileFormula, parseCellAddress } from '@bilig/formula'
import { WorkPaperSheetSizeLimitExceededError, WorkPaperUnableToParseError } from './work-paper-errors.js'
import { isBlankRawCellContent } from './work-paper-runtime-helpers.js'
import type {
  WorkPaperCellAddress,
  WorkPaperCellType,
  WorkPaperCellValueDetailedType,
  WorkPaperCellValueType,
  WorkPaperConfig,
  WorkPaperSheet,
  WorkPaperSheetDimensions,
} from './work-paper-types.js'

const SCALAR_RANGE_FUNCTION_RE =
  /^(?:_XLFN\.)?(?:_XLWS\.)?(?:SUM|COUNT|COUNTA|COUNTBLANK|MIN|MAX|AVERAGE|AVG|SUMIF|COUNTIF|SUMIFS|COUNTIFS|ABS)\(.*\)(?:[+\-*/]\d+(?:\.\d+)?)?$/
const SIMPLE_SCALAR_EXPRESSION_RE = /^\$?[A-Z]+\$?\d+(?:[+\-*/](?:\$?[A-Z]+\$?\d+|\d+(?:\.\d+)?))*$/
const FORMULA_SPILL_PRODUCING_FUNCTION_MARKERS = FORMULA_SPILL_PRODUCING_FUNCTION_NAMES.map((name) => `${name}(`)

export interface WorkPaperSheetInspection {
  readonly hasFormula: boolean
  readonly hasDynamicSpillFormula: boolean
  readonly dimensions: WorkPaperSheetDimensions
  readonly materializedCellCount: number
  readonly maxColumnCount: number
  readonly formulaCellCount: number
}

export interface WorkPaperRuntimeSnapshotSheetList {
  readonly sheets: readonly { readonly name: string }[]
}

export function compareSheetNames(left: string, right: string): number {
  return left.localeCompare(right)
}

export function inspectSheetDimensionsWithinLimits(
  sheetName: string,
  sheet: WorkPaperSheet,
  config: WorkPaperConfig,
): WorkPaperSheetDimensions {
  const height = sheet.length
  let width = 0
  let materializedHeight = 0
  let materializedWidth = 0
  for (let rowIndex = 0; rowIndex < sheet.length; rowIndex += 1) {
    const row = sheet[rowIndex]
    if (!Array.isArray(row)) {
      throw new WorkPaperUnableToParseError({ sheetName, reason: 'Rows must be arrays' })
    }
    width = Math.max(width, row.length)
    let rowHasMaterializedCell = false
    let lastMaterializedCol = -1
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      if (!isBlankRawCellContent(row[colIndex])) {
        rowHasMaterializedCell = true
        lastMaterializedCol = colIndex
      }
    }
    if (rowHasMaterializedCell) {
      materializedHeight = rowIndex + 1
      if (lastMaterializedCol + 1 > materializedWidth) {
        materializedWidth = lastMaterializedCol + 1
      }
    }
  }
  if (height > (config.maxRows ?? MAX_ROWS) || width > (config.maxColumns ?? MAX_COLS)) {
    throw new WorkPaperSheetSizeLimitExceededError()
  }
  return { width: materializedWidth, height: materializedHeight }
}

export function inspectRuntimeSnapshotSheetDimensionsWithinLimits(args: {
  readonly sheetName: string
  readonly snapshotSheet: WorkbookSnapshot['sheets'][number]
  readonly runtimeSheetCells?: {
    readonly coords?: readonly { readonly row: number; readonly col: number }[]
    readonly dimensions?: { readonly width: number; readonly height: number }
    readonly cellCount?: number
  }
  readonly config: WorkPaperConfig
}): WorkPaperSheetDimensions {
  let materializedHeight = 0
  let materializedWidth = 0
  const dimensions = args.runtimeSheetCells?.dimensions
  if (
    dimensions &&
    Number.isInteger(dimensions.width) &&
    Number.isInteger(dimensions.height) &&
    dimensions.width >= 0 &&
    dimensions.height >= 0
  ) {
    materializedWidth = dimensions.width
    materializedHeight = dimensions.height
  } else if (args.runtimeSheetCells?.coords) {
    for (const coords of args.runtimeSheetCells.coords) {
      materializedHeight = Math.max(materializedHeight, coords.row + 1)
      materializedWidth = Math.max(materializedWidth, coords.col + 1)
    }
  } else {
    for (const cell of args.snapshotSheet.cells) {
      const parsed = parseCellAddress(cell.address, args.sheetName)
      materializedHeight = Math.max(materializedHeight, parsed.row + 1)
      materializedWidth = Math.max(materializedWidth, parsed.col + 1)
    }
  }
  if (materializedHeight > (args.config.maxRows ?? MAX_ROWS) || materializedWidth > (args.config.maxColumns ?? MAX_COLS)) {
    throw new WorkPaperSheetSizeLimitExceededError()
  }
  return { width: materializedWidth, height: materializedHeight }
}

export function inspectSheetWithinLimits(sheetName: string, sheet: WorkPaperSheet, config: WorkPaperConfig): WorkPaperSheetInspection {
  const height = sheet.length
  let width = 0
  let materializedHeight = 0
  let materializedWidth = 0
  let materializedCellCount = 0
  let maxColumnCount = 0
  let formulaCellCount = 0
  let hasFormula = false
  let hasDynamicSpillFormula = false
  for (let rowIndex = 0; rowIndex < sheet.length; rowIndex += 1) {
    const row = sheet[rowIndex]
    if (!Array.isArray(row)) {
      throw new WorkPaperUnableToParseError({ sheetName, reason: 'Rows must be arrays' })
    }
    width = Math.max(width, row.length)
    let rowHasMaterializedCell = false
    let lastMaterializedCol = -1
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const cell = row[colIndex]
      if (!isBlankRawCellContent(cell)) {
        materializedCellCount += 1
        rowHasMaterializedCell = true
        lastMaterializedCol = colIndex
      }
      if (typeof cell === 'string' && cellHasFormulaPrefix(cell)) {
        formulaCellCount += 1
        hasFormula = true
        hasDynamicSpillFormula ||= formulaMayResizeDynamically(cell)
      }
    }
    if (rowHasMaterializedCell) {
      materializedHeight = rowIndex + 1
      if (lastMaterializedCol + 1 > materializedWidth) {
        materializedWidth = lastMaterializedCol + 1
      }
    }
    maxColumnCount = Math.max(maxColumnCount, width)
  }
  if (height > (config.maxRows ?? MAX_ROWS) || width > (config.maxColumns ?? MAX_COLS)) {
    throw new WorkPaperSheetSizeLimitExceededError()
  }
  return {
    hasFormula,
    hasDynamicSpillFormula,
    dimensions: { width: materializedWidth, height: materializedHeight },
    materializedCellCount,
    maxColumnCount,
    formulaCellCount,
  }
}

export function workPaperSheetHasDynamicSpillFormula(sheet: WorkPaperSheet): boolean {
  for (let rowIndex = 0; rowIndex < sheet.length; rowIndex += 1) {
    const row = sheet[rowIndex]
    if (!Array.isArray(row)) {
      continue
    }
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const cell = row[colIndex]
      if (typeof cell === 'string' && cellHasFormulaPrefix(cell) && formulaMayResizeDynamically(cell)) {
        return true
      }
    }
  }
  return false
}

export function workbookSnapshotSheetHasDynamicSpillFormula(snapshotSheet: WorkbookSnapshot['sheets'][number]): boolean {
  return snapshotSheet.cells.some((cell) => typeof cell.formula === 'string' && formulaMayResizeDynamically(cell.formula))
}

export function cellHasFormulaPrefix(value: string): boolean {
  const first = value.charCodeAt(0)
  if (first === 61) {
    return true
  }
  if (first !== 32 && first !== 9 && first !== 10 && first !== 13) {
    return false
  }
  return value.trimStart().charCodeAt(0) === 61
}

function formulaMayResizeDynamically(value: string): boolean {
  const formula = stripFormulaPrefix(value)
  if (isDefinitelyScalarFormulaShape(formula)) {
    return false
  }
  try {
    return compileFormula(formula).producesSpill
  } catch {
    return true
  }
}

function isDefinitelyScalarFormulaShape(formula: string): boolean {
  const normalized = normalizeScalarFormulaShape(formula)
  if (normalized.length === 0 || normalized.includes('{') || normalized.includes('#')) {
    return false
  }
  if (SIMPLE_SCALAR_EXPRESSION_RE.test(normalized)) {
    return true
  }
  if (normalized.includes('(') && FORMULA_SPILL_PRODUCING_FUNCTION_MARKERS.some((marker) => normalized.includes(marker))) {
    return false
  }
  return SCALAR_RANGE_FUNCTION_RE.test(normalized)
}

function normalizeScalarFormulaShape(formula: string): string {
  for (let index = 0; index < formula.length; index += 1) {
    const charCode = formula.charCodeAt(index)
    if ((charCode >= 97 && charCode <= 122) || isScalarFormulaWhitespace(charCode)) {
      return formula.replace(/\s+/g, '').toUpperCase()
    }
  }
  return formula
}

function isScalarFormulaWhitespace(charCode: number): boolean {
  return charCode === 32 || charCode === 9 || charCode === 10 || charCode === 11 || charCode === 12 || charCode === 13
}

function stripFormulaPrefix(value: string): string {
  const trimmed = value.trimStart()
  return trimmed.startsWith('=') ? trimmed.slice(1) : trimmed
}

export function classifyWorkPaperCell(input: {
  readonly hasFormula: boolean
  readonly isEmpty: boolean
  readonly isPartOfArray: boolean
}): WorkPaperCellType {
  if (input.isEmpty) {
    return 'EMPTY'
  }
  if (input.isPartOfArray) {
    return 'ARRAY'
  }
  return input.hasFormula ? 'FORMULA' : 'VALUE'
}

export function doesWorkPaperCellHaveSimpleValue(input: { readonly hasFormula: boolean; readonly isEmpty: boolean }): boolean {
  return !input.hasFormula && !input.isEmpty
}

export function workPaperCellValueType(value: CellValue): WorkPaperCellValueType {
  switch (value.tag) {
    case ValueTag.Number:
      return 'NUMBER'
    case ValueTag.String:
      return 'STRING'
    case ValueTag.Boolean:
      return 'BOOLEAN'
    case ValueTag.Error:
      return 'ERROR'
    case ValueTag.Empty:
    default:
      return 'EMPTY'
  }
}

export function workPaperCellValueDetailedType(input: {
  readonly value: CellValue
  readonly format?: string | undefined
}): WorkPaperCellValueDetailedType {
  const type = workPaperCellValueType(input.value)
  if (type !== 'NUMBER') {
    return type
  }
  const format = input.format?.toLowerCase() ?? ''
  if (format.includes('yy') || format.includes('dd')) {
    if (format.includes('h') || format.includes('s')) {
      return 'DATETIME'
    }
    return 'DATE'
  }
  if (format.includes('h') || format.includes('s')) {
    return 'TIME'
  }
  return type
}

export interface WorkPaperSpillRangeLike {
  readonly sheetName: string
  readonly address: string
  readonly rows: number
  readonly cols: number
}

export function workPaperCellIsInsideSpillRange(input: {
  readonly address: WorkPaperCellAddress
  readonly spill: WorkPaperSpillRangeLike
  readonly requireSheetId: (sheetName: string) => number
}): boolean {
  if (input.requireSheetId(input.spill.sheetName) !== input.address.sheet) {
    return false
  }
  const owner = parseCellAddress(input.spill.address, input.spill.sheetName)
  return (
    input.address.row >= owner.row &&
    input.address.row < owner.row + input.spill.rows &&
    input.address.col >= owner.col &&
    input.address.col < owner.col + input.spill.cols
  )
}

export function isWorkPaperCellPartOfArray(input: {
  readonly address: WorkPaperCellAddress
  readonly spillRanges: readonly WorkPaperSpillRangeLike[]
  readonly requireSheetId: (sheetName: string) => number
}): boolean {
  return input.spillRanges.some((spill) =>
    workPaperCellIsInsideSpillRange({
      address: input.address,
      spill,
      requireSheetId: input.requireSheetId,
    }),
  )
}

export function runtimeSnapshotMatchesSheetEntries(
  sheetEntries: readonly (readonly [string, WorkPaperSheet])[],
  runtimeSnapshot: WorkPaperRuntimeSnapshotSheetList,
): boolean {
  if (runtimeSnapshot.sheets.length !== sheetEntries.length) {
    return false
  }
  const matchedNames = new Set<string>()
  const sheetNames = new Set(sheetEntries.map(([sheetName]) => sheetName))
  for (const snapshotSheet of runtimeSnapshot.sheets) {
    if (!sheetNames.has(snapshotSheet.name) || matchedNames.has(snapshotSheet.name)) {
      return false
    }
    matchedNames.add(snapshotSheet.name)
  }
  return true
}

export function validateSheetWithinLimits(sheetName: string, sheet: WorkPaperSheet, config: WorkPaperConfig): void {
  inspectSheetWithinLimits(sheetName, sheet, config)
}
