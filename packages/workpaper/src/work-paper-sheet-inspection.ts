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
const SCALAR_CELL_REF_PATTERN = String.raw`(?:'(?:[^']|'')+'!|[A-Z_][A-Z0-9_ .]*!)?\$?[A-Z]+\$?\d+`
const SCALAR_NUMBER_PATTERN = String.raw`\d+(?:\.\d+)?`
const SIMPLE_SCALAR_EXPRESSION_RE = new RegExp(
  `^(?:${SCALAR_CELL_REF_PATTERN}|${SCALAR_NUMBER_PATTERN})(?:[+\\-*/](?:${SCALAR_CELL_REF_PATTERN}|${SCALAR_NUMBER_PATTERN}))*$`,
)
const SCALAR_ONLY_FUNCTION_NAMES = ['CONCATENATE', 'IF', 'LEN', 'MAX', 'MIN', 'PMT', 'POWER', 'ROUND', 'SQRT'] as const
const SCALAR_ONLY_FUNCTION_MARKERS = SCALAR_ONLY_FUNCTION_NAMES.map((name) => `${name}(`)
const SCALAR_RANGE_FUNCTION_NAMES = new Set([
  'SUM',
  'COUNT',
  'COUNTA',
  'COUNTBLANK',
  'MIN',
  'MAX',
  'AVERAGE',
  'AVG',
  'SUMIF',
  'COUNTIF',
  'SUMIFS',
  'COUNTIFS',
  'ABS',
])
const FORMULA_SPILL_PRODUCING_FUNCTION_MARKERS = FORMULA_SPILL_PRODUCING_FUNCTION_NAMES.map((name) => `${name}(`)

export interface WorkPaperSheetInspection {
  readonly hasFormula: boolean
  readonly hasDynamicSpillFormula: boolean
  readonly dimensions: WorkPaperSheetDimensions
  readonly materializedCellCount: number
  readonly maxColumnCount: number
  readonly formulaCellCount: number
  readonly allMaterializedCellsAreNumbers: boolean
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
      const row = cell.row
      const col = cell.col
      if (typeof row === 'number' && typeof col === 'number' && Number.isInteger(row) && Number.isInteger(col) && row >= 0 && col >= 0) {
        materializedHeight = Math.max(materializedHeight, row + 1)
        materializedWidth = Math.max(materializedWidth, col + 1)
      } else {
        const parsed = parseCellAddress(cell.address, args.sheetName)
        materializedHeight = Math.max(materializedHeight, parsed.row + 1)
        materializedWidth = Math.max(materializedWidth, parsed.col + 1)
      }
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
  let allMaterializedCellsAreNumbers = true
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
        allMaterializedCellsAreNumbers &&= typeof cell === 'number'
      }
      if (typeof cell === 'string' && cellHasFormulaPrefix(cell)) {
        formulaCellCount += 1
        hasFormula = true
        hasDynamicSpillFormula ||= workPaperFormulaMayResizeDynamically(cell)
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
    allMaterializedCellsAreNumbers,
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
      if (typeof cell === 'string' && cellHasFormulaPrefix(cell) && workPaperFormulaMayResizeDynamically(cell)) {
        return true
      }
    }
  }
  return false
}

export function workbookSnapshotSheetHasDynamicSpillFormula(snapshotSheet: WorkbookSnapshot['sheets'][number]): boolean {
  return snapshotSheet.cells.some((cell) => typeof cell.formula === 'string' && workPaperFormulaMayResizeDynamically(cell.formula))
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

export function workPaperFormulaMayResizeDynamically(value: string): boolean {
  const formulaStart = formulaBodyStart(value)
  if (formulaStart >= 0 && isSimpleSameSheetScalarBinarySource(value, formulaStart, value.length)) {
    return false
  }
  if (formulaStart >= 0 && isSimpleScalarExpressionSource(value, formulaStart, value.length)) {
    return false
  }
  const formula = formulaStart < 0 ? stripFormulaPrefix(value) : value.slice(formulaStart)
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
  if (isSimpleScalarExpressionShape(normalized) || isScalarRangeFunctionShape(normalized)) {
    return true
  }
  if (SIMPLE_SCALAR_EXPRESSION_RE.test(normalized)) {
    return true
  }
  if (normalized.includes('(') && FORMULA_SPILL_PRODUCING_FUNCTION_MARKERS.some((marker) => normalized.includes(marker))) {
    return false
  }
  if (SCALAR_RANGE_FUNCTION_RE.test(normalized)) {
    return true
  }
  return !normalized.includes(':') && SCALAR_ONLY_FUNCTION_MARKERS.some((marker) => normalized.includes(marker))
}

function isSimpleSameSheetScalarBinarySource(value: string, start: number, end: number): boolean {
  const leftEnd = readSimpleSameSheetCellRefSource(value, start, end)
  if (leftEnd <= start || leftEnd >= end) {
    return false
  }
  const operator = value.charCodeAt(leftEnd)
  if (operator !== 43 && operator !== 45 && operator !== 42 && operator !== 47) {
    return false
  }
  const rightStart = leftEnd + 1
  const rightCellEnd = readSimpleSameSheetCellRefSource(value, rightStart, end)
  if (rightCellEnd === end) {
    return true
  }
  const rightNumberEnd = readSimpleScalarNumberSource(value, rightStart, end)
  return rightNumberEnd === end
}

function readSimpleSameSheetCellRefSource(value: string, start: number, end: number): number {
  let index = start
  const columnStart = index
  while (index < end && isAsciiAlpha(value.charCodeAt(index))) {
    index += 1
  }
  if (index === columnStart) {
    return start
  }
  const rowStart = index
  while (index < end && isAsciiDigit(value.charCodeAt(index))) {
    index += 1
  }
  return index === rowStart ? start : index
}

function isSimpleScalarExpressionShape(formula: string): boolean {
  let index = readSimpleScalarOperand(formula, 0)
  if (index <= 0) {
    return false
  }
  while (index < formula.length) {
    const operator = formula.charCodeAt(index)
    if (operator !== 43 && operator !== 45 && operator !== 42 && operator !== 47) {
      return false
    }
    index = readSimpleScalarOperand(formula, index + 1)
    if (index <= 0) {
      return false
    }
  }
  return true
}

function formulaBodyStart(value: string): number {
  let index = 0
  while (index < value.length && isScalarFormulaWhitespace(value.charCodeAt(index))) {
    index += 1
  }
  return value.charCodeAt(index) === 61 ? index + 1 : -1
}

function isSimpleScalarExpressionSource(value: string, start: number, end: number): boolean {
  let index = readSimpleScalarOperandSource(value, start, end)
  if (index <= start) {
    return false
  }
  while (index < end) {
    const operator = value.charCodeAt(index)
    if (operator !== 43 && operator !== 45 && operator !== 42 && operator !== 47) {
      return false
    }
    const next = readSimpleScalarOperandSource(value, index + 1, end)
    if (next <= index + 1) {
      return false
    }
    index = next
  }
  return true
}

function readSimpleScalarOperandSource(value: string, start: number, end: number): number {
  const numberEnd = readSimpleScalarNumberSource(value, start, end)
  if (numberEnd > start) {
    return numberEnd
  }
  return readSimpleScalarCellRefSource(value, start, end)
}

function readSimpleScalarNumberSource(value: string, start: number, end: number): number {
  let index = start
  while (index < end && isAsciiDigit(value.charCodeAt(index))) {
    index += 1
  }
  if (index < end && value.charCodeAt(index) === 46) {
    const fractionStart = index + 1
    index = fractionStart
    while (index < end && isAsciiDigit(value.charCodeAt(index))) {
      index += 1
    }
    return index === fractionStart ? start : index
  }
  return index
}

function readSimpleScalarCellRefSource(value: string, start: number, end: number): number {
  let index = skipSimpleSheetQualifier(value, start, end)
  if (value.charCodeAt(index) === 36) {
    index += 1
  }
  const columnStart = index
  while (index < end && isAsciiAlpha(value.charCodeAt(index))) {
    index += 1
  }
  if (index === columnStart) {
    return start
  }
  if (value.charCodeAt(index) === 36) {
    index += 1
  }
  const rowStart = index
  while (index < end && isAsciiDigit(value.charCodeAt(index))) {
    index += 1
  }
  return index === rowStart ? start : index
}

function skipSimpleSheetQualifier(value: string, start: number, end: number): number {
  if (value.charCodeAt(start) === 39) {
    let index = start + 1
    while (index < end) {
      const charCode = value.charCodeAt(index)
      if (charCode === 39) {
        if (value.charCodeAt(index + 1) === 39) {
          index += 2
          continue
        }
        return value.charCodeAt(index + 1) === 33 ? index + 2 : start
      }
      index += 1
    }
    return start
  }

  let index = start
  while (index < end) {
    const charCode = value.charCodeAt(index)
    if (charCode === 33) {
      return index > start ? index + 1 : start
    }
    if (
      charCode === 36 ||
      charCode === 43 ||
      charCode === 45 ||
      charCode === 42 ||
      charCode === 47 ||
      charCode === 40 ||
      charCode === 41 ||
      charCode === 44 ||
      charCode === 58
    ) {
      return start
    }
    index += 1
  }
  return start
}

function readSimpleScalarOperand(formula: string, start: number): number {
  const numberEnd = readSimpleScalarNumber(formula, start)
  if (numberEnd > start) {
    return numberEnd
  }
  return readSimpleScalarCellRef(formula, start)
}

function readSimpleScalarNumber(formula: string, start: number): number {
  let index = start
  while (index < formula.length && isAsciiDigit(formula.charCodeAt(index))) {
    index += 1
  }
  if (index < formula.length && formula.charCodeAt(index) === 46) {
    const fractionStart = index + 1
    index = fractionStart
    while (index < formula.length && isAsciiDigit(formula.charCodeAt(index))) {
      index += 1
    }
    return index === fractionStart ? start : index
  }
  return index
}

function readSimpleScalarCellRef(formula: string, start: number): number {
  let index = start
  if (formula.charCodeAt(index) === 36) {
    index += 1
  }
  const columnStart = index
  while (index < formula.length && isAsciiUpperAlpha(formula.charCodeAt(index))) {
    index += 1
  }
  if (index === columnStart) {
    return start
  }
  if (formula.charCodeAt(index) === 36) {
    index += 1
  }
  const rowStart = index
  while (index < formula.length && isAsciiDigit(formula.charCodeAt(index))) {
    index += 1
  }
  return index === rowStart ? start : index
}

function isScalarRangeFunctionShape(formula: string): boolean {
  const openIndex = formula.indexOf('(')
  if (openIndex <= 0) {
    return false
  }
  const callee = stripExcelFunctionPrefix(formula.slice(0, openIndex))
  if (!SCALAR_RANGE_FUNCTION_NAMES.has(callee)) {
    return false
  }
  const closeIndex = formula.lastIndexOf(')')
  if (closeIndex <= openIndex) {
    return false
  }
  return closeIndex === formula.length - 1 || isScalarNumericOffset(formula, closeIndex + 1)
}

function stripExcelFunctionPrefix(callee: string): string {
  let result = callee
  if (result.startsWith('_XLFN.')) {
    result = result.slice(6)
  }
  if (result.startsWith('_XLWS.')) {
    result = result.slice(6)
  }
  return result
}

function isScalarNumericOffset(formula: string, start: number): boolean {
  const operator = formula.charCodeAt(start)
  if (operator !== 43 && operator !== 45 && operator !== 42 && operator !== 47) {
    return false
  }
  const end = readSimpleScalarNumber(formula, start + 1)
  return end === formula.length
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

function isAsciiDigit(charCode: number): boolean {
  return charCode >= 48 && charCode <= 57
}

function isAsciiUpperAlpha(charCode: number): boolean {
  return charCode >= 65 && charCode <= 90
}

function isAsciiAlpha(charCode: number): boolean {
  return (charCode >= 65 && charCode <= 90) || (charCode >= 97 && charCode <= 122)
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
