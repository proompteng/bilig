import { parseQualifiedA1, recalculateXlsx, type XlsxFormulaRecalcOptions, type XlsxFormulaRecalcResult } from 'xlsx-formula-recalc'

export { WorkPaper, exportXlsx, importXlsx, parseQualifiedA1, parseQualifiedCellTarget, recalculateXlsx } from 'xlsx-formula-recalc'
export type {
  XlsxFormulaRecalcCellValue,
  XlsxFormulaRecalcEdit,
  XlsxFormulaRecalcOptions,
  XlsxFormulaRecalcResult,
} from 'xlsx-formula-recalc'

export interface ExceljsWorkbookLike {
  readonly xlsx: {
    writeBuffer(): Promise<ArrayBuffer | Buffer | Uint8Array>
    load(input: ArrayBuffer | Buffer | Uint8Array): Promise<unknown>
  }
  getWorksheet?(name: string): ExceljsWorksheetLike | undefined
}

export interface ExceljsWorksheetLike {
  getCell(address: string): ExceljsCellLike
}

export interface ExceljsCellLike {
  value: unknown
}

export interface ExceljsFormulaRecalcOptions extends XlsxFormulaRecalcOptions {
  readonly mutateWorkbook?: boolean
}

export interface ExceljsFormulaRecalcResult extends XlsxFormulaRecalcResult {
  readonly workbookMutated: boolean
}

export async function recalculateExceljsWorkbook(
  workbook: ExceljsWorkbookLike,
  options: ExceljsFormulaRecalcOptions = {},
): Promise<ExceljsFormulaRecalcResult> {
  const { mutateWorkbook = true, ...recalcOptions } = options
  const input = await workbook.xlsx.writeBuffer()
  const result = recalculateXlsx(toUint8Array(input), recalcOptions)

  if (mutateWorkbook) {
    await workbook.xlsx.load(result.xlsx)
    patchExceljsReadResults(workbook, result.reads)
  }

  return {
    ...result,
    workbookMutated: mutateWorkbook,
  }
}

export function recalculateExceljsBuffer(
  input: Uint8Array | ArrayBuffer | Buffer,
  options: XlsxFormulaRecalcOptions = {},
): XlsxFormulaRecalcResult {
  return recalculateXlsx(input, options)
}

function toUint8Array(input: Uint8Array | ArrayBuffer | Buffer): Uint8Array {
  if (input instanceof Uint8Array) {
    return input
  }
  return new Uint8Array(input)
}

function patchExceljsReadResults(workbook: ExceljsWorkbookLike, reads: XlsxFormulaRecalcResult['reads']): void {
  if (!workbook.getWorksheet) {
    return
  }

  for (const [target, value] of Object.entries(reads)) {
    const parsed = parseQualifiedA1(target)
    const worksheet = workbook.getWorksheet(parsed.sheetName)
    if (!worksheet) {
      continue
    }
    const cell = worksheet.getCell(`${columnIndexToLetters(parsed.col)}${parsed.row + 1}`)
    const readValue = unwrapReadValue(value)
    if (readValue === undefined) {
      continue
    }
    if (isExceljsFormulaCellValue(cell.value)) {
      cell.value = {
        ...cell.value,
        result: readValue,
      }
    } else {
      cell.value = readValue
    }
  }
}

function unwrapReadValue(value: XlsxFormulaRecalcResult['reads'][string]): unknown {
  if (typeof value === 'object' && value !== null && 'value' in value) {
    return value.value
  }
  return undefined
}

function isExceljsFormulaCellValue(value: unknown): value is { formula: string; result?: unknown } {
  return typeof value === 'object' && value !== null && 'formula' in value && typeof value.formula === 'string'
}

function columnIndexToLetters(columnIndex: number): string {
  let value = columnIndex + 1
  let letters = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    letters = String.fromCharCode(65 + remainder) + letters
    value = Math.floor((value - 1) / 26)
  }
  return letters
}
