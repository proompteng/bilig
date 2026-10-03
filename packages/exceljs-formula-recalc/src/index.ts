import { recalculateXlsx, type XlsxFormulaRecalcOptions, type XlsxFormulaRecalcResult } from '@bilig/xlsx-formula-recalc'

export { recalculateXlsx, recalculateXlsxFileToFile } from '@bilig/xlsx-formula-recalc'
export { WorkPaper, exportXlsx, importXlsx, parseQualifiedCellTarget } from '@bilig/workpaper/xlsx'
export type {
  XlsxFormulaRecalcCellValue,
  XlsxFormulaRecalcEdit,
  XlsxFormulaRecalcOptions,
  XlsxFormulaRecalcResult,
} from '@bilig/xlsx-formula-recalc'

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
  const result = await recalculateXlsx(toUint8Array(input), recalcOptions)

  if (mutateWorkbook) {
    await workbook.xlsx.load(result.xlsx)
    patchExceljsReadResults(workbook, result.reads)
  }

  return {
    ...result,
    workbookMutated: mutateWorkbook,
  }
}

export async function recalculateExceljsBuffer(
  input: Uint8Array | ArrayBuffer | Buffer,
  options: XlsxFormulaRecalcOptions = {},
): Promise<XlsxFormulaRecalcResult> {
  return await recalculateXlsx(input, options)
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

export function parseQualifiedA1(target: string): { readonly sheetName: string; readonly row: number; readonly col: number } {
  const trimmed = target.trim()
  const separator = findSheetSeparator(trimmed)
  if (separator <= 0 || separator >= trimmed.length - 1) {
    throw new Error(`Expected a sheet-qualified A1 target such as Inputs!B2, received: ${target}`)
  }
  const a1 = trimmed
    .slice(separator + 1)
    .replace(/\$/gu, '')
    .toUpperCase()
  const match = /^(?<col>[A-Z]+)(?<row>[1-9][0-9]*)$/u.exec(a1)
  const row = match?.groups?.['row']
  const col = match?.groups?.['col']
  if (!row || !col) {
    throw new Error(`Expected a single A1 cell reference in target ${target}`)
  }
  return {
    sheetName: unquoteSheetName(trimmed.slice(0, separator)),
    row: Number.parseInt(row, 10) - 1,
    col: columnLettersToIndex(col),
  }
}

function findSheetSeparator(target: string): number {
  let inQuote = false
  for (let index = 0; index < target.length; index += 1) {
    const char = target[index]
    if (char === "'") {
      if (inQuote && target[index + 1] === "'") {
        index += 1
      } else {
        inQuote = !inQuote
      }
      continue
    }
    if (char === '!' && !inQuote) {
      return index
    }
  }
  return -1
}

function unquoteSheetName(rawSheetName: string): string {
  const trimmed = rawSheetName.trim()
  return trimmed.startsWith("'") && trimmed.endsWith("'") ? trimmed.slice(1, -1).replace(/''/gu, "'") : trimmed
}

function columnLettersToIndex(letters: string): number {
  let index = 0
  for (const char of letters) {
    index = index * 26 + (char.charCodeAt(0) - 64)
  }
  return index - 1
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
