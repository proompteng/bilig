import { WorkPaper, type RawCellContent, type WorkPaperCellAddress, type WorkPaperChange, type WorkPaperConfig } from '@bilig/headless'
import { exportXlsx, importXlsx } from '@bilig/headless/xlsx'

export { WorkPaper } from '@bilig/headless'
export { exportXlsx, importXlsx } from '@bilig/headless/xlsx'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSnapshot>

export type XlsxFormulaRecalcCellValue = ReturnType<WorkPaperInstance['getCellValue']>

export interface XlsxFormulaRecalcEdit {
  readonly target: string
  readonly value: RawCellContent
}

export interface XlsxFormulaRecalcOptions {
  readonly fileName?: string
  readonly edits?: readonly XlsxFormulaRecalcEdit[]
  readonly reads?: readonly string[]
  readonly config?: WorkPaperConfig
}

export interface XlsxFormulaRecalcResult {
  readonly xlsx: Uint8Array
  readonly warnings: readonly string[]
  readonly sheetNames: readonly string[]
  readonly reads: Readonly<Record<string, XlsxFormulaRecalcCellValue>>
  readonly changes: readonly WorkPaperChange[]
}

export function recalculateXlsx(input: Uint8Array | ArrayBuffer | Buffer, options: XlsxFormulaRecalcOptions = {}): XlsxFormulaRecalcResult {
  const imported = importXlsx(toUint8Array(input), options.fileName ?? 'workbook.xlsx')
  const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
    evaluationTimeoutMs: 30_000,
    useColumnIndex: true,
    ...options.config,
  })

  try {
    const changes: WorkPaperChange[] = []
    for (const edit of options.edits ?? []) {
      changes.push(...workbook.setCellContents(parseQualifiedCellTarget(workbook, edit.target), edit.value))
    }

    const reads: Record<string, XlsxFormulaRecalcCellValue> = {}
    for (const target of options.reads ?? []) {
      reads[target] = workbook.getCellValue(parseQualifiedCellTarget(workbook, target))
    }

    return {
      xlsx: toUint8Array(exportXlsx(workbook.exportSnapshot())),
      warnings: imported.warnings,
      sheetNames: imported.sheetNames,
      reads,
      changes,
    }
  } finally {
    workbook.dispose()
  }
}

export const recalculateSheetjsWorkbook = recalculateXlsx

export function parseQualifiedCellTarget(workbook: WorkPaperInstance, target: string): WorkPaperCellAddress {
  const parsed = parseQualifiedA1(target)
  const sheet = workbook.getSheetId(parsed.sheetName)
  if (sheet === undefined) {
    throw new Error(`Unknown sheet in XLSX formula recalculation target: ${parsed.sheetName}`)
  }
  return {
    sheet,
    row: parsed.row,
    col: parsed.col,
  }
}

export function parseQualifiedA1(target: string): { sheetName: string; row: number; col: number } {
  const trimmed = target.trim()
  const separator = findSheetSeparator(trimmed)
  if (separator <= 0 || separator >= trimmed.length - 1) {
    throw new Error(`Expected a sheet-qualified A1 target such as Inputs!B2, received: ${target}`)
  }

  const sheetName = unquoteSheetName(trimmed.slice(0, separator))
  const a1 = trimmed
    .slice(separator + 1)
    .replace(/\$/gu, '')
    .toUpperCase()
  const match = /^(?<col>[A-Z]+)(?<row>[1-9][0-9]*)$/u.exec(a1)
  if (!match?.groups) {
    throw new Error(`Expected a single A1 cell reference in target ${target}`)
  }

  const row = match.groups['row']
  const col = match.groups['col']
  if (!row || !col) {
    throw new Error(`Expected a single A1 cell reference in target ${target}`)
  }

  return {
    sheetName,
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
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/gu, "'")
  }
  return trimmed
}

function columnLettersToIndex(letters: string): number {
  let index = 0
  for (const char of letters) {
    index = index * 26 + (char.charCodeAt(0) - 64)
  }
  return index - 1
}

function toUint8Array(input: Uint8Array | ArrayBuffer | Buffer): Uint8Array {
  if (input instanceof Uint8Array) {
    return input
  }
  return new Uint8Array(input)
}
