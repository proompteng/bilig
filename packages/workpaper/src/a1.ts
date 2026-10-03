import { WorkPaper } from './work-paper.js'
import { createWorkPaperFromDocument, exportWorkPaperDocument, parseWorkPaperDocument, serializeWorkPaperDocument } from './persistence.js'
import type { PersistedWorkPaperDocument } from './persistence.js'
import type {
  RawCellContent,
  WorkPaperCellAddress,
  WorkPaperCellRange,
  WorkPaperChange,
  WorkPaperConfig,
  WorkPaperFormulaDiagnostic,
  WorkPaperSheets,
} from './work-paper.js'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

export type A1CellInput = RawCellContent
export type A1CellValue = ReturnType<WorkPaperInstance['getCellValue']>

export interface A1WorkPaperOptions {
  readonly defaultSheetName?: string
  readonly writableSheets?: readonly string[]
  readonly limitations?: readonly string[]
}

export interface A1CellRead {
  readonly address: string
  readonly value: A1CellValue
  readonly serialized: RawCellContent
  readonly formula: string | null
  readonly displayValue: string
  readonly formulaDiagnostics: readonly WorkPaperFormulaDiagnostic[]
}

export interface A1RangeRead {
  readonly range: string
  readonly values: A1CellValue[][]
  readonly serialized: RawCellContent[][]
  readonly displayValues: string[][]
  readonly formulaDiagnostics: readonly WorkPaperFormulaDiagnostic[]
  readonly cells: readonly (readonly A1CellRead[])[]
}

export interface A1SetCellAndReadbackOptions {
  readonly readbackRange: string
  readonly requireReadbackChange?: boolean
  readonly includeConfig?: boolean
  readonly includeSerializedDocument?: boolean
  readonly limitations?: readonly string[]
}

export interface A1SetCellAndReadbackProof {
  readonly editedCell: string
  readonly readbackRange: string
  readonly before: A1CellRead
  readonly after: A1CellRead
  readonly beforeReadback: A1RangeRead
  readonly afterReadback: A1RangeRead
  readonly restoredReadback: A1RangeRead
  readonly persistedDocumentBytes: number
  readonly checks: {
    readonly readbackChanged: boolean
    readonly computedReadbackChanged: boolean
    readonly editedFormulaReadbackChanged: boolean
    readonly readbackIncludesEditedCell: boolean
    readonly readbackContainsOnlyEditedCell: boolean
    readonly restoredReadbackMatchesAfter: boolean
    readonly blockingFormulaDiagnosticCount: number
    readonly formulaDiagnostics: readonly WorkPaperFormulaDiagnostic[]
    readonly previousSerialized: RawCellContent
    readonly newSerialized: RawCellContent
  }
  readonly verified: boolean
  readonly limitations: readonly string[]
  readonly changes: readonly WorkPaperChange[]
  readonly serializedDocument?: string
}

export interface A1EditManyAndReadbackProof {
  readonly editedCells: readonly string[]
  readonly readbackRange: string
  readonly before: Readonly<Record<string, A1CellRead>>
  readonly after: Readonly<Record<string, A1CellRead>>
  readonly beforeReadback: A1RangeRead
  readonly afterReadback: A1RangeRead
  readonly restoredReadback: A1RangeRead
  readonly persistedDocumentBytes: number
  readonly checks: {
    readonly readbackChanged: boolean
    readonly computedReadbackChanged: boolean
    readonly editedFormulaReadbackChanged: boolean
    readonly readbackIncludesEditedCells: boolean
    readonly readbackContainsOnlyEditedCells: boolean
    readonly restoredReadbackMatchesAfter: boolean
    readonly blockingFormulaDiagnosticCount: number
    readonly formulaDiagnostics: readonly WorkPaperFormulaDiagnostic[]
    readonly previousSerialized: Readonly<Record<string, RawCellContent>>
    readonly newSerialized: Readonly<Record<string, RawCellContent>>
  }
  readonly verified: boolean
  readonly limitations: readonly string[]
  readonly changes: readonly WorkPaperChange[]
  readonly serializedDocument?: string
}

export interface A1WorkPaper {
  readonly workpaper: WorkPaperInstance
  readCell(address: string): A1CellRead
  get(address: string): A1CellValue
  display(address: string): string
  formula(address: string): string | null
  readRange(range: string): A1RangeRead
  range(range: string): A1CellValue[][]
  setCell(address: string, value: A1CellInput): WorkPaperChange[]
  set(address: string, value: A1CellInput): WorkPaperChange[]
  readMany(addresses: readonly string[]): Readonly<Record<string, A1CellRead>>
  setCells(edits: Readonly<Record<string, A1CellInput>>): WorkPaperChange[]
  setMany(edits: Readonly<Record<string, A1CellInput>>): WorkPaperChange[]
  validateFormula(formula: string): boolean
  setCellAndReadback(address: string, value: A1CellInput, options: A1SetCellAndReadbackOptions): A1SetCellAndReadbackProof
  editAndReadback(address: string, value: A1CellInput, options: A1SetCellAndReadbackOptions): A1SetCellAndReadbackProof
  editManyAndReadback(edits: Readonly<Record<string, A1CellInput>>, options: A1SetCellAndReadbackOptions): A1EditManyAndReadbackProof
  exportDocument(options?: { readonly includeConfig?: boolean }): PersistedWorkPaperDocument
  serialize(options?: { readonly includeConfig?: boolean }): string
  saveJson(options?: { readonly includeConfig?: boolean }): string
  restoreJson(input: string | PersistedWorkPaperDocument): A1WorkPaper
  dispose(): void
}

export function createA1WorkPaper(workpaper: WorkPaperInstance, options: A1WorkPaperOptions = {}): A1WorkPaper {
  return new A1WorkPaperFacade(workpaper, options)
}

export function buildA1WorkPaper(sheets: WorkPaperSheets, config?: WorkPaperConfig, options: A1WorkPaperOptions = {}): A1WorkPaper {
  return createA1WorkPaper(WorkPaper.buildFromSheets(sheets, config), options)
}

export function restoreA1WorkPaper(input: string | PersistedWorkPaperDocument, options: A1WorkPaperOptions = {}): A1WorkPaper {
  const document = typeof input === 'string' ? parseWorkPaperDocument(input) : input
  return createA1WorkPaper(createWorkPaperFromDocument(document), options)
}

class A1WorkPaperFacade implements A1WorkPaper {
  readonly workpaper: WorkPaperInstance
  private readonly defaultSheetName: string | undefined
  private readonly writableSheets: readonly string[] | undefined
  private readonly limitations: readonly string[]

  constructor(workpaper: WorkPaperInstance, options: A1WorkPaperOptions) {
    this.workpaper = workpaper
    this.defaultSheetName = options.defaultSheetName
    this.writableSheets = options.writableSheets
    this.limitations = options.limitations ?? []
  }

  readCell(address: string): A1CellRead {
    const parsed = this.requireCellAddress(address)
    return this.readParsedCell(parsed)
  }

  get(address: string): A1CellValue {
    return this.readCell(address).value
  }

  display(address: string): string {
    return this.readCell(address).displayValue
  }

  formula(address: string): string | null {
    return this.readCell(address).formula
  }

  readRange(range: string): A1RangeRead {
    const parsed = this.requireRange(range)
    const cells = this.readRangeCells(parsed)
    return {
      range: this.formatRange(parsed),
      values: cells.map((row) => row.map((cell) => cell.value)),
      serialized: cells.map((row) => row.map((cell) => cell.serialized)),
      displayValues: cells.map((row) => row.map((cell) => cell.displayValue)),
      formulaDiagnostics: cells.flatMap((row) => row.flatMap((cell) => cell.formulaDiagnostics)),
      cells,
    }
  }

  range(range: string): A1CellValue[][] {
    return this.readRange(range).values
  }

  setCell(address: string, value: A1CellInput): WorkPaperChange[] {
    const parsed = this.requireCellAddress(address)
    this.assertWritable(parsed)
    this.assertFormulaCanParse(value, parsed)
    return this.workpaper.setCellContents(parsed, value)
  }

  set(address: string, value: A1CellInput): WorkPaperChange[] {
    return this.setCell(address, value)
  }

  readMany(addresses: readonly string[]): Readonly<Record<string, A1CellRead>> {
    const reads: Record<string, A1CellRead> = {}
    for (const address of addresses) {
      const read = this.readCell(address)
      reads[read.address] = read
    }
    return reads
  }

  setCells(edits: Readonly<Record<string, A1CellInput>>): WorkPaperChange[] {
    const prepared = this.prepareEdits(edits)
    return this.applyPreparedEdits(prepared)
  }

  setMany(edits: Readonly<Record<string, A1CellInput>>): WorkPaperChange[] {
    return this.setCells(edits)
  }

  validateFormula(formula: string): boolean {
    return this.workpaper.validateFormula(formula)
  }

  setCellAndReadback(address: string, value: A1CellInput, options: A1SetCellAndReadbackOptions): A1SetCellAndReadbackProof {
    const proof = this.editManyAndReadback({ [address]: value }, options)
    const editedCell = proof.editedCells[0]
    if (editedCell === undefined) {
      throw new Error('Expected one edited WorkPaper cell')
    }
    const before = proof.before[editedCell]
    const after = proof.after[editedCell]
    if (before === undefined || after === undefined) {
      throw new Error(`Expected proof reads for ${editedCell}`)
    }

    return {
      editedCell,
      readbackRange: proof.readbackRange,
      before,
      after,
      beforeReadback: proof.beforeReadback,
      afterReadback: proof.afterReadback,
      restoredReadback: proof.restoredReadback,
      persistedDocumentBytes: proof.persistedDocumentBytes,
      checks: {
        readbackChanged: proof.checks.readbackChanged,
        computedReadbackChanged: proof.checks.computedReadbackChanged,
        editedFormulaReadbackChanged: proof.checks.editedFormulaReadbackChanged,
        readbackIncludesEditedCell: proof.checks.readbackIncludesEditedCells,
        readbackContainsOnlyEditedCell: proof.checks.readbackContainsOnlyEditedCells,
        restoredReadbackMatchesAfter: proof.checks.restoredReadbackMatchesAfter,
        blockingFormulaDiagnosticCount: proof.checks.blockingFormulaDiagnosticCount,
        formulaDiagnostics: proof.checks.formulaDiagnostics,
        previousSerialized: before.serialized,
        newSerialized: after.serialized,
      },
      verified: proof.verified,
      limitations: proof.limitations,
      changes: proof.changes,
      ...(options.includeSerializedDocument && proof.serializedDocument !== undefined
        ? { serializedDocument: proof.serializedDocument }
        : {}),
    }
  }

  editAndReadback(address: string, value: A1CellInput, options: A1SetCellAndReadbackOptions): A1SetCellAndReadbackProof {
    return this.setCellAndReadback(address, value, options)
  }

  editManyAndReadback(edits: Readonly<Record<string, A1CellInput>>, options: A1SetCellAndReadbackOptions): A1EditManyAndReadbackProof {
    const requireReadbackChange = options.requireReadbackChange ?? true
    const includeConfig = options.includeConfig ?? true
    const prepared = this.prepareEdits(edits)
    const editedCells = prepared.map((edit) => edit.address)
    const editedCellSet = new Set(editedCells)
    const before = this.readPreparedCells(prepared)
    const beforeReadback = this.readRange(options.readbackRange)
    const changes = this.applyPreparedEdits(prepared)
    const after = this.readPreparedCells(prepared)
    const afterReadback = this.readRange(options.readbackRange)
    const serializedDocument = this.serialize({ includeConfig })
    const restored = restoreA1WorkPaper(serializedDocument, this.facadeOptions())

    try {
      const restoredReadback = restored.readRange(options.readbackRange)
      const persistedDocumentBytes = new TextEncoder().encode(serializedDocument).byteLength
      const readbackChanged = !sameJson(rangeComparison(beforeReadback), rangeComparison(afterReadback))
      const computedReadbackChanged = !sameJson(
        rangeComparisonExcludingEditedCells(beforeReadback, editedCellSet),
        rangeComparisonExcludingEditedCells(afterReadback, editedCellSet),
      )
      const editedFormulaReadbackChanged = readbackChanged && readbackEditedFormulaCells(afterReadback, editedCellSet).length > 0
      const readbackAddresses = rangeCellAddresses(afterReadback)
      const readbackIncludesEditedCells = readbackAddresses.some((readbackAddress) => editedCellSet.has(readbackAddress))
      const readbackContainsOnlyEditedCells =
        readbackAddresses.length > 0 && readbackAddresses.every((readbackAddress) => editedCellSet.has(readbackAddress))
      const restoredReadbackMatchesAfter = sameJson(rangeComparison(afterReadback), rangeComparison(restoredReadback))
      const formulaDiagnostics = afterReadback.formulaDiagnostics
      const blockingFormulaDiagnosticCount = formulaDiagnostics.filter((diagnostic) => diagnostic.severity === 'error').length
      const meaningfulReadbackChanged = computedReadbackChanged || editedFormulaReadbackChanged

      return {
        editedCells,
        readbackRange: afterReadback.range,
        before,
        after,
        beforeReadback,
        afterReadback,
        restoredReadback,
        persistedDocumentBytes,
        checks: {
          readbackChanged,
          computedReadbackChanged,
          editedFormulaReadbackChanged,
          readbackIncludesEditedCells,
          readbackContainsOnlyEditedCells,
          restoredReadbackMatchesAfter,
          blockingFormulaDiagnosticCount,
          formulaDiagnostics,
          previousSerialized: serializedByAddress(before),
          newSerialized: serializedByAddress(after),
        },
        verified:
          persistedDocumentBytes > 0 &&
          restoredReadbackMatchesAfter &&
          blockingFormulaDiagnosticCount === 0 &&
          (!requireReadbackChange || meaningfulReadbackChanged),
        limitations: [...this.limitations, ...(options.limitations ?? [])],
        changes,
        ...(options.includeSerializedDocument ? { serializedDocument } : {}),
      }
    } finally {
      restored.dispose()
    }
  }

  exportDocument(options: { readonly includeConfig?: boolean } = {}): PersistedWorkPaperDocument {
    return exportWorkPaperDocument(this.workpaper, options)
  }

  serialize(options: { readonly includeConfig?: boolean } = {}): string {
    return serializeWorkPaperDocument(this.exportDocument(options))
  }

  saveJson(options: { readonly includeConfig?: boolean } = {}): string {
    return this.serialize(options)
  }

  restoreJson(input: string | PersistedWorkPaperDocument): A1WorkPaper {
    return restoreA1WorkPaper(input, this.facadeOptions())
  }

  dispose(): void {
    this.workpaper.dispose()
  }

  private facadeOptions(): A1WorkPaperOptions {
    return {
      ...(this.defaultSheetName !== undefined ? { defaultSheetName: this.defaultSheetName } : {}),
      ...(this.writableSheets !== undefined ? { writableSheets: this.writableSheets } : {}),
      limitations: this.limitations,
    }
  }

  private readParsedCell(address: WorkPaperCellAddress): A1CellRead {
    const formula = this.workpaper.getCellFormula(address)
    return {
      address: this.formatCell(address),
      value: this.workpaper.getCellValue(address),
      serialized: this.workpaper.getCellSerialized(address),
      formula: formula ?? null,
      displayValue: this.workpaper.getCellDisplayValue(address),
      formulaDiagnostics: this.workpaper.getCellFormulaDiagnostics(address),
    }
  }

  private readRangeCells(range: WorkPaperCellRange): A1CellRead[][] {
    const rows: A1CellRead[][] = []
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      const cells: A1CellRead[] = []
      for (let col = range.start.col; col <= range.end.col; col += 1) {
        cells.push(this.readParsedCell({ sheet: range.start.sheet, row, col }))
      }
      rows.push(cells)
    }
    return rows
  }

  private prepareEdits(edits: Readonly<Record<string, A1CellInput>>): readonly PreparedA1Edit[] {
    const prepared: PreparedA1Edit[] = []
    const seen = new Set<string>()
    for (const [address, value] of Object.entries(edits)) {
      const parsed = this.requireCellAddress(address)
      const formatted = this.formatCell(parsed)
      if (seen.has(formatted)) {
        throw new Error(`Duplicate WorkPaper edit target: ${formatted}`)
      }
      this.assertWritable(parsed)
      this.assertFormulaCanParse(value, parsed)
      seen.add(formatted)
      prepared.push({ address: formatted, parsed, value })
    }
    if (prepared.length === 0) {
      throw new Error('At least one WorkPaper edit is required')
    }
    return prepared
  }

  private applyPreparedEdits(edits: readonly PreparedA1Edit[]): WorkPaperChange[] {
    return this.workpaper.transaction(() => {
      for (const edit of edits) {
        this.workpaper.setCellContents(edit.parsed, edit.value)
      }
    })
  }

  private readPreparedCells(edits: readonly PreparedA1Edit[]): Readonly<Record<string, A1CellRead>> {
    const reads: Record<string, A1CellRead> = {}
    for (const edit of edits) {
      reads[edit.address] = this.readParsedCell(edit.parsed)
    }
    return reads
  }

  private requireCellAddress(address: string): WorkPaperCellAddress {
    this.assertDefaultSheetAllowsUnqualifiedAddress(address, 'cell address')
    const parsed = this.workpaper.simpleCellAddressFromString(address, this.defaultSheetId())
    if (parsed === undefined) {
      throw new Error(`Invalid WorkPaper cell address: ${address}`)
    }
    return parsed
  }

  private requireRange(range: string): WorkPaperCellRange {
    this.assertDefaultSheetAllowsUnqualifiedAddress(range, 'range')
    const parsedRange = this.workpaper.simpleCellRangeFromString(range, this.defaultSheetId())
    if (parsedRange !== undefined) {
      return parsedRange
    }
    const parsedCell = this.workpaper.simpleCellAddressFromString(range, this.defaultSheetId())
    if (parsedCell !== undefined) {
      return {
        start: parsedCell,
        end: parsedCell,
      }
    }
    throw new Error(`Invalid WorkPaper range: ${range}`)
  }

  private assertDefaultSheetAllowsUnqualifiedAddress(value: string, kind: string): void {
    if (this.defaultSheetName === undefined && !value.includes('!')) {
      throw new Error(`Sheet-qualified WorkPaper ${kind} required: ${value}`)
    }
  }

  private defaultSheetId(): number | undefined {
    if (this.defaultSheetName === undefined) {
      return undefined
    }
    const sheetId = this.workpaper.getSheetId(this.defaultSheetName)
    if (sheetId === undefined) {
      throw new Error(`Expected default sheet "${this.defaultSheetName}" to exist`)
    }
    return sheetId
  }

  private assertWritable(address: WorkPaperCellAddress): void {
    if (this.writableSheets === undefined) {
      return
    }
    const sheetName = this.workpaper.getSheetName(address.sheet)
    if (sheetName === undefined || !this.writableSheets.includes(sheetName)) {
      throw new Error(`Sheet "${sheetName ?? address.sheet}" is not writable`)
    }
  }

  private assertFormulaCanParse(value: A1CellInput, address: WorkPaperCellAddress): void {
    if (typeof value !== 'string' || !value.startsWith('=')) {
      return
    }
    if (!this.workpaper.validateFormula(value)) {
      throw new Error(`Invalid WorkPaper formula for ${this.formatCell(address)}: ${value}`)
    }
  }

  private formatCell(address: WorkPaperCellAddress): string {
    return this.workpaper.simpleCellAddressToString(address, {
      includeSheetName: true,
    })
  }

  private formatRange(range: WorkPaperCellRange): string {
    if (sameAddress(range.start, range.end)) {
      return this.formatCell(range.start)
    }
    return this.workpaper.simpleCellRangeToString(range, {
      includeSheetName: true,
    })
  }
}

interface PreparedA1Edit {
  readonly address: string
  readonly parsed: WorkPaperCellAddress
  readonly value: A1CellInput
}

function rangeComparison(read: A1RangeRead): unknown {
  return read.cells.map((row) => row.map(cellComparison))
}

function rangeComparisonExcludingEditedCells(read: A1RangeRead, editedCells: ReadonlySet<string>): unknown {
  return read.cells.map((row) => row.filter((cell) => !editedCells.has(cell.address)).map(cellComparison)).filter((row) => row.length > 0)
}

function readbackEditedFormulaCells(read: A1RangeRead, editedCells: ReadonlySet<string>): A1CellRead[] {
  return read.cells.flatMap((row) => row.filter((cell) => editedCells.has(cell.address) && cell.formula !== null))
}

function rangeCellAddresses(read: A1RangeRead): string[] {
  return read.cells.flatMap((row) => row.map((cell) => cell.address))
}

function cellComparison(read: A1CellRead): unknown {
  return {
    address: read.address,
    displayValue: read.displayValue,
    formula: read.formula,
    serialized: read.serialized,
    value: read.value,
  }
}

function serializedByAddress(reads: Readonly<Record<string, A1CellRead>>): Readonly<Record<string, RawCellContent>> {
  const serialized: Record<string, RawCellContent> = {}
  for (const [address, read] of Object.entries(reads)) {
    serialized[address] = read.serialized
  }
  return serialized
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function sameAddress(left: WorkPaperCellAddress, right: WorkPaperCellAddress): boolean {
  return left.sheet === right.sheet && left.row === right.row && left.col === right.col
}
