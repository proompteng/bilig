import type { LiteralInput, WorkbookRichTextCellSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { toDisplayText } from './workbook-import-helpers.js'
import type { LargeSimpleSharedStringEntry } from './xlsx-large-simple-shared-strings.js'

const initialCellCapacity = 1024
const noPoolId = 0xffffffff
const valueKindEmpty = 0
const valueKindNumber = 1
const valueKindString = 2
const valueKindBoolean = 3
const valueKindNull = 4
const valueKindSharedStringRef = 5
const flagHasFormula = 1 << 0
const initialStyleIndexCapacity = 256

export interface ImportedWorksheetArenaCellInput {
  readonly sheetIndex: number
  readonly row: number
  readonly column: number
  readonly value: LiteralInput | undefined
}

export interface ImportedWorksheetArenaSharedStringCellInput {
  readonly sheetIndex: number
  readonly row: number
  readonly column: number
  readonly sharedStringIndex: number
}

export interface ImportedWorkbookArenaSnapshot {
  readonly sheetIndexes: Uint32Array
  readonly rows: Uint32Array
  readonly columns: Uint32Array
  readonly valueKinds: Uint8Array
  readonly numberValues: Float64Array
  readonly stringIds: Uint32Array
  readonly booleanValues: Uint8Array
  readonly formulaIds: Uint32Array
  readonly styleIds: Uint32Array
  readonly flags: Uint8Array
  readonly strings: readonly string[]
  readonly formulas: readonly string[]
}

export class ImportedWorksheetStyleIndexArena {
  private rows: Uint32Array<ArrayBuffer> = new Uint32Array(initialStyleIndexCapacity)
  private columns: Uint32Array<ArrayBuffer> = new Uint32Array(initialStyleIndexCapacity)
  private styleIndexes: Uint32Array<ArrayBuffer> = new Uint32Array(initialStyleIndexCapacity)
  private length = 0

  get count(): number {
    return this.length
  }

  add(row: number, column: number, styleIndex: number): void {
    this.ensureCapacity(this.length + 1)
    const index = this.length
    this.length += 1
    this.rows[index] = row
    this.columns[index] = column
    this.styleIndexes[index] = styleIndex
  }

  collectRequiredStyleIndexes(output: Set<number>): void {
    for (let index = 0; index < this.length; index += 1) {
      output.add(this.styleIndexes[index] ?? 0)
    }
  }

  forEach(callback: (row: number, column: number, styleIndex: number) => void): void {
    for (let index = 0; index < this.length; index += 1) {
      callback(this.rows[index] ?? 0, this.columns[index] ?? 0, this.styleIndexes[index] ?? 0)
    }
  }

  release(): void {
    this.rows = new Uint32Array(0)
    this.columns = new Uint32Array(0)
    this.styleIndexes = new Uint32Array(0)
    this.length = 0
  }

  private ensureCapacity(nextLength: number): void {
    if (nextLength <= this.rows.length) {
      return
    }
    let nextCapacity = this.rows.length
    while (nextCapacity < nextLength) {
      nextCapacity *= 2
    }
    this.rows = growUint32Array(this.rows, nextCapacity)
    this.columns = growUint32Array(this.columns, nextCapacity)
    this.styleIndexes = growUint32Array(this.styleIndexes, nextCapacity)
  }
}

export class ImportedWorkbookArena {
  private sheetIndexes: Uint32Array<ArrayBuffer> = new Uint32Array(initialCellCapacity)
  private rows: Uint32Array<ArrayBuffer> = new Uint32Array(initialCellCapacity)
  private columns: Uint32Array<ArrayBuffer> = new Uint32Array(initialCellCapacity)
  private valueKinds: Uint8Array<ArrayBuffer> = new Uint8Array(initialCellCapacity)
  private numberValues: Float64Array<ArrayBuffer> = new Float64Array(initialCellCapacity)
  private stringIds: Uint32Array<ArrayBuffer> = filledUint32Array(initialCellCapacity, noPoolId)
  private booleanValues: Uint8Array<ArrayBuffer> = new Uint8Array(initialCellCapacity)
  private formulaIds: Uint32Array<ArrayBuffer> = filledUint32Array(initialCellCapacity, noPoolId)
  private styleIds: Uint32Array<ArrayBuffer> = filledUint32Array(initialCellCapacity, noPoolId)
  private flags: Uint8Array<ArrayBuffer> = new Uint8Array(initialCellCapacity)
  private length = 0
  private readonly strings: string[] = []
  private readonly stringIdsByValue = new Map<string, number>()
  private readonly formulas: string[] = []
  private readonly formulaIdsByValue = new Map<string, number>()
  private readonly previewValues = new Map<string, LiteralInput | string>()

  get cellCount(): number {
    return this.length
  }

  addCell(input: ImportedWorksheetArenaCellInput): number {
    this.ensureCapacity(this.length + 1)
    const index = this.appendCell(input.sheetIndex, input.row, input.column)
    this.addValue(index, input.value)
    if (input.row < 8 && input.column < 6 && input.value !== undefined) {
      this.previewValues.set(previewKey(input.row, input.column), input.value)
    }
    return index
  }

  addSharedStringCell(input: ImportedWorksheetArenaSharedStringCellInput): number {
    this.ensureCapacity(this.length + 1)
    const index = this.appendCell(input.sheetIndex, input.row, input.column)
    this.valueKinds[index] = valueKindSharedStringRef
    this.numberValues[index] = 0
    this.stringIds[index] = input.sharedStringIndex
    this.booleanValues[index] = 0
    return index
  }

  setFormula(cellIndex: number, formula: string): void {
    const formulaId = this.internFormula(formula)
    this.formulaIds[cellIndex] = formulaId
    this.flags[cellIndex] = (this.flags[cellIndex] ?? 0) | flagHasFormula
    const row = this.rows[cellIndex]
    const column = this.columns[cellIndex]
    if (row !== undefined && column !== undefined && row < 8 && column < 6 && !this.previewValues.has(previewKey(row, column))) {
      this.previewValues.set(previewKey(row, column), `=${formula}`)
    }
  }

  materializeSheetCells(sheetIndex: number): WorkbookSnapshot['sheets'][number]['cells'] {
    const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
    cells.length = this.countMaterializedSheetCells(sheetIndex)
    let outputIndex = 0
    for (let index = 0; index < this.length; index += 1) {
      if (this.sheetIndexes[index] !== sheetIndex) {
        continue
      }
      const value = this.materializeValue(index)
      const formulaId = this.formulaIds[index] ?? noPoolId
      const formula = formulaId === noPoolId ? undefined : this.formulas[formulaId]
      if (value === undefined && formula === undefined) {
        continue
      }
      const cell: WorkbookSnapshot['sheets'][number]['cells'][number] = {
        address: encodeCellAddress(this.rows[index] ?? 0, this.columns[index] ?? 0),
      }
      if (value !== undefined) {
        cell.value = value
      }
      if (formula !== undefined) {
        cell.formula = formula
      }
      cells[outputIndex] = cell
      outputIndex += 1
    }
    return cells
  }

  readPreviewText(row: number, column: number): string {
    return toDisplayText(this.previewValues.get(previewKey(row, column)))
  }

  collectSharedStringIndexes(output: Set<number>): void {
    for (let index = 0; index < this.length; index += 1) {
      if ((this.valueKinds[index] ?? valueKindEmpty) !== valueKindSharedStringRef) {
        continue
      }
      const sharedStringIndex = this.stringIds[index] ?? noPoolId
      if (sharedStringIndex !== noPoolId) {
        output.add(sharedStringIndex)
      }
    }
  }

  resolveSharedStrings(sharedStrings: readonly LargeSimpleSharedStringEntry[]): WorkbookRichTextCellSnapshot[] | null {
    const richTextCells: WorkbookRichTextCellSnapshot[] = []
    for (let index = 0; index < this.length; index += 1) {
      if ((this.valueKinds[index] ?? valueKindEmpty) !== valueKindSharedStringRef) {
        continue
      }
      const sharedStringIndex = this.stringIds[index] ?? noPoolId
      const entry = sharedStringIndex === noPoolId ? undefined : sharedStrings[sharedStringIndex]
      if (!entry) {
        return null
      }
      this.valueKinds[index] = valueKindString
      this.stringIds[index] = this.internString(entry.text)
      const row = this.rows[index] ?? 0
      const column = this.columns[index] ?? 0
      if (row < 8 && column < 6) {
        this.previewValues.set(previewKey(row, column), entry.text)
      }
      if (entry.rich) {
        richTextCells.push({
          address: encodeCellAddress(row, column),
          text: entry.text,
          storage: 'sharedString',
          xml: entry.xml ?? '',
        })
      }
    }
    return richTextCells
  }

  snapshot(): ImportedWorkbookArenaSnapshot {
    return {
      sheetIndexes: this.sheetIndexes.subarray(0, this.length),
      rows: this.rows.subarray(0, this.length),
      columns: this.columns.subarray(0, this.length),
      valueKinds: this.valueKinds.subarray(0, this.length),
      numberValues: this.numberValues.subarray(0, this.length),
      stringIds: this.stringIds.subarray(0, this.length),
      booleanValues: this.booleanValues.subarray(0, this.length),
      formulaIds: this.formulaIds.subarray(0, this.length),
      styleIds: this.styleIds.subarray(0, this.length),
      flags: this.flags.subarray(0, this.length),
      strings: this.strings,
      formulas: this.formulas,
    }
  }

  release(): void {
    this.sheetIndexes = new Uint32Array(0)
    this.rows = new Uint32Array(0)
    this.columns = new Uint32Array(0)
    this.valueKinds = new Uint8Array(0)
    this.numberValues = new Float64Array(0)
    this.stringIds = new Uint32Array(0)
    this.booleanValues = new Uint8Array(0)
    this.formulaIds = new Uint32Array(0)
    this.styleIds = new Uint32Array(0)
    this.flags = new Uint8Array(0)
    this.length = 0
    this.strings.length = 0
    this.stringIdsByValue.clear()
    this.formulas.length = 0
    this.formulaIdsByValue.clear()
    this.previewValues.clear()
  }

  private appendCell(sheetIndex: number, row: number, column: number): number {
    const index = this.length
    this.length += 1
    this.sheetIndexes[index] = sheetIndex
    this.rows[index] = row
    this.columns[index] = column
    this.styleIds[index] = noPoolId
    this.formulaIds[index] = noPoolId
    this.flags[index] = 0
    return index
  }

  private addValue(index: number, value: LiteralInput | undefined): void {
    if (value === undefined) {
      this.valueKinds[index] = valueKindEmpty
      this.numberValues[index] = 0
      this.stringIds[index] = noPoolId
      this.booleanValues[index] = 0
      return
    }
    if (value === null) {
      this.valueKinds[index] = valueKindNull
      this.numberValues[index] = 0
      this.stringIds[index] = noPoolId
      this.booleanValues[index] = 0
      return
    }
    if (typeof value === 'number') {
      this.valueKinds[index] = valueKindNumber
      this.numberValues[index] = value
      this.stringIds[index] = noPoolId
      this.booleanValues[index] = 0
      return
    }
    if (typeof value === 'boolean') {
      this.valueKinds[index] = valueKindBoolean
      this.numberValues[index] = 0
      this.stringIds[index] = noPoolId
      this.booleanValues[index] = value ? 1 : 0
      return
    }
    this.valueKinds[index] = valueKindString
    this.numberValues[index] = 0
    this.stringIds[index] = this.internString(value)
    this.booleanValues[index] = 0
  }

  private materializeValue(index: number): LiteralInput | undefined {
    const valueKind = this.valueKinds[index] ?? valueKindEmpty
    switch (valueKind) {
      case valueKindNumber:
        return this.numberValues[index]
      case valueKindString: {
        const stringId = this.stringIds[index] ?? noPoolId
        return stringId === noPoolId ? undefined : this.strings[stringId]
      }
      case valueKindBoolean:
        return (this.booleanValues[index] ?? 0) === 1
      case valueKindNull:
        return null
      default:
        return undefined
    }
  }

  private countMaterializedSheetCells(sheetIndex: number): number {
    let count = 0
    for (let index = 0; index < this.length; index += 1) {
      if (this.sheetIndexes[index] !== sheetIndex) {
        continue
      }
      const formulaId = this.formulaIds[index] ?? noPoolId
      if ((this.valueKinds[index] ?? valueKindEmpty) !== valueKindEmpty || formulaId !== noPoolId) {
        count += 1
      }
    }
    return count
  }

  private internString(value: string): number {
    const existing = this.stringIdsByValue.get(value)
    if (existing !== undefined) {
      return existing
    }
    const next = this.strings.length
    this.strings.push(value)
    this.stringIdsByValue.set(value, next)
    return next
  }

  private internFormula(value: string): number {
    const existing = this.formulaIdsByValue.get(value)
    if (existing !== undefined) {
      return existing
    }
    const next = this.formulas.length
    this.formulas.push(value)
    this.formulaIdsByValue.set(value, next)
    return next
  }

  private ensureCapacity(nextLength: number): void {
    if (nextLength <= this.rows.length) {
      return
    }
    let nextCapacity = this.rows.length
    while (nextCapacity < nextLength) {
      nextCapacity *= 2
    }
    this.sheetIndexes = growUint32Array(this.sheetIndexes, nextCapacity)
    this.rows = growUint32Array(this.rows, nextCapacity)
    this.columns = growUint32Array(this.columns, nextCapacity)
    this.valueKinds = growUint8Array(this.valueKinds, nextCapacity)
    this.numberValues = growFloat64Array(this.numberValues, nextCapacity)
    this.stringIds = growUint32Array(this.stringIds, nextCapacity, noPoolId)
    this.booleanValues = growUint8Array(this.booleanValues, nextCapacity)
    this.formulaIds = growUint32Array(this.formulaIds, nextCapacity, noPoolId)
    this.styleIds = growUint32Array(this.styleIds, nextCapacity, noPoolId)
    this.flags = growUint8Array(this.flags, nextCapacity)
  }
}

export interface ImportedWorksheetCellScan {
  readonly arena: ImportedWorkbookArena
  readonly sheetIndex: number
  readonly richTextCells: WorkbookRichTextCellSnapshot[]
  readonly styleIndexes: ImportedWorksheetStyleIndexArena
  readonly blankStyleCellCount: number
  readonly cellCount: number
  readonly valueCellCount: number
  readonly formulaCellCount: number
  readonly mergeCount?: number
  readonly conditionalFormatCount?: number
  readonly tableCount?: number
  readonly rowCount: number
  readonly columnCount: number
  readonly usedRange: {
    readonly startRow: number
    readonly startColumn: number
    readonly endRow: number
    readonly endColumn: number
  } | null
}

function filledUint32Array(length: number, value: number): Uint32Array<ArrayBuffer> {
  const output = new Uint32Array(length)
  output.fill(value)
  return output
}

function growUint8Array(source: Uint8Array<ArrayBuffer>, nextCapacity: number): Uint8Array<ArrayBuffer> {
  const output = new Uint8Array(nextCapacity)
  output.set(source)
  return output
}

function growUint32Array(source: Uint32Array<ArrayBuffer>, nextCapacity: number, fillValue?: number): Uint32Array<ArrayBuffer> {
  const output = new Uint32Array(nextCapacity)
  output.set(source)
  if (fillValue !== undefined && nextCapacity > source.length) {
    output.fill(fillValue, source.length)
  }
  return output
}

function growFloat64Array(source: Float64Array<ArrayBuffer>, nextCapacity: number): Float64Array<ArrayBuffer> {
  const output = new Float64Array(nextCapacity)
  output.set(source)
  return output
}

function previewKey(row: number, column: number): string {
  return `${String(row)}:${String(column)}`
}

function encodeCellAddress(row: number, column: number): string {
  let value = column + 1
  let columnName = ''
  while (value > 0) {
    value -= 1
    columnName = String.fromCharCode(65 + (value % 26)) + columnName
    value = Math.floor(value / 26)
  }
  return `${columnName}${String(row + 1)}`
}
