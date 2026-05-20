import { translateFormulaReferences } from '@bilig/formula'
import { normalizeImportedFormulaSource } from './xlsx-formula-translation.js'
import { formulaReferencesExternalWorkbook, formulaReferencesVolatileFunction } from './xlsx-import-warnings.js'
import type { ImportedWorkbookArena } from './xlsx-large-simple-arena.js'

const initialFormulaCapacity = 128
const boundedRawFormulaDedupeMaxEntries = 8192
const noPoolId = 0xffffffff
const formulaTypeNormal = 0
const formulaTypeShared = 1

export class LargeSimpleFormulaRecords {
  private cellIndexes: Uint32Array<ArrayBuffer> = new Uint32Array(initialFormulaCapacity)
  private rows: Uint32Array<ArrayBuffer> = new Uint32Array(initialFormulaCapacity)
  private columns: Uint32Array<ArrayBuffer> = new Uint32Array(initialFormulaCapacity)
  private typeCodes: Uint8Array<ArrayBuffer> = new Uint8Array(initialFormulaCapacity)
  private sharedIndexes: Uint32Array<ArrayBuffer> = filledUint32Array(initialFormulaCapacity, noPoolId)
  private rawFormulaIds: Uint32Array<ArrayBuffer> = filledUint32Array(initialFormulaCapacity, noPoolId)
  private readonly rawFormulas: string[] = []
  private readonly rawFormulaIdsByValue = new Map<string, number>()
  private readonly boundedRawFormulaKeys: string[] = []
  private boundedRawFormulaEvictionIndex = 0
  private readonly normalizedFormulas: (string | null | undefined)[] = []
  private length = 0

  constructor(private readonly allowUnsupportedFormulaText = false) {}

  get count(): number {
    return this.length
  }

  get rawFormulaPoolCount(): number {
    return this.rawFormulas.length
  }

  add(cellIndex: number, row: number, column: number, typeCode: number, sharedIndex: number | null, rawFormulaText: string): void {
    this.ensureCapacity(this.length + 1)
    const index = this.length
    this.length += 1
    this.cellIndexes[index] = cellIndex
    this.rows[index] = row
    this.columns[index] = column
    this.typeCodes[index] = typeCode
    this.sharedIndexes[index] = sharedIndex ?? noPoolId
    this.rawFormulaIds[index] = this.internRawFormula(rawFormulaText)
  }

  resolveIntoArena(arena: ImportedWorkbookArena): boolean {
    const sharedBases = new Map<number, SharedFormulaBase>()
    for (let index = 0; index < this.length; index += 1) {
      if (this.typeCodes[index] !== formulaTypeShared || !this.hasRawFormulaText(index)) {
        continue
      }
      const normalized = this.normalizedFormulaAt(index)
      const sharedIndex = this.sharedIndexes[index] ?? noPoolId
      if (normalized === null || sharedIndex === noPoolId) {
        return false
      }
      sharedBases.set(sharedIndex, {
        row: this.rows[index] ?? 0,
        column: this.columns[index] ?? 0,
        formula: normalized,
      })
      arena.setFormula(this.cellIndexes[index] ?? 0, normalized)
    }

    for (let index = 0; index < this.length; index += 1) {
      if (this.typeCodes[index] === formulaTypeShared) {
        if (this.hasRawFormulaText(index)) {
          continue
        }
        const sharedIndex = this.sharedIndexes[index] ?? noPoolId
        const base = sharedIndex === noPoolId ? undefined : sharedBases.get(sharedIndex)
        if (!base) {
          return false
        }
        try {
          arena.setFormula(
            this.cellIndexes[index] ?? 0,
            translateFormulaReferences(base.formula, (this.rows[index] ?? 0) - base.row, (this.columns[index] ?? 0) - base.column),
          )
        } catch {
          return false
        }
        continue
      }
      const normalized = this.normalizedFormulaAt(index)
      if (normalized === null) {
        return false
      }
      arena.setFormula(this.cellIndexes[index] ?? 0, normalized)
    }
    return true
  }

  private hasRawFormulaText(index: number): boolean {
    return this.rawFormulaText(index).trim().length > 0
  }

  private rawFormulaText(index: number): string {
    const rawFormulaId = this.rawFormulaIds[index] ?? noPoolId
    return rawFormulaId === noPoolId ? '' : (this.rawFormulas[rawFormulaId] ?? '')
  }

  private normalizedFormulaAt(index: number): string | null {
    const rawFormulaId = this.rawFormulaIds[index] ?? noPoolId
    if (rawFormulaId === noPoolId) {
      return null
    }
    if (rawFormulaId < this.normalizedFormulas.length && this.normalizedFormulas[rawFormulaId] !== undefined) {
      return this.normalizedFormulas[rawFormulaId] ?? null
    }
    const normalized = normalizeLargeSimpleFormula(this.rawFormulas[rawFormulaId], this.allowUnsupportedFormulaText)
    this.normalizedFormulas[rawFormulaId] = normalized
    return normalized
  }

  private internRawFormula(value: string): number {
    const existing = this.rawFormulaIdsByValue.get(value)
    if (existing !== undefined) {
      return existing
    }
    const next = this.rawFormulas.length
    this.rawFormulas.push(value)
    this.rawFormulaIdsByValue.set(value, next)
    if (this.allowUnsupportedFormulaText) {
      this.rememberBoundedRawFormulaKey(value)
    }
    return next
  }

  private rememberBoundedRawFormulaKey(value: string): void {
    this.boundedRawFormulaKeys.push(value)
    while (this.boundedRawFormulaKeys.length - this.boundedRawFormulaEvictionIndex > boundedRawFormulaDedupeMaxEntries) {
      const evicted = this.boundedRawFormulaKeys[this.boundedRawFormulaEvictionIndex]
      this.boundedRawFormulaEvictionIndex += 1
      if (evicted !== undefined) {
        this.rawFormulaIdsByValue.delete(evicted)
      }
    }
    if (
      this.boundedRawFormulaEvictionIndex > boundedRawFormulaDedupeMaxEntries &&
      this.boundedRawFormulaEvictionIndex * 2 > this.boundedRawFormulaKeys.length
    ) {
      this.boundedRawFormulaKeys.splice(0, this.boundedRawFormulaEvictionIndex)
      this.boundedRawFormulaEvictionIndex = 0
    }
  }

  private ensureCapacity(nextLength: number): void {
    if (nextLength <= this.cellIndexes.length) {
      return
    }
    let nextCapacity = this.cellIndexes.length
    while (nextCapacity < nextLength) {
      nextCapacity *= 2
    }
    this.cellIndexes = growUint32Array(this.cellIndexes, nextCapacity)
    this.rows = growUint32Array(this.rows, nextCapacity)
    this.columns = growUint32Array(this.columns, nextCapacity)
    this.typeCodes = growUint8Array(this.typeCodes, nextCapacity)
    this.sharedIndexes = growUint32Array(this.sharedIndexes, nextCapacity, noPoolId)
    this.rawFormulaIds = growUint32Array(this.rawFormulaIds, nextCapacity, noPoolId)
  }
}

export function readLargeSimpleFormulaTypeCode(type: string | null): number {
  return type === 'shared' ? formulaTypeShared : formulaTypeNormal
}

export function parseLargeSimpleSharedFormulaIndex(value: string | null): number | null {
  if (!value || !/^[0-9]+$/u.test(value)) {
    return null
  }
  const index = Number(value)
  return Number.isSafeInteger(index) ? index : null
}

interface SharedFormulaBase {
  readonly row: number
  readonly column: number
  readonly formula: string
}

function normalizeLargeSimpleFormula(rawFormulaText: string | undefined, allowUnsupportedFormulaText: boolean): string | null {
  const decoded = rawFormulaText === undefined ? undefined : decodeXmlText(rawFormulaText).trim()
  if (decoded === undefined || decoded.length === 0) {
    return null
  }
  const formula = normalizeImportedFormulaSource(decoded)
  return !allowUnsupportedFormulaText &&
    (formulaReferencesExternalWorkbook(formula) || formulaReferencesVolatileFunction(formula) || formulaReferencesStructuredTable(formula))
    ? null
    : formula
}

function formulaReferencesStructuredTable(formula: string): boolean {
  return /\[[#@\w]/u.test(formula)
}

function decodeXmlText(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/gu, (_match, entity: string) => {
    if (entity.startsWith('#x')) {
      const codePoint = Number.parseInt(entity.slice(2), 16)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    if (entity.startsWith('#')) {
      const codePoint = Number.parseInt(entity.slice(1), 10)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    switch (entity) {
      case 'amp':
        return '&'
      case 'lt':
        return '<'
      case 'gt':
        return '>'
      case 'quot':
        return '"'
      case 'apos':
        return "'"
      default:
        return ''
    }
  })
}

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
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
