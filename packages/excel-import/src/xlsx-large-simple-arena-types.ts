import type { LiteralInput } from '@bilig/protocol'

export type ImportedWorkbookArenaDedupeMode = boolean | 'bounded'

export interface ImportedWorkbookArenaOptions {
  readonly deduplicateStrings?: ImportedWorkbookArenaDedupeMode
  readonly deduplicateFormulas?: ImportedWorkbookArenaDedupeMode
  readonly dedupeMaxEntries?: number
}

export interface ImportedWorkbookArenaSnapshot {
  readonly sheetIndex: number | null
  readonly sheetIndexes?: Uint32Array
  readonly rows: Uint32Array
  readonly columns: Uint16Array
  readonly valueKinds: Uint8Array
  readonly numberValues?: Float64Array
  readonly tinyIntegerValues?: Int8Array
  readonly smallIntegerValues?: Int16Array
  readonly sparseSmallIntegerCellIndexes?: Uint32Array
  readonly sparseSmallIntegerValues?: Int16Array
  readonly integerValues?: Int32Array
  readonly sparseIntegerCellIndexes?: Uint32Array
  readonly sparseIntegerValues?: Int32Array
  readonly stringIds?: Uint32Array
  readonly booleanValues?: Uint8Array
  readonly formulaIds?: Uint32Array
  readonly strings: readonly string[]
  readonly formulas: readonly string[]
}

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
