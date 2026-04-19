import type { CellRangeRef, CellValue } from '@bilig/protocol'

export const ENGINE_CELL_PATCH_KIND = 'cell' as const
export const ENGINE_RANGE_INVALIDATION_PATCH_KIND = 'range-invalidation' as const
export const ENGINE_ROW_INVALIDATION_PATCH_KIND = 'row-invalidation' as const
export const ENGINE_COLUMN_INVALIDATION_PATCH_KIND = 'column-invalidation' as const

export interface EngineCellPatch {
  readonly kind: typeof ENGINE_CELL_PATCH_KIND
  readonly cellIndex: number
  readonly address: {
    readonly sheet: number
    readonly row: number
    readonly col: number
  }
  readonly sheetName: string
  readonly a1: string
  readonly newValue: CellValue
}

export interface EngineRangeInvalidationPatch {
  readonly kind: typeof ENGINE_RANGE_INVALIDATION_PATCH_KIND
  readonly range: CellRangeRef
}

export interface EngineRowInvalidationPatch {
  readonly kind: typeof ENGINE_ROW_INVALIDATION_PATCH_KIND
  readonly sheetName: string
  readonly startIndex: number
  readonly endIndex: number
}

export interface EngineColumnInvalidationPatch {
  readonly kind: typeof ENGINE_COLUMN_INVALIDATION_PATCH_KIND
  readonly sheetName: string
  readonly startIndex: number
  readonly endIndex: number
}

export type EnginePatch = EngineCellPatch | EngineRangeInvalidationPatch | EngineRowInvalidationPatch | EngineColumnInvalidationPatch

export function isEngineCellPatch(value: unknown): value is EngineCellPatch {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    Reflect.get(value, 'kind') === ENGINE_CELL_PATCH_KIND &&
    typeof Reflect.get(value, 'cellIndex') === 'number' &&
    typeof Reflect.get(value, 'sheetName') === 'string' &&
    typeof Reflect.get(value, 'a1') === 'string'
  )
}

export function isEngineRangeInvalidationPatch(value: unknown): value is EngineRangeInvalidationPatch {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  const range = Reflect.get(value, 'range')
  return (
    Reflect.get(value, 'kind') === ENGINE_RANGE_INVALIDATION_PATCH_KIND &&
    typeof range === 'object' &&
    range !== null &&
    typeof Reflect.get(range, 'sheetName') === 'string' &&
    typeof Reflect.get(range, 'startAddress') === 'string' &&
    typeof Reflect.get(range, 'endAddress') === 'string'
  )
}

function isAxisInvalidationPatch(
  value: unknown,
  kind: typeof ENGINE_ROW_INVALIDATION_PATCH_KIND | typeof ENGINE_COLUMN_INVALIDATION_PATCH_KIND,
): boolean {
  return (
    typeof value === 'object' &&
    value !== null &&
    Reflect.get(value, 'kind') === kind &&
    typeof Reflect.get(value, 'sheetName') === 'string' &&
    typeof Reflect.get(value, 'startIndex') === 'number' &&
    typeof Reflect.get(value, 'endIndex') === 'number'
  )
}

export function isEngineRowInvalidationPatch(value: unknown): value is EngineRowInvalidationPatch {
  return isAxisInvalidationPatch(value, ENGINE_ROW_INVALIDATION_PATCH_KIND)
}

export function isEngineColumnInvalidationPatch(value: unknown): value is EngineColumnInvalidationPatch {
  return isAxisInvalidationPatch(value, ENGINE_COLUMN_INVALIDATION_PATCH_KIND)
}
