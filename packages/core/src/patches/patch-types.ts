import type { CellValue } from '@bilig/protocol'

export const ENGINE_CELL_PATCH_KIND = 'cell' as const

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

export type EnginePatch = EngineCellPatch

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
