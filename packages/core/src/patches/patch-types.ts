import type { CellValue } from '@bilig/protocol'

export interface EngineCellPatch {
  readonly kind: 'cell'
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
