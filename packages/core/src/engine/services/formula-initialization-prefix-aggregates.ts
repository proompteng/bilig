export type InitialPrefixAggregateKind = 'sum' | 'count' | 'average' | 'min' | 'max'

export interface InitialPrefixAggregateGroup {
  readonly sheetName: string
  readonly col: number
  readonly colEnd: number
  readonly aggregateKind: InitialPrefixAggregateKind
  maxRowEnd: number
  lastRowEnd: number
  formulasAreOrdered: boolean
  readonly formulas: Array<{ cellIndex: number; rowEnd: number; resultOffset?: number }>
}
