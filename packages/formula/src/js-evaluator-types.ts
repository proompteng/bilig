import type { CellValue, ErrorCode } from '@bilig/protocol'
import type { ExcelDateSystem } from './builtins/datetime.js'
import type { LookupBuiltin } from './builtins/lookup.js'
import type { ArrayValue, EvaluationResult } from './runtime-values.js'

export interface EvaluationContext {
  sheetName: string
  workbookName?: string
  currentAddress?: string
  resolveCell: (sheetName: string, address: string) => CellValue
  resolveRange: (sheetName: string, start: string, end: string, refKind: 'cells' | 'rows' | 'cols') => CellValue[]
  resolveName?: (name: string, sheetName?: string) => CellValue
  resolveNameReference?: (name: string, sheetName?: string) => ReferenceOperand | undefined
  resolveFormula?: (sheetName: string, address: string) => string | undefined
  resolvePivotData?: (request: {
    dataField: string
    sheetName: string
    address: string
    filters: ReadonlyArray<{ field: string; item: CellValue }>
  }) => CellValue | undefined
  resolveMultipleOperations?: (request: {
    formulaSheetName: string
    formulaAddress: string
    rowCellSheetName: string
    rowCellAddress: string
    rowReplacementSheetName: string
    rowReplacementAddress: string
    columnCellSheetName?: string
    columnCellAddress?: string
    columnReplacementSheetName?: string
    columnReplacementAddress?: string
  }) => CellValue | undefined
  resolveExactVectorMatch?: (request: {
    lookupValue: CellValue
    sheetName: string
    start: string
    end: string
    startRow: number
    endRow: number
    startCol: number
    endCol: number
    searchMode: 1 | -1
  }) => ExactVectorMatchResult
  resolveApproximateVectorMatch?: (request: {
    lookupValue: CellValue
    sheetName: string
    start: string
    end: string
    startRow: number
    endRow: number
    startCol: number
    endCol: number
    matchMode: 1 | -1
  }) => ApproximateVectorMatchResult
  noteRangeMaterialization?: (cellCount: number) => void
  checkEvaluationBudget?: (stepCost?: number) => void
  noteExactLookupDirect?: () => void
  noteExactLookupFallback?: () => void
  isRowHidden?: (sheetName: string, rowIndex: number) => boolean
  listSheetNames?: () => string[]
  resolveBuiltin?: (name: string) => ((...args: CellValue[]) => EvaluationResult) | undefined
  resolveLookupBuiltin?: (name: string) => LookupBuiltin | undefined
  dateSystem?: ExcelDateSystem
}

export type ExactVectorMatchResult = { handled: false } | { handled: true; position: number | undefined }

export type ApproximateVectorMatchResult = ExactVectorMatchResult

export interface ReferenceOperand {
  kind: 'cell' | 'range' | 'row' | 'col'
  sheetName?: string
  sheetEndName?: string
  address?: string
  start?: string
  end?: string
  refKind?: 'cells' | 'rows' | 'cols'
}

export type JsPlanInstruction =
  | { opcode: 'push-number'; value: number }
  | { opcode: 'push-boolean'; value: boolean }
  | { opcode: 'push-string'; value: string }
  | { opcode: 'push-error'; code: ErrorCode }
  | { opcode: 'push-omitted' }
  | { opcode: 'make-array'; rows: number; cols: number }
  | { opcode: 'push-name'; name: string; sheetName?: string }
  | { opcode: 'push-cell'; sheetName?: string; address: string }
  | {
      opcode: 'push-range'
      sheetName?: string
      sheetEndName?: string
      start: string
      end: string
      refKind: 'cells' | 'rows' | 'cols'
    }
  | {
      opcode: 'lookup-exact-match'
      callee: 'MATCH' | 'XMATCH'
      sheetName?: string
      start: string
      end: string
      startRow: number
      endRow: number
      startCol: number
      endCol: number
      refKind: 'cells'
      searchMode: 1 | -1
    }
  | {
      opcode: 'lookup-approximate-match'
      callee: 'MATCH' | 'XMATCH'
      sheetName?: string
      start: string
      end: string
      startRow: number
      endRow: number
      startCol: number
      endCol: number
      refKind: 'cells'
      matchMode: 1 | -1
    }
  | { opcode: 'push-lambda'; params: string[]; body: JsPlanInstruction[] }
  | { opcode: 'unary'; operator: '+' | '-' }
  | {
      opcode: 'binary'
      operator: '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '>' | '>=' | '<' | '<=' | ':'
    }
  | {
      opcode: 'call'
      callee: string
      argc: number
      argRefs?: Array<ReferenceOperand | undefined>
    }
  | { opcode: 'invoke'; argc: number }
  | { opcode: 'begin-scope' }
  | { opcode: 'bind-name'; name: string }
  | { opcode: 'end-scope' }
  | { opcode: 'jump-if-false'; target: number }
  | { opcode: 'jump'; target: number }
  | { opcode: 'return' }

export type StackValue =
  | { kind: 'scalar'; value: CellValue; blankReference?: boolean }
  | { kind: 'omitted'; source?: 'argument' | 'binding' }
  | {
      kind: 'range'
      values: CellValue[]
      refKind: 'cells' | 'rows' | 'cols'
      rows: number
      cols: number
      sheetName?: string
      sheetEndName?: string
      start?: string
      end?: string
      blankReference?: boolean
    }
  | {
      kind: 'lambda'
      params: string[]
      body: JsPlanInstruction[]
      scopes: Array<Map<string, StackValue>>
    }
  | ArrayValue
