import {
  canTranslateCompiledFormulaWithoutAst,
  columnToIndex,
  indexToColumn,
  rewriteCompiledFormulaForStructuralTransform,
  translateCompiledFormula,
  translateCompiledFormulaWithoutAst,
  type CompiledFormula,
  type StructuralAxisTransform,
} from '@bilig/formula'
import { mapStructuralAxisIndex } from '../engine-structural-utils.js'
import type { FormulaInstanceSnapshot } from './formula-instance-table.js'
import type { FormulaTemplateSnapshot } from './template-bank.js'

export function retargetFormulaInstance(
  record: FormulaInstanceSnapshot,
  next: {
    readonly sheetName?: string
    readonly row?: number
    readonly col?: number
    readonly source?: string
    readonly templateId?: number
  },
): FormulaInstanceSnapshot {
  return {
    cellIndex: record.cellIndex,
    sheetName: next.sheetName ?? record.sheetName,
    row: next.row ?? record.row,
    col: next.col ?? record.col,
    source: next.source ?? record.source,
    ...(next.templateId !== undefined
      ? { templateId: next.templateId }
      : record.templateId !== undefined
        ? { templateId: record.templateId }
        : {}),
  }
}

export interface StructurallyRewrittenTemplate {
  readonly templateId: number
  readonly baseRow: number
  readonly baseCol: number
  readonly source: string
  readonly compiled: CompiledFormula
  readonly reusedProgram: boolean
}

export function rewriteTemplateForStructuralTransform(args: {
  readonly template: FormulaTemplateSnapshot
  readonly ownerSheetName: string
  readonly targetSheetName: string
  readonly transform: StructuralAxisTransform
}): StructurallyRewrittenTemplate | undefined {
  const rewritten = rewriteCompiledFormulaForStructuralTransform(
    args.template.compiled,
    args.ownerSheetName,
    args.targetSheetName,
    args.transform,
  )
  if (rewritten.source === args.template.baseSource) {
    return undefined
  }

  const mappedBaseRow =
    args.ownerSheetName === args.targetSheetName && args.transform.axis === 'row'
      ? mapStructuralAxisIndex(args.template.baseRow, args.transform)
      : args.template.baseRow
  const mappedBaseCol =
    args.ownerSheetName === args.targetSheetName && args.transform.axis === 'column'
      ? mapStructuralAxisIndex(args.template.baseCol, args.transform)
      : args.template.baseCol
  if (mappedBaseRow === undefined || mappedBaseCol === undefined) {
    return undefined
  }

  return {
    templateId: args.template.id,
    baseRow: mappedBaseRow,
    baseCol: mappedBaseCol,
    source: rewritten.source,
    compiled: rewritten.compiled,
    reusedProgram: rewritten.reusedProgram,
  }
}

export function retargetStructurallyRewrittenTemplateInstance(args: {
  readonly rewrittenTemplate: StructurallyRewrittenTemplate
  readonly ownerRow: number
  readonly ownerCol: number
}): { readonly source: string; readonly compiled: CompiledFormula; readonly reusedProgram: boolean } {
  const rowDelta = args.ownerRow - args.rewrittenTemplate.baseRow
  const colDelta = args.ownerCol - args.rewrittenTemplate.baseCol
  if (rowDelta === 0 && colDelta === 0) {
    return {
      source: args.rewrittenTemplate.source,
      compiled: args.rewrittenTemplate.compiled,
      reusedProgram: args.rewrittenTemplate.reusedProgram,
    }
  }

  if (canTranslateCompiledFormulaWithoutAst(args.rewrittenTemplate.compiled)) {
    const source = translateSimpleReferenceSource(args.rewrittenTemplate.source, rowDelta, colDelta)
    if (source !== undefined) {
      const translated = translateCompiledFormulaWithoutAst(args.rewrittenTemplate.compiled, rowDelta, colDelta, source)
      return {
        source: translated.source,
        compiled: translated.compiled,
        reusedProgram: args.rewrittenTemplate.reusedProgram,
      }
    }
  }

  const translated = translateCompiledFormula(args.rewrittenTemplate.compiled, rowDelta, colDelta)
  return {
    source: translated.source,
    compiled: translated.compiled,
    reusedProgram: args.rewrittenTemplate.reusedProgram,
  }
}

const SIMPLE_REFERENCE_TOKEN_RE = /(^|[^A-Z0-9_.$])(\$?)([A-Z]{1,3})(\$?)([1-9][0-9]*)(?=$|[^A-Z0-9_])/g

function translateSimpleReferenceSource(source: string, rowDelta: number, colDelta: number): string | undefined {
  if (
    source.includes('"') ||
    source.includes("'") ||
    source.includes('!') ||
    source.includes('[') ||
    source.includes(']') ||
    source.includes('{') ||
    source.includes('}')
  ) {
    return undefined
  }
  let failed = false
  const translated = source.replace(
    SIMPLE_REFERENCE_TOKEN_RE,
    (token: string, prefix: string, colAbsolute: string, columnText: string, rowAbsolute: string, rowText: string) => {
      const sourceCol = columnToIndex(columnText)
      const sourceRow = Number.parseInt(rowText, 10) - 1
      const nextCol = colAbsolute === '$' ? sourceCol : sourceCol + colDelta
      const nextRow = rowAbsolute === '$' ? sourceRow : sourceRow + rowDelta
      if (nextCol < 0 || nextRow < 0) {
        failed = true
        return token
      }
      return `${prefix}${colAbsolute}${indexToColumn(nextCol)}${rowAbsolute}${nextRow + 1}`
    },
  )
  return failed ? undefined : translated
}
