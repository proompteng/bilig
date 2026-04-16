import {
  buildRelativeFormulaTemplateTokenKey,
  canTranslateCompiledFormulaWithoutAst,
  compileFormulaAst,
  parseFormula,
  translateCompiledFormula,
  translateCompiledFormulaWithoutAst,
  type CompiledFormula,
} from '@bilig/formula'

export interface EngineFormulaTemplateNormalizationService {
  readonly clear: () => void
  readonly compileForCell: (source: string, ownerRow: number, ownerCol: number) => CompiledFormula
}

interface TemplateEntry {
  readonly templateKey: string
  readonly baseRow: number
  readonly baseCol: number
  readonly compiled: CompiledFormula
}

export function createEngineFormulaTemplateNormalizationService(): EngineFormulaTemplateNormalizationService {
  const templates = new Map<string, TemplateEntry>()
  const recentByColumn = new Map<number, TemplateEntry>()
  const translateTemplate = (compiled: CompiledFormula, rowDelta: number, colDelta: number, source: string): CompiledFormula =>
    (canTranslateCompiledFormulaWithoutAst(compiled)
      ? translateCompiledFormulaWithoutAst(compiled, rowDelta, colDelta, source)
      : translateCompiledFormula(compiled, rowDelta, colDelta, source)
    ).compiled

  return {
    clear() {
      templates.clear()
      recentByColumn.clear()
    },
    compileForCell(source, ownerRow, ownerCol) {
      const templateKey = buildRelativeFormulaTemplateTokenKey(source, ownerRow, ownerCol)
      const recent = recentByColumn.get(ownerCol)
      if (recent && recent.templateKey === templateKey) {
        const rowDelta = ownerRow - recent.baseRow
        const colDelta = ownerCol - recent.baseCol
        const translated =
          rowDelta === 0 && colDelta === 0 ? recent.compiled : translateTemplate(recent.compiled, rowDelta, colDelta, source)
        recentByColumn.set(ownerCol, {
          templateKey,
          baseRow: ownerRow,
          baseCol: ownerCol,
          compiled: translated,
        })
        return translated
      }
      const existing = templates.get(templateKey)
      if (!existing) {
        const ast = parseFormula(source)
        const compiled = compileFormulaAst(source, ast)
        templates.set(templateKey, {
          templateKey,
          baseRow: ownerRow,
          baseCol: ownerCol,
          compiled,
        })
        recentByColumn.set(ownerCol, {
          templateKey,
          baseRow: ownerRow,
          baseCol: ownerCol,
          compiled,
        })
        return compiled
      }
      const rowDelta = ownerRow - existing.baseRow
      const colDelta = ownerCol - existing.baseCol
      if (rowDelta === 0 && colDelta === 0) {
        recentByColumn.set(ownerCol, existing)
        return existing.compiled
      }
      const translated = translateTemplate(existing.compiled, rowDelta, colDelta, source)
      recentByColumn.set(ownerCol, {
        templateKey,
        baseRow: ownerRow,
        baseCol: ownerCol,
        compiled: translated,
      })
      return translated
    },
  }
}
