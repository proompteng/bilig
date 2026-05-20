import type { FormulaTemplateResolution } from '../../formula/template-bank.js'
import {
  translateSimpleDirectScalarFormula,
  translateTrustedSimpleDirectScalarFormula,
} from '../../formula/simple-direct-scalar-compile.js'
import {
  translateInitialPrefixSumFormula,
  tryBuildInitialPrefixSumTemplateKey,
  tryBuildInitialSimpleRowRelativeBinaryTemplateKeyInfo,
  type InitialTemplateFormulaCacheEntry,
} from './formula-initialization-template-keys.js'

export function createInitialTemplateFormulaResolver(
  compileTemplateFormula: (source: string, row: number, col: number) => FormulaTemplateResolution,
): (source: string, row: number, col: number) => FormulaTemplateResolution {
  const simpleTemplateCache = new Map<string | number, InitialTemplateFormulaCacheEntry>()
  return (source, row, col) => {
    const simpleTemplateKey = tryBuildInitialSimpleRowRelativeBinaryTemplateKeyInfo(source, row, col)
    const templateKey = simpleTemplateKey?.key
    const cached = templateKey === undefined ? undefined : simpleTemplateCache.get(templateKey)
    if (cached && simpleTemplateKey !== undefined) {
      const anchorRowDelta = row - cached.anchorRow
      const anchorColDelta = col - cached.anchorCol
      const compiled = simpleTemplateKey.usesRowLiteralSuffix
        ? translateSimpleDirectScalarFormula(cached.anchorCompiled, anchorRowDelta, anchorColDelta, source)
        : translateTrustedSimpleDirectScalarFormula(cached.anchorCompiled, anchorRowDelta, anchorColDelta, source)
      if (compiled) {
        return {
          ...cached.resolution,
          compiled,
          translated: cached.resolution.translated || anchorRowDelta !== 0 || anchorColDelta !== 0,
          rowDelta: cached.resolution.rowDelta + anchorRowDelta,
          colDelta: cached.resolution.colDelta + anchorColDelta,
        }
      }
    }
    const sumTemplateKey = tryBuildInitialPrefixSumTemplateKey(source, row, col)
    const cachedSum = sumTemplateKey === undefined ? undefined : simpleTemplateCache.get(sumTemplateKey.key)
    if (cachedSum && sumTemplateKey !== undefined) {
      return translateInitialPrefixSumFormula(cachedSum, source, row, col, sumTemplateKey)
    }
    const resolution = compileTemplateFormula(source, row, col)
    if (templateKey !== undefined) {
      simpleTemplateCache.set(templateKey, {
        resolution,
        anchorRow: row,
        anchorCol: col,
        anchorCompiled: resolution.compiled,
      })
    }
    if (sumTemplateKey !== undefined) {
      simpleTemplateCache.set(sumTemplateKey.key, {
        resolution,
        anchorRow: row,
        anchorCol: col,
        anchorCompiled: resolution.compiled,
      })
    }
    return resolution
  }
}
