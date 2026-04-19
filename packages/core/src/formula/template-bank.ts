import {
  buildRelativeFormulaTemplateTokenKey,
  canTranslateCompiledFormulaWithoutAst,
  compileFormulaAst,
  parseFormula,
  type ParsedRangeReferenceInfo,
  translateCompiledFormula,
  translateCompiledFormulaWithoutAst,
  type CompiledFormula,
} from '@bilig/formula'
import { addEngineCounter, type EngineCounters } from '../perf/engine-counters.js'
import { tryCompileSimpleDirectAggregateFormula } from './simple-direct-aggregate-compile.js'

export interface FormulaTemplateSnapshot {
  readonly id: number
  readonly templateKey: string
  readonly baseSource: string
  readonly baseRow: number
  readonly baseCol: number
  readonly compiled: CompiledFormula
}

export interface FormulaTemplateResolution {
  readonly templateId: number
  readonly templateKey: string
  readonly baseSource: string
  readonly compiled: CompiledFormula
  readonly translated: boolean
  readonly rowDelta: number
  readonly colDelta: number
}

export interface TemplateBank {
  readonly clearTransientCache: () => void
  readonly reset: () => void
  readonly resolve: (source: string, ownerRow: number, ownerCol: number) => FormulaTemplateResolution
  readonly resolveById: (templateId: number, source: string, ownerRow: number, ownerCol: number) => FormulaTemplateResolution | undefined
  readonly get: (templateId: number) => FormulaTemplateSnapshot | undefined
  readonly list: () => FormulaTemplateSnapshot[]
  readonly hydrate: (snapshots: readonly FormulaTemplateSnapshot[]) => void
}

interface MutableTemplateRecord extends FormulaTemplateSnapshot {}

interface AnchoredPrefixAggregateTemplateMatch {
  readonly compiled: CompiledFormula
  readonly templateKey: string
}

function translateTemplate(compiled: CompiledFormula, rowDelta: number, colDelta: number, source: string): CompiledFormula {
  return (
    canTranslateCompiledFormulaWithoutAst(compiled)
      ? translateCompiledFormulaWithoutAst(compiled, rowDelta, colDelta, source)
      : translateCompiledFormula(compiled, rowDelta, colDelta, source)
  ).compiled
}

function resolveTemplateCompiled(template: MutableTemplateRecord, source: string, ownerRow: number, ownerCol: number): CompiledFormula {
  if (source === template.baseSource) {
    return template.compiled
  }
  const anchoredPrefixAggregate = tryMatchAnchoredPrefixAggregateTemplate(source, ownerRow, ownerCol)
  if (anchoredPrefixAggregate && anchoredPrefixAggregate.templateKey === template.templateKey) {
    return anchoredPrefixAggregate.compiled
  }
  const rowDelta = ownerRow - template.baseRow
  const colDelta = ownerCol - template.baseCol
  if (rowDelta === 0 && colDelta === 0) {
    return template.compiled
  }
  try {
    return translateTemplate(template.compiled, rowDelta, colDelta, source)
  } catch {
    return tryCompileSimpleDirectAggregateFormula(source) ?? compileFormulaAst(source, parseFormula(source))
  }
}

function tryMatchAnchoredPrefixAggregateTemplate(
  source: string,
  ownerRow: number,
  ownerCol: number,
): AnchoredPrefixAggregateTemplateMatch | undefined {
  const compiled = tryCompileSimpleDirectAggregateFormula(source)
  if (!compiled) {
    return undefined
  }
  const aggregateKind = compiled.directAggregateCandidate?.aggregateKind
  const range: ParsedRangeReferenceInfo | undefined = compiled.parsedSymbolicRanges?.[0]
  if (!aggregateKind || !range) {
    return undefined
  }
  if (range.startCol !== range.endCol || range.startRow !== 0 || range.endRow !== ownerRow) {
    return undefined
  }
  return {
    compiled,
    templateKey: `anchored-prefix-aggregate:${aggregateKind}:c${range.startCol - ownerCol}`,
  }
}

export function createTemplateBank(args?: { readonly counters?: EngineCounters }): TemplateBank {
  const templatesByKey = new Map<string, MutableTemplateRecord>()
  const templatesById = new Map<number, MutableTemplateRecord>()
  const recentByColumn = new Map<number, MutableTemplateRecord>()
  let nextTemplateId = 1

  const internTemplate = (
    source: string,
    ownerRow: number,
    ownerCol: number,
    templateKeyOverride?: string,
    compiledOverride?: CompiledFormula,
  ): MutableTemplateRecord => {
    const templateKey = templateKeyOverride ?? buildRelativeFormulaTemplateTokenKey(source, ownerRow, ownerCol)
    const existing = templatesByKey.get(templateKey)
    if (existing) {
      return existing
    }
    if (args?.counters) {
      addEngineCounter(args.counters, 'formulasParsed')
    }
    const compiled = compiledOverride ?? tryCompileSimpleDirectAggregateFormula(source) ?? compileFormulaAst(source, parseFormula(source))
    const record: MutableTemplateRecord = {
      id: nextTemplateId,
      templateKey,
      baseSource: source,
      baseRow: ownerRow,
      baseCol: ownerCol,
      compiled,
    }
    nextTemplateId += 1
    templatesByKey.set(templateKey, record)
    templatesById.set(record.id, record)
    return record
  }

  return {
    clearTransientCache() {
      recentByColumn.clear()
    },
    reset() {
      templatesByKey.clear()
      templatesById.clear()
      recentByColumn.clear()
      nextTemplateId = 1
    },
    resolve(source, ownerRow, ownerCol) {
      const anchoredPrefixAggregate = tryMatchAnchoredPrefixAggregateTemplate(source, ownerRow, ownerCol)
      const templateKey = anchoredPrefixAggregate?.templateKey ?? buildRelativeFormulaTemplateTokenKey(source, ownerRow, ownerCol)
      const recent = recentByColumn.get(ownerCol)
      const template =
        recent && recent.templateKey === templateKey
          ? recent
          : (templatesByKey.get(templateKey) ?? internTemplate(source, ownerRow, ownerCol, templateKey, anchoredPrefixAggregate?.compiled))
      const rowDelta = ownerRow - template.baseRow
      const colDelta = ownerCol - template.baseCol
      const translated = rowDelta !== 0 || colDelta !== 0
      const compiled = resolveTemplateCompiled(template, source, ownerRow, ownerCol)
      recentByColumn.set(ownerCol, template)
      return {
        templateId: template.id,
        templateKey,
        baseSource: template.baseSource,
        compiled,
        translated,
        rowDelta,
        colDelta,
      }
    },
    resolveById(templateId, source, ownerRow, ownerCol) {
      const template = templatesById.get(templateId)
      if (!template) {
        return undefined
      }
      const rowDelta = ownerRow - template.baseRow
      const colDelta = ownerCol - template.baseCol
      const translated = rowDelta !== 0 || colDelta !== 0
      const compiled = resolveTemplateCompiled(template, source, ownerRow, ownerCol)
      recentByColumn.set(ownerCol, template)
      return {
        templateId: template.id,
        templateKey: template.templateKey,
        baseSource: template.baseSource,
        compiled,
        translated,
        rowDelta,
        colDelta,
      }
    },
    get(templateId) {
      return templatesById.get(templateId)
    },
    list() {
      return [...templatesById.values()].map((record) => ({
        id: record.id,
        templateKey: record.templateKey,
        baseSource: record.baseSource,
        baseRow: record.baseRow,
        baseCol: record.baseCol,
        compiled: record.compiled,
      }))
    },
    hydrate(snapshots) {
      templatesByKey.clear()
      templatesById.clear()
      recentByColumn.clear()
      let maxId = 0
      snapshots.forEach((snapshot) => {
        const record: MutableTemplateRecord = {
          id: snapshot.id,
          templateKey: snapshot.templateKey,
          baseSource: snapshot.baseSource,
          baseRow: snapshot.baseRow,
          baseCol: snapshot.baseCol,
          compiled: snapshot.compiled,
        }
        templatesByKey.set(record.templateKey, record)
        templatesById.set(record.id, record)
        maxId = Math.max(maxId, record.id)
      })
      nextTemplateId = maxId + 1
    },
  }
}
