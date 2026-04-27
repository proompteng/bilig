import {
  buildRelativeFormulaTemplateTokenKey,
  canTranslateCompiledFormulaWithoutAst,
  columnToIndex,
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
  readonly resolveTrustedById: (
    templateId: number,
    source: string,
    ownerRow: number,
    ownerCol: number,
  ) => FormulaTemplateResolution | undefined
  readonly get: (templateId: number) => FormulaTemplateSnapshot | undefined
  readonly list: () => FormulaTemplateSnapshot[]
  readonly hydrate: (snapshots: readonly FormulaTemplateSnapshot[]) => void
}

interface MutableTemplateRecord extends FormulaTemplateSnapshot {}

interface AnchoredPrefixAggregateTemplateMatch {
  readonly compiled: CompiledFormula
  readonly templateKey: string
}

interface FormulaTemplateSourceKey {
  readonly compiled: CompiledFormula | undefined
  readonly templateKey: string
}

const SIMPLE_ROW_RELATIVE_BINARY_RE = /^([A-Z]+)([1-9]\d*)([+\-*/])(?:([A-Z]+)([1-9]\d*)|(\d+(?:\.\d+)?))$/

function relativeCellToken(columnText: string, rowText: string, ownerRow: number, ownerCol: number): string | undefined {
  const row = Number(rowText) - 1
  if (row !== ownerRow) {
    return undefined
  }
  const col = columnToIndex(columnText)
  if (col < 0) {
    return undefined
  }
  return `cell:.:rc${col - ownerCol}:rr0`
}

function binaryOperatorToken(operator: string): string | undefined {
  switch (operator) {
    case '+':
      return 'tok:plus:"+"'
    case '-':
      return 'tok:minus:"-"'
    case '*':
      return 'tok:star:"*"'
    case '/':
      return 'tok:slash:"/"'
    default:
      return undefined
  }
}

function tryBuildSimpleRowRelativeBinaryTemplateKey(source: string, ownerRow: number, ownerCol: number): string | undefined {
  const match = SIMPLE_ROW_RELATIVE_BINARY_RE.exec(source)
  if (!match) {
    return undefined
  }
  const left = relativeCellToken(match[1]!, match[2]!, ownerRow, ownerCol)
  const operator = binaryOperatorToken(match[3]!)
  if (!left || !operator) {
    return undefined
  }
  const right = match[4] !== undefined ? relativeCellToken(match[4], match[5]!, ownerRow, ownerCol) : `tok:number:"${match[6]!}"`
  return right ? `${left}|${operator}|${right}|eof` : undefined
}

function translateTemplate(compiled: CompiledFormula, rowDelta: number, colDelta: number, source: string): CompiledFormula {
  return (
    canTranslateCompiledFormulaWithoutAst(compiled)
      ? translateCompiledFormulaWithoutAst(compiled, rowDelta, colDelta, source)
      : translateCompiledFormula(compiled, rowDelta, colDelta, source)
  ).compiled
}

function compileSourceFormula(source: string): CompiledFormula {
  return tryCompileSimpleDirectAggregateFormula(source) ?? compileFormulaAst(source, parseFormula(source))
}

function resolveTemplateCompiled(
  template: MutableTemplateRecord,
  source: string,
  ownerRow: number,
  ownerCol: number,
  sourceTemplateKey = buildRelativeFormulaTemplateTokenKey(source, ownerRow, ownerCol),
): CompiledFormula {
  if (source === template.baseSource) {
    return template.compiled
  }
  if (template.templateKey.startsWith('anchored-prefix-aggregate:')) {
    const anchoredPrefixAggregate = tryMatchAnchoredPrefixAggregateTemplate(source, ownerRow, ownerCol)
    if (anchoredPrefixAggregate && anchoredPrefixAggregate.templateKey === template.templateKey) {
      return anchoredPrefixAggregate.compiled
    }
  }
  if (sourceTemplateKey !== template.templateKey) {
    return compileSourceFormula(source)
  }
  const rowDelta = ownerRow - template.baseRow
  const colDelta = ownerCol - template.baseCol
  if (rowDelta === 0 && colDelta === 0) {
    return compileSourceFormula(source)
  }
  try {
    return translateTemplate(template.compiled, rowDelta, colDelta, source)
  } catch {
    return compileSourceFormula(source)
  }
}

function resolveTrustedTemplateCompiled(
  template: MutableTemplateRecord,
  source: string,
  ownerRow: number,
  ownerCol: number,
): CompiledFormula {
  if (source === template.baseSource) {
    return template.compiled
  }
  if (template.templateKey.startsWith('anchored-prefix-aggregate:')) {
    const anchoredPrefixAggregate = tryMatchAnchoredPrefixAggregateTemplate(source, ownerRow, ownerCol)
    if (anchoredPrefixAggregate && anchoredPrefixAggregate.templateKey === template.templateKey) {
      return anchoredPrefixAggregate.compiled
    }
  }
  const rowDelta = ownerRow - template.baseRow
  const colDelta = ownerCol - template.baseCol
  if (rowDelta === 0 && colDelta === 0) {
    return compileSourceFormula(source)
  }
  try {
    return translateTemplate(template.compiled, rowDelta, colDelta, source)
  } catch {
    return compileSourceFormula(source)
  }
}

function resolveTemplateSourceKey(source: string, ownerRow: number, ownerCol: number): FormulaTemplateSourceKey {
  const anchoredPrefixAggregate = tryMatchAnchoredPrefixAggregateTemplate(source, ownerRow, ownerCol)
  return {
    compiled: anchoredPrefixAggregate?.compiled,
    templateKey:
      anchoredPrefixAggregate?.templateKey ??
      tryBuildSimpleRowRelativeBinaryTemplateKey(source, ownerRow, ownerCol) ??
      buildRelativeFormulaTemplateTokenKey(source, ownerRow, ownerCol),
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
    const compiled = compiledOverride ?? compileSourceFormula(source)
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
      const { compiled: compiledOverride, templateKey } = resolveTemplateSourceKey(source, ownerRow, ownerCol)
      const recent = recentByColumn.get(ownerCol)
      const template =
        recent && recent.templateKey === templateKey
          ? recent
          : (templatesByKey.get(templateKey) ?? internTemplate(source, ownerRow, ownerCol, templateKey, compiledOverride))
      const rowDelta = ownerRow - template.baseRow
      const colDelta = ownerCol - template.baseCol
      const translated = rowDelta !== 0 || colDelta !== 0
      const compiled = resolveTemplateCompiled(template, source, ownerRow, ownerCol, templateKey)
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
      const { templateKey } = resolveTemplateSourceKey(source, ownerRow, ownerCol)
      if (templateKey !== template.templateKey) {
        return undefined
      }
      const rowDelta = ownerRow - template.baseRow
      const colDelta = ownerCol - template.baseCol
      const translated = rowDelta !== 0 || colDelta !== 0
      const compiled = resolveTemplateCompiled(template, source, ownerRow, ownerCol, templateKey)
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
    resolveTrustedById(templateId, source, ownerRow, ownerCol) {
      const template = templatesById.get(templateId)
      if (!template) {
        return undefined
      }
      const rowDelta = ownerRow - template.baseRow
      const colDelta = ownerCol - template.baseCol
      const translated = rowDelta !== 0 || colDelta !== 0
      const compiled = resolveTrustedTemplateCompiled(template, source, ownerRow, ownerCol)
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
