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
import { parseA1RowIndex } from './a1-row-number.js'
import {
  translateAnchoredPrefixDirectAggregateFormula,
  translateSimpleDirectAggregateFormula,
  tryCompileSimpleDirectAggregateFormula,
} from './simple-direct-aggregate-compile.js'
import {
  translateSimpleDirectScalarFormula,
  translateTrustedSimpleDirectScalarFormula,
  tryCompileSimpleDirectScalarFormula,
} from './simple-direct-scalar-compile.js'

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

const SIMPLE_ROW_RELATIVE_BINARY_RE = /^=?([A-Z]+)([1-9]\d*)([+\-*/])(?:([A-Z]+)([1-9]\d*)|(\d+(?:\.\d+)?))(?:\+(\d+(?:\.\d+)?))?$/
const SIMPLE_DIRECT_AGGREGATE_TEMPLATE_RE = /^=?(SUM|AVERAGE|AVG|COUNT|MIN|MAX)\s*\(\s*([A-Z]+)([1-9]\d*):([A-Z]+)([1-9]\d*)\s*\)$/i
const SIMPLE_A1_ROW_TOKEN_RE = /(^|[^A-Z0-9_.$])\$?[A-Z]{1,3}\$?([1-9]\d*)(?=$|[^A-Z0-9_])/gi
const GLOBAL_COMPILED_SOURCE_CACHE_LIMIT = 4096
const globalCompiledSourceCache = new Map<string, CompiledFormula>()

function relativeCellToken(columnText: string, rowText: string, ownerRow: number, ownerCol: number): string | undefined {
  const row = parseA1RowIndex(rowText)
  if (row === undefined) {
    return undefined
  }
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
  if (!right) {
    return undefined
  }
  const offset = match[7] === undefined ? '' : `|tok:plus:"+"|tok:number:"${match[7]}"`
  return `${left}|${operator}|${right}${offset}|eof`
}

function tryBuildSimpleDirectAggregateTemplateKey(source: string, ownerRow: number, ownerCol: number): string | undefined {
  const match = SIMPLE_DIRECT_AGGREGATE_TEMPLATE_RE.exec(source)
  if (!match) {
    return undefined
  }
  const aggregateKind = directAggregateTemplateKind(match[1]!)
  const startCol = columnToIndex(match[2]!)
  const startRow = parseA1RowIndex(match[3]!)
  const endCol = columnToIndex(match[4]!)
  const endRow = parseA1RowIndex(match[5]!)
  if (startRow === undefined || endRow === undefined || startCol < 0 || endCol < startCol || endRow < startRow) {
    return undefined
  }
  if (startCol === endCol && startRow === 0 && endRow === ownerRow) {
    return `anchored-prefix-aggregate:${aggregateKind}:c${startCol - ownerCol}`
  }
  if (startCol === endCol) {
    return ['relative-aggregate', aggregateKind, `c${startCol - ownerCol}`, `s${startRow - ownerRow}`, `e${endRow - ownerRow}`].join(':')
  }
  return [
    'relative-aggregate-rect',
    aggregateKind,
    `c${startCol - ownerCol}`,
    `ce${endCol - ownerCol}`,
    `s${startRow - ownerRow}`,
    `e${endRow - ownerRow}`,
  ].join(':')
}

function directAggregateTemplateKind(callee: string): 'sum' | 'average' | 'count' | 'min' | 'max' {
  switch (callee.toUpperCase()) {
    case 'SUM':
      return 'sum'
    case 'AVERAGE':
    case 'AVG':
      return 'average'
    case 'COUNT':
      return 'count'
    case 'MIN':
      return 'min'
    case 'MAX':
      return 'max'
    default:
      throw new Error(`Unsupported aggregate template callee: ${callee}`)
  }
}

function sourceHasUnsafeA1Row(source: string): boolean {
  SIMPLE_A1_ROW_TOKEN_RE.lastIndex = 0
  let match: RegExpExecArray | null
  while ((match = SIMPLE_A1_ROW_TOKEN_RE.exec(source)) !== null) {
    if (parseA1RowIndex(match[2]!) === undefined) {
      return true
    }
  }
  return false
}

function translateTemplate(compiled: CompiledFormula, rowDelta: number, colDelta: number, source: string): CompiledFormula {
  const translatedSimpleScalar = translateSimpleDirectScalarFormula(compiled, rowDelta, colDelta, source)
  if (translatedSimpleScalar) {
    return translatedSimpleScalar
  }
  const translatedSimpleAggregate = translateSimpleDirectAggregateFormula(compiled, rowDelta, colDelta, source)
  if (translatedSimpleAggregate) {
    return translatedSimpleAggregate
  }
  return (
    canTranslateCompiledFormulaWithoutAst(compiled)
      ? translateCompiledFormulaWithoutAst(compiled, rowDelta, colDelta, source)
      : translateCompiledFormula(compiled, rowDelta, colDelta, source)
  ).compiled
}

function cloneCompiledFormula(compiled: CompiledFormula): CompiledFormula {
  return { ...compiled }
}

function rememberGlobalCompiledSource(source: string, compiled: CompiledFormula): void {
  if (globalCompiledSourceCache.has(source)) {
    globalCompiledSourceCache.delete(source)
  } else if (globalCompiledSourceCache.size >= GLOBAL_COMPILED_SOURCE_CACHE_LIMIT) {
    const oldest = globalCompiledSourceCache.keys().next().value
    if (oldest !== undefined) {
      globalCompiledSourceCache.delete(oldest)
    }
  }
  globalCompiledSourceCache.set(source, compiled)
}

function compileSourceFormula(source: string): CompiledFormula {
  const simpleCompiled = tryCompileSimpleDirectScalarFormula(source) ?? tryCompileSimpleDirectAggregateFormula(source)
  if (simpleCompiled !== undefined) {
    return simpleCompiled
  }
  const cached = globalCompiledSourceCache.get(source)
  if (cached !== undefined) {
    globalCompiledSourceCache.delete(source)
    globalCompiledSourceCache.set(source, cached)
    return cloneCompiledFormula(cached)
  }
  const compiled = compileFormulaAst(source, parseFormula(source))
  rememberGlobalCompiledSource(source, compiled)
  return cloneCompiledFormula(compiled)
}

function resolveTemplateCompiled(
  template: MutableTemplateRecord,
  source: string,
  ownerRow: number,
  ownerCol: number,
  sourceTemplateKey = buildRelativeFormulaTemplateTokenKey(source, ownerRow, ownerCol),
  compiledOverride?: CompiledFormula,
): CompiledFormula {
  if (source === template.baseSource) {
    return template.compiled
  }
  if (compiledOverride && sourceTemplateKey === template.templateKey) {
    return sourceTemplateKey.startsWith('relative-aggregate:') ? { ...compiledOverride, astMatchesSource: false } : compiledOverride
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
  if (template.templateKey.startsWith('cell:')) {
    const translated = translateTrustedSimpleDirectScalarFormula(template.compiled, rowDelta, colDelta, source)
    if (translated) {
      return translated
    }
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
  const rowDelta = ownerRow - template.baseRow
  const colDelta = ownerCol - template.baseCol
  if (template.templateKey.startsWith('cell:')) {
    const translated = translateTrustedSimpleDirectScalarFormula(template.compiled, rowDelta, colDelta, source)
    if (translated) {
      return translated
    }
  }
  if (template.templateKey.startsWith('relative-aggregate:') || template.templateKey.startsWith('relative-aggregate-rect:')) {
    const translated = translateSimpleDirectAggregateFormula(template.compiled, rowDelta, colDelta, source)
    if (translated) {
      return translated
    }
  }
  if (template.templateKey.startsWith('anchored-prefix-aggregate:')) {
    const translated = translateAnchoredPrefixDirectAggregateFormula(template.compiled, ownerRow, colDelta, source)
    if (translated) {
      return translated
    }
  }
  const directScalarTemplateKey = tryBuildSimpleRowRelativeBinaryTemplateKey(source, ownerRow, ownerCol)
  if (directScalarTemplateKey !== undefined) {
    if (directScalarTemplateKey === template.templateKey) {
      const translated = translateSimpleDirectScalarFormula(template.compiled, rowDelta, colDelta, source)
      if (translated) {
        return translated
      }
    }
    return compileSourceFormula(source)
  }
  const directAggregateTemplateKey = tryBuildSimpleDirectAggregateTemplateKey(source, ownerRow, ownerCol)
  if (directAggregateTemplateKey !== undefined) {
    if (directAggregateTemplateKey === template.templateKey) {
      if (template.templateKey.startsWith('anchored-prefix-aggregate:')) {
        const anchoredPrefixAggregate = tryMatchAnchoredPrefixAggregateTemplate(source, ownerRow, ownerCol)
        if (anchoredPrefixAggregate && anchoredPrefixAggregate.templateKey === template.templateKey) {
          return anchoredPrefixAggregate.compiled
        }
      } else {
        const translated = translateSimpleDirectAggregateFormula(template.compiled, rowDelta, colDelta, source)
        if (translated) {
          return translated
        }
      }
    }
    return compileSourceFormula(source)
  }
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
  if (sourceHasUnsafeA1Row(source)) {
    return {
      compiled: undefined,
      templateKey: `unsafe-a1:${source}:r${ownerRow}:c${ownerCol}`,
    }
  }
  const simpleRowRelativeBinaryKey = tryBuildSimpleRowRelativeBinaryTemplateKey(source, ownerRow, ownerCol)
  if (simpleRowRelativeBinaryKey !== undefined) {
    return {
      compiled: undefined,
      templateKey: simpleRowRelativeBinaryKey,
    }
  }
  const simpleDirectAggregateKey = tryBuildSimpleDirectAggregateTemplateKey(source, ownerRow, ownerCol)
  if (simpleDirectAggregateKey !== undefined) {
    return {
      compiled: undefined,
      templateKey: simpleDirectAggregateKey,
    }
  }
  const aggregateTemplate = tryMatchAggregateTemplate(source, ownerRow, ownerCol)
  if (aggregateTemplate) {
    return aggregateTemplate
  }
  return {
    compiled: undefined,
    templateKey: buildRelativeFormulaTemplateTokenKey(source, ownerRow, ownerCol),
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

function tryMatchAggregateTemplate(source: string, ownerRow: number, ownerCol: number): AnchoredPrefixAggregateTemplateMatch | undefined {
  const compiled = tryCompileSimpleDirectAggregateFormula(source)
  if (!compiled) {
    return undefined
  }
  const aggregateKind = compiled.directAggregateCandidate?.aggregateKind
  const range: ParsedRangeReferenceInfo | undefined = compiled.parsedSymbolicRanges?.[0]
  if (!aggregateKind || !range || range.startCol !== range.endCol) {
    return undefined
  }
  if (range.startRow === 0 && range.endRow === ownerRow) {
    return {
      compiled,
      templateKey: `anchored-prefix-aggregate:${aggregateKind}:c${range.startCol - ownerCol}`,
    }
  }
  return {
    compiled,
    templateKey: [
      'relative-aggregate',
      aggregateKind,
      `c${range.startCol - ownerCol}`,
      `s${range.startRow - ownerRow}`,
      `e${range.endRow - ownerRow}`,
    ].join(':'),
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
      const compiled = resolveTemplateCompiled(template, source, ownerRow, ownerCol, templateKey, compiledOverride)
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
      const { compiled: compiledOverride, templateKey } = resolveTemplateSourceKey(source, ownerRow, ownerCol)
      if (templateKey !== template.templateKey) {
        return undefined
      }
      const rowDelta = ownerRow - template.baseRow
      const colDelta = ownerCol - template.baseCol
      const translated = rowDelta !== 0 || colDelta !== 0
      const compiled = resolveTemplateCompiled(template, source, ownerRow, ownerCol, templateKey, compiledOverride)
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
