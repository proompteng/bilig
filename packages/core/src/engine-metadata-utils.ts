import {
  ErrorCode,
  type CellValue,
  type LiteralInput,
  type WorkbookCalculationSettingsSnapshot,
  type WorkbookDefinedNameValueSnapshot,
} from '@bilig/protocol'
import {
  hasBuiltin,
  parseCellAddress,
  parseFormula,
  renameFormulaSheetReferences,
  type FormulaNode,
  type ReferenceOperand,
} from '@bilig/formula'
import type { StringPool } from './string-pool.js'
import { errorValue, literalToValue } from './engine-value-utils.js'
import { normalizeDefinedName, normalizeWorkbookObjectName } from './workbook-store.js'

function literalToFormulaNode(input: LiteralInput): FormulaNode | null {
  if (typeof input === 'number') {
    return { kind: 'NumberLiteral', value: input }
  }
  if (typeof input === 'string') {
    return { kind: 'StringLiteral', value: input }
  }
  if (typeof input === 'boolean') {
    return { kind: 'BooleanLiteral', value: input }
  }
  return null
}

function literalFormulaNodeToCellValue(input: FormulaNode, stringPool: StringPool): CellValue | undefined {
  if (input.kind === 'NumberLiteral' || input.kind === 'BooleanLiteral' || input.kind === 'StringLiteral') {
    return literalToValue(input.value, stringPool)
  }
  if (input.kind === 'ErrorLiteral') {
    return errorValue(input.code)
  }
  return undefined
}

function definedNameValueToFormulaNode(input: WorkbookDefinedNameValueSnapshot): FormulaNode | null {
  if (typeof input === 'object' && input !== null && 'kind' in input) {
    switch (input.kind) {
      case 'scalar':
        return literalToFormulaNode(input.value)
      case 'cell-ref':
        return { kind: 'CellRef', ref: input.address, sheetName: input.sheetName }
      case 'range-ref':
        return {
          kind: 'RangeRef',
          refKind: 'cells',
          start: input.startAddress,
          end: input.endAddress,
          sheetName: input.sheetName,
        }
      case 'structured-ref':
        return {
          kind: 'StructuredRef',
          tableName: input.tableName,
          columnName: input.columnName,
        }
      case 'formula':
        try {
          return parseFormula(input.formula)
        } catch {
          return { kind: 'ErrorLiteral', code: ErrorCode.Value }
        }
    }
  }
  if (typeof input === 'string' && input.startsWith('=')) {
    try {
      return parseFormula(input)
    } catch {
      return { kind: 'ErrorLiteral', code: ErrorCode.Value }
    }
  }
  return literalToFormulaNode(input)
}

function renameFormulaTextForSheet(input: string, oldSheetName: string, newSheetName: string): string {
  const hasLeadingEquals = input.startsWith('=')
  const source = hasLeadingEquals ? input.slice(1) : input
  const rewritten = renameFormulaSheetReferences(source, oldSheetName, newSheetName)
  return hasLeadingEquals ? `=${rewritten}` : rewritten
}

function unquoteFormulaSheetName(sheetName: string): string {
  if (!sheetName.startsWith("'") || !sheetName.endsWith("'")) {
    return sheetName
  }
  return sheetName.slice(1, -1).replaceAll("''", "'")
}

export interface MetadataResolutionContext {
  resolveName: (name: string, scopeSheetName?: string) => WorkbookDefinedNameValueSnapshot | undefined
  resolveStructuredReference: (tableName: string, columnName: string) => FormulaNode | undefined
  resolveSpillReference: (sheetName: string | undefined, address: string) => FormulaNode | undefined
}

export type MetadataFormulaValueContext = 'scalar' | 'array'

const ARRAY_CONTEXT_ALL_ARGUMENT_CALLEES = new Set([
  'SUM',
  'AVG',
  'AVERAGE',
  'MIN',
  'MAX',
  'COUNT',
  'COUNTA',
  'COUNTBLANK',
  'PRODUCT',
  'SUMPRODUCT',
  'SUMSQ',
  'GEOMEAN',
  'HARMEAN',
  'STDEV',
  'STDEV.P',
  'STDEV.S',
  'STDEVA',
  'STDEVP',
  'STDEVPA',
  'VAR',
  'VAR.P',
  'VAR.S',
  'VARA',
  'VARP',
  'VARPA',
  'SKEW',
  'SKEW.P',
  'KURT',
  'BYROW',
  'BYCOL',
  'MAP',
  'SCAN',
  'MAKEARRAY',
  'MMULT',
  'SLOPE',
  'INTERCEPT',
  'CORREL',
  'COVAR',
  'COVARIANCE.P',
  'COVARIANCE.S',
  'FORECAST',
  'FORECAST.LINEAR',
  'LINEST',
  'LOGEST',
  'TREND',
  'GROWTH',
])

const ARRAY_CONTEXT_ARGUMENT_INDEXES = new Map<string, ReadonlySet<number>>([
  ['SINGLE', new Set([0])],
  ['SUMIF', new Set([0, 2])],
  ['SUMIFS', new Set([0])],
  ['COUNTIF', new Set([0])],
  ['COUNTIFS', new Set([0])],
  ['AVERAGEIF', new Set([0, 2])],
  ['AVERAGEIFS', new Set([0])],
  ['MINIFS', new Set([0])],
  ['MAXIFS', new Set([0])],
  ['SUBTOTAL', new Set([1])],
  ['AGGREGATE', new Set([2])],
  ['INDEX', new Set([0])],
  ['MATCH', new Set([1])],
  ['XMATCH', new Set([1])],
  ['LOOKUP', new Set([1, 2])],
  ['VLOOKUP', new Set([1])],
  ['HLOOKUP', new Set([1])],
  ['XLOOKUP', new Set([1, 2])],
  ['FILTER', new Set([0, 1])],
  ['SORT', new Set([0])],
  ['SORTBY', new Set([0, 1])],
  ['UNIQUE', new Set([0])],
  ['TAKE', new Set([0])],
  ['DROP', new Set([0])],
  ['CHOOSECOLS', new Set([0])],
  ['CHOOSEROWS', new Set([0])],
  ['TOCOL', new Set([0])],
  ['TOROW', new Set([0])],
  ['WRAPROWS', new Set([0])],
  ['WRAPCOLS', new Set([0])],
  ['TRANSPOSE', new Set([0])],
  ['TRIMRANGE', new Set([0])],
  ['EXPAND', new Set([0])],
  ['FREQUENCY', new Set([0, 1])],
  ['MODE.MULT', new Set([0])],
  ['IRR', new Set([0])],
  ['MIRR', new Set([0])],
  ['XIRR', new Set([0, 1])],
  ['XNPV', new Set([1, 2])],
  ['ROW', new Set([0])],
  ['COLUMN', new Set([0])],
  ['FORMULA', new Set([0])],
  ['FORMULATEXT', new Set([0])],
  ['CELL', new Set([1])],
])

function callArgumentValueContext(
  callee: string,
  argIndex: number,
  parentContext: MetadataFormulaValueContext,
): MetadataFormulaValueContext {
  const normalized = callee.trim().toUpperCase()
  if (
    ARRAY_CONTEXT_ALL_ARGUMENT_CALLEES.has(normalized) ||
    ARRAY_CONTEXT_ARGUMENT_INDEXES.get(normalized)?.has(argIndex) ||
    ((normalized === 'SUMIFS' || normalized === 'AVERAGEIFS' || normalized === 'MINIFS' || normalized === 'MAXIFS') &&
      (argIndex === 0 || argIndex % 2 === 1)) ||
    (normalized === 'COUNTIFS' && argIndex % 2 === 0) ||
    (normalized === 'SUBTOTAL' && argIndex >= 1) ||
    (normalized === 'AGGREGATE' && argIndex >= 2) ||
    (normalized === 'SORTBY' && (argIndex === 0 || argIndex % 2 === 1))
  ) {
    return 'array'
  }
  return parentContext
}

function maybeImplicitIntersectRangeName(node: FormulaNode, valueContext: MetadataFormulaValueContext): FormulaNode {
  if (valueContext !== 'scalar' || node.kind !== 'RangeRef') {
    return node
  }
  return {
    kind: 'CallExpr',
    callee: 'SINGLE',
    args: [node],
  }
}

function hasOwnCalculationSetting<Key extends keyof WorkbookCalculationSettingsSnapshot>(
  settings: Partial<WorkbookCalculationSettingsSnapshot>,
  key: Key,
): boolean {
  return Object.hasOwn(settings, key)
}

export function normalizeWorkbookCalculationSettings(
  settings: Partial<WorkbookCalculationSettingsSnapshot>,
  base?: WorkbookCalculationSettingsSnapshot,
): WorkbookCalculationSettingsSnapshot {
  const normalized: WorkbookCalculationSettingsSnapshot = {
    mode: settings.mode ?? base?.mode ?? 'automatic',
    compatibilityMode: settings.compatibilityMode ?? base?.compatibilityMode ?? 'excel-modern',
  }
  const dateSystem = hasOwnCalculationSetting(settings, 'dateSystem') ? settings.dateSystem : base?.dateSystem
  if (dateSystem !== undefined) {
    normalized.dateSystem = dateSystem
  }
  const iterate = hasOwnCalculationSetting(settings, 'iterate') ? settings.iterate : base?.iterate
  if (iterate !== undefined) {
    normalized.iterate = iterate
  }
  const iterateCount = hasOwnCalculationSetting(settings, 'iterateCount') ? settings.iterateCount : base?.iterateCount
  if (iterateCount !== undefined) {
    normalized.iterateCount = iterateCount
  }
  const iterateDelta = hasOwnCalculationSetting(settings, 'iterateDelta') ? settings.iterateDelta : base?.iterateDelta
  if (iterateDelta !== undefined) {
    normalized.iterateDelta = iterateDelta
  }
  const fullPrecision = hasOwnCalculationSetting(settings, 'fullPrecision') ? settings.fullPrecision : base?.fullPrecision
  if (fullPrecision !== undefined) {
    normalized.fullPrecision = fullPrecision
  }
  const fullCalcOnLoad = hasOwnCalculationSetting(settings, 'fullCalcOnLoad') ? settings.fullCalcOnLoad : base?.fullCalcOnLoad
  if (fullCalcOnLoad !== undefined) {
    normalized.fullCalcOnLoad = fullCalcOnLoad
  }
  const concurrentCalc = hasOwnCalculationSetting(settings, 'concurrentCalc') ? settings.concurrentCalc : base?.concurrentCalc
  if (concurrentCalc !== undefined) {
    normalized.concurrentCalc = concurrentCalc
  }
  return normalized
}

export function calculationSettingsEqual(left: WorkbookCalculationSettingsSnapshot, right: WorkbookCalculationSettingsSnapshot): boolean {
  return (
    left.mode === right.mode &&
    (left.compatibilityMode ?? 'excel-modern') === (right.compatibilityMode ?? 'excel-modern') &&
    left.dateSystem === right.dateSystem &&
    left.iterate === right.iterate &&
    left.iterateCount === right.iterateCount &&
    left.iterateDelta === right.iterateDelta &&
    left.fullPrecision === right.fullPrecision &&
    left.fullCalcOnLoad === right.fullCalcOnLoad &&
    left.concurrentCalc === right.concurrentCalc
  )
}

export function definedNameValuesEqual(left: WorkbookDefinedNameValueSnapshot, right: WorkbookDefinedNameValueSnapshot): boolean {
  if (left === right) {
    return true
  }
  return JSON.stringify(left) === JSON.stringify(right)
}

export function definedNameValueToCellValue(input: WorkbookDefinedNameValueSnapshot, stringPool: StringPool): CellValue {
  if (typeof input === 'object' && input !== null && 'kind' in input) {
    if (input.kind === 'scalar') {
      return literalToValue(input.value, stringPool)
    }
    if (input.kind === 'formula') {
      try {
        const literalValue = literalFormulaNodeToCellValue(parseFormula(input.formula), stringPool)
        return literalValue ?? errorValue(ErrorCode.Value)
      } catch {
        return errorValue(ErrorCode.Value)
      }
    }
    return errorValue(ErrorCode.Value)
  }
  if (typeof input === 'string' && input.startsWith('=')) {
    try {
      const literalValue = literalFormulaNodeToCellValue(parseFormula(input), stringPool)
      return literalValue ?? errorValue(ErrorCode.Value)
    } catch {
      return errorValue(ErrorCode.Value)
    }
  }
  return literalToValue(input, stringPool)
}

function formulaNodeToReferenceOperand(input: FormulaNode): ReferenceOperand | undefined {
  switch (input.kind) {
    case 'CellRef':
      return {
        kind: 'cell',
        ...(input.sheetName !== undefined ? { sheetName: input.sheetName } : {}),
        address: input.ref,
      }
    case 'RangeRef':
      return {
        kind: 'range',
        refKind: input.refKind,
        start: input.start,
        end: input.end,
        ...(input.sheetName !== undefined ? { sheetName: input.sheetName } : {}),
        ...(input.sheetEndName !== undefined ? { sheetEndName: input.sheetEndName } : {}),
      }
    case 'RowRef':
      return {
        kind: 'row',
        refKind: 'rows',
        start: input.ref,
        end: input.ref,
        ...(input.sheetName !== undefined ? { sheetName: input.sheetName } : {}),
      }
    case 'ColumnRef':
      return {
        kind: 'col',
        refKind: 'cols',
        start: input.ref,
        end: input.ref,
        ...(input.sheetName !== undefined ? { sheetName: input.sheetName } : {}),
      }
    case 'BinaryExpr':
    case 'BooleanLiteral':
    case 'CallExpr':
    case 'ErrorLiteral':
    case 'InvokeExpr':
    case 'NameRef':
    case 'NumberLiteral':
    case 'OmittedArgument':
    case 'ArrayConstant':
    case 'SpillRef':
    case 'StringLiteral':
    case 'StructuredRef':
    case 'UnaryExpr':
      return undefined
  }
}

export function definedNameValueToReferenceOperand(input: WorkbookDefinedNameValueSnapshot): ReferenceOperand | undefined {
  if (typeof input === 'object' && input !== null && 'kind' in input) {
    switch (input.kind) {
      case 'cell-ref':
        return { kind: 'cell', sheetName: input.sheetName, address: input.address }
      case 'range-ref':
        return {
          kind: 'range',
          refKind: 'cells',
          sheetName: input.sheetName,
          start: input.startAddress,
          end: input.endAddress,
        }
      case 'formula':
        try {
          return formulaNodeToReferenceOperand(parseFormula(input.formula))
        } catch {
          return undefined
        }
      case 'scalar':
      case 'structured-ref':
        return undefined
    }
  }
  if (typeof input === 'string' && input.startsWith('=')) {
    try {
      return formulaNodeToReferenceOperand(parseFormula(input))
    } catch {
      return undefined
    }
  }
  return undefined
}

export function renameDefinedNameValueSheet(
  input: WorkbookDefinedNameValueSnapshot,
  oldSheetName: string,
  newSheetName: string,
): WorkbookDefinedNameValueSnapshot {
  if (typeof input === 'object' && input !== null && 'kind' in input) {
    switch (input.kind) {
      case 'scalar':
      case 'structured-ref':
        return input
      case 'cell-ref':
        return input.sheetName === oldSheetName ? { ...input, sheetName: newSheetName } : input
      case 'range-ref':
        return input.sheetName === oldSheetName ? { ...input, sheetName: newSheetName } : input
      case 'formula':
        return {
          ...input,
          formula: renameFormulaTextForSheet(input.formula, oldSheetName, newSheetName),
        }
    }
  }
  if (typeof input === 'string' && input.startsWith('=')) {
    return renameFormulaTextForSheet(input, oldSheetName, newSheetName)
  }
  return input
}

export function resolveMetadataReferencesInAst(
  node: FormulaNode,
  context: MetadataResolutionContext,
  activeNames = new Set<string>(),
  valueContext: MetadataFormulaValueContext = 'scalar',
  localNames = new Set<string>(),
): { node: FormulaNode; fullyResolved: boolean; substituted: boolean } {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'CellRef':
    case 'OmittedArgument':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
      return { node, fullyResolved: true, substituted: false }
    case 'ArrayConstant': {
      let fullyResolved = true
      let substituted = false
      const rows = node.rows.map((row) =>
        row.map((entry) => {
          const resolved = resolveMetadataReferencesInAst(entry, context, activeNames, valueContext, localNames)
          fullyResolved = fullyResolved && resolved.fullyResolved
          substituted = substituted || resolved.substituted
          return resolved.node
        }),
      )
      return {
        node: substituted ? { ...node, rows } : node,
        fullyResolved,
        substituted,
      }
    }
    case 'NameRef': {
      const normalized = normalizeDefinedName(node.name)
      if (localNames.has(normalized)) {
        return { node, fullyResolved: true, substituted: false }
      }
      if (activeNames.has(normalized)) {
        return {
          node: { kind: 'ErrorLiteral', code: ErrorCode.Cycle },
          fullyResolved: true,
          substituted: true,
        }
      }
      const literal = context.resolveName(node.name, node.sheetName)
      const replacement =
        literal === undefined
          ? ({ kind: 'ErrorLiteral', code: ErrorCode.Name } satisfies FormulaNode)
          : definedNameValueToFormulaNode(literal)
      if (!replacement) {
        return { node, fullyResolved: false, substituted: false }
      }
      const nextActiveNames = new Set(activeNames)
      nextActiveNames.add(normalized)
      const resolved = resolveMetadataReferencesInAst(replacement, context, nextActiveNames, valueContext, localNames)
      return {
        node: maybeImplicitIntersectRangeName(resolved.node, valueContext),
        fullyResolved: resolved.fullyResolved,
        substituted: true,
      }
    }
    case 'StructuredRef': {
      const replacement =
        context.resolveStructuredReference(node.tableName, node.columnName) ??
        ({ kind: 'ErrorLiteral', code: ErrorCode.Ref } satisfies FormulaNode)
      return { node: replacement, fullyResolved: true, substituted: true }
    }
    case 'SpillRef': {
      const replacement =
        context.resolveSpillReference(node.sheetName, node.ref) ?? ({ kind: 'ErrorLiteral', code: ErrorCode.Ref } satisfies FormulaNode)
      return { node: replacement, fullyResolved: true, substituted: true }
    }
    case 'UnaryExpr': {
      const resolved = resolveMetadataReferencesInAst(node.argument, context, activeNames, valueContext, localNames)
      return {
        node: resolved.substituted ? { ...node, argument: resolved.node } : node,
        fullyResolved: resolved.fullyResolved,
        substituted: resolved.substituted,
      }
    }
    case 'BinaryExpr': {
      const left = resolveMetadataReferencesInAst(node.left, context, activeNames, valueContext, localNames)
      const right = resolveMetadataReferencesInAst(node.right, context, activeNames, valueContext, localNames)
      return {
        node: left.substituted || right.substituted ? { ...node, left: left.node, right: right.node } : node,
        fullyResolved: left.fullyResolved && right.fullyResolved,
        substituted: left.substituted || right.substituted,
      }
    }
    case 'CallExpr': {
      const callee = node.callee.toUpperCase()
      const normalizedCallee = normalizeDefinedName(node.callee)
      if (!hasBuiltin(callee) && !localNames.has(normalizedCallee)) {
        if (activeNames.has(normalizedCallee)) {
          return {
            node: { kind: 'ErrorLiteral', code: ErrorCode.Cycle },
            fullyResolved: true,
            substituted: true,
          }
        }
        const calleeLiteral = context.resolveName(node.callee)
        const calleeNode = calleeLiteral === undefined ? undefined : definedNameValueToFormulaNode(calleeLiteral)
        if (calleeNode?.kind === 'CallExpr' && calleeNode.callee.toUpperCase() === 'LAMBDA') {
          const nextActiveNames = new Set(activeNames)
          nextActiveNames.add(normalizedCallee)
          const resolved = resolveMetadataReferencesInAst(
            { kind: 'InvokeExpr', callee: calleeNode, args: node.args },
            context,
            nextActiveNames,
            valueContext,
            localNames,
          )
          return {
            ...resolved,
            substituted: true,
          }
        }
      }
      if (callee === 'LAMBDA' && node.args.length >= 1) {
        let fullyResolved = true
        let substituted = false
        const scopedLocalNames = new Set(localNames)
        const args = node.args.map((arg, index) => {
          if (index < node.args.length - 1) {
            if (arg.kind === 'NameRef') {
              scopedLocalNames.add(normalizeDefinedName(arg.name))
            }
            return arg
          }
          const resolved = resolveMetadataReferencesInAst(arg, context, activeNames, valueContext, scopedLocalNames)
          fullyResolved = fullyResolved && resolved.fullyResolved
          substituted = substituted || resolved.substituted
          return resolved.node
        })
        return {
          node: substituted ? { ...node, callee, args } : node,
          fullyResolved,
          substituted,
        }
      }
      if (callee === 'LET' && node.args.length >= 3 && node.args.length % 2 === 1) {
        let fullyResolved = true
        let substituted = false
        const scopedLocalNames = new Set(localNames)
        const args: FormulaNode[] = []
        for (let index = 0; index < node.args.length - 1; index += 2) {
          const nameArg = node.args[index]!
          const valueArg = node.args[index + 1]!
          const resolved = resolveMetadataReferencesInAst(
            valueArg,
            context,
            activeNames,
            callArgumentValueContext(node.callee, index + 1, valueContext),
            scopedLocalNames,
          )
          fullyResolved = fullyResolved && resolved.fullyResolved
          substituted = substituted || resolved.substituted
          args.push(nameArg, resolved.node)
          if (nameArg.kind === 'NameRef') {
            scopedLocalNames.add(normalizeDefinedName(nameArg.name))
          }
        }
        const body = resolveMetadataReferencesInAst(node.args[node.args.length - 1]!, context, activeNames, valueContext, scopedLocalNames)
        fullyResolved = fullyResolved && body.fullyResolved
        substituted = substituted || body.substituted
        args.push(body.node)
        return {
          node: substituted ? { ...node, callee, args } : node,
          fullyResolved,
          substituted,
        }
      }
      let fullyResolved = true
      let substituted = false
      const args = node.args.map((arg, index) => {
        const resolved = resolveMetadataReferencesInAst(
          arg,
          context,
          activeNames,
          callArgumentValueContext(node.callee, index, valueContext),
          localNames,
        )
        fullyResolved = fullyResolved && resolved.fullyResolved
        substituted = substituted || resolved.substituted
        return resolved.node
      })
      return {
        node: substituted ? { ...node, args } : node,
        fullyResolved,
        substituted,
      }
    }
    case 'InvokeExpr': {
      const callee = resolveMetadataReferencesInAst(node.callee, context, activeNames, valueContext, localNames)
      let fullyResolved = callee.fullyResolved
      let substituted = callee.substituted
      const args = node.args.map((arg) => {
        const resolved = resolveMetadataReferencesInAst(arg, context, activeNames, valueContext, localNames)
        fullyResolved = fullyResolved && resolved.fullyResolved
        substituted = substituted || resolved.substituted
        return resolved.node
      })
      return {
        node: substituted ? { ...node, callee: callee.node, args } : node,
        fullyResolved,
        substituted,
      }
    }
  }
}

export function tableDependencyKey(name: string): string {
  return normalizeWorkbookObjectName(name, 'Table')
}

export function spillDependencyKey(sheetName: string, address: string): string {
  return `${sheetName}!${parseCellAddress(address, sheetName).text}`
}

export function spillDependencyKeyFromRef(ref: string, ownerSheetName: string): string {
  if (ref.includes('!')) {
    const separator = ref.indexOf('!')
    const sheetName = unquoteFormulaSheetName(ref.slice(0, separator))
    const address = ref.slice(separator + 1)
    return spillDependencyKey(sheetName, address)
  }
  return spillDependencyKey(ownerSheetName, ref)
}
