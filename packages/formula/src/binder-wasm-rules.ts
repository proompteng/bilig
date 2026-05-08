import { MAX_WASM_RANGE_CELLS } from '@bilig/protocol'
import type { FormulaNode } from './ast.js'
import { parseRangeAddress } from './addressing.js'
import { isWasmSafeBuiltinArity } from './binder-wasm-arity.js'
export { isWasmSafeBuiltinArity } from './binder-wasm-arity.js'

const RANGE_SAFE_BUILTINS = new Set([
  'SUM',
  'AVG',
  'AVERAGE',
  'MIN',
  'MAX',
  'COUNT',
  'COUNTA',
  'COUNTBLANK',
  'PRODUCT',
  'GEOMEAN',
  'HARMEAN',
  'SUMSQ',
  'GCD',
  'LCM',
  'MODE',
  'MODE.SNGL',
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
])

const AXIS_AGGREGATE_CODES = new Map<string, number>([
  ['SUM', 1],
  ['AVERAGE', 2],
  ['AVG', 2],
  ['MIN', 3],
  ['MAX', 4],
  ['COUNT', 5],
  ['COUNTA', 6],
])

interface WasmSafeBuiltinArgsDeps {
  isWasmSafe: (node: FormulaNode, allowRange?: boolean) => boolean
}

function isTextLiteralLike(node: FormulaNode): boolean {
  switch (node.kind) {
    case 'StringLiteral':
      return true
    case 'BinaryExpr':
      return node.operator === '&' && isTextLiteralLike(node.left) && isTextLiteralLike(node.right)
    case 'CallExpr':
      return node.callee.toUpperCase() === 'CONCAT' && node.args.every(isTextLiteralLike)
    case 'BooleanLiteral':
    case 'ArrayConstant':
    case 'CellRef':
    case 'ColumnRef':
    case 'ErrorLiteral':
    case 'InvokeExpr':
    case 'NameRef':
    case 'NumberLiteral':
    case 'OmittedArgument':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StructuredRef':
    case 'UnaryExpr':
      return false
    default:
      return false
  }
}

function hasMixedConditionalResultKinds(args: readonly FormulaNode[], resultIndexes: readonly number[]): boolean {
  const resultNodes = resultIndexes.map((index) => args[index]).filter((node): node is FormulaNode => node !== undefined)
  if (resultNodes.length === 0) {
    return false
  }
  const hasTextLiteralResult = resultNodes.some(isTextLiteralLike)
  const hasNonTextLiteralResult = resultNodes.some((node) => node.kind !== 'ErrorLiteral' && !isTextLiteralLike(node))
  return hasTextLiteralResult && hasNonTextLiteralResult
}

export function isCellRangeNode(node: FormulaNode): boolean {
  if (node.kind !== 'RangeRef' || node.sheetEndName !== undefined) {
    return false
  }
  try {
    const sheetPrefix = node.sheetName ? `${node.sheetName}!` : ''
    return parseRangeAddress(`${sheetPrefix}${node.start}:${node.end}`).kind === 'cells'
  } catch {
    return false
  }
}

export function getNativeAxisAggregateCode(node: FormulaNode): number | null {
  if (node.kind !== 'CallExpr' || node.callee.toUpperCase() !== 'LAMBDA' || node.args.length !== 2) {
    return null
  }
  const [param, body] = node.args
  if (param?.kind !== 'NameRef' || body?.kind !== 'CallExpr' || body.args.length !== 1) {
    return null
  }
  const aggregateCode = AXIS_AGGREGATE_CODES.get(body.callee.toUpperCase())
  if (aggregateCode === undefined) {
    return null
  }
  return body.args[0]?.kind === 'NameRef' && body.args[0].name.trim().toUpperCase() === param.name.trim().toUpperCase()
    ? aggregateCode
    : null
}

export function getNativeRunningFoldCode(node: FormulaNode): number | null {
  if (node.kind !== 'CallExpr' || node.callee.toUpperCase() !== 'LAMBDA' || node.args.length !== 3) {
    return null
  }
  const [acc, value, body] = node.args
  if (acc?.kind !== 'NameRef' || value?.kind !== 'NameRef' || body?.kind !== 'BinaryExpr') {
    return null
  }
  const foldCode = body.operator === '+' ? 1 : body.operator === '*' ? 2 : null
  if (foldCode === null) {
    return null
  }
  const left = body.left
  const right = body.right
  const accName = acc.name.trim().toUpperCase()
  const valueName = value.name.trim().toUpperCase()
  return left.kind === 'NameRef' &&
    right.kind === 'NameRef' &&
    ((left.name.trim().toUpperCase() === accName && right.name.trim().toUpperCase() === valueName) ||
      (left.name.trim().toUpperCase() === valueName && right.name.trim().toUpperCase() === accName))
    ? foldCode
    : null
}

export function isNativeMakearraySumLambda(node: FormulaNode): boolean {
  if (node.kind !== 'CallExpr' || node.callee.toUpperCase() !== 'LAMBDA' || node.args.length !== 3) {
    return false
  }
  const [rowParam, colParam, body] = node.args
  if (rowParam?.kind !== 'NameRef' || colParam?.kind !== 'NameRef' || body?.kind !== 'BinaryExpr') {
    return false
  }
  if (body.operator !== '+') {
    return false
  }
  const left = body.left
  const right = body.right
  const rowName = rowParam.name.trim().toUpperCase()
  const colName = colParam.name.trim().toUpperCase()
  return (
    left.kind === 'NameRef' &&
    right.kind === 'NameRef' &&
    ((left.name.trim().toUpperCase() === rowName && right.name.trim().toUpperCase() === colName) ||
      (left.name.trim().toUpperCase() === colName && right.name.trim().toUpperCase() === rowName))
  )
}

export function isCellVectorNode(node: FormulaNode): boolean {
  if (node.kind !== 'RangeRef' || node.sheetEndName !== undefined) {
    return false
  }
  try {
    const sheetPrefix = node.sheetName ? `${node.sheetName}!` : ''
    const range = parseRangeAddress(`${sheetPrefix}${node.start}:${node.end}`)
    return range.kind === 'cells' && (range.start.row === range.end.row || range.start.col === range.end.col)
  } catch {
    return false
  }
}

function isAxisRangeNode(node: FormulaNode): boolean {
  if (node.kind !== 'RangeRef' || node.sheetEndName !== undefined) {
    return false
  }
  try {
    const sheetPrefix = node.sheetName ? `${node.sheetName}!` : ''
    const range = parseRangeAddress(`${sheetPrefix}${node.start}:${node.end}`)
    return range.kind === 'rows' || range.kind === 'cols'
  } catch {
    return false
  }
}

function isWasmSafeCellVectorNode(node: FormulaNode): boolean {
  if (node.kind !== 'RangeRef' || node.sheetEndName !== undefined) {
    return false
  }
  try {
    const sheetPrefix = node.sheetName ? `${node.sheetName}!` : ''
    const range = parseRangeAddress(`${sheetPrefix}${node.start}:${node.end}`)
    if (range.kind !== 'cells') {
      return false
    }
    const rows = range.end.row - range.start.row + 1
    const cols = range.end.col - range.start.col + 1
    return (rows === 1 || cols === 1) && rows * cols <= MAX_WASM_RANGE_CELLS
  } catch {
    return false
  }
}

function isCanonicalAggregateName(node: FormulaNode | undefined, expected: string): boolean {
  return node?.kind === 'NameRef' && node.name.trim().toUpperCase() === expected.trim().toUpperCase()
}

function isIntegerLiteral(node: FormulaNode | undefined, expected: number): boolean {
  return node?.kind === 'NumberLiteral' && Math.trunc(node.value) === expected && node.value === expected
}

export type NativeGroupedArrayKind = 'groupby-sum-canonical' | 'pivotby-sum-canonical'

export function getNativeGroupedArrayKind(node: FormulaNode): NativeGroupedArrayKind | null {
  if (node.kind !== 'CallExpr') {
    return null
  }

  const callee = node.callee.toUpperCase()
  if (
    callee === 'GROUPBY' &&
    node.args.length === 5 &&
    isWasmSafeCellVectorNode(node.args[0]!) &&
    isWasmSafeCellVectorNode(node.args[1]!) &&
    isCanonicalAggregateName(node.args[2], 'SUM') &&
    isIntegerLiteral(node.args[3], 3) &&
    isIntegerLiteral(node.args[4], 1)
  ) {
    return 'groupby-sum-canonical'
  }

  if (
    callee === 'PIVOTBY' &&
    node.args.length === 8 &&
    isWasmSafeCellVectorNode(node.args[0]!) &&
    isWasmSafeCellVectorNode(node.args[1]!) &&
    isWasmSafeCellVectorNode(node.args[2]!) &&
    isCanonicalAggregateName(node.args[3], 'SUM') &&
    isIntegerLiteral(node.args[4], 3) &&
    isIntegerLiteral(node.args[5], 1) &&
    isIntegerLiteral(node.args[6], 0) &&
    isIntegerLiteral(node.args[7], 1)
  ) {
    return 'pivotby-sum-canonical'
  }

  return null
}

export function isWasmSafeBuiltinArgs(callee: string, args: readonly FormulaNode[], deps: WasmSafeBuiltinArgsDeps): boolean {
  const argc = args.length
  const isScalarArg = (arg: FormulaNode): boolean => deps.isWasmSafe(arg)
  const isCellRangeArg = (arg: FormulaNode): boolean => deps.isWasmSafe(arg, true) && isCellRangeNode(arg)
  const isCellVectorArg = (arg: FormulaNode): boolean => deps.isWasmSafe(arg, true) && isCellVectorNode(arg)
  const isCellOrScalarArg = (arg: FormulaNode): boolean => isCellVectorArg(arg) || isScalarArg(arg)
  const isNativeSequenceArg = (arg: FormulaNode): boolean =>
    arg.kind === 'CallExpr' &&
    arg.callee.toUpperCase() === 'SEQUENCE' &&
    isWasmSafeBuiltinArity('SEQUENCE', arg.args.length) &&
    arg.args.every((child) => deps.isWasmSafe(child))

  switch (callee) {
    case 'SUM':
    case 'MIN':
    case 'MAX':
    case 'COUNT':
    case 'COUNTA':
    case 'COUNTBLANK':
      return args.every((arg) => deps.isWasmSafe(arg, true) || isNativeSequenceArg(arg))
    case 'AVERAGE':
    case 'AVG':
      return args.every((arg) => isAxisRangeNode(arg) || isNativeSequenceArg(arg))
    case 'CHOOSE':
      return argc >= 2 && isScalarArg(args[0]!) && args.slice(1).every((arg) => deps.isWasmSafe(arg, true) || isNativeSequenceArg(arg))
    case 'IF':
      return (
        argc === 3 &&
        !hasMixedConditionalResultKinds(args, [1, 2]) &&
        isScalarArg(args[0]!) &&
        args.slice(1).every((arg) => deps.isWasmSafe(arg, true))
      )
    case 'IFS': {
      if (
        argc < 2 ||
        argc % 2 !== 0 ||
        hasMixedConditionalResultKinds(
          args,
          Array.from({ length: argc / 2 }, (_, index) => index * 2 + 1),
        )
      ) {
        return false
      }
      return args.every((arg, index) => (index % 2 === 0 ? isScalarArg(arg) : deps.isWasmSafe(arg, true)))
    }
    case 'COUNTIF':
    case 'USE.THE.COUNTIF':
      return args.length === 2 && isCellRangeArg(args[0]!) && isScalarArg(args[1]!)
    case 'COUNTIFS':
      if (args.length === 0 || args.length % 2 !== 0) {
        return false
      }
      return args.every((arg, index) => (index % 2 === 0 ? isCellRangeArg(arg) : isScalarArg(arg)))
    case 'DAVERAGE':
    case 'DCOUNT':
    case 'DCOUNTA':
    case 'DGET':
    case 'DMAX':
    case 'DMIN':
    case 'DPRODUCT':
    case 'DSTDEV':
    case 'DSTDEVP':
    case 'DSUM':
    case 'DVAR':
    case 'DVARP':
      return argc === 3 && isCellRangeArg(args[0]!) && deps.isWasmSafe(args[1]!, true) && isCellRangeArg(args[2]!)
    case 'CHISQ.TEST':
    case 'CHITEST':
    case 'LEGACY.CHITEST':
    case 'F.TEST':
    case 'FTEST':
      return argc === 2 && args.every((arg) => deps.isWasmSafe(arg, true))
    case 'T.TEST':
    case 'TTEST':
      return argc === 4 && isCellRangeArg(args[0]!) && isCellRangeArg(args[1]!) && isScalarArg(args[2]!) && isScalarArg(args[3]!)
    case 'Z.TEST':
    case 'ZTEST':
      return (argc === 2 || argc === 3) && args.every((arg) => deps.isWasmSafe(arg, true))
    case 'SUMIF':
    case 'AVERAGEIF':
      if (args.length !== 2 && args.length !== 3) {
        return false
      }
      return isCellRangeArg(args[0]!) && isScalarArg(args[1]!) && (args.length === 2 || isCellRangeArg(args[2]!))
    case 'SUMIFS':
    case 'AVERAGEIFS':
      if (args.length < 3 || args.length % 2 === 0) {
        return false
      }
      if (!isCellRangeArg(args[0]!)) {
        return false
      }
      return args.slice(1).every((arg, index) => (index % 2 === 0 ? isCellRangeArg(arg) : isScalarArg(arg)))
    case 'SUMPRODUCT':
      return args.length >= 1 && args.every((arg) => isCellRangeArg(arg))
    case 'MATCH':
      return (
        (args.length === 2 || args.length === 3) &&
        isScalarArg(args[0]!) &&
        isCellVectorArg(args[1]!) &&
        (args.length === 2 || isScalarArg(args[2]!))
      )
    case 'CORREL':
    case 'COVAR':
    case 'PEARSON':
    case 'COVARIANCE.P':
    case 'COVARIANCE.S':
    case 'INTERCEPT':
    case 'RSQ':
    case 'SLOPE':
    case 'STEYX':
      return args.length === 2 && args.every((arg) => deps.isWasmSafe(arg, true))
    case 'MEDIAN':
      return args.length >= 1 && args.every((arg) => deps.isWasmSafe(arg, true))
    case 'MODE.MULT':
      return args.length >= 1 && args.every((arg) => deps.isWasmSafe(arg, true))
    case 'FREQUENCY':
      return args.length === 2 && deps.isWasmSafe(args[0]!, true) && deps.isWasmSafe(args[1]!, true)
    case 'BESSELI':
    case 'BESSELJ':
    case 'BESSELK':
    case 'BESSELY':
      return args.length === 2 && isScalarArg(args[0]!) && isScalarArg(args[1]!)
    case 'SMALL':
    case 'LARGE':
    case 'PERCENTILE':
    case 'PERCENTILE.INC':
    case 'PERCENTILE.EXC':
    case 'QUARTILE':
    case 'QUARTILE.INC':
    case 'QUARTILE.EXC':
      return args.length === 2 && deps.isWasmSafe(args[0]!, true) && isScalarArg(args[1]!)
    case 'PERCENTRANK':
    case 'PERCENTRANK.INC':
    case 'PERCENTRANK.EXC':
      return (
        (args.length === 2 || args.length === 3) &&
        deps.isWasmSafe(args[0]!, true) &&
        isScalarArg(args[1]!) &&
        (args.length === 2 || isScalarArg(args[2]!))
      )
    case 'RANK':
    case 'RANK.EQ':
    case 'RANK.AVG':
      return (
        (args.length === 2 || args.length === 3) &&
        isScalarArg(args[0]!) &&
        deps.isWasmSafe(args[1]!, true) &&
        (args.length === 2 || isScalarArg(args[2]!))
      )
    case 'FORECAST':
    case 'FORECAST.LINEAR':
      return args.length === 3 && isScalarArg(args[0]!) && deps.isWasmSafe(args[1]!, true) && deps.isWasmSafe(args[2]!, true)
    case 'TREND':
    case 'GROWTH':
    case 'LINEST':
    case 'LOGEST':
      return (
        args.length >= 1 &&
        args.length <= 4 &&
        deps.isWasmSafe(args[0]!, true) &&
        (args.length < 2 || deps.isWasmSafe(args[1]!, true)) &&
        (args.length < 3 || deps.isWasmSafe(args[2]!, true)) &&
        (args.length < 4 || isScalarArg(args[3]!))
      )
    case 'XMATCH':
      return (
        args.length >= 2 &&
        args.length <= 4 &&
        isScalarArg(args[0]!) &&
        isCellVectorArg(args[1]!) &&
        args.slice(2).every((arg) => isScalarArg(arg))
      )
    case 'XLOOKUP':
      return (
        args.length >= 3 &&
        args.length <= 6 &&
        isScalarArg(args[0]!) &&
        isCellVectorArg(args[1]!) &&
        isCellVectorArg(args[2]!) &&
        (args.length < 5 || args[4]?.kind === 'OmittedArgument' || isIntegerLiteral(args[4], 0)) &&
        args.slice(3).every((arg) => isScalarArg(arg))
      )
    case 'INDEX':
      return (
        (args.length === 2 || args.length === 3) &&
        isCellRangeArg(args[0]!) &&
        isScalarArg(args[1]!) &&
        (args.length === 2 || isScalarArg(args[2]!))
      )
    case 'VLOOKUP':
    case 'HLOOKUP':
      return (
        (args.length === 3 || args.length === 4) &&
        isScalarArg(args[0]!) &&
        isCellRangeArg(args[1]!) &&
        isScalarArg(args[2]!) &&
        (args.length === 3 || isScalarArg(args[3]!))
      )
    case 'DAYS':
    case 'DAYS360':
    case 'YEARFRAC':
    case 'DISC':
    case 'INTRATE':
    case 'RECEIVED':
    case 'PRICEDISC':
    case 'YIELDDISC':
    case 'COUPDAYBS':
    case 'COUPDAYS':
    case 'COUPDAYSNC':
    case 'COUPNCD':
    case 'COUPNUM':
    case 'COUPPCD':
    case 'PRICEMAT':
    case 'YIELDMAT':
    case 'ODDFPRICE':
    case 'ODDFYIELD':
    case 'ODDLPRICE':
    case 'ODDLYIELD':
    case 'PRICE':
    case 'YIELD':
    case 'DURATION':
    case 'MDURATION':
    case 'TBILLPRICE':
    case 'TBILLYIELD':
    case 'TBILLEQ':
    case 'ISOWEEKNUM':
    case 'TIMEVALUE':
    case 'WEEKNUM':
      return args.every((arg) => isScalarArg(arg))
    case 'EXPAND':
      return (
        argc >= 2 &&
        argc <= 4 &&
        deps.isWasmSafe(args[0]!, true) &&
        isScalarArg(args[1]!) &&
        (argc < 3 || isScalarArg(args[2]!)) &&
        (argc < 4 || isScalarArg(args[3]!))
      )
    case 'WORKDAY':
    case 'NETWORKDAYS':
    case 'WORKDAY.INTL':
    case 'NETWORKDAYS.INTL':
    case 'NUMBERVALUE':
    case 'TEXT':
    case 'VALUETOTEXT':
    case 'TEXTBEFORE':
    case 'TEXTAFTER':
    case 'CHAR':
    case 'CODE':
    case 'UNICODE':
    case 'UNICHAR':
    case 'CLEAN':
    case 'ASC':
    case 'JIS':
    case 'DBCS':
    case 'BAHTTEXT':
    case 'TEXTSPLIT':
    case 'CONVERT':
    case 'EUROCONVERT':
    case 'BASE':
    case 'DECIMAL':
    case 'BIN2DEC':
    case 'BIN2HEX':
    case 'BIN2OCT':
    case 'DEC2BIN':
    case 'DEC2HEX':
    case 'DEC2OCT':
    case 'HEX2BIN':
    case 'HEX2DEC':
    case 'HEX2OCT':
    case 'OCT2BIN':
    case 'OCT2DEC':
    case 'OCT2HEX':
    case 'BITAND':
    case 'BITOR':
    case 'BITXOR':
    case 'BITLSHIFT':
    case 'BITRSHIFT':
      return args.every((arg) => isScalarArg(arg))
    case 'PHONETIC':
      return argc === 1 && (isScalarArg(args[0]!) || isCellRangeNode(args[0]!))
    case 'TEXTJOIN':
      return (
        argc >= 3 &&
        isScalarArg(args[0]!) &&
        isScalarArg(args[1]!) &&
        args.slice(2).every((arg) => deps.isWasmSafe(arg, true) || isNativeSequenceArg(arg))
      )
    case 'REPLACE':
    case 'SUBSTITUTE':
    case 'REPT':
      return args.every((arg) => isScalarArg(arg))
    case 'OFFSET':
      return false
    case 'TAKE':
    case 'DROP':
    case 'CHOOSECOLS':
    case 'CHOOSEROWS':
    case 'SORT':
    case 'TOCOL':
    case 'TOROW':
    case 'WRAPROWS':
    case 'WRAPCOLS':
      if (args.length === 0) {
        return false
      }
      return isCellRangeArg(args[0]!) && args.slice(1).every((arg) => isScalarArg(arg))
    case 'FILTER':
      return (
        (argc === 2 || argc === 3) && isCellRangeArg(args[0]!) && deps.isWasmSafe(args[1]!, true) && (argc === 2 || isScalarArg(args[2]!))
      )
    case 'UNIQUE':
      return argc >= 1 && argc <= 3 && isCellRangeArg(args[0]!) && args.slice(1).every((arg) => isScalarArg(arg))
    case 'TRIMRANGE':
      return argc >= 1 && argc <= 3 && deps.isWasmSafe(args[0]!, true) && args.slice(1).every((arg) => isScalarArg(arg))
    case 'PROB':
      return (
        (argc === 3 || argc === 4) &&
        deps.isWasmSafe(args[0]!, true) &&
        deps.isWasmSafe(args[1]!, true) &&
        isScalarArg(args[2]!) &&
        (argc === 3 || isScalarArg(args[3]!))
      )
    case 'TRIMMEAN':
      return argc === 2 && deps.isWasmSafe(args[0]!, true) && isScalarArg(args[1]!)
    case 'LOOKUP':
      if (argc < 2 || argc > 3) {
        return false
      }
      return isScalarArg(args[0]!) && isCellOrScalarArg(args[1]!) && (argc === 2 || isCellVectorArg(args[2]!) || isScalarArg(args[2]!))
    case 'TRANSPOSE':
      return args.length === 1 && deps.isWasmSafe(args[0]!, true)
    case 'HSTACK':
    case 'VSTACK':
      return args.length >= 1 && args.every((arg) => deps.isWasmSafe(arg, true))
    case 'AREAS':
    case 'COLUMNS':
    case 'ROWS':
      return args.length === 1 && isCellRangeArg(args[0]!)
    case 'ARRAYTOTEXT':
      return (argc === 1 || argc === 2) && (isCellRangeArg(args[0]!) || isScalarArg(args[0]!)) && (argc === 1 || isScalarArg(args[1]!))
    case 'MINIFS':
    case 'MAXIFS':
      if (args.length < 3 || args.length % 2 === 0 || !isCellRangeArg(args[0]!)) {
        return false
      }
      return args.slice(1).every((arg, index) => (index % 2 === 0 ? isCellRangeArg(arg) : isScalarArg(arg)))
    case 'IRR':
      return (argc === 1 || argc === 2) && isCellRangeArg(args[0]!) && (argc === 1 || isScalarArg(args[1]!))
    case 'MIRR':
      return argc === 3 && isCellRangeArg(args[0]!) && isScalarArg(args[1]!) && isScalarArg(args[2]!)
    case 'XNPV':
      return argc === 3 && isScalarArg(args[0]!) && isCellRangeArg(args[1]!) && isCellRangeArg(args[2]!)
    case 'XIRR':
      return (argc === 2 || argc === 3) && isCellRangeArg(args[0]!) && isCellRangeArg(args[1]!) && (argc === 2 || isScalarArg(args[2]!))
    case 'SORTBY':
      if (args.length < 2) {
        return false
      }
      return (
        isCellRangeArg(args[0]!) &&
        args.slice(1).every((arg, index) => (index % 2 === 0 ? isScalarArg(arg) || deps.isWasmSafe(arg, true) : isScalarArg(arg)))
      )
    default:
      return args.every((arg) => deps.isWasmSafe(arg, RANGE_SAFE_BUILTINS.has(callee)))
  }
}
