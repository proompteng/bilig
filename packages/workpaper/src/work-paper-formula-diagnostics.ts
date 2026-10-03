import { ErrorCode, ValueTag, formatCellDisplayValue, formatErrorCode, type CellValue } from '@bilig/protocol'
import { excelSerialToDateParts, parseCellAddress, parseFormula, parseRangeAddress, type FormulaNode } from '@bilig/formula'
import type { WorkPaperCellAddress, WorkPaperCellRange, WorkPaperFormulaDiagnostic } from './work-paper-types.js'

interface WorkPaperFormulaDiagnosticHooks {
  readonly getCellValue: (address: WorkPaperCellAddress) => CellValue
  readonly getCellValueFormat: (address: WorkPaperCellAddress) => string | undefined
  readonly getCellFormula: (address: WorkPaperCellAddress) => string | undefined
  readonly getRangeValues: (range: WorkPaperCellRange) => CellValue[][]
  readonly getSheetId: (sheetName: string) => number | undefined
  readonly getSheetName: (sheetId: number) => string | undefined
  readonly simpleCellAddressToString: (address: WorkPaperCellAddress, options?: { includeSheetName?: boolean }) => string
  readonly simpleCellRangeToString: (range: WorkPaperCellRange, options?: { includeSheetName?: boolean }) => string
}

interface DiagnosticBase {
  readonly address: WorkPaperCellAddress
  readonly sheetName: string
  readonly a1: string
  readonly formula?: string
  readonly functionName?: string
  readonly errorCode?: ErrorCode
  readonly errorText?: string
}

interface ResolvedCells {
  readonly rows: number
  readonly cols: number
  readonly values: readonly CellValue[]
  readonly addresses: readonly WorkPaperCellAddress[]
  readonly reference: string
}

type FinanceFunctionName = 'XIRR' | 'XNPV'
type FinanceCallNode = Extract<FormulaNode, { kind: 'CallExpr' }>
type ProviderBackedFunctionInfo = {
  readonly functionName: string
  readonly adapterSurface: string
  readonly referenceLabel: string
}

const providerBackedFunctionInfo = new Map<string, Omit<ProviderBackedFunctionInfo, 'functionName'>>([
  ['CALL', { adapterSurface: 'add-in', referenceLabel: 'legacy add-in adapter' }],
  ['COPILOT', { adapterSurface: 'host', referenceLabel: 'host AI adapter' }],
  ['CUBEKPIMEMBER', { adapterSurface: 'cube', referenceLabel: 'cube data adapter' }],
  ['CUBEMEMBER', { adapterSurface: 'cube', referenceLabel: 'cube data adapter' }],
  ['CUBEMEMBERPROPERTY', { adapterSurface: 'cube', referenceLabel: 'cube data adapter' }],
  ['CUBERANKEDMEMBER', { adapterSurface: 'cube', referenceLabel: 'cube data adapter' }],
  ['CUBESET', { adapterSurface: 'cube', referenceLabel: 'cube data adapter' }],
  ['CUBESETCOUNT', { adapterSurface: 'cube', referenceLabel: 'cube data adapter' }],
  ['CUBEVALUE', { adapterSurface: 'cube', referenceLabel: 'cube data adapter' }],
  ['DDE', { adapterSurface: 'external-data', referenceLabel: 'external data adapter' }],
  ['DETECTLANGUAGE', { adapterSurface: 'web', referenceLabel: 'language service adapter' }],
  ['FILTERXML', { adapterSurface: 'web', referenceLabel: 'XML fetch/parse adapter' }],
  ['IMAGE', { adapterSurface: 'web', referenceLabel: 'image fetch adapter' }],
  ['IMPORTRANGE', { adapterSurface: 'web', referenceLabel: 'Google Sheets range adapter' }],
  ['INFO', { adapterSurface: 'host', referenceLabel: 'host environment adapter' }],
  ['PY', { adapterSurface: 'python', referenceLabel: 'Python execution adapter' }],
  ['REGISTER.ID', { adapterSurface: 'add-in', referenceLabel: 'legacy add-in adapter' }],
  ['RTD', { adapterSurface: 'external-data', referenceLabel: 'real-time data adapter' }],
  ['STOCKHISTORY', { adapterSurface: 'web', referenceLabel: 'market data adapter' }],
  ['TRANSLATE', { adapterSurface: 'web', referenceLabel: 'translation service adapter' }],
  ['WEBSERVICE', { adapterSurface: 'web', referenceLabel: 'web request adapter' }],
])

function readFinanceFunctionName(callee: string): FinanceFunctionName | undefined {
  const normalized = callee.toUpperCase()
  if (normalized === 'XIRR' || normalized === 'XNPV') {
    return normalized
  }
  return undefined
}

function isFinanceCall(node: FormulaNode): node is FinanceCallNode {
  return node.kind === 'CallExpr' && readFinanceFunctionName(node.callee) !== undefined
}

function buildDiagnostic(
  base: DiagnosticBase,
  diagnostic: Pick<WorkPaperFormulaDiagnostic, 'code' | 'message'> & {
    references?: string[]
    adapterSurface?: string
  },
): WorkPaperFormulaDiagnostic {
  return {
    severity: 'error',
    address: base.address,
    sheetName: base.sheetName,
    a1: base.a1,
    ...(base.formula !== undefined ? { formula: base.formula } : {}),
    ...(base.functionName !== undefined ? { functionName: base.functionName } : {}),
    ...(base.errorCode !== undefined ? { errorCode: base.errorCode } : {}),
    ...(base.errorText !== undefined ? { errorText: base.errorText } : {}),
    ...diagnostic,
  }
}

function providerBackedInfoForName(name: string): ProviderBackedFunctionInfo | undefined {
  const functionName = name.trim().toUpperCase()
  const info = providerBackedFunctionInfo.get(functionName)
  return info ? { functionName, ...info } : undefined
}

function findProviderBackedFunction(node: FormulaNode): ProviderBackedFunctionInfo | undefined {
  if (node.kind === 'CallExpr') {
    const direct = providerBackedInfoForName(node.callee)
    if (direct) {
      return direct
    }
    for (const arg of node.args) {
      const nested = findProviderBackedFunction(arg)
      if (nested) {
        return nested
      }
    }
    return undefined
  }

  if (node.kind === 'InvokeExpr') {
    const callee = findProviderBackedFunction(node.callee)
    if (callee) {
      return callee
    }
    for (const arg of node.args) {
      const nested = findProviderBackedFunction(arg)
      if (nested) {
        return nested
      }
    }
    return undefined
  }

  if (node.kind === 'UnaryExpr') {
    return findProviderBackedFunction(node.argument)
  }

  if (node.kind === 'BinaryExpr') {
    return findProviderBackedFunction(node.left) ?? findProviderBackedFunction(node.right)
  }

  if (node.kind === 'ArrayConstant') {
    for (const row of node.rows) {
      for (const entry of row) {
        const nested = findProviderBackedFunction(entry)
        if (nested) {
          return nested
        }
      }
    }
  }

  return undefined
}

function diagnoseProviderBackedCall(info: ProviderBackedFunctionInfo, base: DiagnosticBase): WorkPaperFormulaDiagnostic[] {
  return [
    buildDiagnostic(
      {
        ...base,
        functionName: info.functionName,
      },
      {
        code: 'provider-backed-adapter-missing',
        adapterSurface: info.adapterSurface,
        message: `${info.functionName} is provider-backed and returned ${base.errorText ?? '#BLOCKED!'} because no ${info.referenceLabel} is installed. Install an adapter for ${info.functionName}, provide the imported data as local WorkPaper inputs, or route this workbook through the original provider before treating the cell as verified.`,
      },
    ),
  ]
}

function displayValue(value: CellValue, format: string | undefined): string {
  const rendered = formatCellDisplayValue(value, format)
  return rendered === '' ? '<blank>' : rendered
}

function describeCellValue(value: CellValue, format: string | undefined): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return 'blank'
    case ValueTag.Boolean:
      return `boolean ${value.value ? 'TRUE' : 'FALSE'}`
    case ValueTag.String:
      return `text "${value.value}"`
    case ValueTag.Error:
      return `error ${formatErrorCode(value.code)}`
    case ValueTag.Number:
      return `number ${displayValue(value, format)}`
  }
}

function resolveRangeNode(
  node: Extract<FormulaNode, { kind: 'RangeRef' }>,
  ownerSheetName: string,
  hooks: WorkPaperFormulaDiagnosticHooks,
): ResolvedCells | undefined {
  if (node.refKind !== 'cells') {
    return undefined
  }
  const parsed = parseRangeAddress(`${node.start}:${node.end}`, node.sheetName ?? ownerSheetName)
  if (parsed.kind !== 'cells') {
    return undefined
  }
  const sheetName = parsed.sheetName ?? ownerSheetName
  const sheet = hooks.getSheetId(sheetName)
  if (sheet === undefined) {
    return undefined
  }
  const range: WorkPaperCellRange = {
    start: { sheet, row: parsed.start.row, col: parsed.start.col },
    end: { sheet, row: parsed.end.row, col: parsed.end.col },
  }
  const matrix = hooks.getRangeValues(range)
  const values: CellValue[] = []
  const addresses: WorkPaperCellAddress[] = []
  for (let rowOffset = 0; rowOffset < matrix.length; rowOffset += 1) {
    const row = matrix[rowOffset] ?? []
    for (let colOffset = 0; colOffset < (row.length || range.end.col - range.start.col + 1); colOffset += 1) {
      values.push(row[colOffset] ?? { tag: ValueTag.Empty })
      addresses.push({ sheet, row: range.start.row + rowOffset, col: range.start.col + colOffset })
    }
  }
  return {
    rows: range.end.row - range.start.row + 1,
    cols: range.end.col - range.start.col + 1,
    values,
    addresses,
    reference: hooks.simpleCellRangeToString(range, { includeSheetName: true }),
  }
}

function resolveCellNode(
  node: Extract<FormulaNode, { kind: 'CellRef' }>,
  ownerSheetName: string,
  hooks: WorkPaperFormulaDiagnosticHooks,
): ResolvedCells | undefined {
  const parsed = parseCellAddress(node.ref, node.sheetName ?? ownerSheetName)
  const sheetName = parsed.sheetName ?? ownerSheetName
  const sheet = hooks.getSheetId(sheetName)
  if (sheet === undefined) {
    return undefined
  }
  const address = { sheet, row: parsed.row, col: parsed.col }
  return {
    rows: 1,
    cols: 1,
    values: [hooks.getCellValue(address)],
    addresses: [address],
    reference: hooks.simpleCellAddressToString(address, { includeSheetName: true }),
  }
}

function resolveCellArgument(
  node: FormulaNode | undefined,
  ownerSheetName: string,
  hooks: WorkPaperFormulaDiagnosticHooks,
): ResolvedCells | undefined {
  if (node?.kind === 'RangeRef') {
    return resolveRangeNode(node, ownerSheetName, hooks)
  }
  if (node?.kind === 'CellRef') {
    return resolveCellNode(node, ownerSheetName, hooks)
  }
  return undefined
}

function resolveScalarNumber(
  node: FormulaNode | undefined,
  ownerSheetName: string,
  hooks: WorkPaperFormulaDiagnosticHooks,
): number | undefined {
  if (node === undefined) {
    return undefined
  }
  if (node.kind === 'NumberLiteral') {
    return Number.isFinite(node.value) ? node.value : undefined
  }
  if (node.kind !== 'CellRef') {
    return undefined
  }
  const resolved = resolveCellNode(node, ownerSheetName, hooks)
  const value = resolved?.values[0]
  return value?.tag === ValueTag.Number && Number.isFinite(value.value) ? value.value : undefined
}

function firstInvalidNumber(cells: ResolvedCells): { value: CellValue; address: WorkPaperCellAddress } | undefined {
  for (let index = 0; index < cells.values.length; index += 1) {
    const value = cells.values[index]!
    if (value.tag !== ValueTag.Number || !Number.isFinite(value.value)) {
      return { value, address: cells.addresses[index]! }
    }
  }
  return undefined
}

function firstInvalidDateSerial(
  cells: ResolvedCells,
): { value: CellValue; address: WorkPaperCellAddress; unsupportedCoercion: boolean } | undefined {
  for (let index = 0; index < cells.values.length; index += 1) {
    const value = cells.values[index]!
    if (value.tag !== ValueTag.Number) {
      return { value, address: cells.addresses[index]!, unsupportedCoercion: value.tag === ValueTag.String }
    }
    if (!Number.isFinite(value.value) || excelSerialToDateParts(value.value) === undefined) {
      return { value, address: cells.addresses[index]!, unsupportedCoercion: false }
    }
  }
  return undefined
}

function numericValues(cells: ResolvedCells): number[] {
  return cells.values.map((value) => (value.tag === ValueTag.Number ? value.value : Number.NaN))
}

function dateSerialValues(cells: ResolvedCells): number[] {
  return cells.values.map((value) => (value.tag === ValueTag.Number ? Math.trunc(value.value) : Number.NaN))
}

function diagnoseFinanceCall(
  call: FinanceCallNode,
  base: DiagnosticBase,
  ownerSheetName: string,
  hooks: WorkPaperFormulaDiagnosticHooks,
): WorkPaperFormulaDiagnostic[] {
  const functionName = readFinanceFunctionName(call.callee)
  if (functionName === undefined) {
    return [
      buildDiagnostic(base, {
        code: 'formula-error',
        message: `Formula evaluated to ${base.errorText ?? 'an error'}.`,
      }),
    ]
  }
  const valuesArg = functionName === 'XNPV' ? call.args[1] : call.args[0]
  const datesArg = functionName === 'XNPV' ? call.args[2] : call.args[1]
  const values = resolveCellArgument(valuesArg, ownerSheetName, hooks)
  const dates = resolveCellArgument(datesArg, ownerSheetName, hooks)
  const financeBase = { ...base, functionName }

  if (!values || !dates) {
    return [
      buildDiagnostic(financeBase, {
        code: 'financial-invalid-argument',
        message: `${functionName} diagnostics require direct cell or cell-range cash-flow and date arguments.`,
      }),
    ]
  }

  if (values.rows !== dates.rows || values.cols !== dates.cols) {
    return [
      buildDiagnostic(financeBase, {
        code: 'financial-mismatched-dimensions',
        message: `${functionName} values range ${values.reference} is ${values.rows}x${values.cols}, but dates range ${dates.reference} is ${dates.rows}x${dates.cols}.`,
        references: [values.reference, dates.reference],
      }),
    ]
  }

  const invalidCashFlow = firstInvalidNumber(values)
  if (invalidCashFlow) {
    const cell = hooks.simpleCellAddressToString(invalidCashFlow.address, { includeSheetName: true })
    return [
      buildDiagnostic(financeBase, {
        code: 'financial-invalid-cash-flow',
        message: `${functionName} cash-flow range ${values.reference} contains ${describeCellValue(invalidCashFlow.value, hooks.getCellValueFormat(invalidCashFlow.address))} at ${cell}; every cash-flow cell must be numeric.`,
        references: [values.reference, cell],
      }),
    ]
  }

  const invalidDate = firstInvalidDateSerial(dates)
  if (invalidDate) {
    const cell = hooks.simpleCellAddressToString(invalidDate.address, { includeSheetName: true })
    const code = invalidDate.unsupportedCoercion ? 'financial-unsupported-date-coercion' : 'financial-invalid-date-range'
    const text = invalidDate.unsupportedCoercion
      ? `${functionName} date range ${dates.reference} contains ${describeCellValue(invalidDate.value, hooks.getCellValueFormat(invalidDate.address))} at ${cell}. Use numeric Excel serial dates; text date coercion is not supported for headless ${functionName}.`
      : `${functionName} date range ${dates.reference} contains ${describeCellValue(invalidDate.value, hooks.getCellValueFormat(invalidDate.address))} at ${cell}; dates must be valid numeric Excel serial dates.`
    return [
      buildDiagnostic(financeBase, {
        code,
        message: text,
        references: [dates.reference, cell],
      }),
    ]
  }

  const cashFlows = numericValues(values)
  if (!cashFlows.some((value) => value > 0)) {
    return [
      buildDiagnostic(financeBase, {
        code: 'financial-missing-positive-cash-flow',
        message: `${functionName} cash-flow range ${values.reference} must contain at least one positive value.`,
        references: [values.reference],
      }),
    ]
  }
  if (!cashFlows.some((value) => value < 0)) {
    return [
      buildDiagnostic(financeBase, {
        code: 'financial-missing-negative-cash-flow',
        message: `${functionName} cash-flow range ${values.reference} must contain at least one negative value.`,
        references: [values.reference],
      }),
    ]
  }

  const dateSerials = dateSerialValues(dates)
  const firstDate = dateSerials[0]
  if (firstDate !== undefined && dateSerials.some((date) => date < firstDate)) {
    return [
      buildDiagnostic(financeBase, {
        code: 'financial-invalid-date-range',
        message: `${functionName} date range ${dates.reference} contains a date before the first date; XIRR/XNPV require later cash-flow dates to be on or after the first date.`,
        references: [dates.reference],
      }),
    ]
  }

  if (functionName === 'XNPV') {
    const rate = resolveScalarNumber(call.args[0], ownerSheetName, hooks)
    if (rate === undefined || rate <= -0.999999999) {
      return [
        buildDiagnostic(financeBase, {
          code: 'financial-invalid-rate',
          message: 'XNPV discount rate must be a finite numeric value greater than -100%.',
        }),
      ]
    }
  }

  if (functionName === 'XIRR' && call.args[2] !== undefined) {
    const guess = resolveScalarNumber(call.args[2], ownerSheetName, hooks)
    if (guess === undefined) {
      return [
        buildDiagnostic(financeBase, {
          code: 'financial-invalid-rate',
          message: 'XIRR optional guess must be a finite numeric value.',
        }),
      ]
    }
  }

  return [
    buildDiagnostic(financeBase, {
      code: 'financial-nonconvergence',
      message: `${functionName} inputs are structurally valid, but the solver did not find a valid rate for the supplied cash flows and dates.`,
      references: [values.reference, dates.reference],
    }),
  ]
}

export function collectWorkPaperFormulaDiagnostics(
  address: WorkPaperCellAddress,
  hooks: WorkPaperFormulaDiagnosticHooks,
): WorkPaperFormulaDiagnostic[] {
  const value = hooks.getCellValue(address)
  const formula = hooks.getCellFormula(address)
  if (value.tag !== ValueTag.Error || formula === undefined) {
    return []
  }

  const sheetName = hooks.getSheetName(address.sheet) ?? String(address.sheet)
  const a1 = hooks.simpleCellAddressToString(address, { includeSheetName: false })
  const base: DiagnosticBase = {
    address,
    sheetName,
    a1,
    formula,
    errorCode: value.code,
    errorText: formatErrorCode(value.code),
  }

  let parsed: FormulaNode
  try {
    parsed = parseFormula(formula)
  } catch {
    return [
      buildDiagnostic(base, {
        code: 'formula-error',
        message: `Formula evaluated to ${formatErrorCode(value.code)}.`,
      }),
    ]
  }

  if (isFinanceCall(parsed)) {
    return diagnoseFinanceCall(parsed, base, sheetName, hooks)
  }

  const providerBacked = value.code === ErrorCode.Blocked ? findProviderBackedFunction(parsed) : undefined
  if (providerBacked) {
    return diagnoseProviderBackedCall(providerBacked, base)
  }

  return [
    buildDiagnostic(base, {
      code: 'formula-error',
      message: `Formula evaluated to ${formatErrorCode(value.code)}.`,
    }),
  ]
}
