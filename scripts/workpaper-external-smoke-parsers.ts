import { isNumberArray, isStringArray, parseJsonRecord, parseRecordValue, sameJson } from './workpaper-external-smoke-parser-helpers.ts'

export {
  parseNodeAgentToolCallOutput,
  parseNodeAgentVerificationOutput,
  parseNodeMcpStdioOutput,
  parseNodeMcpTranscriptOutput,
  parseNodeRevenueScenarioOutput,
} from './workpaper-agent-smoke-parsers.ts'
export type {
  AgentToolCallSummary,
  AgentVerificationSummary,
  McpTranscriptSummary,
  RevenueScenarioSummary,
} from './workpaper-agent-smoke-parsers.ts'
export { parseJsonRecord } from './workpaper-external-smoke-parser-helpers.ts'

export function parseNodeSmokeOutput(output: string): {
  afterAgentEdit: {
    enterpriseArpa: number
    qualifiedCustomerCounts: number[]
    targetRevenue: number
    totalRevenue: number
    westCustomers: number
  }
  initial: {
    targetRevenue: number
    totalRevenue: number
    westCustomers: number
  }
  persistedNamedExpressions: string[]
  persistedSheets: string[]
  restoredGrowthRatePercent: number
} {
  const parsed = parseJsonRecord(output, 'node smoke output')
  const initial = parseRecordValue(parsed.initial, 'node smoke initial output')
  const afterAgentEdit = parseRecordValue(parsed.afterAgentEdit, 'node smoke edited output')
  const persistedSheets = parsed.persistedSheets
  const persistedNamedExpressions = parsed.persistedNamedExpressions
  const qualifiedCustomerCounts = afterAgentEdit.qualifiedCustomerCounts

  if (
    typeof initial.totalRevenue !== 'number' ||
    typeof initial.westCustomers !== 'number' ||
    typeof initial.targetRevenue !== 'number' ||
    typeof afterAgentEdit.totalRevenue !== 'number' ||
    typeof afterAgentEdit.westCustomers !== 'number' ||
    typeof afterAgentEdit.enterpriseArpa !== 'number' ||
    typeof afterAgentEdit.targetRevenue !== 'number' ||
    !isNumberArray(qualifiedCustomerCounts) ||
    !isStringArray(persistedSheets) ||
    !isStringArray(persistedNamedExpressions) ||
    typeof parsed.restoredGrowthRatePercent !== 'number'
  ) {
    throw new Error(`Unexpected node smoke output: ${output}`)
  }

  return {
    initial: {
      totalRevenue: initial.totalRevenue,
      westCustomers: initial.westCustomers,
      targetRevenue: initial.targetRevenue,
    },
    afterAgentEdit: {
      totalRevenue: afterAgentEdit.totalRevenue,
      westCustomers: afterAgentEdit.westCustomers,
      enterpriseArpa: afterAgentEdit.enterpriseArpa,
      targetRevenue: afterAgentEdit.targetRevenue,
      qualifiedCustomerCounts,
    },
    persistedSheets,
    persistedNamedExpressions,
    restoredGrowthRatePercent: parsed.restoredGrowthRatePercent,
  }
}

export function parseNodeHttpJsonSummaryOutput(output: string): {
  computed: {
    committedMrr: number
    largestOpportunityMrr: number
    weightedPipelineMrr: number
    westSeats: number
  }
  sourceRecords: number
  verified: boolean
} {
  const parsed = parseJsonRecord(output, 'node HTTP JSON summary output')
  const computed = parseRecordValue(parsed.computed, 'node HTTP JSON summary computed output')

  if (
    parsed.verified !== true ||
    parsed.sourceRecords !== 3 ||
    computed.committedMrr !== 39600 ||
    computed.weightedPipelineMrr !== 43400 ||
    computed.westSeats !== 27 ||
    computed.largestOpportunityMrr !== 21600
  ) {
    throw new Error(`Unexpected node HTTP JSON summary output: ${output}`)
  }

  return {
    computed: {
      committedMrr: computed.committedMrr,
      largestOpportunityMrr: computed.largestOpportunityMrr,
      weightedPipelineMrr: computed.weightedPipelineMrr,
      westSeats: computed.westSeats,
    },
    sourceRecords: parsed.sourceRecords,
    verified: parsed.verified,
  }
}

export function parseNodeJsonFileOutput(output: string): {
  computed: {
    committedMrr: number
    largestOpportunityMrr: number
    weightedPipelineMrr: number
    westSeats: number
  }
  source: string
  sourceRecords: number
  verified: boolean
} {
  const parsed = parseJsonRecord(output, 'node JSON file output')
  const computed = parseRecordValue(parsed.computed, 'node JSON file computed output')

  if (
    parsed.verified !== true ||
    parsed.source !== 'fixtures/opportunities.json' ||
    parsed.sourceRecords !== 3 ||
    computed.committedMrr !== 39600 ||
    computed.weightedPipelineMrr !== 43400 ||
    computed.westSeats !== 27 ||
    computed.largestOpportunityMrr !== 21600
  ) {
    throw new Error(`Unexpected node JSON file output: ${output}`)
  }

  return {
    computed: {
      committedMrr: computed.committedMrr,
      largestOpportunityMrr: computed.largestOpportunityMrr,
      weightedPipelineMrr: computed.weightedPipelineMrr,
      westSeats: computed.westSeats,
    },
    source: parsed.source,
    sourceRecords: parsed.sourceRecords,
    verified: parsed.verified,
  }
}

export function parseNodeFormulaDiagnosticsOutput(output: string): {
  invalidDiagnostics: {
    code: string
    errorText: string
    functionName: string
    references: string[]
  }[]
  invalidDisplay: string
  validDisplay: string
  validValue: number
  verified: boolean
} {
  const parsed = parseJsonRecord(output, 'node formula diagnostics output')
  const invalidDiagnostics = parseFormulaDiagnostics(parsed.invalidDiagnostics)

  if (
    parsed.verified !== true ||
    parsed.invalidDisplay !== '#VALUE!' ||
    invalidDiagnostics.length !== 1 ||
    invalidDiagnostics[0]?.code !== 'financial-unsupported-date-coercion' ||
    invalidDiagnostics[0]?.functionName !== 'XIRR' ||
    invalidDiagnostics[0]?.errorText !== '#VALUE!' ||
    JSON.stringify(invalidDiagnostics[0]?.references) !== JSON.stringify(['Tax!D2:D5', 'Tax!D2']) ||
    parsed.validDisplay !== '0.02256857579464' ||
    parsed.validValue !== 0.02256857579463996
  ) {
    throw new Error(`Unexpected node formula diagnostics output: ${output}`)
  }

  return {
    invalidDiagnostics,
    invalidDisplay: parsed.invalidDisplay,
    validDisplay: parsed.validDisplay,
    validValue: parsed.validValue,
    verified: parsed.verified,
  }
}

export function parseNodeMarkdownReportOutput(output: string): {
  report: string
  verified: boolean
} {
  const parsed = parseJsonRecord(output, 'node Markdown report output')
  const expectedReport = [
    '| Metric | Value |',
    '| --- | ---: |',
    '| Committed MRR | $39,600 |',
    '| Weighted pipeline MRR | $43,400 |',
    '| Target gap | $10,400 |',
  ].join('\n')

  if (parsed.verified !== true || parsed.report !== expectedReport) {
    throw new Error(`Unexpected node Markdown report output: ${output}`)
  }

  return {
    report: parsed.report,
    verified: parsed.verified,
  }
}

export function parseNodeSnapshotDiffOutput(output: string): {
  afterSerializedInput: number
  beforeSerializedInput: number
  changedCell: string
  changedSummaryValues: {
    after: {
      annualizedArr: number
      netMrr: number
    }
    before: {
      annualizedArr: number
      netMrr: number
    }
  }
  documentBytes: {
    after: number
    before: number
  }
  verified: boolean
} {
  const parsed = parseJsonRecord(output, 'node snapshot diff output')
  const changedSummaryValues = parseRecordValue(parsed.changedSummaryValues, 'node snapshot diff changed summary values')
  const before = parseSnapshotDiffSummary(changedSummaryValues.before, 'node snapshot diff before summary')
  const after = parseSnapshotDiffSummary(changedSummaryValues.after, 'node snapshot diff after summary')
  const documentBytes = parseRecordValue(parsed.documentBytes, 'node snapshot diff document bytes')

  if (
    parsed.verified !== true ||
    parsed.changedCell !== 'Revenue!B2' ||
    parsed.beforeSerializedInput !== 12000 ||
    parsed.afterSerializedInput !== 15000 ||
    before.netMrr !== 14200 ||
    before.annualizedArr !== 170400 ||
    after.netMrr !== 17200 ||
    after.annualizedArr !== 206400 ||
    typeof documentBytes.before !== 'number' ||
    documentBytes.before <= 0 ||
    typeof documentBytes.after !== 'number' ||
    documentBytes.after <= 0
  ) {
    throw new Error(`Unexpected node snapshot diff output: ${output}`)
  }

  return {
    afterSerializedInput: parsed.afterSerializedInput,
    beforeSerializedInput: parsed.beforeSerializedInput,
    changedCell: parsed.changedCell,
    changedSummaryValues: { after, before },
    documentBytes: {
      after: documentBytes.after,
      before: documentBytes.before,
    },
    verified: parsed.verified,
  }
}

export function parseNodeRangeReadbackOutput(output: string): {
  range: string
  serializedReadback: unknown[][]
  valueReadback: unknown[][]
  verified: boolean
} {
  const parsed = parseJsonRecord(output, 'node range readback output')
  const expectedValueReadback = [
    ['Metric', 'Value'],
    ['Total MRR', 31500],
    ['West Customers', 20],
  ]
  const expectedSerializedReadback = [
    ['Metric', 'Value'],
    ['Total MRR', '=SUM(Revenue!D2:D3)'],
    ['West Customers', '=Revenue!B2'],
  ]

  if (
    parsed.verified !== true ||
    parsed.range !== 'Summary!A1:B3' ||
    !sameJson(parsed.valueReadback, expectedValueReadback) ||
    !sameJson(parsed.serializedReadback, expectedSerializedReadback)
  ) {
    throw new Error(`Unexpected node range readback output: ${output}`)
  }

  return {
    range: parsed.range,
    serializedReadback: expectedSerializedReadback,
    valueReadback: expectedValueReadback,
    verified: parsed.verified,
  }
}

export function parseNodeMcpChallengeOutput(output: string): {
  editedCell: string
  dependentCell: string
  before: number
  after: number
  afterRestart: number
  displayValue: string
  toolNames: string[]
  resourceUris: string[]
  promptNames: string[]
  persistence: {
    persisted: boolean
    serializedBytes: number
  }
  verified: boolean
} {
  const parsed = parseJsonRecord(output, 'node MCP challenge output')
  const checks = parseRecordValue(parsed.checks, 'node MCP challenge checks')
  const persistence = parseRecordValue(parsed.persistence, 'node MCP challenge persistence')
  const toolNames = parsed.tools
  const resourceUris = parsed.resources
  const promptNames = parsed.prompts

  if (
    parsed.verified !== true ||
    parsed.editedCell !== 'Inputs!B3' ||
    parsed.dependentCell !== 'Summary!B3' ||
    parsed.before !== 60000 ||
    parsed.after !== 96000 ||
    parsed.afterRestart !== 96000 ||
    parsed.displayValue !== '96000' ||
    !isStringArray(toolNames) ||
    !sameJson(toolNames, [
      'list_sheets',
      'read_range',
      'read_cell',
      'set_cell_contents',
      'get_cell_display_value',
      'export_workpaper_document',
      'validate_formula',
    ]) ||
    !isStringArray(resourceUris) ||
    !sameJson(resourceUris, [
      'bilig://workpaper/manifest',
      'bilig://workpaper/agent-handoff',
      'bilig://workpaper/sheets',
      'bilig://workpaper/current-document',
    ]) ||
    !isStringArray(promptNames) ||
    !sameJson(promptNames, ['edit_and_verify_workpaper', 'debug_workpaper_formula']) ||
    persistence.persisted !== true ||
    typeof persistence.serializedBytes !== 'number' ||
    persistence.serializedBytes <= 0 ||
    checks.listedFileBackedTools !== true ||
    checks.listedResourcesAndPrompts !== true ||
    checks.formulaValidationPassed !== true ||
    checks.dependentCellChanged !== true ||
    checks.persistedToDisk !== true ||
    checks.exportContainsWorkPaperDocument !== true ||
    checks.restartReadbackMatchesAfter !== true ||
    checks.displayValueRead !== true
  ) {
    throw new Error(`Unexpected node MCP challenge output: ${output}`)
  }

  return {
    editedCell: parsed.editedCell,
    dependentCell: parsed.dependentCell,
    before: parsed.before,
    after: parsed.after,
    afterRestart: parsed.afterRestart,
    displayValue: parsed.displayValue,
    toolNames,
    resourceUris,
    promptNames,
    persistence: {
      persisted: persistence.persisted,
      serializedBytes: persistence.serializedBytes,
    },
    verified: parsed.verified,
  }
}

export function parseNodeSheetInspectionOutput(output: string): {
  lookup: {
    dimensions: {
      height: number
      width: number
    }
    query: string
    sheetId: number
    sheetName: string
  }
  restoredSheets: string[]
  verified: boolean
} {
  const parsed = parseJsonRecord(output, 'node sheet inspection output')
  const lookup = parseRecordValue(parsed.lookup, 'node sheet inspection lookup')
  const dimensions = parseRecordValue(lookup.dimensions, 'node sheet inspection dimensions')

  if (
    parsed.verified !== true ||
    !sameJson(parsed.restoredSheets, ['Inputs', 'Summary']) ||
    lookup.query !== 'Summary' ||
    lookup.sheetId !== 2 ||
    lookup.sheetName !== 'Summary' ||
    dimensions.width !== 2 ||
    dimensions.height !== 3
  ) {
    throw new Error(`Unexpected node sheet inspection output: ${output}`)
  }

  return {
    lookup: {
      dimensions: {
        height: dimensions.height,
        width: dimensions.width,
      },
      query: lookup.query,
      sheetId: lookup.sheetId,
      sheetName: lookup.sheetName,
    },
    restoredSheets: ['Inputs', 'Summary'],
    verified: parsed.verified,
  }
}

function parseFormulaDiagnostics(value: unknown): {
  code: string
  errorText: string
  functionName: string
  references: string[]
}[] {
  if (!Array.isArray(value)) {
    throw new Error(`Expected formula diagnostics to be an array: ${JSON.stringify(value)}`)
  }

  return value.map((entry, index) => {
    const diagnostic = parseRecordValue(entry, `formula diagnostic ${String(index + 1)}`)
    const references = diagnostic.references
    if (
      typeof diagnostic.code !== 'string' ||
      typeof diagnostic.errorText !== 'string' ||
      typeof diagnostic.functionName !== 'string' ||
      !isStringArray(references)
    ) {
      throw new Error(`Unexpected formula diagnostic output: ${JSON.stringify(entry)}`)
    }

    return {
      code: diagnostic.code,
      errorText: diagnostic.errorText,
      functionName: diagnostic.functionName,
      references,
    }
  })
}

function parseSnapshotDiffSummary(
  value: unknown,
  context: string,
): {
  annualizedArr: number
  netMrr: number
} {
  const parsed = parseRecordValue(value, context)
  if (typeof parsed.netMrr !== 'number' || typeof parsed.annualizedArr !== 'number') {
    throw new Error(`Unexpected ${context}: ${JSON.stringify(value)}`)
  }
  return {
    annualizedArr: parsed.annualizedArr,
    netMrr: parsed.netMrr,
  }
}

export function parseNodeSnapshotImportOutput(output: string): {
  currencyLabel: string
  firstPeriod: number
  secondPeriod: number
  totalValue: number
  updatedFirstPeriod: number
} {
  const parsed = parseJsonRecord(output, 'node snapshot import output')
  const currencyLabel = parsed.currencyLabel
  const firstPeriod = parsed.firstPeriod
  const secondPeriod = parsed.secondPeriod
  const totalValue = parsed.totalValue
  const updatedFirstPeriod = parsed.updatedFirstPeriod
  if (currencyLabel !== 'USD  000s' || firstPeriod !== 0 || secondPeriod !== 1 || totalValue !== 225 || updatedFirstPeriod !== 2) {
    throw new Error(`Unexpected node snapshot import output: ${output}`)
  }
  return {
    currencyLabel,
    firstPeriod,
    secondPeriod,
    totalValue,
    updatedFirstPeriod,
  }
}

export function parseNodeXlsxImportOutput(output: string): {
  currencyLabel: string
  editedTotalValue: number
  exportedBytes: number
  headerPeriod: number
  heightFeet: number
  firstPeriod: number
  roundTripSheetNames: string[]
  roundTripTotalValue: number
  secondPeriod: number
  totalValue: number
  translatedStructuredRefs: boolean
} {
  const parsed = parseJsonRecord(output, 'node XLSX import output')
  const currencyLabel = parsed.currencyLabel
  const editedTotalValue = parsed.editedTotalValue
  const exportedBytes = parsed.exportedBytes
  const headerPeriod = parsed.headerPeriod
  const heightFeet = parsed.heightFeet
  const firstPeriod = parsed.firstPeriod
  const roundTripSheetNames = parsed.roundTripSheetNames
  const roundTripTotalValue = parsed.roundTripTotalValue
  const secondPeriod = parsed.secondPeriod
  const totalValue = parsed.totalValue
  const translatedStructuredRefs = parsed.translatedStructuredRefs
  if (
    currencyLabel !== 'USD  000s' ||
    editedTotalValue !== 300 ||
    typeof exportedBytes !== 'number' ||
    exportedBytes <= 0 ||
    headerPeriod !== 1 ||
    typeof heightFeet !== 'number' ||
    Math.abs(heightFeet - (5 + 7 / 12)) > 1e-12 ||
    firstPeriod !== 0 ||
    !isStringArray(roundTripSheetNames) ||
    !sameJson(roundTripSheetNames, ['Constants', 'Imports', 'PlayerData']) ||
    roundTripTotalValue !== 300 ||
    secondPeriod !== 1 ||
    totalValue !== 225 ||
    translatedStructuredRefs !== true
  ) {
    throw new Error(`Unexpected node XLSX import output: ${output}`)
  }
  return {
    currencyLabel,
    editedTotalValue,
    exportedBytes,
    headerPeriod,
    heightFeet,
    firstPeriod,
    roundTripSheetNames,
    roundTripTotalValue,
    secondPeriod,
    totalValue,
    translatedStructuredRefs,
  }
}

export function parseNodePersistenceOutput(output: string): {
  afterRestoreAndEdit: {
    annualizedRunRate: number
    expansionAdjustedArr: number
    quarterNetMrr: number
  }
  beforeSave: {
    annualizedRunRate: number
    expansionAdjustedArr: number
    quarterNetMrr: number
  }
  persistedNamedExpressions: string[]
  persistedSheets: string[]
  saveFileBytes: number
} {
  const parsed = parseJsonRecord(output, 'node persistence output')
  const beforeSave = parseRecordValue(parsed.beforeSave, 'node persistence before-save output')
  const afterRestoreAndEdit = parseRecordValue(parsed.afterRestoreAndEdit, 'node persistence restored output')
  const persistedSheets = parsed.persistedSheets
  const persistedNamedExpressions = parsed.persistedNamedExpressions

  if (
    typeof beforeSave.quarterNetMrr !== 'number' ||
    typeof beforeSave.annualizedRunRate !== 'number' ||
    typeof beforeSave.expansionAdjustedArr !== 'number' ||
    typeof afterRestoreAndEdit.quarterNetMrr !== 'number' ||
    typeof afterRestoreAndEdit.annualizedRunRate !== 'number' ||
    typeof afterRestoreAndEdit.expansionAdjustedArr !== 'number' ||
    !isStringArray(persistedSheets) ||
    !isStringArray(persistedNamedExpressions) ||
    typeof parsed.saveFileBytes !== 'number' ||
    parsed.saveFileBytes <= 0
  ) {
    throw new Error(`Unexpected node persistence output: ${output}`)
  }

  return {
    beforeSave: {
      quarterNetMrr: beforeSave.quarterNetMrr,
      annualizedRunRate: beforeSave.annualizedRunRate,
      expansionAdjustedArr: beforeSave.expansionAdjustedArr,
    },
    afterRestoreAndEdit: {
      quarterNetMrr: afterRestoreAndEdit.quarterNetMrr,
      annualizedRunRate: afterRestoreAndEdit.annualizedRunRate,
      expansionAdjustedArr: afterRestoreAndEdit.expansionAdjustedArr,
    },
    persistedSheets,
    persistedNamedExpressions,
    saveFileBytes: parsed.saveFileBytes,
  }
}
