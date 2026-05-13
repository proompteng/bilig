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
  headerPeriod: number
  heightFeet: number
  firstPeriod: number
  secondPeriod: number
  totalValue: number
  translatedStructuredRefs: boolean
} {
  const parsed = parseJsonRecord(output, 'node XLSX import output')
  const currencyLabel = parsed.currencyLabel
  const headerPeriod = parsed.headerPeriod
  const heightFeet = parsed.heightFeet
  const firstPeriod = parsed.firstPeriod
  const secondPeriod = parsed.secondPeriod
  const totalValue = parsed.totalValue
  const translatedStructuredRefs = parsed.translatedStructuredRefs
  if (
    currencyLabel !== 'USD  000s' ||
    headerPeriod !== 1 ||
    typeof heightFeet !== 'number' ||
    Math.abs(heightFeet - (5 + 7 / 12)) > 1e-12 ||
    firstPeriod !== 0 ||
    secondPeriod !== 1 ||
    totalValue !== 225 ||
    translatedStructuredRefs !== true
  ) {
    throw new Error(`Unexpected node XLSX import output: ${output}`)
  }
  return {
    currencyLabel,
    headerPeriod,
    heightFeet,
    firstPeriod,
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

export type RevenueScenarioSummary = {
  annualRunRate: number
  enterpriseNetMrr: number
  expansionTarget: number
  scenarios: {
    conservativeNetMrr: number
    expansionNetMrr: number
    stretchNetMrr: number
  }
  totalNetMrr: number
}

type AgentVerificationProjection = {
  annualizedArr: number
  arrTargetDelta: number
  customers: number
  expansionMrr: number
  grossMrr: number
}

export type AgentVerificationSummary = {
  after: AgentVerificationProjection
  before: AgentVerificationProjection
  edits: {
    after: number
    before: number
    cell: string
  }[]
  formulaContracts: {
    annualizedArr: string
    arrTargetDelta: string
    customers: string
    expansionMrr: string
    grossMrr: string
  }
  restored: AgentVerificationProjection
  verified: {
    formulasPersisted: boolean
    formulasUnchanged: boolean
    restoredMatchesAfter: boolean
    serializedBytes: number
  }
}

export type AgentToolCallSummary = {
  toolCall: {
    arguments: {
      address: string
      reason: string
      sheetName: string
      value: number
    }
    toolName: string
  }
  toolResult: {
    after: AgentToolCallProjection
    before: AgentToolCallProjection
    editedCell: string
    restored: AgentToolCallProjection
    verified: {
      expectedArrImproved: boolean
      formulasPersisted: boolean
      newValue: number
      previousValue: number
      restoredMatchesAfter: boolean
      serializedBytes: number
      targetGapClosed: boolean
    }
  }
}

type AgentToolCallProjection = {
  expectedArr: number
  expectedCustomers: number
  expansionArr: number
  targetGap: number
}

export function parseNodeRevenueScenarioOutput(output: string): {
  afterEdit: RevenueScenarioSummary
  beforeEdit: RevenueScenarioSummary
  persistedSheets: string[]
  serializedBytes: number
} {
  const parsed = parseJsonRecord(output, 'node revenue scenario output')
  const beforeEdit = parseRevenueScenarioSummary(parsed.beforeEdit, 'before-edit revenue scenario output')
  const afterEdit = parseRevenueScenarioSummary(parsed.afterEdit, 'after-edit revenue scenario output')
  const persistedSheets = parsed.persistedSheets

  if (!isStringArray(persistedSheets) || typeof parsed.serializedBytes !== 'number' || parsed.serializedBytes <= 0) {
    throw new Error(`Unexpected node revenue scenario output: ${output}`)
  }

  return {
    beforeEdit,
    afterEdit,
    persistedSheets,
    serializedBytes: parsed.serializedBytes,
  }
}

function parseRevenueScenarioSummary(value: unknown, context: string): RevenueScenarioSummary {
  const parsed = parseRecordValue(value, context)
  const scenarios = parseRecordValue(parsed.scenarios, `${context} scenarios`)

  if (
    typeof parsed.totalNetMrr !== 'number' ||
    typeof parsed.annualRunRate !== 'number' ||
    typeof parsed.enterpriseNetMrr !== 'number' ||
    typeof parsed.expansionTarget !== 'number' ||
    typeof scenarios.conservativeNetMrr !== 'number' ||
    typeof scenarios.expansionNetMrr !== 'number' ||
    typeof scenarios.stretchNetMrr !== 'number'
  ) {
    throw new Error(`Unexpected ${context}: ${JSON.stringify(value)}`)
  }

  return {
    totalNetMrr: parsed.totalNetMrr,
    annualRunRate: parsed.annualRunRate,
    enterpriseNetMrr: parsed.enterpriseNetMrr,
    expansionTarget: parsed.expansionTarget,
    scenarios: {
      conservativeNetMrr: scenarios.conservativeNetMrr,
      expansionNetMrr: scenarios.expansionNetMrr,
      stretchNetMrr: scenarios.stretchNetMrr,
    },
  }
}

export function parseNodeAgentToolCallOutput(output: string): AgentToolCallSummary {
  const parsed = parseJsonRecord(output, 'node agent tool-call output')
  const toolCall = parseRecordValue(parsed.toolCall, 'node agent tool-call request')
  const toolCallArguments = parseRecordValue(toolCall.arguments, 'node agent tool-call arguments')
  const toolResult = parseRecordValue(parsed.toolResult, 'node agent tool-call result')
  const before = parseAgentToolCallProjection(toolResult.before, 'node agent tool-call before output')
  const after = parseAgentToolCallProjection(toolResult.after, 'node agent tool-call after output')
  const restored = parseAgentToolCallProjection(toolResult.restored, 'node agent tool-call restored output')
  const verified = parseRecordValue(toolResult.verified, 'node agent tool-call flags')

  if (
    toolCall.toolName !== 'setInputCell' ||
    toolCallArguments.sheetName !== 'Inputs' ||
    toolCallArguments.address !== 'B3' ||
    toolCallArguments.value !== 0.4 ||
    typeof toolCallArguments.reason !== 'string' ||
    toolResult.editedCell !== 'Inputs!B3' ||
    before.expectedCustomers !== 5 ||
    before.expectedArr !== 60000 ||
    before.expansionArr !== 66000 ||
    before.targetGap !== -34000 ||
    after.expectedCustomers !== 8 ||
    after.expectedArr !== 96000 ||
    after.expansionArr !== 105600 ||
    after.targetGap !== 5600 ||
    restored.expectedCustomers !== after.expectedCustomers ||
    restored.expectedArr !== after.expectedArr ||
    restored.expansionArr !== after.expansionArr ||
    restored.targetGap !== after.targetGap ||
    verified.previousValue !== 0.25 ||
    verified.newValue !== 0.4 ||
    verified.formulasPersisted !== true ||
    verified.restoredMatchesAfter !== true ||
    verified.expectedArrImproved !== true ||
    verified.targetGapClosed !== true ||
    typeof verified.serializedBytes !== 'number' ||
    verified.serializedBytes <= 0
  ) {
    throw new Error(`Unexpected node agent tool-call output: ${output}`)
  }

  return {
    toolCall: {
      arguments: {
        address: toolCallArguments.address,
        reason: toolCallArguments.reason,
        sheetName: toolCallArguments.sheetName,
        value: toolCallArguments.value,
      },
      toolName: toolCall.toolName,
    },
    toolResult: {
      after,
      before,
      editedCell: toolResult.editedCell,
      restored,
      verified: {
        expectedArrImproved: verified.expectedArrImproved,
        formulasPersisted: verified.formulasPersisted,
        newValue: verified.newValue,
        previousValue: verified.previousValue,
        restoredMatchesAfter: verified.restoredMatchesAfter,
        serializedBytes: verified.serializedBytes,
        targetGapClosed: verified.targetGapClosed,
      },
    },
  }
}

export function parseNodeMcpStdioOutput(
  output: string,
  options: { expectedServerName?: string } = {},
): {
  editedCell: string
  initialized: boolean
  toolNames: string[]
  verified: {
    expectedArrChanged: boolean
    formulasPersisted: boolean
    newValue: number
    previousValue: number
    restoredMatchesAfter: boolean
    serializedBytes: number
  }
} {
  const expectedServerName = options.expectedServerName ?? 'bilig-headless-workpaper-example'
  const responses = output
    .split(/\r?\n/)
    .filter((line) => line.trim().length > 0)
    .map((line, index) => parseJsonRecord(line, `node MCP stdio response ${index + 1}`))

  const initializeResponse = requireJsonRpcResponse(responses, 1, 'node MCP stdio initialize response')
  const listResponse = requireJsonRpcResponse(responses, 2, 'node MCP stdio tools/list response')
  const writeResponse = requireJsonRpcResponse(responses, 3, 'node MCP stdio tools/call response')
  const initializeResult = parseRecordValue(initializeResponse.result, 'node MCP stdio initialize result')
  const listResult = parseRecordValue(listResponse.result, 'node MCP stdio tools/list result')
  const writeResult = parseRecordValue(writeResponse.result, 'node MCP stdio tools/call result')
  const structuredContent = parseRecordValue(writeResult.structuredContent, 'node MCP stdio structured content')
  const before = parseAgentToolCallProjection(structuredContent.before, 'node MCP stdio before output')
  const after = parseAgentToolCallProjection(structuredContent.after, 'node MCP stdio after output')
  const restored = parseAgentToolCallProjection(structuredContent.restored, 'node MCP stdio restored output')
  const checks = parseRecordValue(structuredContent.checks, 'node MCP stdio checks')
  const tools = parseToolList(listResult.tools)

  if (
    initializeResult.protocolVersion !== '2025-06-18' ||
    parseRecordValue(initializeResult.serverInfo, 'node MCP stdio server info').name !== expectedServerName ||
    !sameJson(tools, ['read_workpaper_summary', 'set_workpaper_input_cell']) ||
    writeResult.isError !== false ||
    structuredContent.editedCell !== 'Inputs!B3' ||
    before.expectedCustomers !== 5 ||
    before.expectedArr !== 60000 ||
    before.expansionArr !== 66000 ||
    before.targetGap !== -34000 ||
    after.expectedCustomers !== 8 ||
    after.expectedArr !== 96000 ||
    after.expansionArr !== 105600 ||
    after.targetGap !== 5600 ||
    !sameJson(restored, after) ||
    checks.previousValue !== 0.25 ||
    checks.newValue !== 0.4 ||
    checks.formulasPersisted !== true ||
    checks.restoredMatchesAfter !== true ||
    checks.expectedArrChanged !== true ||
    typeof checks.serializedBytes !== 'number' ||
    checks.serializedBytes <= 0
  ) {
    throw new Error(`Unexpected node MCP stdio output: ${output}`)
  }

  return {
    editedCell: structuredContent.editedCell,
    initialized: true,
    toolNames: tools,
    verified: {
      expectedArrChanged: checks.expectedArrChanged,
      formulasPersisted: checks.formulasPersisted,
      newValue: checks.newValue,
      previousValue: checks.previousValue,
      restoredMatchesAfter: checks.restoredMatchesAfter,
      serializedBytes: checks.serializedBytes,
    },
  }
}

function requireJsonRpcResponse(responses: Record<string, unknown>[], id: number, context: string): Record<string, unknown> {
  const response = responses.find((entry) => entry.id === id)
  if (response === undefined || response.jsonrpc !== '2.0') {
    throw new Error(`Missing ${context}: ${JSON.stringify(responses)}`)
  }
  return response
}

function parseToolList(value: unknown): string[] {
  if (!Array.isArray(value)) {
    throw new Error(`Expected MCP tools to be an array: ${JSON.stringify(value)}`)
  }
  return value.map((entry, index) => {
    const tool = parseRecordValue(entry, `MCP tool ${index + 1}`)
    if (typeof tool.name !== 'string') {
      throw new Error(`Unexpected MCP tool: ${JSON.stringify(entry)}`)
    }
    return tool.name
  })
}

function parseAgentToolCallProjection(value: unknown, context: string): AgentToolCallProjection {
  const parsed = parseRecordValue(value, context)
  if (
    typeof parsed.expectedCustomers !== 'number' ||
    typeof parsed.expectedArr !== 'number' ||
    typeof parsed.expansionArr !== 'number' ||
    typeof parsed.targetGap !== 'number'
  ) {
    throw new Error(`Unexpected ${context}: ${JSON.stringify(value)}`)
  }

  return {
    expectedArr: parsed.expectedArr,
    expectedCustomers: parsed.expectedCustomers,
    expansionArr: parsed.expansionArr,
    targetGap: parsed.targetGap,
  }
}

export function parseNodeAgentVerificationOutput(output: string): AgentVerificationSummary {
  const parsed = parseJsonRecord(output, 'node agent verification output')
  const before = parseAgentVerificationProjection(parsed.before, 'node agent verification before output')
  const after = parseAgentVerificationProjection(parsed.after, 'node agent verification after output')
  const restored = parseAgentVerificationProjection(parsed.restored, 'node agent verification restored output')
  const edits = parseAgentVerificationEdits(parsed.edits)
  const formulaContracts = parseAgentVerificationFormulaContracts(parsed.formulaContracts)
  const verified = parseRecordValue(parsed.verified, 'node agent verification flags')

  if (
    before.customers !== 40 ||
    before.grossMrr !== 9600 ||
    before.expansionMrr !== 10560 ||
    before.annualizedArr !== 126720 ||
    before.arrTargetDelta !== -23280 ||
    after.customers !== 65 ||
    after.grossMrr !== 15600 ||
    after.expansionMrr !== 18720 ||
    after.annualizedArr !== 224640 ||
    after.arrTargetDelta !== 74640 ||
    restored.customers !== after.customers ||
    restored.grossMrr !== after.grossMrr ||
    restored.expansionMrr !== after.expansionMrr ||
    restored.annualizedArr !== after.annualizedArr ||
    restored.arrTargetDelta !== after.arrTargetDelta ||
    formulaContracts.customers !== '=Assumptions!B2*Assumptions!B3' ||
    formulaContracts.grossMrr !== '=B2*Assumptions!B4' ||
    formulaContracts.expansionMrr !== '=B3*Assumptions!B5' ||
    formulaContracts.annualizedArr !== '=B4*12' ||
    formulaContracts.arrTargetDelta !== '=Plan!B5-150000' ||
    verified.formulasUnchanged !== true ||
    verified.formulasPersisted !== true ||
    verified.restoredMatchesAfter !== true ||
    typeof verified.serializedBytes !== 'number' ||
    verified.serializedBytes <= 0
  ) {
    throw new Error(`Unexpected node agent verification output: ${output}`)
  }

  return {
    after,
    before,
    edits,
    formulaContracts,
    restored,
    verified: {
      formulasPersisted: verified.formulasPersisted,
      formulasUnchanged: verified.formulasUnchanged,
      restoredMatchesAfter: verified.restoredMatchesAfter,
      serializedBytes: verified.serializedBytes,
    },
  }
}

function parseAgentVerificationProjection(value: unknown, context: string): AgentVerificationProjection {
  const parsed = parseRecordValue(value, context)
  if (
    typeof parsed.customers !== 'number' ||
    typeof parsed.grossMrr !== 'number' ||
    typeof parsed.expansionMrr !== 'number' ||
    typeof parsed.annualizedArr !== 'number' ||
    typeof parsed.arrTargetDelta !== 'number'
  ) {
    throw new Error(`Unexpected ${context}: ${JSON.stringify(value)}`)
  }
  return {
    annualizedArr: parsed.annualizedArr,
    arrTargetDelta: parsed.arrTargetDelta,
    customers: parsed.customers,
    expansionMrr: parsed.expansionMrr,
    grossMrr: parsed.grossMrr,
  }
}

function parseAgentVerificationFormulaContracts(value: unknown): AgentVerificationSummary['formulaContracts'] {
  const parsed = parseRecordValue(value, 'node agent verification formula contracts')
  if (
    typeof parsed.customers !== 'string' ||
    typeof parsed.grossMrr !== 'string' ||
    typeof parsed.expansionMrr !== 'string' ||
    typeof parsed.annualizedArr !== 'string' ||
    typeof parsed.arrTargetDelta !== 'string'
  ) {
    throw new Error(`Unexpected node agent verification formula contracts: ${JSON.stringify(value)}`)
  }
  return {
    annualizedArr: parsed.annualizedArr,
    arrTargetDelta: parsed.arrTargetDelta,
    customers: parsed.customers,
    expansionMrr: parsed.expansionMrr,
    grossMrr: parsed.grossMrr,
  }
}

function parseAgentVerificationEdits(value: unknown): AgentVerificationSummary['edits'] {
  if (!Array.isArray(value)) {
    throw new Error(`Expected node agent verification edits to be an array: ${JSON.stringify(value)}`)
  }
  const edits = value.map((entry, index) => {
    const edit = parseRecordValue(entry, `node agent verification edit ${index + 1}`)
    if (typeof edit.cell !== 'string' || typeof edit.before !== 'number' || typeof edit.after !== 'number') {
      throw new Error(`Unexpected node agent verification edit: ${JSON.stringify(entry)}`)
    }
    return {
      after: edit.after,
      before: edit.before,
      cell: edit.cell,
    }
  })
  const expected = [
    { after: 650, before: 500, cell: 'Assumptions!B2' },
    { after: 0.1, before: 0.08, cell: 'Assumptions!B3' },
    { after: 1.2, before: 1.1, cell: 'Assumptions!B5' },
  ]
  if (JSON.stringify(edits) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected node agent verification edits: ${JSON.stringify(value)}`)
  }
  return edits
}

export function parseJsonRecord(serialized: string, context: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(serialized)
  return parseRecordValue(parsed, context)
}

function parseRecordValue(candidate: unknown, context: string): Record<string, unknown> {
  const parsed = candidate
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new Error(`Expected ${context} to be a JSON object`)
  }
  const record: Record<string, unknown> = {}
  for (const [key, value] of Object.entries(parsed)) {
    record[key] = value
  }
  return record
}

function isNumberArray(value: unknown): value is number[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === 'number')
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === 'string')
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
