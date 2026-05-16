import { isStringArray, parseJsonRecord, parseRecordValue, sameJson } from './workpaper-external-smoke-parser-helpers.ts'

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
