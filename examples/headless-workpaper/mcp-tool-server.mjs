import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const server = createMcpWorkPaperToolServer(buildWorkbook())

const output = {
  capabilities: server.capabilities,
  listResponse: server.handleJsonRpc({
    jsonrpc: '2.0',
    id: 1,
    method: 'tools/list',
  }),
  readResponse: server.handleJsonRpc({
    jsonrpc: '2.0',
    id: 2,
    method: 'tools/call',
    params: {
      name: 'read_workpaper_summary',
      arguments: {
        range: 'Summary!A1:B5',
      },
    },
  }),
  writeResponse: server.handleJsonRpc({
    jsonrpc: '2.0',
    id: 3,
    method: 'tools/call',
    params: {
      name: 'set_workpaper_input_cell',
      arguments: {
        sheetName: 'Inputs',
        address: 'B3',
        value: 0.4,
      },
    },
  }),
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

function buildWorkbook() {
  return WorkPaper.buildFromSheets({
    Inputs: [
      ['Metric', 'Value'],
      ['Qualified opportunities', 20],
      ['Win rate', 0.25],
      ['Average ARR', 12000],
      ['Expansion multiplier', 1.1],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Expected customers', '=Inputs!B2*Inputs!B3'],
      ['Expected ARR', '=B2*Inputs!B4'],
      ['Expansion ARR', '=B3*Inputs!B5'],
      ['Target gap', '=B4-100000'],
    ],
  })
}

function createMcpWorkPaperToolServer(workbook) {
  const workPaperTools = createWorkPaperTools(workbook)
  const toolDefinitions = [
    {
      name: 'read_workpaper_summary',
      description: 'Read computed WorkPaper summary values for a small range.',
      inputSchema: {
        type: 'object',
        properties: {
          range: {
            type: 'string',
            description: 'A1 range with an optional sheet name.',
            default: 'Summary!A1:B5',
          },
        },
        additionalProperties: false,
      },
    },
    {
      name: 'set_workpaper_input_cell',
      description: 'Set one validated WorkPaper input cell and return formula readback.',
      inputSchema: {
        type: 'object',
        required: ['sheetName', 'address', 'value'],
        properties: {
          sheetName: {
            type: 'string',
            const: 'Inputs',
          },
          address: {
            type: 'string',
            description: 'A1 cell address in the Inputs sheet.',
          },
          value: {
            type: ['string', 'number', 'boolean', 'null'],
          },
        },
        additionalProperties: false,
      },
    },
  ]

  return {
    capabilities: {
      tools: {
        listChanged: false,
      },
    },

    handleJsonRpc(request) {
      if (request?.jsonrpc !== '2.0') {
        throw new Error('Expected JSON-RPC 2.0 request')
      }

      if (request.method === 'tools/list') {
        return {
          jsonrpc: '2.0',
          id: request.id,
          result: {
            tools: toolDefinitions,
          },
        }
      }

      if (request.method === 'tools/call') {
        const toolName = request.params?.name
        const args = request.params?.arguments ?? {}
        const structuredContent = callTool(workPaperTools, toolName, args)

        return {
          jsonrpc: '2.0',
          id: request.id,
          result: {
            content: [
              {
                type: 'text',
                text: JSON.stringify(structuredContent),
              },
            ],
            structuredContent,
            isError: false,
          },
        }
      }

      throw new Error(`Unsupported MCP method: ${String(request.method)}`)
    },
  }
}

function callTool(workPaperTools, toolName, args) {
  if (toolName === 'read_workpaper_summary') {
    return workPaperTools.readWorkPaperSummary(args.range)
  }

  if (toolName === 'set_workpaper_input_cell') {
    return workPaperTools.setWorkPaperInputCell(args)
  }

  throw new Error(`Unknown WorkPaper tool: ${String(toolName)}`)
}

function createWorkPaperTools(workbook) {
  const summarySheet = requireSheet(workbook, 'Summary')

  return {
    readWorkPaperSummary(range = 'Summary!A1:B5') {
      const parsedRange = workbook.simpleCellRangeFromString(range, summarySheet)
      if (parsedRange === undefined) {
        throw new Error(`Invalid readable range: ${range}`)
      }

      return {
        range,
        values: workbook.getRangeValues(parsedRange),
        serialized: workbook.getRangeSerialized(parsedRange),
      }
    },

    setWorkPaperInputCell({ sheetName, address, value }) {
      if (sheetName !== 'Inputs') {
        throw new Error(`This example only permits Inputs edits, received ${sheetName}`)
      }

      const target = requireCellAddress(workbook, sheetName, address)
      const before = readSummary(workbook, summarySheet)
      const formulaContracts = readFormulaContracts(workbook, summarySheet)
      const previousValue = workbook.getCellSerialized(target)

      workbook.setCellContents(target, value)

      const after = readSummary(workbook, summarySheet)
      const serialized = serializeWorkbook(workbook)
      const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
      const restoredSummarySheet = requireSheet(restored, 'Summary')
      const restoredSummary = readSummary(restored, restoredSummarySheet)
      const restoredFormulaContracts = readFormulaContracts(restored, restoredSummarySheet)

      return {
        editedCell: workbook.simpleCellAddressToString(target, {
          includeSheetName: true,
        }),
        before,
        after,
        restored: restoredSummary,
        formulaContracts,
        checks: {
          previousValue,
          newValue: workbook.getCellSerialized(target),
          formulasPersisted: sameJson(formulaContracts, restoredFormulaContracts),
          restoredMatchesAfter: sameJson(after, restoredSummary),
          expectedArrChanged: after.expectedArr > before.expectedArr,
          serializedBytes: Buffer.byteLength(serialized, 'utf8'),
        },
      }
    },
  }
}

function requireSheet(workpaper, sheetName) {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function requireCellAddress(workpaper, sheetName, a1Address) {
  const sheetId = requireSheet(workpaper, sheetName)
  const parsed = workpaper.simpleCellAddressFromString(a1Address, sheetId)

  if (parsed === undefined || parsed.sheet !== sheetId) {
    throw new Error(`Invalid cell address: ${sheetName}!${a1Address}`)
  }

  return parsed
}

function readSummary(workpaper, summary) {
  return {
    expectedCustomers: readNumber(workpaper, summary, 1, 1, 'expected customers'),
    expectedArr: readNumber(workpaper, summary, 2, 1, 'expected ARR'),
    expansionArr: readNumber(workpaper, summary, 3, 1, 'expansion ARR'),
    targetGap: readNumber(workpaper, summary, 4, 1, 'target gap'),
  }
}

function readFormulaContracts(workpaper, summary) {
  return {
    expectedCustomers: readFormula(workpaper, summary, 1, 1, 'expected customers'),
    expectedArr: readFormula(workpaper, summary, 2, 1, 'expected ARR'),
    expansionArr: readFormula(workpaper, summary, 3, 1, 'expansion ARR'),
    targetGap: readFormula(workpaper, summary, 4, 1, 'target gap'),
  }
}

function readNumber(workpaper, sheet, row, col, label) {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function readFormula(workpaper, sheet, row, col, label) {
  const formula = workpaper.getCellFormula({ sheet, row, col })
  if (formula === undefined) {
    throw new Error(`Expected ${label} to be a formula`)
  }
  return formula
}

function serializeWorkbook(workpaper) {
  return serializeWorkPaperDocument(
    exportWorkPaperDocument(workpaper, {
      includeConfig: true,
    }),
  )
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right)
}

function assertOutput(actual) {
  const toolNames = actual.listResponse.result.tools.map((tool) => tool.name)
  if (!sameJson(toolNames, ['read_workpaper_summary', 'set_workpaper_input_cell'])) {
    throw new Error(`Unexpected MCP tool list: ${JSON.stringify(toolNames)}`)
  }

  const writeResult = actual.writeResponse.result.structuredContent
  const expectedBefore = {
    expectedCustomers: 5,
    expectedArr: 60000,
    expansionArr: 66000,
    targetGap: -34000,
  }
  const expectedAfter = {
    expectedCustomers: 8,
    expectedArr: 96000,
    expansionArr: 105600,
    targetGap: 5600,
  }
  const expectedFormulaContracts = {
    expectedCustomers: '=Inputs!B2*Inputs!B3',
    expectedArr: '=B2*Inputs!B4',
    expansionArr: '=B3*Inputs!B5',
    targetGap: '=B4-100000',
  }

  if (
    actual.capabilities.tools.listChanged !== false ||
    actual.readResponse.result.content[0].type !== 'text' ||
    actual.writeResponse.result.content[0].type !== 'text' ||
    writeResult.editedCell !== 'Inputs!B3' ||
    !sameJson(writeResult.before, expectedBefore) ||
    !sameJson(writeResult.after, expectedAfter) ||
    !sameJson(writeResult.restored, expectedAfter) ||
    !sameJson(writeResult.formulaContracts, expectedFormulaContracts) ||
    writeResult.checks.previousValue !== 0.25 ||
    writeResult.checks.newValue !== 0.4 ||
    writeResult.checks.formulasPersisted !== true ||
    writeResult.checks.restoredMatchesAfter !== true ||
    writeResult.checks.expectedArrChanged !== true ||
    writeResult.checks.serializedBytes <= 0
  ) {
    throw new Error(`Unexpected MCP adapter result: ${JSON.stringify(actual)}`)
  }
}
