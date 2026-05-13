import { describe, expect, it } from 'vitest'

import { createMcpDemoOutput } from '../../examples/headless-workpaper/mcp-tool-server.ts'

describe('headless WorkPaper MCP tool server contract', () => {
  it('lists WorkPaper tools with input schemas and returns structured happy-path output', () => {
    const output = createMcpDemoOutput()

    expect(output.listResponse.result.tools).toEqual([
      expect.objectContaining({
        name: 'read_workpaper_summary',
        inputSchema: expect.objectContaining({
          type: 'object',
          properties: expect.objectContaining({
            range: expect.objectContaining({ type: 'string' }),
          }),
        }),
      }),
      expect.objectContaining({
        name: 'set_workpaper_input_cell',
        inputSchema: expect.objectContaining({
          type: 'object',
          required: ['sheetName', 'address', 'value'],
          properties: expect.objectContaining({
            sheetName: expect.objectContaining({ const: 'Inputs' }),
            address: expect.objectContaining({ type: 'string' }),
            value: expect.objectContaining({ type: ['string', 'number', 'boolean', 'null'] }),
          }),
        }),
      }),
    ])

    expect(output.readResponse.result).toEqual(
      expect.objectContaining({
        isError: false,
        content: [
          expect.objectContaining({
            type: 'text',
            text: expect.any(String),
          }),
        ],
        structuredContent: expect.objectContaining({
          range: 'Summary!A1:B5',
          values: expect.any(Array),
          serialized: expect.any(Array),
        }),
      }),
    )

    expect(output.writeResponse.result).toEqual(
      expect.objectContaining({
        isError: false,
        content: [
          expect.objectContaining({
            type: 'text',
            text: expect.any(String),
          }),
        ],
        structuredContent: expect.objectContaining({
          editedCell: 'Inputs!B3',
          before: expect.objectContaining({ expectedArr: 60000 }),
          after: expect.objectContaining({ expectedArr: 96000 }),
          restored: expect.objectContaining({ expectedArr: 96000 }),
          formulaContracts: expect.objectContaining({
            expectedCustomers: '=Inputs!B2*Inputs!B3',
            expectedArr: '=B2*Inputs!B4',
            expansionArr: '=B3*Inputs!B5',
            targetGap: '=B4-100000',
          }),
          checks: expect.objectContaining({
            formulasPersisted: true,
            restoredMatchesAfter: true,
            expectedArrChanged: true,
          }),
        }),
      }),
    )
  })
})
