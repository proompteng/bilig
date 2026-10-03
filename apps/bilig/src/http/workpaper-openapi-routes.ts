import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify'
import { WORKPAPER_VERSION } from '@bilig/workpaper'
import { buildDemoWorkPaper, createFileBackedWorkPaperMcpToolServer } from '@bilig/workpaper/mcp'

const OPENAPI_SPEC_ENDPOINTS = ['/openapi/workpaper', '/openapi/workpaper.json', '/openapi/workpaper/openapi.json'] as const
const OPENAPI_OPERATION_ENDPOINTS = [
  '/openapi/workpaper/list-sheets',
  '/openapi/workpaper/read-range',
  '/openapi/workpaper/set-cell-and-readback',
] as const
const OPENAPI_ALLOWED_METHODS = 'POST, GET, OPTIONS'

interface OpenApiToolCallSuccess {
  readonly jsonrpc: '2.0'
  readonly result: {
    readonly structuredContent?: unknown
  }
}

interface ListSheetsBody {
  readonly includeDemoLimitations?: boolean
}

interface ReadRangeBody {
  readonly range?: string
  readonly sheetName?: string
}

interface SetCellAndReadbackBody {
  readonly sheetName?: string
  readonly address?: string
  readonly value?: unknown
  readonly readbackRange?: string
  readonly readbackSheetName?: string
}

export function registerWorkPaperOpenApiRoutes(app: FastifyInstance): void {
  for (const endpoint of OPENAPI_SPEC_ENDPOINTS) {
    app.get(endpoint, async (_request, reply) => handleWorkPaperOpenApiSpec(reply))
    app.options(endpoint, async (_request, reply) => handleWorkPaperOpenApiOptions(reply))
  }

  for (const endpoint of OPENAPI_OPERATION_ENDPOINTS) {
    app.options(endpoint, async (_request, reply) => handleWorkPaperOpenApiOptions(reply))
  }

  app.post('/openapi/workpaper/list-sheets', async (request: FastifyRequest<{ Body: ListSheetsBody }>, reply) =>
    handleWorkPaperOpenApiToolCall(
      reply,
      'list_sheets',
      request.body?.includeDemoLimitations === undefined ? {} : { includeDemoLimitations: request.body.includeDemoLimitations },
    ),
  )
  app.post('/openapi/workpaper/read-range', async (request: FastifyRequest<{ Body: ReadRangeBody }>, reply) =>
    handleWorkPaperOpenApiToolCall(reply, 'read_range', {
      range: request.body?.range ?? 'Summary!A1:B5',
      ...(request.body?.sheetName ? { sheetName: request.body.sheetName } : {}),
    }),
  )
  app.post('/openapi/workpaper/set-cell-and-readback', async (request: FastifyRequest<{ Body: SetCellAndReadbackBody }>, reply) =>
    handleWorkPaperOpenApiToolCall(reply, 'set_cell_contents_and_readback', {
      sheetName: request.body?.sheetName ?? 'Inputs',
      address: request.body?.address ?? 'B3',
      value: request.body?.value ?? 0.4,
      readbackRange: request.body?.readbackRange ?? 'Summary!A1:B3',
      ...(request.body?.readbackSheetName ? { readbackSheetName: request.body.readbackSheetName } : {}),
    }),
  )
}

function handleWorkPaperOpenApiSpec(reply: FastifyReply): Record<string, unknown> {
  applyOpenApiHeaders(reply)
  reply.header('cache-control', 'public, max-age=300')
  return createWorkPaperOpenApiSpec()
}

function handleWorkPaperOpenApiOptions(reply: FastifyReply) {
  applyOpenApiHeaders(reply)
  reply.header('access-control-allow-methods', OPENAPI_ALLOWED_METHODS)
  reply.header('access-control-allow-headers', 'accept, content-type')
  reply.header('access-control-max-age', '600')
  return reply.code(204).send()
}

function handleWorkPaperOpenApiToolCall(reply: FastifyReply, toolName: string, args: Record<string, unknown>): unknown {
  try {
    applyOpenApiHeaders(reply)
    reply.header('cache-control', 'no-store')
    return callDemoTool(toolName, args)
  } catch (error) {
    applyOpenApiHeaders(reply)
    reply.code(400)
    reply.header('cache-control', 'no-store')
    return {
      verified: false,
      error: error instanceof Error ? error.message : 'Invalid WorkPaper OpenAPI request',
    }
  }
}

function applyOpenApiHeaders(reply: FastifyReply): void {
  reply.header('access-control-allow-origin', '*')
  reply.header('content-type', 'application/json; charset=utf-8')
}

function callDemoTool(toolName: string, args: Record<string, unknown>): unknown {
  const server = createFileBackedWorkPaperMcpToolServer({
    workbook: buildDemoWorkPaper(),
    writable: false,
  })
  const response = server.handleJsonRpc({
    jsonrpc: '2.0',
    id: 'openapi',
    method: 'tools/call',
    params: {
      name: toolName,
      arguments: args,
    },
  })

  if (!isToolCallSuccess(response)) {
    throw new Error(`Expected ${toolName} to return structured content`)
  }

  return response.result.structuredContent
}

function isToolCallSuccess(value: unknown): value is OpenApiToolCallSuccess {
  if (!isJsonObject(value) || value['jsonrpc'] !== '2.0') {
    return false
  }
  const result = value['result']
  if (!isJsonObject(result)) {
    return false
  }
  return Object.hasOwn(result, 'structuredContent')
}

function isJsonObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function createWorkPaperOpenApiSpec(): Record<string, unknown> {
  return {
    openapi: '3.1.0',
    info: {
      title: 'Bilig WorkPaper OpenAPI Tool Server',
      version: WORKPAPER_VERSION,
      description:
        'Stateless hosted WorkPaper tools for Open WebUI: discover sheets, read calculated ranges, and edit an input while returning dependent formula readback proof.',
    },
    servers: [
      {
        url: 'https://bilig.proompteng.ai',
        description: 'Hosted stateless WorkPaper demo',
      },
    ],
    paths: {
      '/openapi/workpaper/list-sheets': {
        post: {
          operationId: 'list_workpaper_sheets',
          summary: 'List demo WorkPaper sheets',
          description: 'Returns sheet names and used dimensions for the hosted stateless demo WorkPaper.',
          requestBody: jsonRequestBody({
            type: 'object',
            properties: {
              includeDemoLimitations: {
                type: 'boolean',
                description: 'Optional no-op flag for clients that require a request body.',
              },
            },
            additionalProperties: false,
          }),
          responses: jsonResponses('Sheet metadata for the demo WorkPaper.'),
        },
      },
      '/openapi/workpaper/read-range': {
        post: {
          operationId: 'read_workpaper_range',
          summary: 'Read calculated values from a demo WorkPaper range',
          description: 'Returns calculated values and serialized formulas/inputs for an A1 range.',
          requestBody: jsonRequestBody({
            type: 'object',
            properties: {
              range: {
                type: 'string',
                description: 'A1 range such as Summary!A1:B5.',
                default: 'Summary!A1:B5',
              },
              sheetName: {
                type: 'string',
                description: 'Optional sheet name when range does not include one.',
              },
            },
            additionalProperties: false,
          }),
          responses: jsonResponses('Calculated values plus serialized formulas and inputs.'),
        },
      },
      '/openapi/workpaper/set-cell-and-readback': {
        post: {
          operationId: 'set_workpaper_cell_and_readback',
          summary: 'Edit one input and verify dependent formula readback',
          description:
            'Writes one demo WorkPaper input, recalculates dependent formulas, reads a dependent range before and after, serializes/restores the WorkPaper, and returns verification checks.',
          requestBody: jsonRequestBody({
            type: 'object',
            properties: {
              sheetName: {
                type: 'string',
                default: 'Inputs',
              },
              address: {
                type: 'string',
                default: 'B3',
              },
              value: {
                oneOf: [{ type: 'number' }, { type: 'string' }, { type: 'boolean' }, { type: 'null' }],
                default: 0.4,
              },
              readbackRange: {
                type: 'string',
                default: 'Summary!A1:B3',
              },
              readbackSheetName: {
                type: 'string',
                description: 'Optional sheet name when readbackRange does not include one.',
              },
            },
            additionalProperties: false,
          }),
          responses: jsonResponses('Before/after/restored proof for the edited cell and dependent readback range.'),
        },
      },
    },
  }
}

function jsonRequestBody(schema: Record<string, unknown>): Record<string, unknown> {
  return {
    required: false,
    content: {
      'application/json': {
        schema,
      },
    },
  }
}

function jsonResponses(description: string): Record<string, unknown> {
  return {
    '200': {
      description,
      content: {
        'application/json': {
          schema: {
            type: 'object',
            additionalProperties: true,
          },
        },
      },
    },
    '400': {
      description: 'Invalid WorkPaper OpenAPI request.',
      content: {
        'application/json': {
          schema: {
            type: 'object',
            required: ['verified', 'error'],
            properties: {
              verified: {
                type: 'boolean',
              },
              error: {
                type: 'string',
              },
            },
            additionalProperties: false,
          },
        },
      },
    },
  }
}
