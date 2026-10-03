import { basename, resolve } from 'node:path'

import type { WorkPaperMcpToolServer } from './work-paper-mcp-server.js'
import { buildWorkbookCompatibilityReportFromFile, type WorkbookCompatibilityReport } from '@bilig/xlsx/workbook-compatibility-report'

type JsonObject = Record<string, unknown>
type JsonRpcId = string | number | null
type WorkPaperMcpJsonRpcResponse = ReturnType<WorkPaperMcpToolServer['handleJsonRpc']>
type XlsxRiskInspectLimit = number | 'all'

const defaultXlsxWorkbookRiskInspectLimit = 2000

const xlsxWorkbookRiskToolDefinition = {
  name: 'analyze_workbook_risk',
  title: 'Analyze XLSX Workbook Risk',
  description:
    'Analyze the XLSX file that started this MCP server and return workbook risk indicators before the workflow trusts the imported WorkPaper. This diagnostic does not certify Excel compatibility.',
  inputSchema: {
    type: 'object',
    properties: {
      inspectLimit: {
        type: 'string',
        default: defaultXlsxWorkbookRiskInspectLimit.toString(),
        description: 'Formula cells to inspect with the native file-backed scanner. Use all only for explicit deep diagnostics.',
      },
    },
    additionalProperties: false,
  },
  outputSchema: {
    type: 'object',
    required: ['schemaVersion', 'verified', 'input', 'workbook', 'findings', 'risk', 'excelParity', 'limitations'],
    properties: {
      schemaVersion: {
        type: 'string',
      },
      verified: {
        type: 'boolean',
      },
      input: {
        type: 'object',
      },
      workbook: {
        type: 'object',
      },
      findings: {
        type: 'object',
      },
      risk: {
        type: 'object',
      },
      excelParity: {
        const: 'not_proven',
      },
      limitations: {
        type: 'array',
      },
    },
    additionalProperties: true,
  },
  annotations: {
    title: 'Analyze XLSX Workbook Risk',
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: false,
  },
} as const

function withXlsxWorkbookRiskTool(
  server: WorkPaperMcpToolServer,
  input: {
    readonly xlsxPath: string
  },
): WorkPaperMcpToolServer {
  const xlsxPath = resolve(input.xlsxPath)

  return {
    capabilities: server.capabilities,

    handleJsonRpc(request: unknown): WorkPaperMcpJsonRpcResponse {
      const parsedRequest = parseJsonRpcRequest(request)
      if (parsedRequest?.method === 'tools/list') {
        return appendRiskTool(server.handleJsonRpc(request))
      }

      if (parsedRequest?.method === 'tools/call') {
        const params = isRecord(parsedRequest.params) ? parsedRequest.params : {}
        if (params['name'] === xlsxWorkbookRiskToolDefinition.name) {
          const args = isRecord(params['arguments']) ? params['arguments'] : {}
          const report = buildWorkbookCompatibilityReportFromFile(xlsxPath, {
            fileName: basename(xlsxPath),
            inspectLimit: parseInspectLimit(args['inspectLimit'] ?? defaultXlsxWorkbookRiskInspectLimit),
          })
          return {
            jsonrpc: '2.0',
            id: parsedRequest.id,
            result: {
              content: [
                {
                  type: 'text',
                  text: renderWorkbookRiskSummary(report),
                },
              ],
              structuredContent: report,
              isError: false,
            },
          }
        }
      }

      return server.handleJsonRpc(request)
    },
  }
}

function appendRiskTool(response: WorkPaperMcpJsonRpcResponse): WorkPaperMcpJsonRpcResponse {
  const result = response.result
  if (!isRecord(result) || !Array.isArray(result['tools'])) {
    return response
  }
  return {
    ...response,
    result: {
      ...result,
      tools: [...result['tools'], xlsxWorkbookRiskToolDefinition],
    },
  }
}

function renderWorkbookRiskSummary(report: WorkbookCompatibilityReport): string {
  return [
    `Workbook risk level: ${report.risk.level.toUpperCase()}`,
    `Formula cells: ${report.workbook.formulaCellCount.toString()}`,
    `Unsupported functions: ${report.findings.unsupportedFunctions.length.toString()}`,
    `External links: ${report.findings.externalLinks.count.toString()}`,
    'This is a preflight diagnostic, not an Excel compatibility certification.',
  ].join('\n')
}

function parseJsonRpcRequest(value: unknown):
  | {
      readonly id: JsonRpcId | undefined
      readonly method: string
      readonly params?: unknown
    }
  | undefined {
  if (!isRecord(value) || value['jsonrpc'] !== '2.0' || typeof value['method'] !== 'string') {
    return undefined
  }
  const id = value['id']
  return {
    id: id === undefined || id === null || typeof id === 'string' || typeof id === 'number' ? id : undefined,
    method: value['method'],
    params: value['params'],
  }
}

function parseInspectLimit(value: unknown): XlsxRiskInspectLimit {
  if (value === 'all') {
    return value
  }
  if (typeof value === 'number' && Number.isInteger(value) && value > 0) {
    return value
  }
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (trimmed === 'all') {
      return trimmed
    }
    const numberValue = Number(trimmed)
    if (Number.isInteger(numberValue) && numberValue > 0) {
      return numberValue
    }
  }
  throw new Error(`Expected analyze_workbook_risk inspectLimit to be "all" or a positive integer, received: ${String(value)}`)
}

function isRecord(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null
}

export { withXlsxWorkbookRiskTool }
