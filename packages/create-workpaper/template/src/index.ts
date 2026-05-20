import { createServer, type IncomingMessage } from 'node:http'
import { pathToFileURL } from 'node:url'

import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

type QuoteInput = {
  units: number
  listPrice: number
  discount: number
  unitCost: number
  minimumMargin: number
}

type QuoteSummary = {
  listRevenue: number
  discountAmount: number
  netRevenue: number
  totalCost: number
  grossMargin: number
  decision: string
}

type WorkbookStorage = {
  loadWorkbookJson(): Promise<string> | string
  saveWorkbookJson(nextWorkbookJson: string): Promise<void> | void
}

const inputCells = {
  units: 'Inputs!B2',
  listPrice: 'Inputs!B3',
  discount: 'Inputs!B4',
  unitCost: 'Inputs!B5',
  minimumMargin: 'Inputs!B6',
} as const

export function createQuoteApprovalRequestHandler(storage: WorkbookStorage) {
  return async function handleQuoteApprovalRequest(request: Request): Promise<Response> {
    const url = new URL(request.url)

    if (request.method === 'GET' && url.pathname === '/api/quote/approval') {
      const workbook = await loadWorkbook(storage)
      return json({ summary: readQuoteSummary(workbook), inputCells })
    }

    if (request.method === 'POST' && url.pathname === '/api/quote/approval') {
      try {
        const input = parseQuoteInput(await request.json())
        const workbook = await loadWorkbook(storage)
        const before = readQuoteSummary(workbook)
        writeQuoteInputs(workbook, input)
        const after = readQuoteSummary(workbook)
        const workbookJson = serializeWorkbook(workbook)
        await storage.saveWorkbookJson(workbookJson)

        const restored = createWorkPaperFromDocument(parseWorkPaperDocument(workbookJson))
        const restoredSummary = readQuoteSummary(restored)

        return json({
          input,
          inputCells,
          before,
          after,
          restored: restoredSummary,
          checks: {
            decisionChanged: before.decision !== after.decision,
            formulasPersisted: workbookJson.includes('=IF(B6>=Inputs!B6'),
            restoredMatchesAfter: JSON.stringify(restoredSummary) === JSON.stringify(after),
            serializedBytes: Buffer.byteLength(workbookJson, 'utf8'),
          },
        })
      } catch (error) {
        return json({ error: error instanceof Error ? error.message : String(error) }, 400)
      }
    }

    return json({ error: 'not found' }, 404)
  }
}

function createQuoteApprovalWorkbook(): WorkPaperInstance {
  return WorkPaper.buildFromSheets({
    Inputs: [
      ['Metric', 'Value'],
      ['Units', 40],
      ['List price', 1200],
      ['Discount', 0.1],
      ['Unit cost', 760],
      ['Minimum margin', 0.3],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['List revenue', '=Inputs!B2*Inputs!B3'],
      ['Discount amount', '=B2*Inputs!B4'],
      ['Net revenue', '=B2-B3'],
      ['Total cost', '=Inputs!B2*Inputs!B5'],
      ['Gross margin', '=(B4-B5)/B4'],
      ['Decision', '=IF(B6>=Inputs!B6,"approved","review")'],
    ],
  })
}

function createMemoryStorage(): WorkbookStorage {
  let workbookJson = serializeWorkbook(createQuoteApprovalWorkbook())
  return {
    loadWorkbookJson() {
      return workbookJson
    },
    saveWorkbookJson(nextWorkbookJson) {
      workbookJson = nextWorkbookJson
    },
  }
}

function writeQuoteInputs(workbook: WorkPaperInstance, input: QuoteInput): void {
  const sheet = requireSheet(workbook, 'Inputs')
  workbook.setCellContents({ sheet, row: 1, col: 1 }, input.units)
  workbook.setCellContents({ sheet, row: 2, col: 1 }, input.listPrice)
  workbook.setCellContents({ sheet, row: 3, col: 1 }, input.discount)
  workbook.setCellContents({ sheet, row: 4, col: 1 }, input.unitCost)
  workbook.setCellContents({ sheet, row: 5, col: 1 }, input.minimumMargin)
}

async function loadWorkbook(storage: WorkbookStorage): Promise<WorkPaperInstance> {
  return createWorkPaperFromDocument(parseWorkPaperDocument(await storage.loadWorkbookJson()))
}

function serializeWorkbook(workbook: WorkPaperInstance): string {
  return serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
}

function readQuoteSummary(workbook: WorkPaperInstance): QuoteSummary {
  const sheet = requireSheet(workbook, 'Summary')
  return {
    listRevenue: readNumber(workbook, sheet, 1, 1, 'List revenue'),
    discountAmount: readNumber(workbook, sheet, 2, 1, 'Discount amount'),
    netRevenue: readNumber(workbook, sheet, 3, 1, 'Net revenue'),
    totalCost: readNumber(workbook, sheet, 4, 1, 'Total cost'),
    grossMargin: readRoundedNumber(workbook, sheet, 5, 1, 'Gross margin'),
    decision: readString(workbook, sheet, 6, 1, 'Decision'),
  }
}

function parseQuoteInput(value: unknown): QuoteInput {
  const record = readRecord(value, 'request body')
  return {
    units: readBoundedNumber(record.units, 'units', 1),
    listPrice: readBoundedNumber(record.listPrice, 'listPrice', 0),
    discount: readBoundedNumber(record.discount, 'discount', 0, 0.95),
    unitCost: readBoundedNumber(record.unitCost, 'unitCost', 0),
    minimumMargin: readBoundedNumber(record.minimumMargin, 'minimumMargin', 0, 1),
  }
}

function requireSheet(workbook: WorkPaperInstance, name: string): number {
  const sheet = workbook.getSheetId(name)
  if (sheet === undefined) {
    throw new Error(`missing sheet: ${name}`)
  }
  return sheet
}

function readNumber(workbook: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  return Math.round(readCellNumber(workbook, sheet, row, col, label) * 100) / 100
}

function readRoundedNumber(workbook: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  return Math.round(readCellNumber(workbook, sheet, row, col, label) * 10_000) / 10_000
}

function readCellNumber(workbook: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  const cell: unknown = workbook.getCellValue({ sheet, row, col })
  if (!isRecord(cell) || typeof cell.value !== 'number') {
    throw new Error(`expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function readString(workbook: WorkPaperInstance, sheet: number, row: number, col: number, label: string): string {
  const cell: unknown = workbook.getCellValue({ sheet, row, col })
  if (!isRecord(cell) || typeof cell.value !== 'string') {
    throw new Error(`expected ${label} to be text, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function readBoundedNumber(value: unknown, label: string, min: number, max = Number.POSITIVE_INFINITY): number {
  const numberValue = Number(value)
  if (!Number.isFinite(numberValue) || numberValue < min || numberValue > max) {
    throw new Error(`${label} must be a finite number between ${min.toString()} and ${max.toString()}`)
  }
  return numberValue
}

function readRecord(value: unknown, label: string): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new Error(`${label} must be a JSON object`)
  }
  return value
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function json(payload: unknown, status = 200): Response {
  return Response.json(payload, { status, headers: { 'cache-control': 'no-store' } })
}

async function runSmoke(): Promise<void> {
  const handler = createQuoteApprovalRequestHandler(createMemoryStorage())
  const before = await requestJson(handler, '/api/quote/approval')
  const edit = await requestJson(handler, '/api/quote/approval', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({
      units: 40,
      listPrice: 1200,
      discount: 0.05,
      unitCost: 760,
      minimumMargin: 0.3,
    }),
  })

  const output = {
    before,
    edit,
    verified: true,
    nextStep:
      'If this proof matches your service or agent workflow, star or bookmark Bilig: https://github.com/proompteng/bilig/stargazers',
  }
  assertSmokeOutput(output)
  console.log(JSON.stringify(output, null, 2))
}

async function requestJson(handler: (request: Request) => Promise<Response>, path: string, init?: RequestInit): Promise<unknown> {
  const response = await handler(new Request(`http://localhost:8788${path}`, init))
  const body: unknown = await response.json()
  if (!response.ok) {
    throw new Error(`request failed: ${response.status.toString()} ${JSON.stringify(body)}`)
  }
  return body
}

function assertSmokeOutput(value: unknown): void {
  const output = readRecord(value, 'smoke output')
  const edit = readRecord(output.edit, 'smoke edit')
  const checks = readRecord(edit.checks, 'smoke checks')
  const after = readRecord(edit.after, 'smoke after')
  const restored = readRecord(edit.restored, 'smoke restored')

  if (
    output.verified !== true ||
    after.decision !== 'approved' ||
    JSON.stringify(after) !== JSON.stringify(restored) ||
    checks.decisionChanged !== true ||
    checks.formulasPersisted !== true ||
    checks.restoredMatchesAfter !== true ||
    Number(checks.serializedBytes) <= 0
  ) {
    throw new Error(`unexpected smoke output: ${JSON.stringify(value)}`)
  }
}

async function toWebRequest(incoming: IncomingMessage): Promise<Request> {
  const origin = `http://${incoming.headers.host ?? 'localhost:8788'}`
  const headers = new Headers()

  for (const [name, value] of Object.entries(incoming.headers)) {
    if (Array.isArray(value)) {
      headers.set(name, value.join(', '))
    } else if (value !== undefined) {
      headers.set(name, value)
    }
  }

  return new Request(new URL(incoming.url ?? '/', origin), {
    method: incoming.method,
    headers,
    body: incoming.method === 'GET' || incoming.method === 'HEAD' ? undefined : await readIncomingBody(incoming),
    duplex: 'half',
  } as RequestInit & { duplex: 'half' })
}

function readIncomingBody(incoming: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = []
    incoming.on('data', (chunk: Buffer | string) => {
      chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : chunk)
    })
    incoming.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')))
    incoming.on('error', reject)
  })
}

if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  if (process.argv.includes('--serve')) {
    const handler = createQuoteApprovalRequestHandler(createMemoryStorage())
    createServer(async (incoming, outgoing) => {
      try {
        const response = await handler(await toWebRequest(incoming))
        outgoing.writeHead(response.status, Object.fromEntries(response.headers))
        outgoing.end(Buffer.from(await response.arrayBuffer()))
      } catch (error) {
        outgoing.writeHead(500, { 'content-type': 'application/json; charset=utf-8' })
        outgoing.end(`${JSON.stringify({ error: error instanceof Error ? error.message : String(error) })}\n`)
      }
    }).listen(8788, () => {
      console.log('Quote approval WorkPaper API listening on http://localhost:8788')
    })
  } else {
    await runSmoke()
  }
}
