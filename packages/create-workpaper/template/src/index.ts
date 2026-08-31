import { createServer, type IncomingMessage } from 'node:http'
import { pathToFileURL } from 'node:url'

import { buildA1WorkPaper, restoreA1WorkPaper, type A1EditManyAndReadbackProof, type A1WorkPaper } from '@bilig/workpaper'

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

const outputCells = {
  listRevenue: 'Summary!B2',
  discountAmount: 'Summary!B3',
  netRevenue: 'Summary!B4',
  totalCost: 'Summary!B5',
  grossMargin: 'Summary!B6',
  decision: 'Summary!B7',
} as const

const decisionFormula = '=IF(B6>=Inputs!B6,"approved","review")'

export function createQuoteApprovalRequestHandler(storage: WorkbookStorage) {
  return async function handleQuoteApprovalRequest(request: Request): Promise<Response> {
    const url = new URL(request.url)

    if (request.method === 'GET' && url.pathname === '/api/quote/approval') {
      const workbook = await loadWorkbook(storage)
      try {
        return json({ summary: readQuoteSummary(workbook), inputCells, outputCells })
      } finally {
        workbook.dispose()
      }
    }

    if (request.method === 'POST' && url.pathname === '/api/quote/approval') {
      try {
        const input = parseQuoteInput(await request.json())
        const workbook = await loadWorkbook(storage)
        try {
          const before = readQuoteSummary(workbook)
          const proof = workbook.editManyAndReadback(
            {
              [inputCells.units]: input.units,
              [inputCells.listPrice]: input.listPrice,
              [inputCells.discount]: input.discount,
              [inputCells.unitCost]: input.unitCost,
              [inputCells.minimumMargin]: input.minimumMargin,
            },
            {
              includeConfig: true,
              includeSerializedDocument: true,
              readbackRange: 'Summary!B2:B7',
            },
          )
          const after = readQuoteSummary(workbook)
          const workbookJson = proof.serializedDocument ?? workbook.saveJson({ includeConfig: true })
          await storage.saveWorkbookJson(workbookJson)

          const restored = workbook.restoreJson(workbookJson)
          try {
            const restoredSummary = readQuoteSummary(restored)

            return json({
              input,
              inputCells,
              outputCells,
              before,
              after,
              restored: restoredSummary,
              proof: summarizeProof(proof),
              checks: {
                decisionChanged: before.decision !== after.decision,
                formulasPersisted: restored.formula(outputCells.decision) === decisionFormula,
                restoredMatchesAfter: JSON.stringify(restoredSummary) === JSON.stringify(after),
                serializedBytes: Buffer.byteLength(workbookJson, 'utf8'),
              },
            })
          } finally {
            restored.dispose()
          }
        } finally {
          workbook.dispose()
        }
      } catch (error) {
        return json({ error: error instanceof Error ? error.message : String(error) }, 400)
      }
    }

    return json({ error: 'not found' }, 404)
  }
}

function createQuoteApprovalWorkbook(): A1WorkPaper {
  const workbook = buildA1WorkPaper(
    {
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
        ['Decision', decisionFormula],
      ],
    },
    undefined,
    { writableSheets: ['Inputs'] },
  )
  if (!workbook.validateFormula(decisionFormula)) {
    workbook.dispose()
    throw new Error(`invalid decision formula: ${decisionFormula}`)
  }
  return workbook
}

function createMemoryStorage(): WorkbookStorage {
  const initialWorkbook = createQuoteApprovalWorkbook()
  let workbookJson: string
  try {
    workbookJson = initialWorkbook.saveJson({ includeConfig: true })
  } finally {
    initialWorkbook.dispose()
  }
  return {
    loadWorkbookJson() {
      return workbookJson
    },
    saveWorkbookJson(nextWorkbookJson) {
      workbookJson = nextWorkbookJson
    },
  }
}

async function loadWorkbook(storage: WorkbookStorage): Promise<A1WorkPaper> {
  return restoreA1WorkPaper(await storage.loadWorkbookJson(), { writableSheets: ['Inputs'] })
}

function readQuoteSummary(workbook: A1WorkPaper): QuoteSummary {
  return {
    listRevenue: readNumber(workbook, outputCells.listRevenue, 'List revenue'),
    discountAmount: readNumber(workbook, outputCells.discountAmount, 'Discount amount'),
    netRevenue: readNumber(workbook, outputCells.netRevenue, 'Net revenue'),
    totalCost: readNumber(workbook, outputCells.totalCost, 'Total cost'),
    grossMargin: readRoundedNumber(workbook, outputCells.grossMargin, 'Gross margin'),
    decision: readString(workbook, outputCells.decision, 'Decision'),
  }
}

function summarizeProof(proof: A1EditManyAndReadbackProof) {
  return {
    editedCells: proof.editedCells,
    readbackRange: proof.readbackRange,
    afterReadback: proof.afterReadback.displayValues,
    restoredReadback: proof.restoredReadback.displayValues,
    persistedDocumentBytes: proof.persistedDocumentBytes,
    checks: {
      computedReadbackChanged: proof.checks.computedReadbackChanged,
      blockingFormulaDiagnosticCount: proof.checks.blockingFormulaDiagnosticCount,
      restoredReadbackMatchesAfter: proof.checks.restoredReadbackMatchesAfter,
    },
    verified: proof.verified,
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

function readNumber(workbook: A1WorkPaper, address: string, label: string): number {
  return Math.round(readCellNumber(workbook, address, label) * 100) / 100
}

function readRoundedNumber(workbook: A1WorkPaper, address: string, label: string): number {
  return Math.round(readCellNumber(workbook, address, label) * 10_000) / 10_000
}

function readCellNumber(workbook: A1WorkPaper, address: string, label: string): number {
  const cell: unknown = workbook.get(address)
  if (!isRecord(cell) || typeof cell.value !== 'number') {
    throw new Error(`expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function readString(workbook: A1WorkPaper, address: string, label: string): string {
  const cell: unknown = workbook.get(address)
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
  const proof = readRecord(edit.proof, 'smoke proof')
  const after = readRecord(edit.after, 'smoke after')
  const restored = readRecord(edit.restored, 'smoke restored')

  if (
    output.verified !== true ||
    after.decision !== 'approved' ||
    JSON.stringify(after) !== JSON.stringify(restored) ||
    checks.decisionChanged !== true ||
    checks.formulasPersisted !== true ||
    checks.restoredMatchesAfter !== true ||
    !Array.isArray(proof.editedCells) ||
    proof.editedCells.length !== Object.keys(inputCells).length ||
    !proof.editedCells.includes(inputCells.minimumMargin) ||
    proof.verified !== true ||
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
