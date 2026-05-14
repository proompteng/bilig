import { pathToFileURL } from 'node:url'

import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const state = {
  workbookJson: serializeWorkbook(createInitialWorkbook()),
}

export const dynamic = 'force-dynamic'
export const runtime = 'nodejs'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>
type RevenueModel = {
  arpa: number
  customers: number
  revenue: number
}
type RouteEditResponse = {
  input: {
    cell: 'Inputs!B2'
    customers: number
  }
  before: RevenueModel
  formulaReadback: {
    cell: 'Summary!B2'
    revenue: number
  }
  persistence: {
    formulasPersisted: boolean
    inputPersisted: boolean
    persistedRevenue: number
    serializedBytes: number
  }
}

export async function GET(): Promise<Response> {
  const workbook = loadWorkbook()
  return json({
    model: readRevenueModel(workbook),
    sheets: workbook.getSheetNames(),
  })
}

export async function POST(request: Request): Promise<Response> {
  let customers: number
  try {
    customers = readCustomersInput(await request.json())
  } catch (error) {
    return json({ error: error instanceof Error ? error.message : String(error) }, 400)
  }

  const workbook = loadWorkbook()
  const before = readRevenueModel(workbook)
  const inputs = requireSheet(workbook, 'Inputs')
  workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, customers)

  const after = readRevenueModel(workbook)
  const workbookJson = serializeWorkbook(workbook)
  state.workbookJson = workbookJson

  const persistedWorkbook = loadWorkbook()
  const persisted = readRevenueModel(persistedWorkbook)
  const output: RouteEditResponse = {
    input: {
      cell: 'Inputs!B2',
      customers,
    },
    before,
    formulaReadback: {
      cell: 'Summary!B2',
      revenue: after.revenue,
    },
    persistence: {
      formulasPersisted: workbookJson.includes('=Inputs!B2*Inputs!B3'),
      inputPersisted: persisted.customers === customers,
      persistedRevenue: persisted.revenue,
      serializedBytes: Buffer.byteLength(workbookJson, 'utf8'),
    },
  }

  return json(output)
}

export async function createNextRouteHandlerDemoOutput() {
  const before = await requestJson(
    GET(),
    (value) => ({ model: readRevenueModelPayload(readJsonRecord(value, 'summary response').model, 'summary response model') }),
    'next route summary before',
  )
  const edit = await requestJson(
    POST(
      new Request('http://localhost:3000/api/workpaper/model', {
        method: 'POST',
        headers: {
          'content-type': 'application/json',
        },
        body: JSON.stringify({ customers: 65 }),
      }),
    ),
    parseRouteEditResponse,
    'next route JSON edit',
  )
  const after = await requestJson(
    GET(),
    (value) => ({ model: readRevenueModelPayload(readJsonRecord(value, 'summary response').model, 'summary response model') }),
    'next route summary after',
  )

  const output = {
    route: 'Next.js Route Handler JSON',
    runtime,
    dynamic,
    before: before.model,
    edit,
    after: after.model,
    verified: true,
  }

  assertOutput(output)
  return output
}

if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  console.log(JSON.stringify(await createNextRouteHandlerDemoOutput(), null, 2))
}

function createInitialWorkbook(): WorkPaperInstance {
  return WorkPaper.buildFromSheets({
    Inputs: [
      ['Metric', 'Value'],
      ['Customers', 20],
      ['ARPA', 1200],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Revenue', '=Inputs!B2*Inputs!B3'],
    ],
  })
}

function loadWorkbook(): WorkPaperInstance {
  return createWorkPaperFromDocument(parseWorkPaperDocument(state.workbookJson))
}

function serializeWorkbook(workbook: WorkPaperInstance): string {
  return serializeWorkPaperDocument(
    exportWorkPaperDocument(workbook, {
      includeConfig: true,
    }),
  )
}

function readRevenueModel(workbook: WorkPaperInstance): RevenueModel {
  const inputs = requireSheet(workbook, 'Inputs')
  const summary = requireSheet(workbook, 'Summary')
  return {
    arpa: readNumberCell(workbook, inputs, 2, 1, 'Inputs!B3'),
    customers: readNumberCell(workbook, inputs, 1, 1, 'Inputs!B2'),
    revenue: readNumberCell(workbook, summary, 1, 1, 'Summary!B2'),
  }
}

function requireSheet(workbook: WorkPaperInstance, name: string): number {
  const sheet = workbook.getSheetId(name)
  if (sheet === undefined) {
    throw new Error(`missing sheet: ${name}`)
  }
  return sheet
}

function readNumberCell(workbook: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  const cell = workbook.getCellValue({ sheet, row, col })
  if (typeof cell !== 'object' || cell === null || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function readCustomersInput(value: unknown): number {
  const record = readJsonRecord(value, 'request body')
  const customers = Number(record.customers)
  if (!Number.isFinite(customers) || customers < 0) {
    throw new Error('customers must be a non-negative number')
  }
  return customers
}

async function requestJson<T>(responsePromise: Promise<Response>, parse: (value: unknown) => T, label: string): Promise<T> {
  const response = await responsePromise
  const body: unknown = await response.json()
  if (!response.ok) {
    throw new Error(`${label} failed: ${response.status} ${JSON.stringify(body)}`)
  }
  return parse(body)
}

function parseRouteEditResponse(value: unknown): RouteEditResponse {
  const record = readJsonRecord(value, 'edit response')
  const input = readJsonRecord(record.input, 'edit response input')
  const formulaReadback = readJsonRecord(record.formulaReadback, 'edit response formulaReadback')
  const persistence = readJsonRecord(record.persistence, 'edit response persistence')
  return {
    input: {
      cell: readLiteral(input.cell, 'Inputs!B2', 'edit response input cell'),
      customers: readNumber(input.customers, 'edit response input customers'),
    },
    before: readRevenueModelPayload(record.before, 'edit response before'),
    formulaReadback: {
      cell: readLiteral(formulaReadback.cell, 'Summary!B2', 'edit response formulaReadback cell'),
      revenue: readNumber(formulaReadback.revenue, 'edit response formulaReadback revenue'),
    },
    persistence: {
      formulasPersisted: readBoolean(persistence.formulasPersisted, 'edit response formulasPersisted'),
      inputPersisted: readBoolean(persistence.inputPersisted, 'edit response inputPersisted'),
      persistedRevenue: readNumber(persistence.persistedRevenue, 'edit response persistedRevenue'),
      serializedBytes: readNumber(persistence.serializedBytes, 'edit response serializedBytes'),
    },
  }
}

function readRevenueModelPayload(value: unknown, label: string): RevenueModel {
  const record = readJsonRecord(value, label)
  return {
    arpa: readNumber(record.arpa, `${label} arpa`),
    customers: readNumber(record.customers, `${label} customers`),
    revenue: readNumber(record.revenue, `${label} revenue`),
  }
}

function readJsonRecord(value: unknown, label: string): Record<string, unknown> {
  if (!isJsonRecord(value)) {
    throw new Error(`${label} must be an object`)
  }
  return value
}

function isJsonRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readNumber(value: unknown, label: string): number {
  if (typeof value !== 'number') {
    throw new Error(`${label} must be a number`)
  }
  return value
}

function readBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`${label} must be a boolean`)
  }
  return value
}

function readLiteral<T extends string>(value: unknown, expected: T, label: string): T {
  if (value !== expected) {
    throw new Error(`${label} must be ${expected}`)
  }
  return expected
}

function json(payload: unknown, status = 200): Response {
  return Response.json(payload, {
    status,
    headers: {
      'cache-control': 'no-store',
    },
  })
}

function assertOutput(actual: Awaited<ReturnType<typeof createNextRouteHandlerDemoOutput>>): void {
  const expectedBefore = {
    arpa: 1200,
    customers: 20,
    revenue: 24000,
  }
  const expectedAfter = {
    arpa: 1200,
    customers: 65,
    revenue: 78000,
  }

  if (
    actual.route !== 'Next.js Route Handler JSON' ||
    actual.runtime !== 'nodejs' ||
    actual.dynamic !== 'force-dynamic' ||
    JSON.stringify(actual.before) !== JSON.stringify(expectedBefore) ||
    actual.edit.input.cell !== 'Inputs!B2' ||
    actual.edit.input.customers !== 65 ||
    JSON.stringify(actual.edit.before) !== JSON.stringify(expectedBefore) ||
    actual.edit.formulaReadback.cell !== 'Summary!B2' ||
    actual.edit.formulaReadback.revenue !== expectedAfter.revenue ||
    !actual.edit.persistence.formulasPersisted ||
    !actual.edit.persistence.inputPersisted ||
    actual.edit.persistence.persistedRevenue !== expectedAfter.revenue ||
    actual.edit.persistence.serializedBytes <= 0 ||
    JSON.stringify(actual.after) !== JSON.stringify(expectedAfter)
  ) {
    throw new Error(`unexpected Next.js route handler JSON output: ${JSON.stringify(actual)}`)
  }
}
