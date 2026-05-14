import { pathToFileURL } from 'node:url'

import { createInMemoryWorkbookStorage, createWorkPaperRequestHandler } from './route.ts'

const handleWorkPaperRequest = createWorkPaperRequestHandler(createInMemoryWorkbookStorage())

/**
 * Vercel Functions can expose web-standard method handlers from files under
 * `/api`. Keep each exported function thin and let the shared WorkPaper route
 * own formula evaluation, persistence, and JSON readback checks.
 */
export function GET(request: Request): Promise<Response> {
  return handleWorkPaperRequest(request)
}

export function POST(request: Request): Promise<Response> {
  return handleWorkPaperRequest(request)
}

/**
 * Some Vercel runtimes also accept a single Fetch-style default export. Use it
 * when one function file should own both WorkPaper route methods.
 */
const vercelFetchEntrypoint = {
  fetch(request: Request): Promise<Response> {
    return handleWorkPaperRequest(request)
  },
}

export default vercelFetchEntrypoint

export async function createVercelFunctionDemoOutput() {
  const before = await requestJson(
    GET(new Request('https://workpaper.example.vercel.app/api/workpaper/summary')),
    parseSummaryResponse,
    'vercel summary before',
  )
  const edit = await requestJson(
    POST(
      new Request('https://workpaper.example.vercel.app/api/workpaper/revenue', {
        method: 'POST',
        headers: {
          'content-type': 'application/json',
        },
        body: JSON.stringify({
          records: [
            { region: 'West', customers: 20, arpa: 1200 },
            { region: 'East', customers: 30, arpa: 250 },
            { region: 'Central', customers: 18, arpa: 300 },
            { region: 'North', customers: 65, arpa: 180 },
          ],
        }),
      }),
    ),
    parseEditResponse,
    'vercel revenue edit',
  )
  const after = await requestJson(
    vercelFetchEntrypoint.fetch(new Request('https://workpaper.example.vercel.app/api/workpaper/summary')),
    parseSummaryResponse,
    'vercel summary after',
  )

  const output = {
    route: 'Vercel Function',
    entrypoints: ['GET', 'POST', 'default.fetch'],
    before: before.summary,
    edit: {
      records: edit.records,
      after: edit.after,
      checks: edit.checks,
    },
    after: after.summary,
    verified: true,
  }

  assertOutput(output)
  return output
}

type Summary = {
  largestDeal: number
  totalRevenue: number
  westCustomers: number
}

type SummaryResponse = {
  summary: Summary
}

type EditResponse = {
  after: Summary
  checks: {
    formulasPersisted: boolean
    serializedBytes: number
    totalRevenueChanged: boolean
  }
  records: number
}

if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  console.log(JSON.stringify(await createVercelFunctionDemoOutput(), null, 2))
}

async function requestJson<T>(responsePromise: Promise<Response>, parse: (value: unknown) => T, label: string): Promise<T> {
  const response = await responsePromise
  const body: unknown = await response.json()
  if (!response.ok) {
    throw new Error(`${label} failed: ${response.status} ${JSON.stringify(body)}`)
  }
  return parse(body)
}

function parseSummaryResponse(value: unknown): SummaryResponse {
  const record = readJsonRecord(value, 'summary response')
  return {
    summary: readSummary(record.summary, 'summary response summary'),
  }
}

function parseEditResponse(value: unknown): EditResponse {
  const record = readJsonRecord(value, 'edit response')
  const checks = readJsonRecord(record.checks, 'edit response checks')
  return {
    after: readSummary(record.after, 'edit response after'),
    checks: {
      formulasPersisted: readBoolean(checks.formulasPersisted, 'edit response formulasPersisted'),
      serializedBytes: readNumber(checks.serializedBytes, 'edit response serializedBytes'),
      totalRevenueChanged: readBoolean(checks.totalRevenueChanged, 'edit response totalRevenueChanged'),
    },
    records: readNumber(record.records, 'edit response records'),
  }
}

function readSummary(value: unknown, label: string): Summary {
  const record = readJsonRecord(value, label)
  return {
    largestDeal: readNumber(record.largestDeal, `${label} largestDeal`),
    totalRevenue: readNumber(record.totalRevenue, `${label} totalRevenue`),
    westCustomers: readNumber(record.westCustomers, `${label} westCustomers`),
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

function assertOutput(actual: Awaited<ReturnType<typeof createVercelFunctionDemoOutput>>): void {
  const expectedBefore = {
    largestDeal: 24000,
    totalRevenue: 36900,
    westCustomers: 20,
  }
  const expectedAfter = {
    largestDeal: 24000,
    totalRevenue: 48600,
    westCustomers: 20,
  }

  if (
    actual.route !== 'Vercel Function' ||
    !actual.entrypoints.includes('GET') ||
    !actual.entrypoints.includes('POST') ||
    !actual.entrypoints.includes('default.fetch') ||
    JSON.stringify(actual.before) !== JSON.stringify(expectedBefore) ||
    JSON.stringify(actual.edit.after) !== JSON.stringify(expectedAfter) ||
    JSON.stringify(actual.after) !== JSON.stringify(expectedAfter) ||
    actual.edit.records !== 4 ||
    !actual.edit.checks.totalRevenueChanged ||
    !actual.edit.checks.formulasPersisted ||
    actual.edit.checks.serializedBytes <= 0
  ) {
    throw new Error(`unexpected Vercel Function output: ${JSON.stringify(actual)}`)
  }
}
