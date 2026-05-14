import { pathToFileURL } from 'node:url'

import { createInMemoryWorkbookStorage, createWorkPaperRequestHandler } from './route.ts'

type RevenueRecordInput = {
  region: string
  customers: number
  arpa: number
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
  before: Summary
  checks: {
    formulasPersisted: boolean
    serializedBytes: number
    totalRevenueChanged: boolean
  }
  records: number
}

const handleWorkPaperRequest = createWorkPaperRequestHandler(createInMemoryWorkbookStorage())

export async function readRevenueSummaryAction(): Promise<SummaryResponse> {
  'use server'

  return requestJson('/api/workpaper/summary', parseSummaryResponse, 'Next.js Server Action summary read')
}

export async function updateRevenueRecordsAction(records: readonly RevenueRecordInput[]): Promise<EditResponse> {
  'use server'

  return requestJson('/api/workpaper/revenue', parseEditResponse, 'Next.js Server Action revenue edit', {
    body: JSON.stringify({ records }),
    headers: {
      'content-type': 'application/json',
    },
    method: 'POST',
  })
}

export async function createNextServerActionDemoOutput() {
  const before = await readRevenueSummaryAction()
  const edit = await updateRevenueRecordsAction([
    { region: 'West', customers: 20, arpa: 1200 },
    { region: 'East', customers: 30, arpa: 250 },
    { region: 'Central', customers: 18, arpa: 300 },
    { region: 'North', customers: 65, arpa: 180 },
  ])
  const after = await readRevenueSummaryAction()

  const output = {
    action: 'Next.js Server Action',
    before: before.summary,
    edit: {
      before: edit.before,
      after: edit.after,
      checks: edit.checks,
      records: edit.records,
    },
    after: after.summary,
    verified: true,
  }

  assertOutput(output)
  return output
}

if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  console.log(JSON.stringify(await createNextServerActionDemoOutput(), null, 2))
}

async function requestJson<T>(path: string, parse: (value: unknown) => T, label: string, init?: RequestInit): Promise<T> {
  const response = await handleWorkPaperRequest(new Request(`http://localhost:3000${path}`, init))
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
    before: readSummary(record.before, 'edit response before'),
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

function assertOutput(actual: Awaited<ReturnType<typeof createNextServerActionDemoOutput>>): void {
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
    actual.action !== 'Next.js Server Action' ||
    JSON.stringify(actual.before) !== JSON.stringify(expectedBefore) ||
    JSON.stringify(actual.edit.before) !== JSON.stringify(expectedBefore) ||
    JSON.stringify(actual.edit.after) !== JSON.stringify(expectedAfter) ||
    JSON.stringify(actual.after) !== JSON.stringify(expectedAfter) ||
    actual.edit.records !== 4 ||
    !actual.edit.checks.totalRevenueChanged ||
    !actual.edit.checks.formulasPersisted ||
    actual.edit.checks.serializedBytes <= 0
  ) {
    throw new Error(`unexpected Next.js Server Action output: ${JSON.stringify(actual)}`)
  }
}
