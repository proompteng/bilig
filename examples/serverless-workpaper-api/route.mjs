import { createServer } from 'node:http'
import { Readable } from 'node:stream'
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

export async function handleWorkPaperRequest(request) {
  const url = new URL(request.url)

  if (request.method === 'GET' && url.pathname === '/api/workpaper/summary') {
    const workbook = loadWorkbook()
    return json({
      summary: readSummary(workbook),
      sheets: workbook.getSheetNames(),
    })
  }

  if (request.method === 'POST' && url.pathname === '/api/workpaper/revenue') {
    let records
    try {
      const body = await request.json()
      records = normalizeRevenueRecords(body.records)
    } catch (error) {
      return json({ error: error instanceof Error ? error.message : String(error) }, 400)
    }

    const before = readSummary(loadWorkbook())
    const workbook = buildRevenueWorkbook(records)
    const after = readSummary(workbook)
    state.workbookJson = serializeWorkbook(workbook)

    return json({
      records: records.length,
      before,
      after,
      checks: {
        totalRevenueChanged: before.totalRevenue !== after.totalRevenue,
        formulasPersisted: state.workbookJson.includes('=SUM(Revenue!D2:D'),
        serializedBytes: Buffer.byteLength(state.workbookJson, 'utf8'),
      },
    })
  }

  return json({ error: 'not found' }, 404)
}

function createInitialWorkbook() {
  return buildRevenueWorkbook([
    { region: 'West', customers: 20, arpa: 1200 },
    { region: 'East', customers: 30, arpa: 250 },
    { region: 'Central', customers: 18, arpa: 300 },
  ])
}

function buildRevenueWorkbook(records) {
  const dataRows = records.map((record, index) => {
    const spreadsheetRow = index + 2
    return [record.region, record.customers, record.arpa, `=B${spreadsheetRow}*C${spreadsheetRow}`]
  })
  const lastDataRow = records.length + 1

  return WorkPaper.buildFromSheets({
    Revenue: [['Region', 'Customers', 'ARPA', 'Revenue'], ...dataRows],
    Summary: [
      ['Metric', 'Value'],
      ['Total revenue', `=SUM(Revenue!D2:D${lastDataRow})`],
      ['West customers', `=SUMIF(Revenue!A2:A${lastDataRow},"West",Revenue!B2:B${lastDataRow})`],
      ['Largest deal', `=MAX(Revenue!D2:D${lastDataRow})`],
    ],
  })
}

function loadWorkbook() {
  return createWorkPaperFromDocument(parseWorkPaperDocument(state.workbookJson))
}

function serializeWorkbook(workbook) {
  return serializeWorkPaperDocument(
    exportWorkPaperDocument(workbook, {
      includeConfig: true,
    }),
  )
}

function normalizeRevenueRecords(value) {
  if (!Array.isArray(value) || value.length === 0) {
    throw new Error('records must be a non-empty array')
  }

  return value.map((record, index) => {
    if (typeof record !== 'object' || record === null) {
      throw new Error(`record ${index + 1} must be an object`)
    }

    const region = record.region
    const customers = Number(record.customers)
    const arpa = Number(record.arpa)
    if (typeof region !== 'string' || region.trim() === '') {
      throw new Error(`record ${index + 1} needs a region`)
    }
    if (!Number.isFinite(customers) || customers < 0) {
      throw new Error(`record ${index + 1} needs non-negative customers`)
    }
    if (!Number.isFinite(arpa) || arpa < 0) {
      throw new Error(`record ${index + 1} needs non-negative arpa`)
    }

    return {
      region: region.trim(),
      customers,
      arpa,
    }
  })
}

function readSummary(workbook) {
  const summary = requireSheet(workbook, 'Summary')
  return {
    totalRevenue: readNumber(workbook, summary, 1, 1, 'Total revenue'),
    westCustomers: readNumber(workbook, summary, 2, 1, 'West customers'),
    largestDeal: readNumber(workbook, summary, 3, 1, 'Largest deal'),
  }
}

function requireSheet(workbook, name) {
  const sheet = workbook.getSheetId(name)
  if (sheet === undefined) {
    throw new Error(`missing sheet: ${name}`)
  }
  return sheet
}

function readNumber(workbook, sheet, row, col, label) {
  const cell = workbook.getCellValue({ sheet, row, col })
  if (typeof cell !== 'object' || cell === null || typeof cell.value !== 'number') {
    throw new Error(`expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function json(payload, status = 200) {
  return Response.json(payload, {
    status,
    headers: {
      'cache-control': 'no-store',
    },
  })
}

if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  createServer(async (incoming, outgoing) => {
    try {
      const request = toWebRequest(incoming)
      const response = await handleWorkPaperRequest(request)
      outgoing.writeHead(response.status, Object.fromEntries(response.headers))
      outgoing.end(Buffer.from(await response.arrayBuffer()))
    } catch (error) {
      outgoing.writeHead(500, { 'content-type': 'application/json; charset=utf-8' })
      outgoing.end(`${JSON.stringify({ error: error instanceof Error ? error.message : String(error) })}\n`)
    }
  }).listen(8787, () => {
    console.log('WorkPaper API route listening on http://localhost:8787')
  })
}

function toWebRequest(incoming) {
  const origin = `http://${incoming.headers.host ?? 'localhost:8787'}`
  const url = new URL(incoming.url ?? '/', origin)
  const headers = new Headers()

  for (const [name, value] of Object.entries(incoming.headers)) {
    if (Array.isArray(value)) {
      headers.set(name, value.join(', '))
    } else if (value !== undefined) {
      headers.set(name, value)
    }
  }

  return new Request(url, {
    method: incoming.method,
    headers,
    body: incoming.method === 'GET' || incoming.method === 'HEAD' ? undefined : Readable.toWeb(incoming),
    duplex: 'half',
  })
}
