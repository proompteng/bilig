import { createServer } from 'node:http'
import { WorkPaper } from '@bilig/headless'

const opportunityRecords = [
  {
    account: 'Acme Manufacturing',
    region: 'West',
    stage: 'Committed',
    seats: 12,
    arpa: 1800,
    probability: 1,
  },
  {
    account: 'Beacon Health',
    region: 'East',
    stage: 'Pipeline',
    seats: 8,
    arpa: 950,
    probability: 0.5,
  },
  {
    account: 'Cobalt Finance',
    region: 'West',
    stage: 'Committed',
    seats: 15,
    arpa: 1200,
    probability: 1,
  },
]

const server = createServer(async (request, response) => {
  try {
    if (request.method !== 'POST' || request.url !== '/summary') {
      writeJson(response, 404, { error: 'not_found' })
      return
    }

    const records = parseOpportunityRecords(await readRequestBody(request))
    writeJson(response, 200, summarizeOpportunities(records))
  } catch (error) {
    writeJson(response, 400, {
      error: error instanceof Error ? error.message : 'invalid_request',
    })
  }
})

await listen(server)

try {
  const { port } = server.address()
  const result = await fetch(`http://127.0.0.1:${port}/summary`, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
    },
    body: JSON.stringify(opportunityRecords),
  })

  if (!result.ok) {
    throw new Error(`Unexpected HTTP status ${result.status}`)
  }

  const output = await result.json()
  assertOutput(output)
  console.log(JSON.stringify(output, null, 2))
} finally {
  await close(server)
}

function summarizeOpportunities(records) {
  const opportunityRows = records.map((record, index) => {
    const spreadsheetRow = index + 2

    return [
      record.account,
      record.region,
      record.stage,
      record.seats,
      record.arpa,
      record.probability,
      `=D${spreadsheetRow}*E${spreadsheetRow}`,
      `=G${spreadsheetRow}*F${spreadsheetRow}`,
    ]
  })

  const workbook = WorkPaper.buildFromSheets({
    Opportunities: [['Account', 'Region', 'Stage', 'Seats', 'ARPA', 'Probability', 'Gross MRR', 'Weighted MRR'], ...opportunityRows],
    Summary: [
      ['Metric', 'Value'],
      ['Committed MRR', '=SUMIFS(Opportunities!G2:G4,Opportunities!C2:C4,"Committed")'],
      ['Weighted pipeline MRR', '=SUM(Opportunities!H2:H4)'],
      ['West seats', '=SUMIF(Opportunities!B2:B4,"West",Opportunities!D2:D4)'],
      ['Largest opportunity MRR', '=MAX(Opportunities!G2:G4)'],
    ],
  })

  const summarySheet = requireSheet(workbook, 'Summary')

  return {
    verified: true,
    sourceRecords: records.length,
    computed: {
      committedMrr: readNumber(workbook, summarySheet, 1, 1, 'Committed MRR'),
      weightedPipelineMrr: readNumber(workbook, summarySheet, 2, 1, 'Weighted pipeline MRR'),
      westSeats: readNumber(workbook, summarySheet, 3, 1, 'West seats'),
      largestOpportunityMrr: readNumber(workbook, summarySheet, 4, 1, 'Largest opportunity MRR'),
    },
  }
}

function parseOpportunityRecords(body) {
  const records = JSON.parse(body)
  if (!Array.isArray(records)) {
    throw new Error('expected JSON array')
  }

  return records.map((record, index) => {
    if (!record || typeof record !== 'object') {
      throw new Error(`record ${index + 1} must be an object`)
    }

    const parsed = {
      account: readString(record, 'account', index),
      region: readString(record, 'region', index),
      stage: readString(record, 'stage', index),
      seats: readNumberField(record, 'seats', index),
      arpa: readNumberField(record, 'arpa', index),
      probability: readNumberField(record, 'probability', index),
    }

    if (parsed.probability < 0 || parsed.probability > 1) {
      throw new Error(`record ${index + 1} probability must be between 0 and 1`)
    }

    return parsed
  })
}

function readString(record, field, index) {
  const value = record[field]
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`record ${index + 1} ${field} must be a non-empty string`)
  }
  return value
}

function readNumberField(record, field, index) {
  const value = record[field]
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`record ${index + 1} ${field} must be a finite number`)
  }
  return value
}

function requireSheet(workpaper, sheetName) {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function readNumber(workpaper, sheet, row, col, label) {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function assertOutput(actual) {
  const expected = {
    verified: true,
    sourceRecords: 3,
    computed: {
      committedMrr: 39600,
      weightedPipelineMrr: 43400,
      westSeats: 27,
      largestOpportunityMrr: 21600,
    },
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected HTTP JSON summary result: ${JSON.stringify(actual)}`)
  }
}

function readRequestBody(request) {
  return new Promise((resolve, reject) => {
    let body = ''

    request.setEncoding('utf8')
    request.on('data', (chunk) => {
      body += chunk
    })
    request.on('end', () => resolve(body))
    request.on('error', reject)
  })
}

function writeJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    'content-type': 'application/json; charset=utf-8',
  })
  response.end(JSON.stringify(payload))
}

function listen(httpServer) {
  return new Promise((resolve, reject) => {
    httpServer.once('error', reject)
    httpServer.listen(0, '127.0.0.1', () => {
      httpServer.off('error', reject)
      resolve()
    })
  })
}

function close(httpServer) {
  return new Promise((resolve, reject) => {
    httpServer.close((error) => {
      if (error) {
        reject(error)
        return
      }
      resolve()
    })
  })
}
