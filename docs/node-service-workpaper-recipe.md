# WorkPaper Node Service Recipe

This recipe shows how to put `@bilig/headless` behind a small Node service
boundary. It uses Node's built-in `node:http` module so evaluators can see the
service shape without adopting a web framework.

Use this when a backend job, queue worker, API route, or agent tool needs
formula-backed workbook state with controlled edits and persistence. Start with
the package contract in
[`packages/headless/README.md`](../packages/headless/README.md).

## Setup

```sh
mkdir bilig-workpaper-service
cd bilig-workpaper-service
npm init -y
npm pkg set type=module
npm pkg set scripts.start="node service.mjs"
npm install @bilig/headless
```

Create `service.mjs`:

```js
import { createServer } from 'node:http'
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

let persistedWorkbook = serializeWorkbook(createInitialWorkbook())

const server = createServer(async (request, response) => {
  try {
    const url = new URL(request.url ?? '/', 'http://localhost:8787')

    if (request.method === 'GET' && url.pathname === '/summary') {
      const workbook = loadWorkbook()
      sendJson(response, 200, {
        summary: readSummary(workbook),
        sheets: workbook.getSheetNames(),
      })
      return
    }

    if (request.method === 'POST' && url.pathname === '/assumptions/customers') {
      const body = await readJsonBody(request)
      const customers = body.customers
      if (typeof customers !== 'number' || !Number.isFinite(customers) || customers < 0) {
        sendJson(response, 400, { error: 'customers must be a non-negative number' })
        return
      }

      const workbook = loadWorkbook()
      const before = readSummary(workbook)
      setAssumption(workbook, 'Customers', customers)
      const after = readSummary(workbook)
      persistedWorkbook = serializeWorkbook(workbook)

      sendJson(response, 200, {
        before,
        after,
        checks: {
          grossMrrChanged: before.grossMrr !== after.grossMrr,
          annualizedArrChanged: before.annualizedArr !== after.annualizedArr,
          serializedBytes: Buffer.byteLength(persistedWorkbook, 'utf8'),
        },
      })
      return
    }

    sendJson(response, 404, { error: 'not found' })
  } catch (error) {
    sendJson(response, 500, {
      error: error instanceof Error ? error.message : String(error),
    })
  }
})

server.listen(8787, () => {
  console.log('WorkPaper service listening on http://localhost:8787')
})

function createInitialWorkbook() {
  return WorkPaper.buildFromSheets({
    Assumptions: [
      ['Metric', 'Value'],
      ['Customers', 40],
      ['ARPA', 240],
      ['Expansion factor', 1.1],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Gross MRR', '=Assumptions!B2*Assumptions!B3'],
      ['Expansion MRR', '=B2*Assumptions!B4'],
      ['Annualized ARR', '=B3*12'],
    ],
  })
}

function loadWorkbook() {
  return createWorkPaperFromDocument(parseWorkPaperDocument(persistedWorkbook))
}

function serializeWorkbook(workbook) {
  return serializeWorkPaperDocument(
    exportWorkPaperDocument(workbook, {
      includeConfig: true,
    }),
  )
}

function readSummary(workbook) {
  const summary = requireSheet(workbook, 'Summary')
  return {
    grossMrr: readNumber(workbook, summary, 1, 1, 'Gross MRR'),
    expansionMrr: readNumber(workbook, summary, 2, 1, 'Expansion MRR'),
    annualizedArr: readNumber(workbook, summary, 3, 1, 'Annualized ARR'),
  }
}

function setAssumption(workbook, metricName, value) {
  const assumptions = requireSheet(workbook, 'Assumptions')
  const rows = workbook.getSheetSerialized(assumptions)

  for (let row = 1; row < rows.length; row += 1) {
    if (rows[row]?.[0] === metricName) {
      workbook.setCellContents({ sheet: assumptions, row, col: 1 }, value)
      return
    }
  }

  throw new Error(`unknown assumption: ${metricName}`)
}

function requireSheet(workbook, sheetName) {
  const sheet = workbook.getSheetId(sheetName)
  if (sheet === undefined) {
    throw new Error(`missing sheet: ${sheetName}`)
  }
  return sheet
}

function readNumber(workbook, sheet, row, col, label) {
  const value = workbook.getCellValue({ sheet, row, col })
  if (typeof value !== 'object' || value === null || typeof value.value !== 'number') {
    throw new Error(`expected ${label} to be numeric, received ${JSON.stringify(value)}`)
  }
  return Math.round(value.value * 100) / 100
}

async function readJsonBody(request) {
  let body = ''
  for await (const chunk of request) {
    body += chunk
  }
  return body ? JSON.parse(body) : {}
}

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    'content-type': 'application/json; charset=utf-8',
  })
  response.end(`${JSON.stringify(payload, null, 2)}\n`)
}
```

Run it:

```sh
npm start
```

From another terminal:

```sh
curl -s http://localhost:8787/summary
curl -s -X POST http://localhost:8787/assumptions/customers \
  -H 'content-type: application/json' \
  -d '{"customers":65}'
curl -s http://localhost:8787/summary
```

The edit response should include a computed before/after check:

```json
{
  "before": {
    "grossMrr": 9600,
    "expansionMrr": 10560,
    "annualizedArr": 126720
  },
  "after": {
    "grossMrr": 15600,
    "expansionMrr": 17160,
    "annualizedArr": 205920
  },
  "checks": {
    "grossMrrChanged": true,
    "annualizedArrChanged": true,
    "serializedBytes": 1097
  }
}
```

`serializedBytes` can change as the persisted document schema evolves. Treat it
as a positive persistence signal, not a golden value.

## Service Boundary Notes

- Keep the WorkPaper object inside the service boundary. External callers
  should send narrow business inputs and receive computed summaries or validation
  errors.
- Use `WorkPaper.buildFromSheets()` for hand-authored service fixtures and
  `WorkPaper.buildFromSnapshot()` for importer-produced snapshots.
- Use `exportWorkPaperDocument()`, `serializeWorkPaperDocument()`,
  `parseWorkPaperDocument()`, and `createWorkPaperFromDocument()` for persisted
  state.
- Return computed values after every controlled edit. A successful HTTP status
  only proves the route ran; readback proves the workbook recalculated.
- Use public `@bilig/headless` exports only. Do not import from this monorepo's
  internal `src/` or `dist/` paths in a consumer service.

## Validation

For the standalone copy-paste recipe, run:

```sh
npm start
curl -s -X POST http://localhost:8787/assumptions/customers \
  -H 'content-type: application/json' \
  -d '{"customers":65}'
```

For a documentation patch in this repository, run:

```sh
pnpm docs:discovery:check
pnpm run ci
```
