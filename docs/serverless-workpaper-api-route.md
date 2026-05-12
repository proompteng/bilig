# Serverless WorkPaper API Route

This recipe shows how to put `@bilig/headless` behind a small API route using
web-standard `Request` and `Response` objects. Use it when a serverless
function, route handler, queue worker, or coding-agent tool needs spreadsheet
formulas without keeping a browser grid open.

For a clone-and-run copy of this route, start with
[`examples/serverless-workpaper-api`](../examples/serverless-workpaper-api).

The example also includes a tiny Node adapter so you can run it locally before
moving the route into Vercel, Cloudflare Workers, Fastify, Hono, or another HTTP
surface.

## Setup

```sh
mkdir bilig-serverless-workpaper
cd bilig-serverless-workpaper
npm init -y
npm pkg set type=module
npm pkg set scripts.start="node route.mjs"
npm install @bilig/headless
```

Create `route.mjs`:

```js
import { createServer } from 'node:http'
import { Readable } from 'node:stream'
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
    const body = await request.json()
    const records = normalizeRevenueRecords(body.records)
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

if (import.meta.url === `file://${process.argv[1]}`) {
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
```

Run it:

```sh
npm start
```

From another terminal:

```sh
curl -s http://localhost:8787/api/workpaper/summary
curl -s -X POST http://localhost:8787/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
curl -s http://localhost:8787/api/workpaper/summary
```

The edit response should include formula readback and persistence checks:

```json
{
  "before": {
    "totalRevenue": 36900,
    "westCustomers": 20,
    "largestDeal": 24000
  },
  "after": {
    "totalRevenue": 48600,
    "westCustomers": 20,
    "largestDeal": 24000
  },
  "checks": {
    "totalRevenueChanged": true,
    "formulasPersisted": true,
    "serializedBytes": 1195
  }
}
```

`serializedBytes` can change as the persisted document schema evolves. Treat it
as a positive persistence signal, not a golden value.

## Moving Into A Framework

Keep the exported `handleWorkPaperRequest()` function as the stable boundary:

- In a Vercel or Next.js route handler, call it from `GET()` and `POST()`.
- In Cloudflare Workers, call it from `fetch(request)`.
- In Hono, Fastify, or Express, adapt the framework request into a web-standard
  `Request`, then return or write the `Response`.
- Persist `state.workbookJson` in your durable store instead of module memory
  when the route needs to survive cold starts and multiple instances.

The important part is not the framework. The route accepts JSON records, writes
them into a workbook, recalculates formulas, persists the document, and returns
computed values that prove the write took effect.

## Next.js App Router Adapter

For a Next.js App Router project, keep the WorkPaper code in a shared module and
make the route files thin.

Create `app/api/workpaper/workpaper-route.js` from the shared route code above:

- keep the `@bilig/headless` imports
- keep `state`, `handleWorkPaperRequest()`, and every workbook helper
- omit `createServer()`, `toWebRequest()`, and the `if (import.meta.url === ...)`
  local Node adapter block

Then create `app/api/workpaper/summary/route.js`:

```js
import { handleWorkPaperRequest } from '../workpaper-route.js'

export const dynamic = 'force-dynamic'
export const runtime = 'nodejs'

export async function GET(request) {
  return handleWorkPaperRequest(request)
}
```

Create `app/api/workpaper/revenue/route.js`:

```js
import { handleWorkPaperRequest } from '../workpaper-route.js'

export const dynamic = 'force-dynamic'
export const runtime = 'nodejs'

export async function POST(request) {
  return handleWorkPaperRequest(request)
}
```

The route path stays the same as the standalone recipe:

```sh
curl -s http://localhost:3000/api/workpaper/summary
curl -s -X POST http://localhost:3000/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

Do not put the workbook helpers into both route files. The adapter should stay
small; the shared handler is what keeps formula evaluation, persistence, and
readback behavior identical between local Node, Next.js, and other web-standard
route surfaces.

## Cloudflare Worker Adapter

Cloudflare Workers use the same web-standard `Request` and `Response` shape as
the shared route handler. Keep the WorkPaper code in a module such as
`workpaper-route.js`, then make `src/index.js` small:

```js
import { handleWorkPaperRequest } from './workpaper-route.js'

export default {
  async fetch(request, env, ctx) {
    return handleWorkPaperRequest(request)
  },
}
```

The Worker entrypoint stays a pass-through. The shared handler still owns the
route paths:

```sh
curl -s https://example.workers.dev/api/workpaper/summary
curl -s -X POST https://example.workers.dev/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

Do not rely on module memory for durable workbook state on the edge. The
standalone recipe keeps `state.workbookJson` in memory so the example is small,
but a real Worker should load and save the serialized WorkPaper document through
KV, Durable Objects, D1, R2, or another storage boundary that matches the
workflow.

## Durable JSON Persistence Variant

For production routes, keep storage behind two small functions and pass them to
the shared handler. The storage provider can be a database row, object storage,
KV, a Durable Object, or a queue-backed persistence layer; the WorkPaper API only
needs serialized JSON.

```js
import { createWorkPaperRequestHandler, createInMemoryWorkbookStorage } from './workpaper-route.js'

const fallbackStorage = createInMemoryWorkbookStorage()

const storage = {
  async loadWorkbookJson() {
    const stored = await loadTextFromYourStore('workpaper/revenue.json')
    return stored ?? fallbackStorage.loadWorkbookJson()
  },
  async saveWorkbookJson(workbookJson) {
    await saveTextToYourStore('workpaper/revenue.json', workbookJson)
  },
}

export const handleWorkPaperRequest = createWorkPaperRequestHandler(storage)
```

The POST route does not need to know where the document lives. It loads the
current serialized WorkPaper JSON, writes the new records into a fresh workbook,
calculates formulas, saves the next serialized document, and returns the same
summary and verification checks as the in-memory demo:

```json
{
  "checks": {
    "totalRevenueChanged": true,
    "formulasPersisted": true,
    "serializedBytes": 1195
  }
}
```

Use module memory only for local demos and smoke tests. If the route can run in
multiple instances, cold starts, edge isolates, or background workers, the
serialized WorkPaper document should come from durable storage before reads and
be written back after accepted mutations.

## Hono Adapter

Hono exposes the raw Fetch `Request`, so the adapter does not need to translate
headers, bodies, or responses. Keep the WorkPaper logic in the shared handler and
make the framework route file small:

```js
import { Hono } from 'hono'
import { handleWorkPaperRequest } from './workpaper-route.js'

const app = new Hono()

app.get('/api/workpaper/summary', (c) => handleWorkPaperRequest(c.req.raw))
app.post('/api/workpaper/revenue', (c) => handleWorkPaperRequest(c.req.raw))

export default app
```

The same smoke calls apply when the Hono app is running locally:

```sh
curl -s http://localhost:3000/api/workpaper/summary
curl -s -X POST http://localhost:3000/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

If the Hono app is deployed to a worker or serverless runtime, pair this adapter
with the durable storage variant above instead of relying on module memory.

## SvelteKit Adapter

SvelteKit route handlers already receive a web-standard `Request`, so the
adapter should stay as small as the Hono and Next.js variants. Keep the
WorkPaper logic in a shared module such as `src/lib/server/workpaper-route.js`,
then call it from the route files.

Create `src/routes/api/workpaper/summary/+server.js`:

```js
import { handleWorkPaperRequest } from '$lib/server/workpaper-route.js'

export async function GET({ request }) {
  return handleWorkPaperRequest(request)
}
```

Create `src/routes/api/workpaper/revenue/+server.js`:

```js
import { handleWorkPaperRequest } from '$lib/server/workpaper-route.js'

export async function POST({ request }) {
  return handleWorkPaperRequest(request)
}
```

The shared module should keep the `@bilig/headless` imports, `state`,
`handleWorkPaperRequest()`, and the workbook helpers from the standalone route
recipe. Omit the local `createServer()` adapter and the `toWebRequest()`
translation helper because SvelteKit already hands your route a Fetch
`Request`.

The route path stays the same:

```sh
curl -s http://localhost:5173/api/workpaper/summary
curl -s -X POST http://localhost:5173/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

Keep the framework files thin. The workbook route should still load the
serialized WorkPaper document, apply the accepted edit, recalculate formulas,
persist the next JSON document, and return summary proof through the shared
handler instead of duplicating those details in each SvelteKit endpoint.

## Validation

For the standalone recipe:

```sh
npm start
curl -s -X POST http://localhost:8787/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

For a documentation patch in this repository:

```sh
pnpm docs:discovery:check
pnpm run ci
```
