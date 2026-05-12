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

## Deno Deploy Adapter

Deno's HTTP server uses Fetch `Request` and `Response` objects, so the adapter is
the same shape as a Worker or Hono route. Keep the WorkPaper code in a shared
module and let the Deno entrypoint delegate to it:

```js
import { handleWorkPaperRequest } from './workpaper-route.js'

Deno.serve((request) => handleWorkPaperRequest(request))
```

For `deno serve` or Deno Deploy entrypoints that expect a default export, expose
the same function through `fetch`:

```js
import { handleWorkPaperRequest } from './workpaper-route.js'

export default {
  fetch(request) {
    return handleWorkPaperRequest(request)
  },
}
```

If the shared WorkPaper module runs directly in Deno instead of a bundled build,
import the published package with Deno's npm specifier:

```js
import { WorkPaper } from 'npm:@bilig/headless'
```

The same route paths apply when the Deno server is running locally:

```sh
curl -s http://localhost:8000/api/workpaper/summary
curl -s -X POST http://localhost:8000/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

Use the durable storage variant above for deployed Deno services. Module memory
is fine for the local smoke test, but it is not a durable workbook store across
deploys, isolates, or concurrent service instances.

## SvelteKit Endpoint Adapter

SvelteKit `+server.js` handlers receive a request event and return a web-standard
`Response`. Put the shared WorkPaper route in a server-only module, then keep
each endpoint file as a pass-through.

Create `src/lib/server/workpaper-route.js` from the shared route code above:

- keep the `@bilig/headless` imports
- keep `state`, `handleWorkPaperRequest()`, and every workbook helper
- omit `createServer()`, `toWebRequest()`, and the local Node adapter block

Then create `src/routes/api/workpaper/summary/+server.js`:

```js
import { handleWorkPaperRequest } from '$lib/server/workpaper-route.js'

export const prerender = false

export function GET({ request }) {
  return handleWorkPaperRequest(request)
}
```

Create `src/routes/api/workpaper/revenue/+server.js`:

```js
import { handleWorkPaperRequest } from '$lib/server/workpaper-route.js'

export const prerender = false

export function POST({ request }) {
  return handleWorkPaperRequest(request)
}
```

The same route paths apply when the SvelteKit app is running locally:

```sh
curl -s http://localhost:5173/api/workpaper/summary
curl -s -X POST http://localhost:5173/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

Keep the route files thin. The shared handler is what preserves identical
formula evaluation, persisted document shape, and readback checks between the
standalone Node server, SvelteKit, and other serverless surfaces.

## Bun.serve Adapter

`Bun.serve()` receives a web-standard `Request` and can return a `Response`, so
the adapter does not need to translate the route body, headers, or status code.
Keep the WorkPaper code in a shared module and let the Bun entrypoint pass each
request through.

```js
import { handleWorkPaperRequest } from './workpaper-route.js'

Bun.serve({
  port: 3000,
  fetch(request) {
    return handleWorkPaperRequest(request)
  },
})
```

The shared handler still owns the route paths:

```sh
curl -s http://localhost:3000/api/workpaper/summary
curl -s -X POST http://localhost:3000/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

Use the durable storage variant above for deployed Bun services. The local
example's module memory is intentionally small, but it should not be treated as
the workbook store for multiple processes, containers, or regions.

## Fastify Adapter

Fastify uses its own `request` and `reply` objects, so keep the adapter focused
on translating those objects to and from the web-standard boundary. The WorkPaper
handler should still own the route paths, formula writes, persistence, and
readback checks.

```js
import { handleWorkPaperRequest } from './workpaper-route.js'

fastify.get('/api/workpaper/summary', async (request, reply) => {
  return writeWebResponse(reply, await handleWorkPaperRequest(toWebRequest(request)))
})

fastify.post('/api/workpaper/revenue', async (request, reply) => {
  return writeWebResponse(reply, await handleWorkPaperRequest(toWebRequest(request)))
})

function toWebRequest(request) {
  const protocol = request.protocol ?? 'http'
  const host = request.hostname ?? request.headers.host ?? 'localhost:3000'
  const url = new URL(request.url, `${protocol}://${host}`)
  const headers = new Headers()

  for (const [name, value] of Object.entries(request.headers)) {
    if (Array.isArray(value)) {
      headers.set(name, value.join(', '))
    } else if (value !== undefined) {
      headers.set(name, String(value))
    }
  }

  return new Request(url, {
    method: request.method,
    headers,
    body:
      request.method === 'GET' || request.method === 'HEAD'
        ? undefined
        : JSON.stringify(request.body ?? {}),
  })
}

async function writeWebResponse(reply, response) {
  response.headers.forEach((value, name) => reply.header(name, value))
  reply.code(response.status)
  return reply.send(Buffer.from(await response.arrayBuffer()))
}
```

This adapter assumes Fastify parsed the JSON body before the handler runs. If a
route accepts raw uploads or non-JSON payloads, preserve the raw body in
`toWebRequest()` instead of serializing `request.body`.

The same smoke calls apply when the Fastify app is running locally:

```sh
curl -s http://localhost:3000/api/workpaper/summary
curl -s -X POST http://localhost:3000/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

Use the durable storage variant above when the app runs in multiple Node
processes, serverless instances, or any deployment where module memory is not a
safe source of truth.

## Express Adapter

Express apps need the same thin translation layer as Fastify. Parse JSON before
the route, convert the incoming request into the shared web-standard shape, then
copy the returned `Response` back to Express.

```js
import express from 'express'
import { handleWorkPaperRequest } from './workpaper-route.js'

const app = express()

app.use(express.json())
app.get('/api/workpaper/summary', runWorkPaperRoute)
app.post('/api/workpaper/revenue', runWorkPaperRoute)

async function runWorkPaperRoute(req, res, next) {
  try {
    const response = await handleWorkPaperRequest(toWebRequest(req))
    await writeWebResponse(res, response)
  } catch (error) {
    next(error)
  }
}

function toWebRequest(req) {
  const protocol = req.protocol ?? 'http'
  const host = req.get('host') ?? req.headers.host ?? 'localhost:3000'
  const url = new URL(req.originalUrl ?? req.url, `${protocol}://${host}`)
  const headers = new Headers()

  for (const [name, value] of Object.entries(req.headers)) {
    if (Array.isArray(value)) {
      headers.set(name, value.join(', '))
    } else if (value !== undefined) {
      headers.set(name, String(value))
    }
  }

  return new Request(url, {
    method: req.method,
    headers,
    body:
      req.method === 'GET' || req.method === 'HEAD'
        ? undefined
        : JSON.stringify(req.body ?? {}),
  })
}

async function writeWebResponse(res, response) {
  response.headers.forEach((value, name) => res.setHeader(name, value))
  res.status(response.status).send(Buffer.from(await response.arrayBuffer()))
}
```

Keep `runWorkPaperRoute()` pointed at both routes. The shared handler still
decides which workbook operation runs from the request method and pathname.

The same smoke calls apply when the Express app is running locally:

```sh
curl -s http://localhost:3000/api/workpaper/summary
curl -s -X POST http://localhost:3000/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

If the Express service runs behind a reverse proxy, configure Express proxy
trust as you normally would before relying on `req.protocol` for absolute URLs.
For non-JSON payloads, preserve the raw request body in `toWebRequest()` instead
of serializing `req.body`.

## Koa Adapter

Koa middleware receives a `ctx` object, but the shared WorkPaper route still only
needs a web-standard `Request`. Mount the adapter before any middleware that
consumes the request body, then let Koa own the final response fields.

```js
import { Readable } from 'node:stream'
import Koa from 'koa'
import { handleWorkPaperRequest } from './workpaper-route.js'

const app = new Koa()

app.use(async (ctx, next) => {
  if (!isWorkPaperRoute(ctx)) {
    return next()
  }

  const response = await handleWorkPaperRequest(toWebRequest(ctx))
  await writeWebResponse(ctx, response)
})

function isWorkPaperRoute(ctx) {
  return (
    (ctx.method === 'GET' && ctx.path === '/api/workpaper/summary') ||
    (ctx.method === 'POST' && ctx.path === '/api/workpaper/revenue')
  )
}

function toWebRequest(ctx) {
  const headers = new Headers()

  for (const [name, value] of Object.entries(ctx.headers)) {
    if (Array.isArray(value)) {
      headers.set(name, value.join(', '))
    } else if (value !== undefined) {
      headers.set(name, String(value))
    }
  }

  return new Request(new URL(ctx.originalUrl ?? ctx.url, ctx.origin), {
    method: ctx.method,
    headers,
    body:
      ctx.method === 'GET' || ctx.method === 'HEAD'
        ? undefined
        : Readable.toWeb(ctx.req),
    duplex: 'half',
  })
}

async function writeWebResponse(ctx, response) {
  response.headers.forEach((value, name) => ctx.set(name, value))
  ctx.status = response.status
  ctx.body = Buffer.from(await response.arrayBuffer())
}

app.listen(3000)
```

If another Koa middleware has already parsed the JSON body, replace
`Readable.toWeb(ctx.req)` with `JSON.stringify(ctx.request.body ?? {})` and keep
the same method, URL, and header mapping.

The same smoke calls apply when the Koa app is running locally:

```sh
curl -s http://localhost:3000/api/workpaper/summary
curl -s -X POST http://localhost:3000/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

Use the durable storage variant above when the Koa service runs in more than one
process, container, or serverless instance.

## Hapi Adapter

Hapi route handlers receive a framework request plus the response toolkit `h`.
Keep the WorkPaper handler framework-agnostic by translating the Hapi request
into a web-standard `Request`, then translate the returned `Response` through
`h.response()`.

```js
import Hapi from '@hapi/hapi'
import { handleWorkPaperRequest } from './workpaper-route.js'

const server = Hapi.server({ port: 3000 })

server.route([
  {
    method: 'GET',
    path: '/api/workpaper/summary',
    handler: runWorkPaperRoute,
  },
  {
    method: 'POST',
    path: '/api/workpaper/revenue',
    handler: runWorkPaperRoute,
  },
])

async function runWorkPaperRoute(request, h) {
  const response = await handleWorkPaperRequest(toWebRequest(request))
  return writeWebResponse(h, response)
}

function toWebRequest(request) {
  const protocol = request.headers['x-forwarded-proto'] ?? 'http'
  const host = request.headers.host ?? request.info.host ?? 'localhost:3000'
  const path = request.raw.req.url ?? request.url?.href ?? request.path
  const headers = new Headers()

  for (const [name, value] of Object.entries(request.headers)) {
    if (Array.isArray(value)) {
      headers.set(name, value.join(', '))
    } else if (value !== undefined) {
      headers.set(name, String(value))
    }
  }

  return new Request(new URL(path, `${protocol}://${host}`), {
    method: request.method,
    headers,
    body:
      request.method === 'GET' || request.method === 'HEAD'
        ? undefined
        : normalizePayload(request.payload),
  })
}

function normalizePayload(payload) {
  if (payload === undefined || payload === null) {
    return undefined
  }

  if (typeof payload === 'string' || payload instanceof Uint8Array) {
    return payload
  }

  return JSON.stringify(payload)
}

async function writeWebResponse(h, response) {
  const reply = h.response(Buffer.from(await response.arrayBuffer())).code(response.status)
  response.headers.forEach((value, name) => reply.header(name, value))
  return reply
}

await server.start()
```

This adapter assumes Hapi parses JSON payloads before the handler runs. If a
route needs the exact raw request stream, configure that route's Hapi payload
options for raw input and pass the raw payload through `normalizePayload()`.

The same smoke calls apply when the Hapi server is running locally:

```sh
curl -s http://localhost:3000/api/workpaper/summary
curl -s -X POST http://localhost:3000/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

Use the durable storage variant above when the Hapi service runs in more than
one process, container, or serverless instance.

## AWS Lambda Function URL Adapter

Lambda Function URLs use the API Gateway payload format version 2.0. Keep the
Lambda handler small: convert the event into a web-standard `Request`, run the
shared WorkPaper handler, then return a Lambda proxy response.

```js
import { handleWorkPaperRequest } from './workpaper-route.js'

export async function handler(event) {
  const response = await handleWorkPaperRequest(toWebRequest(event))
  return toLambdaResult(response)
}

function toWebRequest(event) {
  const headers = new Headers()

  for (const [name, value] of Object.entries(event.headers ?? {})) {
    if (value !== undefined) {
      headers.set(name, String(value))
    }
  }

  const method = event.requestContext?.http?.method ?? event.httpMethod ?? 'GET'
  const protocol = headers.get('x-forwarded-proto') ?? 'https'
  const host = event.requestContext?.domainName ?? headers.get('host') ?? 'localhost'
  const path = event.rawPath ?? event.path ?? '/'
  const query =
    event.rawQueryString === undefined || event.rawQueryString === ''
      ? ''
      : `?${event.rawQueryString}`
  const body =
    event.body === undefined || event.body === null
      ? undefined
      : event.isBase64Encoded
        ? Buffer.from(event.body, 'base64')
        : event.body

  return new Request(new URL(`${path}${query}`, `${protocol}://${host}`), {
    method,
    headers,
    body: method === 'GET' || method === 'HEAD' ? undefined : body,
  })
}

async function toLambdaResult(response) {
  return {
    statusCode: response.status,
    headers: Object.fromEntries(response.headers),
    body: Buffer.from(await response.arrayBuffer()).toString('utf8'),
    isBase64Encoded: false,
  }
}
```

This example returns UTF-8 JSON, so `isBase64Encoded` stays `false`. If you add a
binary export route later, base64-encode that route's response body and set
`isBase64Encoded: true`.

The same route paths apply behind the function URL:

```sh
curl -s https://example.lambda-url.us-east-1.on.aws/api/workpaper/summary
curl -s -X POST https://example.lambda-url.us-east-1.on.aws/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

Pair the Lambda handler with the durable storage variant above. Module memory is
not a reliable persistence boundary across Lambda cold starts, concurrent
instances, or redeploys.

## Azure Functions HTTP Adapter

Azure Functions' Node.js v4 model registers HTTP functions with `app.http()`.
The function handler receives an `HttpRequest` with a full URL, method, headers,
and body readers, then returns an HTTP response object. Keep the adapter thin and
let the shared WorkPaper handler own the route behavior.

```js
import { app } from '@azure/functions'
import { handleWorkPaperRequest } from './workpaper-route.js'

app.http('workpaperSummary', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'workpaper/summary',
  handler: runWorkPaperRoute,
})

app.http('workpaperRevenue', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'workpaper/revenue',
  handler: runWorkPaperRoute,
})

async function runWorkPaperRoute(request) {
  const response = await handleWorkPaperRequest(await toWebRequest(request))
  return toAzureResponse(response)
}

async function toWebRequest(request) {
  return new Request(request.url, {
    method: request.method,
    headers: request.headers,
    body:
      request.method === 'GET' || request.method === 'HEAD'
        ? undefined
        : await request.arrayBuffer(),
  })
}

async function toAzureResponse(response) {
  return {
    status: response.status,
    headers: Object.fromEntries(response.headers),
    body: Buffer.from(await response.arrayBuffer()),
  }
}
```

The local Azure Functions route prefix usually includes `/api`:

```sh
curl -s http://localhost:7071/api/workpaper/summary
curl -s -X POST http://localhost:7071/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

Use the durable storage variant above for deployed function apps. Module memory
is not a reliable workbook store across cold starts, scale-out instances, or
regional deployments.

## Netlify Functions Adapter

Netlify's current Functions runtime passes a web-standard `Request` into the
default export and expects a `Response` back. Put the shared WorkPaper route in a
module such as `workpaper-route.js`, then keep the function file as a direct
pass-through.

Create `netlify/functions/workpaper.mjs`:

```js
import { handleWorkPaperRequest } from './workpaper-route.js'

export default async function workpaper(request) {
  return handleWorkPaperRequest(request)
}

export const config = {
  path: ['/api/workpaper/summary', '/api/workpaper/revenue'],
}
```

If your Netlify project is using the Lambda-compatible named `handler` export,
adapt the event into a `Request` and convert the `Response` back into a function
result:

```js
import { handleWorkPaperRequest } from './workpaper-route.js'

export async function handler(event) {
  const response = await handleWorkPaperRequest(toWebRequest(event))
  return toNetlifyResult(response)
}

function toWebRequest(event) {
  const headers = new Headers(event.headers ?? {})
  const protocol = headers.get('x-forwarded-proto') ?? 'https'
  const host = headers.get('host') ?? 'localhost'
  const path = event.rawUrl ?? event.path ?? '/api/workpaper/summary'
  const body = event.body === undefined || event.body === null
    ? undefined
    : event.isBase64Encoded
      ? Buffer.from(event.body, 'base64')
      : event.body

  return new Request(new URL(path, `${protocol}://${host}`), {
    method: event.httpMethod ?? 'GET',
    headers,
    body:
      event.httpMethod === 'GET' || event.httpMethod === 'HEAD'
        ? undefined
        : body,
  })
}

async function toNetlifyResult(response) {
  return {
    statusCode: response.status,
    headers: Object.fromEntries(response.headers),
    body: Buffer.from(await response.arrayBuffer()).toString('utf8'),
    isBase64Encoded: false,
  }
}
```

The same route paths apply when `netlify dev` is running locally:

```sh
curl -s http://localhost:8888/api/workpaper/summary
curl -s -X POST http://localhost:8888/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

Use the durable storage variant above for deployed Netlify Functions. Function
instances can be short-lived or scaled out, so module memory is only appropriate
for a local smoke test.

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
