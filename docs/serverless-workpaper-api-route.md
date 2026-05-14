# Serverless WorkPaper API Route

This recipe shows how to put `@bilig/headless` behind a small API route using
web-standard `Request` and `Response` objects. Use it when a serverless
function, route handler, queue worker, or coding-agent tool needs spreadsheet
formulas without keeping a browser grid open.

For a clone-and-run copy of this route, start with
[`examples/serverless-workpaper-api`](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api).

The example also includes a tiny Node adapter so you can run it locally before
moving the route into Vercel Functions, Cloudflare Workers, Supabase Edge
Functions, Fastify, Hono, or another HTTP surface.

If the route is going into an existing Node service, use the runnable
[`Express, Fastify, and Hono adapter guide`](./node-framework-workpaper-adapters.md)
with `npm run framework-adapters`.

For a copyable Next.js App Router boundary, the same example ships a runnable
`npm run next-route-handler` smoke that exports `GET()` and `POST()` around the
shared WorkPaper handler.

## Setup

```sh
mkdir bilig-serverless-workpaper
cd bilig-serverless-workpaper
npm init -y
npm pkg set type=module
npm pkg set scripts.start="tsx route.ts"
npm install @bilig/headless
npm install --save-dev tsx typescript @types/node
```

Create `route.ts`:

```ts
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

- In a Vercel Function, export `GET()` and `POST()` from files under `/api`.
- In a Next.js route handler, call it from `GET()` and `POST()`.
- In Cloudflare Workers, call it from `fetch(request)`.
- In Cloudflare Pages Functions, call it from `onRequestGet()` and
  `onRequestPost()`.
- In Supabase Edge Functions, call it from `Deno.serve()` and strip the function
  name prefix before handing the request to the shared route.
- In Remix resource routes, return it from a `loader()` for `GET` and an
  `action()` for `POST`.
- In Nitro, wrap it with H3's `fromWebHandler()` and export that from
  method-specific route files.
- In NestJS, adapt the Express request and response inside a thin controller.
- In Elysia, adapt the route context's web request and parsed body, then return
  the shared `Response`.
- In Firebase Functions, adapt the HTTPS function request into a web-standard
  `Request`, then write the shared `Response` through the Express-style
  response object.
- In Hono, Fastify, or Express, adapt the framework request into a web-standard
  `Request`, then return or write the `Response`.
- Persist `state.workbookJson` in your durable store instead of module memory
  when the route needs to survive cold starts and multiple instances.

The important part is not the framework. The route accepts JSON records, writes
them into a workbook, recalculates formulas, persists the document, and returns
computed values that prove the write took effect.

## Next.js App Router JSON Route Handler

Use this shape when a client, agent, or server component posts JSON to a
Next.js App Router endpoint and expects a deterministic formula readback. The
runnable example keeps this dependency-free and does not require a running Next
app for the smoke proof:

```sh
cd examples/serverless-workpaper-api
npm install
npm run next-route-handler
# or the focused example test
npm run test
```

The example exports route constants and methods matching an
`app/api/workpaper/model/route.ts` file:

```ts
export const dynamic = 'force-dynamic'
export const runtime = 'nodejs'

export async function POST(request: Request): Promise<Response> {
  const { customers } = await request.json()
  const workbook = loadWorkbook()
  const inputs = requireSheet(workbook, 'Inputs')

  workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, Number(customers))

  const after = readRevenueModel(workbook)
  const workbookJson = serializeWorkbook(workbook)
  state.workbookJson = workbookJson
  const persisted = readRevenueModel(loadWorkbook())

  return Response.json({
    input: { cell: 'Inputs!B2', customers: after.customers },
    formulaReadback: { cell: 'Summary!B2', revenue: after.revenue },
    persistence: {
      formulasPersisted: workbookJson.includes('=Inputs!B2*Inputs!B3'),
      inputPersisted: persisted.customers === after.customers,
      persistedRevenue: persisted.revenue,
      serializedBytes: Buffer.byteLength(workbookJson, 'utf8'),
    },
  })
}
```

The route accepts JSON such as:

```sh
curl -s -X POST http://localhost:3000/api/workpaper/model \
  -H 'content-type: application/json' \
  -d '{"customers":65}'
```

The smoke updates `Inputs!B2`, reads back the dependent `Summary!B2` revenue
formula, serializes the WorkPaper document JSON, reloads it, and checks that the
input value and formula result survived persistence:

```json
{
  "input": {
    "cell": "Inputs!B2",
    "customers": 65
  },
  "formulaReadback": {
    "cell": "Summary!B2",
    "revenue": 78000
  },
  "persistence": {
    "formulasPersisted": true,
    "inputPersisted": true,
    "persistedRevenue": 78000,
    "serializedBytes": 989
  }
}
```

Keep workbook construction, parsing, serialization, and durable storage in a
shared module. The Next.js `route.ts` file should stay framework-specific: parse
the web `Request`, call the WorkPaper helper, and return a deterministic JSON
`Response`.

## Next.js Server Action Adapter

Use a Server Action when a form or mutation should update a WorkPaper directly
from the server-side action instead of posting through an API route. The
repository example keeps this dependency-free and runnable:

```sh
cd examples/serverless-workpaper-api
npm install
npm run next-server-action
```

The example exports small action functions:

```ts
export async function readRevenueSummaryAction() {
  'use server'

  return requestJson('/api/workpaper/summary', parseSummaryResponse)
}

export async function updateRevenueRecordsAction(records) {
  'use server'

  return requestJson('/api/workpaper/revenue', parseEditResponse, {
    body: JSON.stringify({ records }),
    headers: { 'content-type': 'application/json' },
    method: 'POST',
  })
}
```

In a real Next.js app, keep `requestJson()` as a tiny wrapper around the shared
WorkPaper handler or route module. The smoke test prints `verified: true` only
after the action reads the original summary, writes the revenue records, reads
the recalculated summary, and confirms formulas survived the saved document.

## Next.js Server Action FormData Adapter

Use the FormData variant when a server-side form should update a WorkPaper
without a client-side JSON adapter:

```sh
cd examples/serverless-workpaper-api
npm install
npm run next-server-action-formdata
```

The example accepts repeated fields from a form submission:

```ts
export async function updateRevenueRecordsFromFormDataAction(formData) {
  'use server'

  const records = readRevenueRecordsFromFormData(formData)
  return requestJson('/api/workpaper/revenue', parseEditResponse, {
    body: JSON.stringify({ records }),
    headers: { 'content-type': 'application/json' },
    method: 'POST',
  })
}
```

Each `region`, `customers`, and `arpa` field becomes one typed revenue record.
The smoke output includes `action: "Next.js Server Action FormData"` and
`verified: true` only after formula readback and saved-document formula
persistence both pass.

## Vercel Function Adapter

Plain Vercel Functions can use the same web-standard `Request` and `Response`
objects as the shared WorkPaper handler. This is different from the Next.js App
Router layout above: the function files live under `/api` and do not need a
Next.js route segment.

Put the shared WorkPaper route in `api/workpaper-route.ts`:

- keep the `@bilig/headless` imports
- keep `state`, `handleWorkPaperRequest()`, and every workbook helper
- omit `createServer()`, `toWebRequest()`, and the local Node adapter block

Then create `api/workpaper/summary.ts`:

```ts
import { handleWorkPaperRequest } from '../workpaper-route.ts'

export function GET(request) {
  return handleWorkPaperRequest(request)
}
```

Create `api/workpaper/revenue.ts`:

```ts
import { handleWorkPaperRequest } from '../workpaper-route.ts'

export function POST(request) {
  return handleWorkPaperRequest(request)
}
```

Vercel also supports a single `fetch` web handler when one file should own every
method:

```ts
import { handleWorkPaperRequest } from './workpaper-route.ts'

export default {
  fetch(request) {
    return handleWorkPaperRequest(request)
  },
}
```

For older Vercel projects that still use a Node-style `(request, response)`
handler, keep the compatibility layer at the edge of the route and convert back
to the Vercel response object:

```ts
import { Readable } from 'node:stream'
import { handleWorkPaperRequest } from './workpaper-route.ts'

export default async function handler(request, response) {
  const webResponse = await handleWorkPaperRequest(toWebRequest(request))
  response.status(webResponse.status)

  for (const [name, value] of webResponse.headers) {
    response.setHeader(name, value)
  }

  response.send(Buffer.from(await webResponse.arrayBuffer()))
}

function toWebRequest(request) {
  const protocol = request.headers['x-forwarded-proto'] ?? 'https'
  const host = request.headers.host ?? 'localhost'
  const url = new URL(request.url ?? '/', `${protocol}://${host}`)
  const headers = new Headers()

  for (const [name, value] of Object.entries(request.headers)) {
    if (Array.isArray(value)) {
      for (const item of value) headers.append(name, item)
    } else if (value !== undefined) {
      headers.set(name, String(value))
    }
  }

  return new Request(url, {
    method: request.method,
    headers,
    body: request.method === 'GET' || request.method === 'HEAD' ? undefined : Readable.toWeb(request),
    duplex: 'half',
  })
}
```

Use the modern web handler form for new Vercel Functions. The Node-style bridge
exists only for older projects where Vercel has already handed you a
request/response pair. In either form, keep durable workbook storage behind the
`createWorkPaperRequestHandler(storage)` boundary before deploying.

## Cloudflare Worker Adapter

Cloudflare Workers use the same web-standard `Request` and `Response` shape as
the shared route handler. Keep the WorkPaper code in a module such as
`workpaper-route.ts`, then make `src/index.ts` small:

```ts
import { handleWorkPaperRequest } from './workpaper-route.ts'

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

## Cloudflare Pages Functions Adapter

Cloudflare Pages Functions use file-based routes under `/functions`. The
function receives a context object with a web-standard `request`, so the adapter
can stay as thin as the Worker adapter.

Put the shared WorkPaper route in `src/workpaper-route.ts`:

- keep the `@bilig/headless` imports
- keep `state`, `handleWorkPaperRequest()`, and every workbook helper
- omit `createServer()`, `toWebRequest()`, and the local Node adapter block

Then create `functions/api/workpaper/summary.ts`:

```ts
import { handleWorkPaperRequest } from '../../../src/workpaper-route.ts'

export function onRequestGet({ request }) {
  return handleWorkPaperRequest(request)
}
```

Create `functions/api/workpaper/revenue.ts`:

```ts
import { handleWorkPaperRequest } from '../../../src/workpaper-route.ts'

export function onRequestPost({ request }) {
  return handleWorkPaperRequest(request)
}
```

The route paths match the file names generated by Pages Functions:

```sh
curl -s https://example.pages.dev/api/workpaper/summary
curl -s -X POST https://example.pages.dev/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

This does not require a Cloudflare SDK dependency in the route file. If your
Pages project uses TypeScript, Cloudflare can type these files as
`PagesFunction` entries after you generate runtime types, but the runtime
boundary is still just `{ request }` in and a `Response` out.

Use the durable storage variant below for deployed Pages Functions. Module
memory is useful for a local smoke test, but it is not a reliable workbook store
across cold starts, edge isolates, or concurrent deployments.

## Durable JSON Persistence Variant

For production routes, keep storage behind two small functions and pass them to
the shared handler. The storage provider can be a database row, object storage,
KV, a Durable Object, or a queue-backed persistence layer; the WorkPaper API only
needs serialized JSON. For a blob-store version of this boundary, see the
[object storage adapter](persisting-formula-backed-workpaper-documents-in-node.md#object-storage-adapter).

```ts
import { createWorkPaperRequestHandler, createInMemoryWorkbookStorage } from './workpaper-route.ts'

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

```ts
import { Hono } from 'hono'
import { handleWorkPaperRequest } from './workpaper-route.ts'

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

```ts
import { handleWorkPaperRequest } from './workpaper-route.ts'

Deno.serve((request) => handleWorkPaperRequest(request))
```

For `deno serve` or Deno Deploy entrypoints that expect a default export, expose
the same function through `fetch`:

```ts
import { handleWorkPaperRequest } from './workpaper-route.ts'

export default {
  fetch(request) {
    return handleWorkPaperRequest(request)
  },
}
```

If the shared WorkPaper module runs directly in Deno instead of a bundled build,
import the published package with Deno's npm specifier:

```ts
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

## Supabase Edge Function Adapter

Supabase Edge Functions run on Deno and receive Fetch `Request` objects, so the
shared WorkPaper handler can stay the route boundary. Put the shared route in
`supabase/functions/workpaper/workpaper-route.ts`:

- keep the `@bilig/headless` imports, using Deno's npm specifier:
  `npm:@bilig/headless`
- keep `state`, `handleWorkPaperRequest()`, and every workbook helper
- omit `createServer()`, `toWebRequest()`, and the local Node adapter block

Then create `supabase/functions/workpaper/index.ts`:

```ts
import { handleWorkPaperRequest } from './workpaper-route.ts'

Deno.serve((request) => {
  return handleWorkPaperRequest(toWorkPaperRouteRequest(request))
})

function toWorkPaperRouteRequest(request: Request): Request {
  const url = new URL(request.url)
  url.pathname = stripFunctionNamePrefix(url.pathname, 'workpaper')

  return new Request(url.toString(), {
    method: request.method,
    headers: request.headers,
    body: request.method === 'GET' || request.method === 'HEAD' ? undefined : request.body,
  })
}

function stripFunctionNamePrefix(pathname: string, functionName: string): string {
  const prefix = `/${functionName}`
  return pathname === prefix || pathname.startsWith(`${prefix}/`) ? pathname.slice(prefix.length) || '/' : pathname
}
```

When the function is called directly, include the function name before the
WorkPaper route path:

```sh
curl -s https://PROJECT.supabase.co/functions/v1/workpaper/api/workpaper/summary \
  -H 'apikey: SUPABASE_PUBLISHABLE_KEY'
curl -s -X POST https://PROJECT.supabase.co/functions/v1/workpaper/api/workpaper/revenue \
  -H 'apikey: SUPABASE_PUBLISHABLE_KEY' \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

Keep Supabase auth, CORS, rate limits, and durable storage at the function edge
or behind the `createWorkPaperRequestHandler(storage)` boundary. Module memory
is fine for a local smoke test, but deployed Supabase functions should load and
save the serialized WorkPaper document through Postgres, Storage, or another
durable service before returning formula readback.

## SvelteKit Endpoint Adapter

SvelteKit `+server.ts` handlers receive a request event and return a web-standard
`Response`. Put the shared WorkPaper route in a server-only module, then keep
each endpoint file as a pass-through.

Create `src/lib/server/workpaper-route.ts` from the shared route code above:

- keep the `@bilig/headless` imports
- keep `state`, `handleWorkPaperRequest()`, and every workbook helper
- omit `createServer()`, `toWebRequest()`, and the local Node adapter block

Then create `src/routes/api/workpaper/summary/+server.ts`:

```ts
import { handleWorkPaperRequest } from '$lib/server/workpaper-route.ts'

export const prerender = false

export function GET({ request }) {
  return handleWorkPaperRequest(request)
}
```

Create `src/routes/api/workpaper/revenue/+server.ts`:

```ts
import { handleWorkPaperRequest } from '$lib/server/workpaper-route.ts'

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

## Remix Resource Route Adapter

Remix resource routes are route modules with no default component. They can
return any `Response`, which makes them a small adapter for the shared WorkPaper
handler.

Create `app/workpaper-route.server.ts` from the shared route code above:

- keep the `@bilig/headless` imports
- keep `state`, `handleWorkPaperRequest()`, and every workbook helper
- omit `createServer()`, `toWebRequest()`, and the local Node adapter block

Then create `app/routes/api.workpaper.summary.ts`:

```ts
import { handleWorkPaperRequest } from '../workpaper-route.server.ts'

export function loader({ request }) {
  return handleWorkPaperRequest(request)
}
```

Create `app/routes/api.workpaper.revenue.ts`:

```ts
import { handleWorkPaperRequest } from '../workpaper-route.server.ts'

export function action({ request }) {
  return handleWorkPaperRequest(request)
}
```

The route files intentionally export no React component. Remix will use the
`loader()` response for `GET /api/workpaper/summary` and the `action()` response
for `POST /api/workpaper/revenue`:

```sh
curl -s http://localhost:3000/api/workpaper/summary
curl -s -X POST http://localhost:3000/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

If the files are TypeScript, annotate the arguments with `LoaderFunctionArgs`
and `ActionFunctionArgs` from the Remix runtime package used by the app. The
runtime value that matters here is still the incoming Fetch `Request`; the
shared handler returns the `Response` directly.

Use the durable storage variant above for deployed Remix apps. Module memory is
fine for a local smoke test, but it is not a durable workbook store across
server restarts, serverless cold starts, or multiple app instances.

## Nitro Event Handler Adapter

Nitro maps files in `api/` or `routes/` to H3 route handlers. H3 can return a
web-standard `Response`, and its `fromWebHandler()` helper converts a
`Request => Response` function into an event handler, so the WorkPaper route can
stay framework-agnostic.

Create `workpaper-route.ts` from the shared route code above:

- keep the `@bilig/headless` imports
- keep `state`, `handleWorkPaperRequest()`, and every workbook helper
- omit `createServer()`, `toWebRequest()`, and the local Node adapter block

Then create `routes/api/workpaper/summary.get.ts`:

```ts
import { fromWebHandler } from 'h3'
import { handleWorkPaperRequest } from '../../../workpaper-route.ts'

export default fromWebHandler((request) => handleWorkPaperRequest(request))
```

Create `routes/api/workpaper/revenue.post.ts`:

```ts
import { fromWebHandler } from 'h3'
import { handleWorkPaperRequest } from '../../../workpaper-route.ts'

export default fromWebHandler((request) => handleWorkPaperRequest(request))
```

Nitro's method suffixes keep the route intent visible:

```sh
curl -s http://localhost:3000/api/workpaper/summary
curl -s -X POST http://localhost:3000/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

For Nuxt projects, put the same files under `server/routes/api/workpaper/`.
For a standalone Nitro project, `routes/api/...` keeps the `/api/...` path clear
even on platforms that reserve a top-level `api/` directory.

Use the durable storage variant above for deployed Nitro services. Module
memory is acceptable for the local smoke test only; real workbook state should
come from a database, object store, KV binding, or Nitro storage mount before
reads and be saved after accepted mutations.

## Bun.serve Adapter

`Bun.serve()` receives a web-standard `Request` and can return a `Response`, so
the adapter does not need to translate the route body, headers, or status code.
Keep the WorkPaper code in a shared module and let the Bun entrypoint pass each
request through.

```ts
import { handleWorkPaperRequest } from './workpaper-route.ts'

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

## Elysia Route Adapter

Elysia route handlers receive a context object that includes the web-standard
`request` and a parsed `body`. Keep the Elysia layer as a small bridge: rebuild
the WorkPaper request at the route edge, then return the shared `Response`
directly.

```ts
import { Elysia } from 'elysia'
import { handleWorkPaperRequest } from './workpaper-route.ts'

const app = new Elysia()
  .get('/api/workpaper/summary', ({ request }) => {
    return handleWorkPaperRequest(toWorkPaperRequest(request))
  })
  .post('/api/workpaper/revenue', ({ request, body }) => {
    return handleWorkPaperRequest(toWorkPaperRequest(request, body))
  })
  .listen(3000)

console.log(`WorkPaper API route listening on ${app.server?.url}`)

function toWorkPaperRequest(request, body) {
  return new Request(request.url, {
    method: request.method,
    headers: request.headers,
    body: request.method === 'GET' || request.method === 'HEAD' ? undefined : JSON.stringify(body ?? {}),
  })
}
```

The same route paths apply when the Elysia app is running locally:

```sh
curl -s http://localhost:3000/api/workpaper/summary
curl -s -X POST http://localhost:3000/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

This adapter assumes Elysia has parsed the JSON request body before the handler
runs. If a route accepts raw uploads or signed webhooks, preserve the raw request
body instead of serializing `body`.

Use the durable storage variant above for deployed Elysia services. Module
memory is fine for a local smoke test, but it is not a durable workbook store
across restarts, scale-out instances, or background workers.

## NestJS Controller Adapter

NestJS controllers route requests with decorators such as `@Get()` and
`@Post()`. Keep those controller methods thin: convert the platform request into
a web-standard `Request`, pass it to the shared WorkPaper handler, then copy the
returned `Response` back to Nest's response object.

Create `workpaper-route.ts` from the shared route code above:

- keep the `@bilig/headless` imports
- keep `state`, `handleWorkPaperRequest()`, and every workbook helper
- omit `createServer()`, `toWebRequest()`, and the local Node adapter block

Then create `workpaper.controller.ts`:

```ts
import { Controller, Get, Post, Req, Res } from '@nestjs/common'
import type { Request, Response } from 'express'
import { handleWorkPaperRequest } from './workpaper-route.ts'

@Controller('api/workpaper')
export class WorkPaperController {
  @Get('summary')
  summary(@Req() request: Request, @Res() response: Response) {
    return runWorkPaperRoute(request, response)
  }

  @Post('revenue')
  revenue(@Req() request: Request, @Res() response: Response) {
    return runWorkPaperRoute(request, response)
  }
}

async function runWorkPaperRoute(request: Request, response: Response) {
  const routeResponse = await handleWorkPaperRequest(toWebRequest(request))

  routeResponse.headers.forEach((value, name) => {
    response.setHeader(name, value)
  })
  response.status(routeResponse.status)
  response.send(Buffer.from(await routeResponse.arrayBuffer()))
}

function toWebRequest(request: Request) {
  const protocol = request.protocol ?? 'http'
  const host = request.get('host') ?? 'localhost:3000'
  const url = new URL(request.originalUrl ?? request.url, `${protocol}://${host}`)
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
    body: request.method === 'GET' || request.method === 'HEAD' ? undefined : JSON.stringify(request.body ?? {}),
  })
}
```

The controller exposes the same route paths:

```sh
curl -s http://localhost:3000/api/workpaper/summary
curl -s -X POST http://localhost:3000/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

This adapter assumes the default Nest Express setup has already parsed the JSON
body. If the route needs raw uploads or webhook signature verification, preserve
the raw body instead of serializing `request.body`.

Use the durable storage variant above for deployed NestJS services. Module
memory is fine for a local smoke test, but it is not a durable workbook store
across restarts, replicas, or background workers.

## Fastify Adapter

Fastify uses its own `request` and `reply` objects, so keep the adapter focused
on translating those objects to and from the web-standard boundary. The WorkPaper
handler should still own the route paths, formula writes, persistence, and
readback checks.

```ts
import { handleWorkPaperRequest } from './workpaper-route.ts'

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
    body: request.method === 'GET' || request.method === 'HEAD' ? undefined : JSON.stringify(request.body ?? {}),
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

```ts
import express from 'express'
import { handleWorkPaperRequest } from './workpaper-route.ts'

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
    body: req.method === 'GET' || req.method === 'HEAD' ? undefined : JSON.stringify(req.body ?? {}),
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

```ts
import { Readable } from 'node:stream'
import Koa from 'koa'
import { handleWorkPaperRequest } from './workpaper-route.ts'

const app = new Koa()

app.use(async (ctx, next) => {
  if (!isWorkPaperRoute(ctx)) {
    return next()
  }

  const response = await handleWorkPaperRequest(toWebRequest(ctx))
  await writeWebResponse(ctx, response)
})

function isWorkPaperRoute(ctx) {
  return (ctx.method === 'GET' && ctx.path === '/api/workpaper/summary') || (ctx.method === 'POST' && ctx.path === '/api/workpaper/revenue')
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
    body: ctx.method === 'GET' || ctx.method === 'HEAD' ? undefined : Readable.toWeb(ctx.req),
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

```ts
import Hapi from '@hapi/hapi'
import { handleWorkPaperRequest } from './workpaper-route.ts'

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
    body: request.method === 'GET' || request.method === 'HEAD' ? undefined : normalizePayload(request.payload),
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

```ts
import { handleWorkPaperRequest } from './workpaper-route.ts'

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
  const query = event.rawQueryString === undefined || event.rawQueryString === '' ? '' : `?${event.rawQueryString}`
  const body =
    event.body === undefined || event.body === null ? undefined : event.isBase64Encoded ? Buffer.from(event.body, 'base64') : event.body

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

```ts
import { app } from '@azure/functions'
import { handleWorkPaperRequest } from './workpaper-route.ts'

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
    body: request.method === 'GET' || request.method === 'HEAD' ? undefined : await request.arrayBuffer(),
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
module such as `workpaper-route.ts`, then keep the function file as a direct
pass-through.

Create `netlify/functions/workpaper.ts`:

```ts
import { handleWorkPaperRequest } from './workpaper-route.ts'

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

```ts
import { handleWorkPaperRequest } from './workpaper-route.ts'

export async function handler(event) {
  const response = await handleWorkPaperRequest(toWebRequest(event))
  return toNetlifyResult(response)
}

function toWebRequest(event) {
  const headers = new Headers(event.headers ?? {})
  const protocol = headers.get('x-forwarded-proto') ?? 'https'
  const host = headers.get('host') ?? 'localhost'
  const path = event.rawUrl ?? event.path ?? '/api/workpaper/summary'
  const body =
    event.body === undefined || event.body === null ? undefined : event.isBase64Encoded ? Buffer.from(event.body, 'base64') : event.body

  return new Request(new URL(path, `${protocol}://${host}`), {
    method: event.httpMethod ?? 'GET',
    headers,
    body: event.httpMethod === 'GET' || event.httpMethod === 'HEAD' ? undefined : body,
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

## Firebase Functions HTTPS Adapter

Firebase HTTPS functions receive an Express-style request and response object.
Keep Firebase at the edge of the route: convert the incoming function request to
a web-standard `Request`, pass it to the shared WorkPaper handler, then copy the
returned `Response` back to Firebase.

Create `functions/workpaper-route.ts` from the shared route code above:

- keep the `@bilig/headless` imports
- keep `state`, `handleWorkPaperRequest()`, and every workbook helper
- omit `createServer()`, `toWebRequest()`, and the local Node adapter block

Then create `functions/index.ts`:

```ts
import { onRequest } from 'firebase-functions/v2/https'
import { handleWorkPaperRequest } from './workpaper-route.ts'

export const workpaper = onRequest(async (request, response) => {
  const routeResponse = await handleWorkPaperRequest(toWebRequest(request))
  await writeFirebaseResponse(response, routeResponse)
})

function toWebRequest(request) {
  const protocol = request.get('x-forwarded-proto') ?? request.protocol ?? 'https'
  const host = request.get('host') ?? 'localhost'
  const url = new URL(request.originalUrl ?? request.url ?? '/', `${protocol}://${host}`)
  url.pathname = stripFunctionNamePrefix(url.pathname, 'workpaper')

  const headers = new Headers()
  for (const [name, value] of Object.entries(request.headers)) {
    if (Array.isArray(value)) {
      for (const item of value) headers.append(name, item)
    } else if (value !== undefined) {
      headers.set(name, String(value))
    }
  }

  return new Request(url, {
    method: request.method,
    headers,
    body: request.method === 'GET' || request.method === 'HEAD' ? undefined : (request.rawBody ?? JSON.stringify(request.body ?? {})),
  })
}

function stripFunctionNamePrefix(pathname, functionName) {
  const prefix = `/${functionName}`
  return pathname === prefix || pathname.startsWith(`${prefix}/`) ? pathname.slice(prefix.length) || '/' : pathname
}

async function writeFirebaseResponse(response, routeResponse) {
  routeResponse.headers.forEach((value, name) => {
    response.set(name, value)
  })
  response.status(routeResponse.status)
  response.send(Buffer.from(await routeResponse.arrayBuffer()))
}
```

When the function is called directly, include the exported function name before
the WorkPaper route path:

```sh
curl -s https://REGION-PROJECT.cloudfunctions.net/workpaper/api/workpaper/summary
curl -s -X POST https://REGION-PROJECT.cloudfunctions.net/workpaper/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
```

If Firebase Hosting rewrites `/api/workpaper/**` to the function, the same
adapter also works without the `workpaper` prefix because it only strips the
prefix when Firebase includes it in the request path.

Use the durable storage variant above for deployed Firebase functions. Module
memory is fine for a local emulator smoke test, but it is not a durable workbook
store across cold starts, scaled instances, or function redeploys.

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
