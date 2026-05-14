---
title: Express, Fastify, and Hono adapters for a WorkPaper API
published: true
description: Copyable TypeScript adapters for serving @bilig/headless WorkPaper formulas from Express, Fastify, Hono, Next.js, Vercel Functions, and Fetch-style route handlers.
tags: typescript, node, spreadsheet, express
canonical_url: https://proompteng.github.io/bilig/node-framework-workpaper-adapters.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Express, Fastify, and Hono adapters for a WorkPaper API

Most Node framework code should not know how the workbook is built. Keep the
spreadsheet logic behind one web-standard `Request -> Response` handler, then
adapt the framework edge around it.

The runnable example is in
[`examples/serverless-workpaper-api`](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api).
It builds a small revenue workbook, writes records into a `Revenue` sheet,
reads summary formulas, saves the WorkPaper document JSON, and verifies that
the computed total survives the framework boundary.

## Run the adapter smoke

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/serverless-workpaper-api
npm install
npm run framework-adapters
```

Expected output:

```json
{
  "adapters": ["fetch", "hono", "express", "fastify"],
  "before": {
    "fetch": {
      "totalRevenue": 36900,
      "westCustomers": 20,
      "largestDeal": 24000
    },
    "hono": {
      "totalRevenue": 36900,
      "westCustomers": 20,
      "largestDeal": 24000
    }
  },
  "express": {
    "status": 200,
    "edit": {
      "records": 4,
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
  },
  "fastify": {
    "status": 200,
    "summary": {
      "totalRevenue": 48600,
      "westCustomers": 20,
      "largestDeal": 24000
    }
  },
  "verified": true
}
```

## Shared route shape

The example keeps the WorkPaper handler framework-neutral:

```ts
import { handleWorkPaperRequest } from './route.ts'

export const GET = handleWorkPaperRequest
export const POST = handleWorkPaperRequest
```

That shape works directly in Fetch-style runtimes and is easy to wrap in
frameworks that use their own request and response objects.

## Next.js Route Handler JSON

For App Router endpoints that accept JSON, keep the Next-specific file thin and
return web-standard `Response` objects. The runnable example proves the route
can parse JSON, update an input cell, read back a dependent formula, and reload
the persisted WorkPaper document:

```sh
cd examples/serverless-workpaper-api
npm install
npm run test
```

The copyable route shape is:

```ts
export const dynamic = 'force-dynamic'
export const runtime = 'nodejs'

export async function POST(request: Request) {
  const { customers } = await request.json()
  const result = updateRevenueInputCell(Number(customers))

  return Response.json({
    input: { cell: 'Inputs!B2', customers: result.customers },
    formulaReadback: { cell: 'Summary!B2', revenue: result.revenue },
    persistence: result.persistence,
  })
}
```

Use this for Next.js route handlers; use the generic adapters below when the
framework gives you Express, Fastify, Hono, or another request wrapper.

## Express

```ts
import express from 'express'
import { createExpressWorkPaperHandler } from './framework-adapters.ts'

const app = express()

app.use(express.json())
app.get('/api/workpaper/summary', createExpressWorkPaperHandler())
app.post('/api/workpaper/revenue', createExpressWorkPaperHandler())

app.listen(8787)
```

## Fastify

```ts
import Fastify from 'fastify'
import { createFastifyWorkPaperHandler } from './framework-adapters.ts'

const app = Fastify()
const workpaper = createFastifyWorkPaperHandler()

app.get('/api/workpaper/summary', workpaper)
app.post('/api/workpaper/revenue', workpaper)

await app.listen({ port: 8787 })
```

## Hono

```ts
import { Hono } from 'hono'
import { createHonoWorkPaperHandler } from './framework-adapters.ts'

const app = new Hono()
const workpaper = createHonoWorkPaperHandler()

app.get('/api/workpaper/summary', workpaper)
app.post('/api/workpaper/revenue', workpaper)

export default app
```

## What the wrapper must preserve

The adapter should do only four things:

- preserve the HTTP method and path
- pass JSON request bodies through as JSON
- copy response status and headers back to the framework response
- keep storage outside the framework handler when the workbook must survive
  cold starts or multiple instances

The workbook logic stays in
[`route.ts`](https://github.com/proompteng/bilig/blob/main/examples/serverless-workpaper-api/route.ts).
The adapters live in
[`framework-adapters.ts`](https://github.com/proompteng/bilig/blob/main/examples/serverless-workpaper-api/framework-adapters.ts).
Run `npm run smoke` and `npm run framework-adapters` before moving the handler
into your own service.
