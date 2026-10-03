---
title: Express, Fastify, Hono, Oak, Hapi, AdonisJS, and tRPC adapters for a WorkPaper API
published: true
description: Copyable TypeScript adapters for serving @bilig/workpaper WorkPaper formulas from Express, Fastify, Hono, Oak, Hapi, AdonisJS, tRPC, Next.js, Vercel Functions, and Fetch-style route handlers.
tags: typescript, node, spreadsheet, express
canonical_url: https://proompteng.github.io/bilig/node-framework-workpaper-adapters.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Express, Fastify, Hono, Oak, Hapi, AdonisJS, and tRPC adapters for a WorkPaper API

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
cd bilig
pnpm --dir examples/serverless-workpaper-api install --ignore-workspace
pnpm --dir examples/serverless-workpaper-api run framework-adapters
```

Expected output:

```json
{
  "adapters": ["fetch", "hono", "oak", "adonis", "hapi", "express", "fastify"],
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
    },
    "oak": {
      "totalRevenue": 36900,
      "westCustomers": 20,
      "largestDeal": 24000
    },
    "adonis": {
      "totalRevenue": 36900,
      "westCustomers": 20,
      "largestDeal": 24000
    },
    "hapi": {
      "totalRevenue": 36900,
      "westCustomers": 20,
      "largestDeal": 24000
    }
  },
  "oak": {
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
        "serializedBytes": 1194
      }
    }
  },
  "adonis": {
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
        "serializedBytes": 1194
      }
    }
  },
  "hapi": {
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
        "serializedBytes": 1194
      }
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
        "serializedBytes": 1194
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
pnpm --dir examples/serverless-workpaper-api install --ignore-workspace
pnpm --dir examples/serverless-workpaper-api run test
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
framework gives you Express, Fastify, Hono, Oak, AdonisJS, Hapi, or another request
wrapper.

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

Smoke the Express boundary with the same revenue update used by the runnable
adapter proof:

```sh
curl -X POST http://localhost:8787/api/workpaper/revenue \
  -H 'content-type: application/json' \
  --data '{
    "records": [
      { "region": "West", "customers": 20, "arpa": 1200 },
      { "region": "East", "customers": 30, "arpa": 250 },
      { "region": "Central", "customers": 18, "arpa": 300 },
      { "region": "Enterprise", "customers": 12, "arpa": 475 }
    ]
  }'
```

The response is the shared WorkPaper `Response` copied back through Express
after the formulas recalculate:

```json
{
  "records": 4,
  "after": {
    "totalRevenue": 48600,
    "westCustomers": 20,
    "largestDeal": 24000
  },
  "checks": {
    "totalRevenueChanged": true,
    "formulasPersisted": true,
    "serializedBytes": 1194
  }
}
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

## Oak

```ts
import { Application, Router } from '@oak/oak'
import { createOakWorkPaperRoutes } from './framework-adapters.ts'

const app = new Application()
const router = new Router()
const [summaryRoute, revenueRoute] = createOakWorkPaperRoutes()

router.get(summaryRoute.path, summaryRoute.handler)
router.post(revenueRoute.path, revenueRoute.handler)

app.use(router.routes())
app.use(router.allowedMethods())

await app.listen({ port: 8787 })
```

## AdonisJS

```ts
import router from '@adonisjs/core/services/router'
import { createAdonisWorkPaperRoutes } from './framework-adapters.ts'

const [summaryRoute, revenueRoute] = createAdonisWorkPaperRoutes()

router.get(summaryRoute.path, summaryRoute.handler)
router.post(revenueRoute.path, revenueRoute.handler)
```

## Hapi

```ts
import Hapi from '@hapi/hapi'
import { createHapiWorkPaperRoutes } from './framework-adapters.ts'

const server = Hapi.server({ port: 8787 })

server.route([...createHapiWorkPaperRoutes()])

await server.start()
```

## tRPC Procedure Smoke

tRPC procedures should keep the same boundary as the framework adapters: parse
typed service input, call the shared WorkPaper route or helper, then return
computed readback fields. Do not copy workbook-building logic into each
procedure.

This compact shape maps a nested `workpaper` router onto the same
[`examples/serverless-workpaper-api`](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api)
handler used by the route and framework examples:

```ts
import { initTRPC } from '@trpc/server'
import { z } from 'zod'
import { handleWorkPaperRequest } from './route.ts'

const t = initTRPC.create()

const revenueRecordInput = z.object({
  region: z.string().min(1),
  customers: z.number().nonnegative(),
  arpa: z.number().nonnegative(),
})

async function callWorkPaper(path: 'summary' | 'revenue', init?: RequestInit) {
  const response = await handleWorkPaperRequest(new Request(`https://workpaper.local/api/workpaper/${path}`, init))

  if (!response.ok) {
    throw new Error(await response.text())
  }

  return response.json()
}

export const appRouter = t.router({
  workpaper: t.router({
    summary: t.procedure.query(() => callWorkPaper('summary')),
    updateRevenue: t.procedure.input(z.object({ records: z.array(revenueRecordInput).min(1) })).mutation(({ input }) =>
      callWorkPaper('revenue', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(input),
      }),
    ),
  }),
})
```

Calling `workpaper.updateRevenue` with the same records used by the runnable
adapter smoke returns deterministic formula readback:

```json
{
  "procedure": "workpaper.updateRevenue",
  "records": 4,
  "after": {
    "totalRevenue": 48600,
    "westCustomers": 20,
    "largestDeal": 24000
  },
  "checks": {
    "totalRevenueChanged": true,
    "formulasPersisted": true,
    "serializedBytes": 1194
  },
  "verified": true
}
```

Use `workpaper.summary` for read-only callers that only need `totalRevenue`,
`westCustomers`, and `largestDeal`. Use `workpaper.updateRevenue` for mutations
that accept a JSON `{ "records": [...] }` input and must prove formulas
recalculated before returning.

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
