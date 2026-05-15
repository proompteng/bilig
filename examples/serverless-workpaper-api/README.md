# Serverless WorkPaper API Example

This is a runnable `@bilig/headless` example for HTTP routes, serverless
functions, queue workers, and agent tools that need spreadsheet formulas behind
a JSON boundary.

The route accepts revenue records, writes them into a WorkPaper workbook,
calculates summary formulas, persists the workbook as JSON, and returns
computed values that prove the write took effect.

Run it outside the monorepo with the published package:

```sh
npm install
npm run test
npm run next-route-handler
npm run next-server-action
npm run next-server-action-formdata
npm run framework-adapters
npm run persistence-adapters
```

Expected smoke output:

```json
{
  "before": {
    "totalRevenue": 36900,
    "westCustomers": 20,
    "largestDeal": 24000
  },
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
  },
  "after": {
    "totalRevenue": 48600,
    "westCustomers": 20,
    "largestDeal": 24000
  },
  "verified": true
}
```

To run it as a local HTTP server:

```sh
npm start
```

Then call it from another terminal:

```sh
curl -s http://localhost:8787/api/workpaper/summary
curl -s -X POST http://localhost:8787/api/workpaper/revenue \
  -H 'content-type: application/json' \
  -d '{"records":[{"region":"West","customers":20,"arpa":1200},{"region":"East","customers":30,"arpa":250},{"region":"Central","customers":18,"arpa":300},{"region":"North","customers":65,"arpa":180}]}'
curl -s http://localhost:8787/api/workpaper/summary
```

The exported `handleWorkPaperRequest(request)` function uses web-standard
`Request` and `Response` objects, so it can be adapted to framework route
handlers without changing the workbook logic.

The lower-level `createWorkPaperRequestHandler(storage)` helper accepts
`loadWorkbookJson()` and `saveWorkbookJson(workbookJson)` functions. Use that
shape when the serialized WorkPaper document should live in a database, object
store, KV namespace, Durable Object, or another durable service instead of
module memory.

## Vercel Function Smoke

Run the Vercel Function-shaped smoke when you want a copyable `/api` function
boundary for a deployed Node runtime:

```sh
npm run vercel-function
```

The script exports web-standard `GET(request)` and `POST(request)` handlers,
plus a `default.fetch(request)` entrypoint for single-file Vercel Functions.
Each entrypoint delegates to the shared WorkPaper request handler, writes a
revenue update, reads the saved workbook again, and prints `verified: true`
only after the formula totals recalculate and the persisted document still
contains formulas.

Expected output:

```json
{
  "route": "Vercel Function",
  "entrypoints": ["GET", "POST", "default.fetch"],
  "before": {
    "largestDeal": 24000,
    "totalRevenue": 36900,
    "westCustomers": 20
  },
  "edit": {
    "records": 4,
    "after": {
      "largestDeal": 24000,
      "totalRevenue": 48600,
      "westCustomers": 20
    },
    "checks": {
      "formulasPersisted": true,
      "serializedBytes": 1195,
      "totalRevenueChanged": true
    }
  },
  "after": {
    "largestDeal": 24000,
    "totalRevenue": 48600,
    "westCustomers": 20
  },
  "verified": true
}
```

Run the focused proof used by this example with:

```sh
npm run test
```

## Next.js App Router Smoke

Run the Next.js-shaped Route Handler smoke when you want a copyable App Router
boundary that accepts JSON and updates one WorkPaper input cell without adding a
full Next app to this example:

```sh
npm run next-route-handler
# or run the focused example test
npm run test
```

The script exports the same route constants a Next.js `route.ts` file expects,
including `runtime = 'nodejs'` and `dynamic = 'force-dynamic'`. Its `POST()`
handler parses `{ "customers": 65 }`, writes that value into `Inputs!B2`, reads
back the dependent `Summary!B2` revenue formula, persists the WorkPaper document
JSON, and reloads it to prove both the input and formula survived.

Expected output:

```json
{
  "route": "Next.js Route Handler JSON",
  "runtime": "nodejs",
  "dynamic": "force-dynamic",
  "before": {
    "arpa": 1200,
    "customers": 20,
    "revenue": 24000
  },
  "edit": {
    "input": {
      "cell": "Inputs!B2",
      "customers": 65
    },
    "before": {
      "arpa": 1200,
      "customers": 20,
      "revenue": 24000
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
  },
  "after": {
    "arpa": 1200,
    "customers": 65,
    "revenue": 78000
  },
  "verified": true
}
```

`serializedBytes` can change as the persisted document schema evolves. Treat it
as a positive persistence signal, not a golden value.

## Next.js Server Action Smoke

Run the Server Action-shaped smoke when a form or mutation should call
WorkPaper logic directly instead of going through a route handler:

```sh
npm run next-server-action
```

The script exports dependency-free `readRevenueSummaryAction()` and
`updateRevenueRecordsAction()` functions with the same `'use server'` boundary a
Next.js app would use. The actions call the shared WorkPaper request handler,
send a revenue update, read the summary again, and print `verified: true` only
after formulas recalculate and the saved document still contains formulas.

Expected output:

```json
{
  "action": "Next.js Server Action",
  "before": {
    "largestDeal": 24000,
    "totalRevenue": 36900,
    "westCustomers": 20
  },
  "edit": {
    "before": {
      "largestDeal": 24000,
      "totalRevenue": 36900,
      "westCustomers": 20
    },
    "after": {
      "largestDeal": 24000,
      "totalRevenue": 48600,
      "westCustomers": 20
    },
    "checks": {
      "formulasPersisted": true,
      "serializedBytes": 1195,
      "totalRevenueChanged": true
    },
    "records": 4
  },
  "after": {
    "largestDeal": 24000,
    "totalRevenue": 48600,
    "westCustomers": 20
  },
  "verified": true
}
```

## Next.js Server Action FormData Smoke

Run the FormData-shaped Server Action smoke when a form submission should feed
typed records into WorkPaper formulas:

```sh
npm run next-server-action-formdata
```

The script exports `updateRevenueRecordsFromFormDataAction(formData)`, parses
repeated `region`, `customers`, and `arpa` fields into revenue records, calls
the shared WorkPaper request handler, and prints `verified: true` only after
the formulas recalculate and the saved document still contains formulas.

Expected output:

```json
{
  "action": "Next.js Server Action FormData",
  "input": {
    "fields": ["region", "customers", "arpa"],
    "records": 4
  },
  "before": {
    "largestDeal": 24000,
    "totalRevenue": 36900,
    "westCustomers": 20
  },
  "edit": {
    "before": {
      "largestDeal": 24000,
      "totalRevenue": 36900,
      "westCustomers": 20
    },
    "after": {
      "largestDeal": 24000,
      "totalRevenue": 48600,
      "westCustomers": 20
    },
    "checks": {
      "formulasPersisted": true,
      "serializedBytes": 1195,
      "totalRevenueChanged": true
    },
    "records": 4
  },
  "after": {
    "largestDeal": 24000,
    "totalRevenue": 48600,
    "westCustomers": 20
  },
  "verified": true
}
```

## Framework Adapters

Run the adapter smoke when you want copyable TypeScript wrappers for common
Node service frameworks:

```sh
npm run framework-adapters
```

The script exercises the same WorkPaper route through Fetch-style handlers,
Hono-style `context.req.raw`, Oak-style `context.request.source`, AdonisJS-style `HttpContext`, Hapi-style `request` plus `h.response()`,
Express-style `(req, res, next)`, and Fastify-style `(request, reply)` adapters.
It reads the summary and writes the revenue update through the Oak, AdonisJS, and Hapi
routes, keeps the existing Express and Fastify wrappers covered, and prints
`verified: true` only after the calculated summary matches.

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
        "serializedBytes": 1195
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
        "serializedBytes": 1195
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
        "serializedBytes": 1195
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

## Persistence Adapters

Run the persistence smoke when the route needs durable state instead of module
memory:

```sh
npm run persistence-adapters
```

The script exercises the same WorkPaper request handler through four typed
storage adapters:

- Postgres JSONB, using a `query(sql, values)` client shape compatible with
  `pg`-style clients.
- SQLite, using parameterized SQL and a workbook id.
- Redis or another string KV store, using `get(key)` and `set(key, value)`.
- Object storage such as S3, R2, GCS, or Azure Blob, using small text load and
  save functions.

Each adapter starts from an empty store, handles a summary read, accepts the
revenue write, creates a fresh handler, then reads the restored workbook from
the saved document. `verified: true` means formulas survived persistence and
the cold read returned the recalculated total. These adapters persist the
serialized WorkPaper document state, not an XLSX file cache.

Expected output shape:

```json
{
  "adapters": ["postgres-jsonb", "sqlite", "redis", "object-storage"],
  "postgres": {
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
    "verified": true
  },
  "sqlite": {
    "after": {
      "totalRevenue": 48600,
      "westCustomers": 20,
      "largestDeal": 24000
    },
    "verified": true
  },
  "redis": {
    "after": {
      "totalRevenue": 48600,
      "westCustomers": 20,
      "largestDeal": 24000
    },
    "verified": true
  },
  "objectStorage": {
    "after": {
      "totalRevenue": 48600,
      "westCustomers": 20,
      "largestDeal": 24000
    },
    "verified": true
  },
  "verified": true
}
```

## Moving Into A Real Route

- In a Next.js app route, call `handleWorkPaperRequest()` from `GET()` and
  `POST()`.
- In a Next.js Server Action, call the shared WorkPaper handler from small
  `'use server'` action functions like `readRevenueSummaryAction()` and
  `updateRevenueRecordsAction()`.
- In a Vercel Function, export small web-standard `GET()` and `POST()` handlers
  from files under `/api`; the runnable `vercel-function.ts` smoke shows the deployed Node function shape.
- In a Cloudflare Worker, call it from `fetch(request)`.
- In Cloudflare Pages Functions, call it from small `onRequestGet()` and
  `onRequestPost()` files under `/functions`.
- In a Supabase Edge Function, call it from `Deno.serve()` in
  `supabase/functions/workpaper/index.ts`; if the public URL includes
  `/functions/v1/workpaper/api/workpaper/...`, strip the leading `workpaper`
  function segment before passing the request to the shared handler.
- In Hono, pass `c.req.raw` directly to the shared handler.
- In Oak, adapt `ctx.request.source` into the shared handler, then write the
  returned status, headers, and body back through `ctx.response`. The
  `framework-adapters.ts` smoke includes both `/api/workpaper/summary` and
  `/api/workpaper/revenue` Oak routes.
- In Deno, return `handleWorkPaperRequest(request)` from `Deno.serve()` or a
  `fetch` default export.
- In SvelteKit, return `handleWorkPaperRequest(request)` from thin `+server.ts`
  `GET` and `POST` handlers.
- In Remix, return `handleWorkPaperRequest(request)` from resource-route
  `loader` and `action` functions.
- In Nitro, export an H3 `fromWebHandler()` wrapper from method-specific route
  files.
- In Bun, return `handleWorkPaperRequest(request)` from the `Bun.serve()`
  `fetch` handler.
- In Elysia, adapt the route context's web request and parsed body, then return
  the shared `Response`.
- In NestJS, adapt the Express request into a web-standard `Request`, then copy
  the returned `Response` back through `@Res()`.
- In Fastify, adapt the framework request into a web-standard `Request`, then
  write the returned `Response` through `reply`.
- In Express, adapt the framework request into a web-standard `Request`, then
  write the returned `Response` through `res`.
- In Koa, adapt `ctx` into a web-standard `Request`, then write status, headers,
  and body back to the Koa context.
- In AdonisJS, adapt the `HttpContext` request into a web-standard `Request`,
  then write status, headers, and body back through `ctx.response`. The
  `framework-adapters.ts` smoke includes both `/api/workpaper/summary` and
  `/api/workpaper/revenue` AdonisJS routes.
- In Hapi, adapt the framework request into a web-standard `Request`, then
  return the shared response through `h.response()`. The
  `framework-adapters.ts` smoke includes both `/api/workpaper/summary` and
  `/api/workpaper/revenue` Hapi routes.
- In an AWS Lambda Function URL handler, adapt the event into a web-standard
  `Request`, then return a Lambda proxy response.
- In Azure Functions, adapt the HTTP trigger request into a web-standard
  `Request`, then return an HTTP response object.
- In Netlify Functions, return `handleWorkPaperRequest(request)` from the
  default export, or use the Lambda-compatible adapter for named `handler`
  functions.
- In Firebase Functions, adapt the HTTPS request into a web-standard `Request`,
  then write the returned `Response` through the function response object.
- Replace the in-memory `state.workbookJson` with your durable store when the
  workbook needs to survive cold starts or multiple instances; start from the
  typed adapters in `persistence-adapters.ts` for Postgres JSONB, Redis/KV, or
  object storage.

For the longer walkthrough, see
[`docs/serverless-workpaper-api-route.md`](../../docs/serverless-workpaper-api-route.md).
