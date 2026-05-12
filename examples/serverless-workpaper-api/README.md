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
npm run smoke
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

## Moving Into A Real Route

- In a Next.js app route, call `handleWorkPaperRequest()` from `GET()` and
  `POST()`.
- In a Cloudflare Worker, call it from `fetch(request)`.
- In Hono, pass `c.req.raw` directly to the shared handler.
- In Deno, return `handleWorkPaperRequest(request)` from `Deno.serve()` or a
  `fetch` default export.
- In SvelteKit, return `handleWorkPaperRequest(request)` from thin `+server.js`
  `GET` and `POST` handlers.
- In Fastify, adapt the framework request into a web-standard `Request`, then
  write the returned `Response` through `reply`.
- In Express, adapt the framework request into a web-standard `Request`, then
  write the returned `Response` through `res`.
- In Koa, adapt `ctx` into a web-standard `Request`, then write status, headers,
  and body back to the Koa context.
- In an AWS Lambda Function URL handler, adapt the event into a web-standard
  `Request`, then return a Lambda proxy response.
- Replace the in-memory `state.workbookJson` with your durable store when the
  workbook needs to survive cold starts or multiple instances.

For the longer walkthrough, see
[`docs/serverless-workpaper-api-route.md`](../../docs/serverless-workpaper-api-route.md).
