---
title: Create a Bilig WorkPaper starter
published: true
description: Track the @bilig/create-workpaper starter package and use the maintained quote approval WorkPaper API example while npm package creation is being finalized.
tags: typescript, node, spreadsheet, formulas, opensource
canonical_url: https://proompteng.github.io/bilig/create-bilig-workpaper.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Create a Bilig WorkPaper starter

Use this path when you want a runnable project instead of a pasted snippet.

The starter package is being prepared as `@bilig/create-workpaper`, exposed by
`npm create @bilig/workpaper@latest`. Until that package is visible on npm, use
the maintained flagship example below; it exercises the same quote approval
WorkPaper API shape against the published `@bilig/headless` package.

The starter creates a quote approval API with `@bilig/headless`. It writes
quote inputs through an API-style handler, recalculates workbook formulas,
persists the WorkPaper as JSON, restores it, and verifies that the restored
formula output still matches the live result.

## Run It

Current published-package path:

```sh
git clone --depth 1 https://github.com/proompteng/bilig.git
cd bilig/examples/serverless-workpaper-api
npm install
npm run smoke
```

Prepared generated-project path:

```sh
npm create @bilig/workpaper@latest pricing-workpaper
cd pricing-workpaper
npm install
npm run smoke
```

Expected output includes:

```json
{
  "verified": true,
  "nextStep": "If this proof matches your service or agent workflow, star or bookmark Bilig: https://github.com/proompteng/bilig/stargazers"
}
```

The generated project also includes a local API server:

```sh
npm run dev
curl http://localhost:8788/api/quote/approval
curl -X POST http://localhost:8788/api/quote/approval \
  -H 'content-type: application/json' \
  -d '{"units":40,"listPrice":1200,"discount":0.05,"unitCost":760,"minimumMargin":0.3}'
```

## What It Proves

- a real service shape, not only a formula evaluator snippet;
- input writes into named workbook cells;
- formula readback after mutation;
- JSON persistence and restore;
- a narrow API surface that an agent tool or backend route can own.

## Source

- package source:
  [`packages/create-workpaper`](https://github.com/proompteng/bilig/tree/main/packages/create-workpaper)
- generated API source:
  [`packages/create-workpaper/template/src/index.ts`](https://github.com/proompteng/bilig/blob/main/packages/create-workpaper/template/src/index.ts)
- full flagship example:
  [quote approval WorkPaper API](quote-approval-workpaper-api.md)

If this starter matches a service or agent workflow you maintain, star or
bookmark the repo so the package is easier to find later:
<https://github.com/proompteng/bilig/stargazers>.
