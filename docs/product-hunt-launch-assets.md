---
title: Product Hunt launch assets for bilig
description: Gallery images, thumbnail, copy, and launch-day checklist for a Product Hunt draft.
published: true
---

# Product Hunt launch assets for bilig

Product Hunt's current posting guide recommends a `240x240` square thumbnail and
`1270x760` gallery images. This asset pack keeps those files in-repo so a draft
can be assembled without ad hoc image work.

Do not schedule a launch until there is a clear support window for comments and
one more concrete product proof item to point people at. Creating a draft is
fine; launching blind is not.

## Generated assets

Run:

```sh
pnpm docs:launch-assets:generate
pnpm docs:launch-assets:check
```

Files:

- thumbnail:
  [`docs/assets/product-hunt-thumbnail.png`](assets/product-hunt-thumbnail.png)
- gallery image 1:
  [`docs/assets/product-hunt-gallery-01-workbook-api.png`](assets/product-hunt-gallery-01-workbook-api.png)
- gallery image 2:
  [`docs/assets/product-hunt-gallery-02-agent-readback.png`](assets/product-hunt-gallery-02-agent-readback.png)
- gallery image 3:
  [`docs/assets/product-hunt-gallery-03-node-service.png`](assets/product-hunt-gallery-03-node-service.png)

![Product Hunt gallery image showing the workbook API asset](assets/product-hunt-gallery-01-workbook-api.png)

![Product Hunt gallery image showing the agent readback asset](assets/product-hunt-gallery-02-agent-readback.png)

![Product Hunt gallery image showing the Node service asset](assets/product-hunt-gallery-03-node-service.png)

## Draft fields

Name:

```text
bilig
```

Tagline:

```text
headless spreadsheet engine for services and agents
```

Website:

```text
https://github.com/proompteng/bilig
```

Topics:

```text
Developer Tools, Open Source, Artificial Intelligence
```

Description:

```text
bilig is an open-source TypeScript workbook runtime for Node services and coding agents. Create formula-backed sheets, read computed values, persist documents, and verify workbook edits without screen scraping.
```

First comment:

```text
i built bilig because spreadsheet automation needs a better boundary than a grid screenshot.

the public package is @bilig/headless. it lets a node service or agent create workbook state, write formulas, read computed values, serialize the document, restore it, and verify the same output again.

the useful feedback is concrete: missing formula semantics, api friction, xlsx import/export expectations, or a real workbook case that should become a fixture.

repo: https://github.com/proompteng/bilig
npm: https://www.npmjs.com/package/@bilig/headless
proof article: https://dev.to/gregkonush/why-agents-need-workbook-apis-instead-of-spreadsheet-screenshots-3d61
```

## Launch checklist

- Use a personal Product Hunt account, not a company account.
- Create a draft first; do not pick a launch date until the support window is
  known.
- Upload the generated thumbnail and at least two gallery images.
- Use the GitHub repository as the primary URL, not a blog post.
- Keep the launch copy honest: early infrastructure, not a finished Excel clone.
- Link the compatibility caveats:
  <https://proompteng.github.io/bilig/where-bilig-is-not-excel-compatible-yet.html>.
- Link the feedback discussion:
  <https://github.com/proompteng/bilig/discussions/115>.
- After launch, log the Product Hunt URL in the feedback discussion and compare
  GitHub stars, npm downloads, and GitHub traffic after `24h`, `48h`, and `7d`.

## Sources

- Product Hunt posting guide:
  <https://help.producthunt.com/en/articles/479557-how-to-post-a-product>
- Product Hunt featuring guidelines:
  <https://help.producthunt.com/en/articles/9883485-product-hunt-featuring-guidelines>
