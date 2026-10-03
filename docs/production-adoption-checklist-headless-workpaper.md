---
title: Production adoption checklist for headless WorkPaper
published: true
description: A service-readiness checklist for adopting @bilig/workpaper in Node services and agent tools without overstating Excel parity or performance claims.
tags: typescript, node, spreadsheet, production, agents
canonical_url: https://proompteng.github.io/bilig/production-adoption-checklist-headless-workpaper.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Production Adoption Checklist For Headless WorkPaper

`@bilig/workpaper` is useful when workbook formulas are product logic and the
runtime boundary is a Node service, queue worker, serverless route, test, or
agent tool. This checklist is for deciding whether to move from evaluation to a
controlled production use.

It is deliberately narrower than "is bilig a complete spreadsheet platform?"
The adoption decision should be based on your workbook shape, your rollback
path, and the evidence you keep in CI.

## Start With A Low-Blast-Radius Workflow

Good first production candidates:

- quote approval rules where the source inputs are typed JSON
- payout or budget checks that can be recomputed from source data
- import validation where unsupported formulas can reject the import
- internal agent tools that need read-after-write proof
- service tests that need workbook-shaped fixtures

Avoid first using it for:

- irreversible money movement without a shadow-mode period
- regulated workflows without independent review and audit trails
- arbitrary uploaded Excel workbooks without compatibility gating
- workflows that require native macros, desktop Excel automation, or complete
  Excel UI behavior

## Required Service Checks

Before promoting a WorkPaper-backed workflow, keep a small fixture in your own
repository that proves:

1. The package version is pinned.
2. The workbook can be built or restored from a checked fixture.
3. Every formula family used by the workflow is covered by your fixture.
4. The service writes typed inputs through the WorkPaper API.
5. The service reads calculated outputs after the edit.
6. The WorkPaper document serializes to JSON.
7. The serialized document restores into the same computed state.
8. Unsupported formulas fail with a visible diagnostic instead of being treated
   as valid.
9. The workflow can run in shadow mode beside the current implementation.
10. Rollback means switching the caller back to the previous implementation or
    previous pinned package version.

The smallest proof looks like:

```sh
npm install @bilig/workpaper
npm install -D tsx typescript @types/node
curl -fsSLo quickstart.ts https://proompteng.github.io/bilig/npm-eval.ts
npx tsx quickstart.ts
```

The expected signal is `verified: true`, plus matching `after` and
`afterRestore` values.

## Runtime Boundaries

Use WorkPaper as a service boundary, not as hidden global state.

- Accept typed business inputs at the edge of the service.
- Map those inputs to explicit cell writes.
- Read explicit output cells by address or named workflow convention.
- Persist WorkPaper JSON only when persisted workbook state is part of the
  product behavior.
- Keep source business records as the authority when the workbook is just a
  calculation model.
- Treat imported XLSX cached formula values as hints, not as an accuracy oracle.

For persistence, start with
[`persisting-formula-backed-workpaper-documents-in-node.md`](persisting-formula-backed-workpaper-documents-in-node.md).

## Compatibility Gates

Do not rely on broad spreadsheet claims. Gate the exact workbook shape you use.

- Read
  [`where-bilig-is-not-excel-compatible-yet.md`](where-bilig-is-not-excel-compatible-yet.md)
  before accepting real Excel workbooks.
- Use the XLSX corpus verifier when import/export compatibility matters.
- Add a minimal fixture for each formula edge case that matters to your
  workflow.
- Keep macro execution, desktop automation, and interactive UI behavior out of
  scope unless you have a separate product path for them.
- Treat volatile or environment-dependent formulas as policy decisions, not
  silent successes.

## Performance And Regression Gates

The repository benchmark is evidence, not a blanket promise. Use it as a
starting point, then measure your own workload.

Repository checks:

```sh
pnpm import-export:fidelity:check
pnpm large-workbook:slo:check
pnpm public:evidence:check
```

Application checks:

- one representative small fixture
- one representative large fixture if workbook size matters
- one restore-after-save fixture
- one unsupported-formula fixture
- one rollback smoke test using the previous service path

Keep benchmark losses visible. If a row loses on p95, document whether that row
matters to your workload instead of hiding the caveat.

## Agent Tool Gates

For agent workflows, require proof objects rather than prose.

The tool call should return:

- cell writes requested
- formula cells read after the writes
- display values and raw values where relevant
- validation errors or unsupported formula diagnostics
- serialized document size or persistence revision if state is saved

For an MCP shape, start with
[`mcp-workpaper-tool-server.md`](mcp-workpaper-tool-server.md). For direct
tool-calling, start with
[`agent-workpaper-tool-calling-recipe.md`](agent-workpaper-tool-calling-recipe.md).

## Release And Operations Gates

Before upgrading a production workflow:

- verify npm package provenance and registry signatures with
  [`npm-provenance-package-trust.md`](npm-provenance-package-trust.md)
- inspect the
  [OpenSSF Scorecard](https://scorecard.dev/viewer/?uri=github.com/proompteng/bilig)
  result for repository security posture
- read the package changelog
- rerun your workflow fixture on the new package version
- rerun restore-after-save proof against existing saved documents
- confirm the security and support policies still match your risk requirements
- deploy behind a feature flag or caller-level routing rule
- keep previous package/version rollback instructions close to the deployment
  runbook

Security reports belong in
[`SECURITY.md`](../SECURITY.md). Ordinary support requests and reproducible
workflow issues belong in [`SUPPORT.md`](../SUPPORT.md).

## Decision

Use `@bilig/workpaper` in production first where you can pin the package, own the
workflow fixture, run shadow mode, and roll back at the caller boundary.

Do not use it as a default for arbitrary customer-critical durable execution
until your own replay, restore, compatibility, and rollback evidence is boring.
