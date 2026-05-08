# New Contributor Guide

Use this guide for your first `bilig` pull request. The shortest path is a
small docs, example, fixture, or test change with one clear validation command.

## Pick A Scoped Task

Start with [`docs/starter-issues.md`](starter-issues.md). Comment on the issue
before opening a pull request so a maintainer can confirm the scope is still
current.

If this is your first patch to this repository, prefer the
[`first-timers-only`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only)
filter. Those issues are reserved for tasks that should be possible from the
issue body, linked docs, and one focused validation command.

A useful claim comment says:

```md
I can take this.

Plan:
- files I expect to touch:
- validation command:
- question or assumption:
```

If the issue already has an assignee, ask whether help is still wanted before
starting work.

## Get Local Feedback Quickly

Use Node `24+`, Bun, and `pnpm@10.32.1`.

```bash
pnpm install
pnpm wasm:build
```

Then run the narrowest check that matches the change:

- docs or examples: `pnpm docs:discovery:check`
- formula behavior: `pnpm test:correctness:formula`
- WorkPaper or engine behavior: `pnpm test:correctness:core`
- import/export behavior: `pnpm test:correctness:corpus`
- browser UI behavior: `pnpm test:browser`

Run `pnpm run ci` before asking for review when the change touches runtime
behavior, generated artifacts, benchmarks, browser flows, or multiple packages.

## Know Where To Look

- Public headless API: `packages/headless/README.md`
- Runnable examples: `examples/headless-workpaper/`
- Formula fixtures: `packages/excel-fixtures/src/`
- Formula runtime checks: `packages/formula/src/__tests__/` and
  `packages/core/src/__tests__/formula-runtime-correctness.test.ts`
- Import/export checks: `packages/excel-import/src/__tests__/`
- Browser workbook shell: `apps/web/` and `packages/grid/`
- Agent protocol surfaces: `packages/agent-api/` and `docs/agent-api.md`

Prefer public package exports in examples. Do not import from `src/` or `dist/`
unless the issue is specifically about package internals.

## Keep The PR Easy To Merge

- Keep one issue per pull request.
- Add or tighten a focused test before changing behavior.
- Include the exact command output or fixture proof in the PR description.
- Link the issue with `Fixes #...` when the PR fully closes it.
- Open a draft PR early if setup, scope, or expected behavior is unclear.

The best first contribution gives a future user a clearer path to evaluate
`@bilig/headless` or gives maintainers a small regression proof they can keep in
CI.
