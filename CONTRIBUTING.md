# Contributing to bilig

Thanks for contributing. `bilig` is an engine-heavy TypeScript monorepo, so the
best contributions are small, tested, and explicit about which runtime behavior
they change.

## Start Here

- For a first patch, pick a current scoped ticket from
  [`docs/starter-issues.md`](docs/starter-issues.md).
- If this is your first contribution to this repository, start with the
  [`first-timers-only`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only)
  queue.
- If this is your first `bilig` pull request, use
  [`docs/new-contributor-guide.md`](docs/new-contributor-guide.md) for the
  shortest setup, code-map, and PR-proof path.
- If you are reporting a formula or workbook bug, include the exact formula,
  workbook shape, expected result, actual result, and the smallest command or
  fixture that reproduces it.
- If you are changing public behavior, add or tighten a regression test before
  changing implementation.

## First Patch Flow

1. Pick a starter issue and comment with the files you expect to touch, the
   validation command, and any open assumption.
2. Run the smallest useful local check before opening a pull request.
3. Keep the PR tied to one issue, one package or doc path, and one proof command
   whenever possible.
4. Open a draft PR early if setup, expected behavior, or scope is unclear.

## Local Setup

Use Node `24+`, Bun, and `pnpm@10.32.1`.

```bash
pnpm install
pnpm wasm:build
pnpm typecheck
pnpm test
```

For the app shell:

```bash
pnpm dev:web-local
```

For targeted changes, run the smallest useful gate first:

- Formula or WorkPaper behavior: `pnpm test:correctness:core`
- Formula package changes: `pnpm --filter @bilig/formula build`
- Headless package changes: `pnpm --filter @bilig/headless build`
- Import/export changes: `pnpm test:correctness:corpus`
- Browser shell changes: `pnpm test:browser`
- Docs discovery changes: `pnpm docs:discovery:check`

## Before You Open a PR

Run the narrowest checks that cover your change, then run the full gate when the
change is ready to publish.

```bash
pnpm lint
pnpm typecheck
pnpm test
pnpm test:browser
pnpm run ci
```

If a narrower gate is enough for a docs-only or fixture-only patch, say why in
the pull request.

If you edit generated protocol, formula inventory, workspace-resolution, or
benchmark-baseline sources, regenerate and commit the generated output.

## Good First Areas

- Add formula fixtures and tests for missing Excel-compatible semantics.
- Turn architecture docs into smaller runnable examples.
- Improve grid accessibility, keyboard behavior, and focus handling.
- Add WorkPaper benchmark cases that describe the real spreadsheet pattern they
  represent.
- Tighten engine correctness tests around mutation, snapshot, undo/redo, and
  dependency behavior.

## Formula Parity Fixture Walkthrough

Start by reading the existing fixture shape before adding a new case:

- `packages/excel-fixtures/src/` is the canonical formula and workbook-semantics
  corpus used by `@bilig/formula` and engine runtime checks.
- `packages/formula/src/compatibility.ts` records implementation status and
  generated inventory metadata for formula families.
- `packages/formula/src/__tests__/fixture-harness.test.ts` executes implemented
  canonical fixtures through the JavaScript evaluator.
- `packages/core/src/__tests__/formula-runtime-correctness.test.ts` covers the
  production runtime path for fixtures that should run through the engine.
- `packages/headless/fixtures/xlsx-corpus/` holds checked-in XLSX cached-result
  reductions for public workbook compatibility regressions.

A minimal formula-parity contribution should:

1. Add or tighten one small fixture with an Excel-observed expected result.
2. Update the compatibility entry only to the status that the implementation
   actually supports.
3. Add focused package tests when the fixture exposes behavior not already
   covered by the harness.
4. Run the generated checks before opening a PR:

```bash
pnpm formula-inventory:check
pnpm formula:dominance:check
pnpm test:correctness:formula
```

For cached-result workbook reductions, regenerate or verify the headless XLSX
fixture corpus instead of hand-editing binary evidence:

```bash
pnpm workpaper:xlsx-corpus:fixtures:generate
pnpm workpaper:xlsx-corpus:fixtures:check
```

Before claiming the work is ready, run `pnpm run ci`. Do not describe a fixture
as full Excel parity; the checked-in fixtures prove only the named formulas,
inputs, cached workbook results, and runtime paths they actually cover.

## Contribution Rules

- Keep public APIs boring and stable. Prefer `is...`, `allows...`, `on...`, and
  `on...Change` naming.
- Keep formula semantics in JavaScript first. Promote to WASM only after parity
  and differential tests are green.
- Avoid `any`; lint fails on weak typing and floating promises.
- Use explicit `.js` suffixes where nearby ESM imports already do.
- Do not mix UI rendering, behavior policy, and workbook engine logic in one
  component when a hook, controller, or package boundary can own one concern.
- Keep benchmark claims tied to commands, artifacts, counters, or checked-in
  fixtures.
- Do not weaken parity, benchmark, generated-file, or clean-diff checks to make
  a patch pass.

## PR Description

Include:

- what changed
- why the change belongs in this package
- commands run
- benchmark output or screenshots when behavior is visual or performance-related
- known risk or follow-up work

Small pull requests are easier to review and merge than broad refactors.

## Security

Do not post private workbook data, credentials, tokens, or vulnerability details
in public issues. Follow [`SECURITY.md`](SECURITY.md) for private security
reports.

## Source of Truth

Forgejo `origin` is the primary repo workflow for maintainers. GitHub mirrors
the public verification contract and public collaboration surface.
