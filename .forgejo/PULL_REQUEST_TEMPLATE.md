## Summary

- What changed?
- Why was it needed?
- What user-visible or engine-visible behavior should reviewers expect?

## Scope

- Area: `formula` / `core` / `wasm-kernel` / `renderer` / `grid` / `import` / `docs` / other
- Type: feature / bug fix / refactor / test-only / docs-only / perf

## Validation

- [ ] `pnpm protocol:check`
- [ ] `pnpm formula-inventory:check`
- [ ] `pnpm lint`
- [ ] `pnpm wasm:build`
- [ ] `pnpm typecheck`
- [ ] `pnpm coverage`
- [ ] `pnpm test:browser`
- [ ] `pnpm release:check`
- [ ] `pnpm run ci`

List any commands you actually ran and any intentionally skipped checks:

```text
```

## Formula And Engine Impact

- New formulas implemented:
- Existing formulas changed:
- JS-only / WASM / protocol impact:
- Metadata impact: names / tables / spills / structured refs / snapshots / external adapters

## Generated Or Contract Files

- [ ] No generated files changed
- [ ] Generated files were regenerated and checked in

If generated files changed, list them:

```text
```

## Performance And Release Risk

- Recalc/perf impact:
- Binary-size / release-budget impact:
- Browser/runtime risk:

## Notes For Reviewers

- Key files to review:
- Follow-up work:
- Known limitations:
