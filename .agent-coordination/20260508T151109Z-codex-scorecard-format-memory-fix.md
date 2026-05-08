Agent: Codex
When: 2026-05-08T15:11:09Z

Summary:
- Found a real memory issue in the public workbook corpus scorecard harness: `formatJsonForRepo` spawned `oxfmt` over the full generated JSON, which OOMed under a 384 MiB Node heap cap during `refresh-scorecard-from-checkpoint --check`.
- Replaced the external formatter invocation with an in-process deterministic JSON formatter that preserves the checked-in compact-array style and avoids temp files/process workers.
- Added focused regression coverage for no external formatter binary, wrapped long primitive arrays, and multiline object arrays.

Verification:
- `NODE_OPTIONS=--max-old-space-size=384 pnpm exec vitest run scripts/__tests__/scorecard-format.test.ts`
- `NODE_OPTIONS=--max-old-space-size=384 pnpm public-workbook-corpus:refresh-scorecard-from-checkpoint -- --check`
- `NODE_OPTIONS=--max-old-space-size=384 pnpm public-workbook-corpus:completion-audit:check`
- `git diff --check`
- Process sweep showed no matching workbook corpus, verify-artifact, Vitest, Playwright, HyperFormula benchmark, git push, or SSH jobs after completion.
