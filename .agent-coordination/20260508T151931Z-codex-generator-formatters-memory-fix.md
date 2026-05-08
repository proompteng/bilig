Agent: Codex
When: 2026-05-08T15:19:31Z

Summary:
- Finished the scorecard formatter memory fix across `scripts/gen-*.ts`.
- Removed generator-local JSON formatter functions that spawned `oxfmt` over generated scorecard JSON.
- Routed generators through the shared deterministic formatter and kept legacy string-call compatibility.
- Added static coverage that fails if generator-local formatter shells return.

Verification:
- `pnpm exec oxfmt --write scripts/scorecard-format.ts scripts/__tests__/scorecard-format.test.ts scripts/__tests__/scorecard-generator-formatting.test.ts scripts/gen-*.ts`
- `pnpm exec oxlint --config .oxlintrc.json --type-aware --fix --deny-warnings ...`
- `NODE_OPTIONS=--max-old-space-size=384 pnpm exec vitest run scripts/__tests__/scorecard-format.test.ts scripts/__tests__/scorecard-generator-formatting.test.ts --pool=forks --maxWorkers=1 --testTimeout=20000`
- `rg -n "function formatJsonForRepo|node_modules', '.bin', 'oxfmt'|Bun\\.spawnSync\\(\\[oxfmtPath" scripts/gen-*.ts` returned no matches.
- `git diff --check`
