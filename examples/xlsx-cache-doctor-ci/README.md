# XLSX cache doctor CI example

This example is for repositories that keep `.xlsx` fixtures, pricing models, or
report templates under version control and want CI to catch stale cached formula
values before tests or services read the wrong number.

The fixture at `fixtures/stale-pricing.xlsx` has sixty formula cells and one
intentionally stale formula cache after the old 50-cell inspection cutoff:

- `Sheet1!A61` is `60`.
- `Sheet1!B61` contains `=A61*10`.
- The saved cached value for `Sheet1!B61` is `999`, but the recalculated value is
  `600`.

Run the same diagnostic locally:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- \
  xlsx-cache-doctor fixtures/stale-pricing.xlsx --json
```

Expected output shape is committed at
`reports/stale-pricing.cache-doctor.json`. The important fields are:

```json
{
  "formulaCellCount": 60,
  "inspectedFormulaCellCount": 60,
  "uninspectedFormulaCellCount": 0,
  "staleCachedFormulaCount": 1,
  "suggestedReads": ["Sheet1!B2", "...", "Sheet1!B61"],
  "commandSucceeded": true,
  "inspectionCompleted": true
}
```

## GitHub Actions

Copy `.github/workflows/xlsx-cache-doctor.yml` into a repository that keeps XLSX
fixtures. Start with `fail-on-stale: "false"` if you want a non-blocking report,
then switch it to `"true"` when stale formula caches should fail pull requests.

Or generate the same workflow from npm:

```sh
mkdir -p .github/workflows
npm exec --package @bilig/xlsx-formula-recalc@latest -- \
  xlsx-cache-doctor --print-github-action "fixtures/**/*.xlsx" \
  --changed-files-only false \
  > .github/workflows/xlsx-cache-doctor.yml
```

That generated workflow is report-only by default. Add `--fail-on-stale true`
after the first report is clean enough to make the check blocking.

The workflow is intentionally read-only: it checks out the repo, inspects one
or more matched workbooks, writes a job summary, and uploads the JSON report as
an artifact. The Action also writes the same human-readable summary to a
Markdown report path, so teams can upload it, paste it into a PR comment from
their own workflow, or attach it to a release note without giving the detector a
write token. It does not comment on pull requests, rewrite workbooks, require
secrets, or use a write token.

To see the same shape as a live GitHub reviewer, inspect the
[demo pull request](https://github.com/proompteng/xlsx-cache-doctor-demo/pull/1).
Its green Action run proves 1 workbook, 60 formula cells, and 1 stale cached
formula value, then uploads the JSON report artifact.

When a stale cache is found, the job summary includes the exact cell, cached
value, recalculated value, and a follow-up `xlsx-recalc --read ... --json`
command. The Actions log also gets warning annotations for the first stale
cells, so reviewers can see the stale value without downloading the artifact.

## Regenerate The Fixture

The fixture is deliberately tiny. To regenerate it from source:

```sh
pnpm --filter @bilig/xlsx-formula-recalc build
bun scripts/build-stale-pricing-fixture.ts
bun ../../packages/xlsx-formula-recalc/src/cache-doctor-cli.ts \
  fixtures/stale-pricing.xlsx --json > reports/stale-pricing.cache-doctor.json
```

Use the follow-up proof only after the detector finds the formula cells that
matter to your service:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- \
  xlsx-recalc fixtures/stale-pricing.xlsx \
  --read Sheet1!B61 \
  --out fixtures/stale-pricing.recalculated.xlsx \
  --json
```
