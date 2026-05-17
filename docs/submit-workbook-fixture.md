# Submit A Workbook Fixture

Status: public path for turning evaluator blockers into tests, examples, and
docs

Use this when Bilig is close to useful for your workbook, service, or agent
workflow, but a real case is missing from the public proof set.

The best reports are small enough to review and specific enough to become one
of these artifacts:

- a formula regression test
- an XLSX import/export corpus case
- a WorkPaper JSON persistence fixture
- a service-route example
- an MCP or agent-tool transcript
- a compatibility note with a runnable workaround

Open the fixture form:
<https://github.com/proompteng/bilig/issues/new?template=workbook_fixture.yml>.

If you want to discuss the shape before opening a fixture issue, use the public
fixture discussion:
<https://github.com/proompteng/bilig/discussions/414>.

## What To Send

Send the smallest public case that proves the behavior.

Good fixture reports include:

- the `@bilig/headless` version or commit you tested
- a reduced workbook, public gist, or pasted sheet data
- exact sheet names, cells, ranges, and formulas
- expected output from Excel, another system, or a manual check
- actual Bilig output, error, or missing API
- one command or script maintainers can run

Do not send confidential spreadsheets, customer data, private financial models,
or files whose license does not allow redistribution. If the original workbook
is private, replace names and numbers with neutral values while keeping the same
formula shape.

## Quick Checks

For a formula or WorkPaper API case, reduce it to a script first. A small script
beats a large workbook because it can become a test directly.

```sh
mkdir bilig-fixture-check
cd bilig-fixture-check
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install --save-dev tsx typescript
```

For an XLSX case inside this repository, prefer the corpus checker:

```sh
pnpm workpaper:xlsx-corpus:check ./path/to/reduced.xlsx -- --from-import-snapshot
```

For a service workflow, start from the quote approval example and change only
the inputs, formulas, or expected response that matter:

```sh
cd examples/serverless-workpaper-api
npm install
npm run quote-approval-api
```

For an MCP or agent-tool workflow, include the JSON-RPC or tool-call transcript
showing the write, recalculation, and readback.

## What Maintainers Do With It

If the case is in scope and public, maintainers can turn it into a committed
fixture, test, docs page, or example. When it lands, the issue should link the
commit or release note so future evaluators can see the exact blocker that was
converted into public proof.

If the case is out of scope, the issue should still end with a concrete answer:
unsupported Excel behavior, missing package boundary, too-private fixture,
runtime limitation, or a better tool for that job.

## Short Version

If Bilig almost works for a real workbook, do not describe the whole private
system. Send one reduced public case with expected output:
<https://github.com/proompteng/bilig/issues/new?template=workbook_fixture.yml>.

Discussion: <https://github.com/proompteng/bilig/discussions/414>.
