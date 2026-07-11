# Custom Automation

`@bilig/workpaper` is built for workbook-shaped automation: inputs, formulas,
calculated outputs, and a saved JSON state that can run inside a TypeScript
service, queue worker, serverless route, GitHub Action, coding-agent tool, or
MCP server.

If your team already has a spreadsheet-like workflow and wants it to run as
software, MichelleBao can help turn it into a small working integration.

## Good Fits

- CSV or Excel-style import validation with calculated pass/fail checks.
- Pricing, quote approval, commission, payout, budget, or capacity calculators.
- Weekly or monthly report checks where inputs change but formulas stay stable.
- GitHub Actions that recalculate a workbook and publish a deterministic result.
- Agent or MCP tools that need read-after-write proof from workbook formulas.

## Not A Fit

- Private data extraction without permission.
- CAPTCHA bypass, account automation, scraping behind login walls, or bot abuse.
- Desktop Excel macros or manual spreadsheet editing workflows.
- Financial trading automation or high-risk production decisions without review.

## Small Fixed-Scope Packages

These are starting points for scoping, not a promise that every request fits
every package.

| Package | Scope | Typical deliverable |
| --- | --- | --- |
| Starter | One narrow script or example | TypeScript script, sample input, README |
| Workflow | One repeatable automation | Script/API route, validation, usage notes |
| Integration | Production-adjacent setup | GitHub Actions, service boundary, tests/docs |

For a first pass, send a sanitized sample input plus the expected output shape.
The first reply should answer:

1. What do you do manually today?
2. What is the input format?
3. What output do you need?
4. Is this one-time cleanup or a repeatable workflow?

Contact: <mxbao063@gmail.com>
