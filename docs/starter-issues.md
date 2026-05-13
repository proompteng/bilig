# Starter Issues

This page is the stable contributor on-ramp for small public `bilig` tasks. It
links directly to current scoped issues instead of relying on GitHub issue
search indexing.

Current starter queue as of May 13, 2026:

- 40 open `first-timers-only` issues.
- 34 issues are generally available for a new contributor to claim.
- 6 issues already have active pull requests; comment before duplicating one of
  those patches.

## Available Starter Picks

### Agent And Tooling Docs

- [#153: docs(agent): add WorkPaper tool result contract table](https://github.com/proompteng/bilig/issues/153)
- [#163: docs(agent): add coding-agent workbook automation recipe](https://github.com/proompteng/bilig/issues/163)

### Example Recipes

- [#134: docs(examples): add XLSX formula-cache roundtrip example](https://github.com/proompteng/bilig/issues/134)
- [#193: docs(examples): add append-rows WorkPaper example](https://github.com/proompteng/bilig/issues/193)
- [#194: docs(examples): add changed-cells audit log example](https://github.com/proompteng/bilig/issues/194)
- [#195: docs(examples): add range-to-records export example](https://github.com/proompteng/bilig/issues/195)
- [#196: docs(examples): add validation-error JSON response example](https://github.com/proompteng/bilig/issues/196)
- [#207: docs(examples): add inventory reorder point WorkPaper example](https://github.com/proompteng/bilig/issues/207)
- [#208: docs(examples): add accounts receivable aging WorkPaper example](https://github.com/proompteng/bilig/issues/208)
- [#209: docs(examples): add usage-based billing tier WorkPaper example](https://github.com/proompteng/bilig/issues/209)
- [#210: docs(examples): add support SLA breach summary example](https://github.com/proompteng/bilig/issues/210)
- [#211: docs(examples): add weighted sales pipeline forecast example](https://github.com/proompteng/bilig/issues/211)
- [#212: docs(examples): add headcount capacity forecast example](https://github.com/proompteng/bilig/issues/212)
- [#217: docs(examples): add loan amortization WorkPaper example](https://github.com/proompteng/bilig/issues/217)
- [#218: docs(examples): add cohort retention WorkPaper example](https://github.com/proompteng/bilig/issues/218)
- [#219: docs(examples): add commission payout WorkPaper example](https://github.com/proompteng/bilig/issues/219)
- [#220: docs(examples): add cash runway WorkPaper example](https://github.com/proompteng/bilig/issues/220)
- [#221: docs(examples): add currency sensitivity WorkPaper example](https://github.com/proompteng/bilig/issues/221)
- [#222: docs(examples): add payroll accrual WorkPaper example](https://github.com/proompteng/bilig/issues/222)
- [#223: docs(examples): add deferred revenue roll-forward WorkPaper example](https://github.com/proompteng/bilig/issues/223)
- [#257: docs(examples): add a runnable Hono WorkPaper route smoke](https://github.com/proompteng/bilig/issues/257)
- [#258: docs(examples): add Cloudflare KV WorkPaper persistence snippet](https://github.com/proompteng/bilig/issues/258)
- [#260: docs(examples): add Fastify WorkPaper route smoke snippet](https://github.com/proompteng/bilig/issues/260)

### Contributor Docs

- [#156: docs(contributor): add first PR description template](https://github.com/proompteng/bilig/issues/156)
- [#158: docs(contributor): add fuzz replay walkthrough](https://github.com/proompteng/bilig/issues/158)
- [#198: docs(contributor): add first issue local verification matrix](https://github.com/proompteng/bilig/issues/198)
- [#255: docs(serverless): explain TypeScript local import suffixes](https://github.com/proompteng/bilig/issues/255)
- [#256: test(docs): enforce public snippets stay TypeScript-first](https://github.com/proompteng/bilig/issues/256)

### Headless Package, Benchmarks, And Troubleshooting

- [#154: docs(comparison): add headless engine use-case chooser](https://github.com/proompteng/bilig/issues/154)
- [#155: docs(troubleshooting): add formula diagnostic output examples](https://github.com/proompteng/bilig/issues/155)
- [#159: docs(troubleshooting): document CSV delimiter autodetection](https://github.com/proompteng/bilig/issues/159)
- [#162: docs(benchmarks): add WorkPaper benchmark reproduction notes](https://github.com/proompteng/bilig/issues/162)
- [#197: docs(headless): add WorkPaper persistence decision table](https://github.com/proompteng/bilig/issues/197)
- [#259: docs(service): add Prisma-backed WorkPaper JSON persistence recipe](https://github.com/proompteng/bilig/issues/259)

## Already In Review

These are still open, but a contributor has an active pull request attached.
Pick one only after checking the PR thread or asking whether the scope is free.

- [#231: test(examples): cover MCP stdio JSON-RPC error responses](https://github.com/proompteng/bilig/issues/231) via [PR #237](https://github.com/proompteng/bilig/pull/237)
- [#233: docs(examples): add a copy-paste MCP stdio output transcript](https://github.com/proompteng/bilig/issues/233) via [PR #234](https://github.com/proompteng/bilig/pull/234)
- [#247: docs(agent): add OpenAI WorkPaper tool-calling recipe](https://github.com/proompteng/bilig/issues/247) via [PR #254](https://github.com/proompteng/bilig/pull/254)
- [#248: docs(service): add downstream GitHub Actions smoke snippet](https://github.com/proompteng/bilig/issues/248) via [PR #253](https://github.com/proompteng/bilig/pull/253)
- [#249: docs(examples): add AWS Lambda Function URL WorkPaper adapter note](https://github.com/proompteng/bilig/issues/249) via [PR #252](https://github.com/proompteng/bilig/pull/252)
- [#250: docs(headless): add package-size and cold-start evaluation note](https://github.com/proompteng/bilig/issues/250) via [PR #251](https://github.com/proompteng/bilig/pull/251)

The list intentionally excludes closed issues and broad corpus/parity epics. Add
new starter tickets only when the expected patch can stay small, has a clear
acceptance test, and does not require understanding the whole workbook runtime.

## Claim A Starter Issue

Comment on the issue before opening a pull request. If the issue is unassigned,
a maintainer can assign it to you and keep the scope reserved while you work.
If it already has an assignee, pick another starter ticket or ask whether the
current assignee still wants help.

For a first patch, keep the pull request focused on the issue's acceptance
proof. Include the command you ran, mention the issue number, and open a draft
pull request early if any requirement is unclear. The
[new contributor guide](new-contributor-guide.md) gives the shortest setup,
code-map, and
[first-time command checklist](new-contributor-guide.md#first-time-command-checklist).

Useful filters:

- [`good first issue`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22good%20first%20issue%22)
- [`first-timers-only`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only)
- [`area: import-export`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22area%3A%20import-export%22)
- [`area: formula`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22area%3A%20formula%22)
- [`needs reduced fixture`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22needs%20reduced%20fixture%22)

GitHub surfaces issues labeled `good first issue` in contributor discovery
paths, per
[GitHub's label guidance](https://docs.github.com/articles/helping-new-contributors-find-your-project-with-labels),
so starter tickets should stay genuinely scoped and current. Do not use that
label for cross-cutting formula, import/export, or runtime changes that require
broad architectural context.

Use `first-timers-only` only for issues that are ready for someone making their
first contribution to this repository. Those issues should name the expected
files, a copyable validation command, and a narrow acceptance proof in the issue
body.

## What Makes A Good Starter Patch

- Keep the change small enough to review in one sitting.
- Use public `@bilig/headless` exports in examples.
- Prefer runnable recipes over abstract prose.
- Link back to the relevant package README or benchmark evidence note.
- Include the focused validation command in the issue or PR description.

## Maintainer Checklist

When opening a starter task, use the `Starter task` issue template and include:

- the likely files or directories to inspect first
- a suggested implementation approach
- the exact command or artifact that proves completion
- any out-of-scope behavior that should not be pulled into the first PR

Add `good first issue` only after the task has enough context for a newcomer to
make progress without learning the whole workbook runtime. Add
`first-timers-only` only when the issue can be completed from the issue body and
linked docs without maintainer-only context.

Before opening a PR, read
[`CONTRIBUTING.md`](https://github.com/proompteng/bilig/blob/main/CONTRIBUTING.md)
and run the smallest relevant local check first. If the change touches package
behavior, run `pnpm run ci` before asking for review.
