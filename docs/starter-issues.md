# Starter Issues

This page is the stable contributor on-ramp for small public `bilig` tasks. It
intentionally stays short: GitHub's `good first issue`, `first-timers-only`, and
`help wanted` labels should point to work that is current, scoped, and credible
for someone opening the repository cold.

Current starter queue as of May 16, 2026:

- 15 open `good first issue` issues.
- 15 open `first-timers-only` issues.
- 15 open `help wanted` issues.
- 9 starter issues are code or test tasks.
- 6 starter issues are focused docs or integration transcript tasks.
- 0 starter issues are currently under active review.

## Start Here This Week

If you are opening the queue cold, pick one of these before browsing the full
issue list. They are small, current, and map to the public adoption path for
`@bilig/headless`.

- [#360: test(headless): cover display-value readback after JSON restore](https://github.com/proompteng/bilig/issues/360)
  proves the persistence readback path in a focused headless test.
- [#361: test(headless): cover range readback after an input edit](https://github.com/proompteng/bilig/issues/361)
  proves the service/agent write-then-read workflow.
- [#362: test(examples): guard the headless README command index against missing scripts](https://github.com/proompteng/bilig/issues/362)
  keeps the npm README commands from drifting.
- [#363: test(examples): add invalid-request proof to the HTTP JSON summary smoke](https://github.com/proompteng/bilig/issues/363)
  gives service examples a clear error-path check.
- [#273: docs(examples): add Express WorkPaper route smoke](https://github.com/proompteng/bilig/issues/273)
  adds the most familiar Node service entry point.
- [#300: docs(examples): add tRPC WorkPaper procedure smoke](https://github.com/proompteng/bilig/issues/300)
  covers a common TypeScript RPC integration.
- [#334: docs(agent): add OpenAI Responses streaming tool-call transcript](https://github.com/proompteng/bilig/issues/334)
  helps agent builders see the tool-call loop.
- [#358: docs(agent): add AI SDK onStepFinish WorkPaper transcript](https://github.com/proompteng/bilig/issues/358)
  connects the WorkPaper proof loop to a common TypeScript agent stack.

## Code And Test Starters

- [#360: test(headless): cover display-value readback after JSON restore](https://github.com/proompteng/bilig/issues/360)
- [#361: test(headless): cover range readback after an input edit](https://github.com/proompteng/bilig/issues/361)
- [#362: test(examples): guard the headless README command index against missing scripts](https://github.com/proompteng/bilig/issues/362)
- [#363: test(examples): add invalid-request proof to the HTTP JSON summary smoke](https://github.com/proompteng/bilig/issues/363)
- [#366: test(headless): cover changed named expressions after WorkPaper restore](https://github.com/proompteng/bilig/issues/366)
- [#367: test(headless): cover dense sheet range read with sparse values](https://github.com/proompteng/bilig/issues/367)
- [#368: test(headless): cover two-column formula tiling in fill ranges](https://github.com/proompteng/bilig/issues/368)
- [#369: test(headless): cover tab-indented formula prefix detection](https://github.com/proompteng/bilig/issues/369)
- [#371: test(examples): add deterministic markdown-report output test](https://github.com/proompteng/bilig/issues/371)

## Integration Docs Starters

- [#273: docs(examples): add Express WorkPaper route smoke](https://github.com/proompteng/bilig/issues/273)
- [#283: docs(mcp): add Cursor MCP config for the WorkPaper stdio server](https://github.com/proompteng/bilig/issues/283)
- [#285: docs(mcp): add MCP Inspector smoke-test transcript for the WorkPaper server](https://github.com/proompteng/bilig/issues/285)
- [#300: docs(examples): add tRPC WorkPaper procedure smoke](https://github.com/proompteng/bilig/issues/300)
- [#334: docs(agent): add OpenAI Responses streaming tool-call transcript](https://github.com/proompteng/bilig/issues/334)
- [#358: docs(agent): add AI SDK onStepFinish WorkPaper transcript](https://github.com/proompteng/bilig/issues/358)

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
Read
[CONTRIBUTING.md](https://github.com/proompteng/bilig/blob/main/CONTRIBUTING.md)
before opening the pull request.

Useful filters:

- [`good first issue`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22good%20first%20issue%22)
- [`first-timers-only`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only)
- [`help wanted`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22help%20wanted%22)

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

Add `help wanted` only when an external contributor can make progress without
private context or maintainer-only systems.
