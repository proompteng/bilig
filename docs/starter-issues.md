# Starter Issues

This page is the stable contributor on-ramp for small public `bilig` tasks. It
links directly to current scoped issues instead of relying on GitHub issue
search indexing.

Current starter tickets as of May 12, 2026:

- [#134: docs(examples): add XLSX formula-cache roundtrip example](https://github.com/proompteng/bilig/issues/134)
- [#141: docs(examples): add HTTP JSON summary WorkPaper example](https://github.com/proompteng/bilig/issues/141)
- [#142: docs(examples): add JSON file input WorkPaper example](https://github.com/proompteng/bilig/issues/142)
- [#143: docs(examples): add markdown report WorkPaper output example](https://github.com/proompteng/bilig/issues/143)
- [#144: docs(examples): add formula diagnostics WorkPaper example](https://github.com/proompteng/bilig/issues/144)
- [#145: docs(examples): add snapshot diff WorkPaper example](https://github.com/proompteng/bilig/issues/145)
- [#146: docs(readme): add first-time contributor command checklist](https://github.com/proompteng/bilig/issues/146)
- [#147: docs(examples): add npm script for CSV-shaped input example](https://github.com/proompteng/bilig/issues/147)
- [#148: docs(examples): add range readback WorkPaper example](https://github.com/proompteng/bilig/issues/148)
- [#149: docs(examples): add sheet inspection WorkPaper example](https://github.com/proompteng/bilig/issues/149)
- [#150: docs(examples): add headless example command index](https://github.com/proompteng/bilig/issues/150)
- [#151: docs(headless): map npm visitor needs to examples](https://github.com/proompteng/bilig/issues/151)
- [#152: test(docs): guard package README links for npm readers](https://github.com/proompteng/bilig/issues/152)
- [#153: docs(agent): add WorkPaper tool result contract table](https://github.com/proompteng/bilig/issues/153)
- [#154: docs(comparison): add headless engine use-case chooser](https://github.com/proompteng/bilig/issues/154)
- [#155: docs(troubleshooting): add formula diagnostic output examples](https://github.com/proompteng/bilig/issues/155)
- [#156: docs(contributor): add first PR description template](https://github.com/proompteng/bilig/issues/156)
- [#158: docs(contributor): add fuzz replay walkthrough](https://github.com/proompteng/bilig/issues/158)
- [#159: docs(troubleshooting): document CSV delimiter autodetection](https://github.com/proompteng/bilig/issues/159)
- [#160: docs(headless): add WorkPaper API read/write cheat sheet](https://github.com/proompteng/bilig/issues/160)
- [#162: docs(benchmarks): add WorkPaper benchmark reproduction notes](https://github.com/proompteng/bilig/issues/162)
- [#163: docs(agent): add coding-agent workbook automation recipe](https://github.com/proompteng/bilig/issues/163)
- [#181: docs(serverless): add Remix resource route adapter](https://github.com/proompteng/bilig/issues/181)
- [#182: docs(serverless): add Nitro event handler adapter](https://github.com/proompteng/bilig/issues/182)
- [#183: docs(serverless): add NestJS controller adapter](https://github.com/proompteng/bilig/issues/183)
- [#184: docs(serverless): add Elysia route adapter](https://github.com/proompteng/bilig/issues/184)
- [#185: docs(serverless): add Vercel Function adapter](https://github.com/proompteng/bilig/issues/185)
- [#186: docs(serverless): add Firebase Functions HTTPS adapter](https://github.com/proompteng/bilig/issues/186)

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
code-map, and PR-proof path.

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

Before opening a PR, read [`CONTRIBUTING.md`](../CONTRIBUTING.md) and run the
smallest relevant local check first. If the change touches package behavior,
run `pnpm run ci` before asking for review.
