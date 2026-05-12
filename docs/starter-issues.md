# Starter Issues

This page is the stable contributor on-ramp for small public `bilig` tasks. It
links directly to current scoped issues instead of relying on GitHub issue
search indexing.

Current starter tickets as of May 12, 2026:

- [#134: docs(examples): add XLSX formula-cache roundtrip example](https://github.com/proompteng/bilig/issues/134)
- [#153: docs(agent): add WorkPaper tool result contract table](https://github.com/proompteng/bilig/issues/153)
- [#154: docs(comparison): add headless engine use-case chooser](https://github.com/proompteng/bilig/issues/154)
- [#155: docs(troubleshooting): add formula diagnostic output examples](https://github.com/proompteng/bilig/issues/155)
- [#156: docs(contributor): add first PR description template](https://github.com/proompteng/bilig/issues/156)
- [#158: docs(contributor): add fuzz replay walkthrough](https://github.com/proompteng/bilig/issues/158)
- [#159: docs(troubleshooting): document CSV delimiter autodetection](https://github.com/proompteng/bilig/issues/159)
- [#162: docs(benchmarks): add WorkPaper benchmark reproduction notes](https://github.com/proompteng/bilig/issues/162)
- [#163: docs(agent): add coding-agent workbook automation recipe](https://github.com/proompteng/bilig/issues/163)
- [#193: docs(examples): add append-rows WorkPaper example](https://github.com/proompteng/bilig/issues/193)
- [#194: docs(examples): add changed-cells audit log example](https://github.com/proompteng/bilig/issues/194)
- [#195: docs(examples): add range-to-records export example](https://github.com/proompteng/bilig/issues/195)
- [#196: docs(examples): add validation-error JSON response example](https://github.com/proompteng/bilig/issues/196)
- [#197: docs(headless): add WorkPaper persistence decision table](https://github.com/proompteng/bilig/issues/197)
- [#198: docs(contributor): add first issue local verification matrix](https://github.com/proompteng/bilig/issues/198)
- [#199: test(docs): check headless example README scripts exist](https://github.com/proompteng/bilig/issues/199)

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

Before opening a PR, read [`CONTRIBUTING.md`](../CONTRIBUTING.md) and run the
smallest relevant local check first. If the change touches package behavior,
run `pnpm run ci` before asking for review.
