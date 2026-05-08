# Starter Issues

This page is the stable contributor on-ramp for small public `bilig` tasks. It
links directly to current scoped issues instead of relying on GitHub issue
search indexing.

Current starter tickets as of May 8, 2026:

- [#134: docs(examples): add XLSX formula-cache roundtrip example](https://github.com/proompteng/bilig/issues/134)
- [#138: docs(examples): add named-expression change WorkPaper example](https://github.com/proompteng/bilig/issues/138)
- [#141: docs(examples): add HTTP JSON summary WorkPaper example](https://github.com/proompteng/bilig/issues/141)
- [#142: docs(examples): add JSON file input WorkPaper example](https://github.com/proompteng/bilig/issues/142)
- [#143: docs(examples): add markdown report WorkPaper output example](https://github.com/proompteng/bilig/issues/143)
- [#144: docs(examples): add formula diagnostics WorkPaper example](https://github.com/proompteng/bilig/issues/144)
- [#145: docs(examples): add snapshot diff WorkPaper example](https://github.com/proompteng/bilig/issues/145)
- [#146: docs(readme): add first-time contributor command checklist](https://github.com/proompteng/bilig/issues/146)

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
