# Starter Issues

This page is the stable contributor on-ramp for small public `bilig` tasks. It
links directly to current scoped issues instead of relying on GitHub issue
search indexing.

Current starter tickets as of May 7, 2026:

No scoped starter tickets are currently open. Add a new one when the expected
patch is small, testable, and safe for a first-time contributor.

The list intentionally excludes closed issues and broad corpus/parity epics. Add
new starter tickets only when the expected patch can stay small, has a clear
acceptance test, and does not require understanding the whole workbook runtime.

Useful filters:

- [`good first issue`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22good%20first%20issue%22)
- [`area: import-export`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22area%3A%20import-export%22)
- [`area: formula`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22area%3A%20formula%22)
- [`needs reduced fixture`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22needs%20reduced%20fixture%22)

GitHub surfaces issues labeled `good first issue` in contributor discovery
paths, per
[GitHub's label guidance](https://docs.github.com/articles/helping-new-contributors-find-your-project-with-labels),
so starter tickets should stay genuinely scoped and current. Do not use that
label for cross-cutting formula, import/export, or runtime changes that require
broad architectural context.

## What Makes A Good Starter Patch

- Keep the change small enough to review in one sitting.
- Use public `@bilig/headless` exports in examples.
- Prefer runnable recipes over abstract prose.
- Link back to the relevant package README or benchmark evidence note.
- Include the focused validation command in the issue or PR description.

Before opening a PR, read [`CONTRIBUTING.md`](../CONTRIBUTING.md) and run the
smallest relevant local check first. If the change touches package behavior,
run `pnpm run ci` before asking for review.
