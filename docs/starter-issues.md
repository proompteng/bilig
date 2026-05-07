# Starter Issues

This page is the stable contributor on-ramp for small public `bilig` tasks. It
links directly to current scoped issues instead of relying on GitHub issue
search indexing.

Current starter tickets as of May 7, 2026:

- [`#19 docs(headless): add a Node service recipe`](https://github.com/proompteng/bilig/issues/19)
- [`#20 docs(headless): add a CSV-shaped input recipe`](https://github.com/proompteng/bilig/issues/20)
- [`#21 docs(agents): add a WorkPaper tool-calling recipe`](https://github.com/proompteng/bilig/issues/21)
- [`#22 docs(benchmarks): add a local benchmark walkthrough`](https://github.com/proompteng/bilig/issues/22)
- [`#23 docs(headless): add an unsupported-formula troubleshooting recipe`](https://github.com/proompteng/bilig/issues/23)
- [`#29 workbook import dispatcher rejects MIME types with parameters or different case`](https://github.com/proompteng/bilig/issues/29)
- [`#67 XLSX import/export drops worksheet tab colors`](https://github.com/proompteng/bilig/issues/67)

Step-up tickets with a small but production-facing implementation surface:

- [`#63 XLSX import hangs on corrupt zip-backed workbook instead of rejecting`](https://github.com/proompteng/bilig/issues/63)

Useful filters:

- [`good first issue`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22good%20first%20issue%22)
- [`area: import-export`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22area%3A%20import-export%22)
- [`area: formula`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22area%3A%20formula%22)
- [`needs reduced fixture`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22needs%20reduced%20fixture%22)

## What Makes A Good Starter Patch

- Keep the change small enough to review in one sitting.
- Use public `@bilig/headless` exports in examples.
- Prefer runnable recipes over abstract prose.
- Link back to the relevant package README or benchmark evidence note.
- Include the focused validation command in the issue or PR description.

Before opening a PR, read [`CONTRIBUTING.md`](../CONTRIBUTING.md) and run the
smallest relevant local check first. If the change touches package behavior,
run `pnpm run ci` before asking for review.
