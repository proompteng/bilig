# Bilig Community Launch Pack

This is the working growth text for moving `proompteng/bilig` from a tiny
public footprint to `1,000` legitimate GitHub stars without fake accounts,
paid stars, spam, or inflated claims.

Baseline verified on May 8, 2026:

- GitHub repo: <https://github.com/proompteng/bilig>
- Current GitHub surface: public repo, `8` stars, `1` fork, `39` open issues
- npm package: <https://www.npmjs.com/package/@bilig/headless>
- Current npm surface: `@bilig/headless@0.10.62`, MIT, TypeScript types, npm
  API downloads: `2,399` last week and `16,491` last month
- Existing Hacker News launch:
  <https://news.ycombinator.com/item?id=48045427>, `2` points and `1`
  self-comment after roughly one day

## Goal Text

Grow `proompteng/bilig` from `8` to `1,000` legitimate GitHub stars by making
`@bilig/headless` the obvious headless spreadsheet engine for Node services,
coding agents, and workbook automation.

The star goal is a proxy for useful adoption, not the product. The real target
is a repeatable external-user loop:

1. A developer sees a concrete workbook-automation problem.
2. They find a short `@bilig/headless` example that matches the problem.
3. They run it without cloning the monorepo.
4. They see formula or persistence proof, not a toy printout.
5. They bookmark the repo, open an issue, or ask for a missing workflow.

## Research Inputs

- GitHub's Open Source Guide says promotion starts with a clear message, a
  single home URL, going where the exact audience already is, asking for
  feedback, and helping people before asking for attention:
  <https://opensource.guide/finding-users/>
- GitHub Docs says repository topics help people find projects to use and
  contribute to, and topic choices should describe purpose, subject area,
  community, and language:
  <https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/classifying-your-repository-with-topics>
- Current HN evidence says the generic Show HN post did not break out. The next
  move should be targeted, evidence-led replies and adjacent problem-solving,
  not reposting the same launch pitch.

## Growth Hacks That Are Allowed

These are aggressive but legitimate. They optimize distribution around a real
library and real proof.

### 1. Piggyback On Existing Search Intent

Create narrowly titled pages and examples for exact searches people already
make:

- `headless spreadsheet engine for Node`
- `HyperFormula alternative`
- `spreadsheet engine for agents`
- `evaluate Excel formulas in Node`
- `persist formula backed workbook JSON`
- `server side spreadsheet automation`
- `XLSX formula cache verifier`

Every page should have one runnable command, one expected output block, and one
link back to the repo star URL. Do not ask for stars before the example proves
something.

### 2. Turn Benchmarks Into A Distribution Loop

The `46/46` WorkPaper mean-win claim should become small, quotable artifacts:

- one chart image for X, Bluesky, LinkedIn, and README embeds
- one short post explaining the p95 caveat honestly
- one "run this benchmark locally" post
- one "what would make this benchmark unfair?" discussion prompt

The hack is to invite skeptics into the measurement instead of only announcing
a win. Spreadsheet-engine users trust reproducible evidence more than launch
copy.

### 3. Mine Competitor And Adjacent Issues For Real Problems

Search public issues and discussions in adjacent tools for unsolved workflows:
HyperFormula, ExcelJS, SheetJS, formula.js, Handsontable, AG Grid, agent
frameworks, LangChain, LlamaIndex, and Vercel AI SDK examples.

When someone has a problem `@bilig/headless` genuinely solves, answer with:

1. the relevant diagnosis,
2. a tiny runnable snippet,
3. a disclosure that you maintain `bilig`,
4. the repo link only after the solution.

Do not necro-bump old issues unless the answer is materially useful.

### 4. Package Every Example As A Shareable Gist-Sized Artifact

For each high-intent workflow, keep the example small enough to paste:

- revenue model with formulas and restore
- JSON records to WorkPaper
- CSV-shaped input to WorkPaper
- agent writeback and verification
- named-expression update
- formula cache mismatch verifier
- unsupported formula troubleshooting

Each example needs a title that names the user problem, not the internal
package architecture.

### 5. Use Feedback Loops Instead Of Launch Blasts

Post asks that are easy to answer:

- "What spreadsheet automation workflow would make this worth trying?"
- "Which Excel formulas would block you from replacing a headless engine?"
- "Does this WorkPaper API feel too custom, or boring enough to script?"
- "What benchmark row looks suspicious?"

The point is to make experts correct the project in public. Useful corrections
create better issues, better docs, and more credible follow-up posts.

### 6. Convert npm Demand Into GitHub Stars

The npm page already has visible downloads while GitHub stars are low. Add or
keep star/bookmark CTAs only after value proof:

- after the quickstart expected output
- after the benchmark evidence link
- after the external-consumer example
- after troubleshooting recipes

Good CTA:

> If this saves you a workbook-automation spike, star the repo so the package is
> easier to find later: <https://github.com/proompteng/bilig>

Avoid generic "please star" copy at the top of technical docs.

### 7. Build A Contributor Discovery Flywheel

Keep at least three current starter issues alive. Each starter issue should
ship one new shareable example or compatibility reduction. That turns
contributor onboarding into public marketing:

- new issue names a real user workflow
- PR adds a runnable example
- release notes mention the example
- docs link it from package README
- social post explains the workflow

## Channel Copy

### X / Bluesky

```text
I am building @bilig/headless: a TypeScript WorkPaper engine for Node services
and coding agents that need spreadsheet semantics without opening a browser
grid.

It does formula evaluation, structural edits, range reads, JSON persistence,
and post-write verification.

The current benchmark artifact shows 46/46 mean wins against HyperFormula-style
headless workloads, with the p95 caveat documented instead of hidden.

Repo: https://github.com/proompteng/bilig
npm: https://www.npmjs.com/package/@bilig/headless

Question for people who automate spreadsheets: what workflow would you need to
see before trying a new headless workbook engine?
```

### Hacker News Follow-Up Comment

Do not repost the same Show HN. Add a follow-up only when there is new proof or
a concrete question:

```text
Follow-up after the Show HN: the most useful signal I am looking for is not
"nice project", but which headless spreadsheet workflows are missing.

The current package is aimed at Node services and agents that need formula
evaluation, structural edits, range reads, JSON persistence, and post-write
verification without opening a browser grid.

If you have used HyperFormula, ExcelJS, or SheetJS for this kind of automation:
what would block you from trying a WorkPaper-style API?
```

### Reddit / Community Post

```text
I am looking for feedback from people who automate spreadsheets from Node.

Project: @bilig/headless
Repo: https://github.com/proompteng/bilig

It is a TypeScript headless workbook engine for services and coding agents:
formula evaluation, structural row/column edits, range reads, JSON persistence,
and verification after writes.

The docs include a quick npm-only smoke test and benchmark evidence against
HyperFormula-style workloads. It is early infrastructure, not a complete Excel
clone.

What I am trying to learn: which workflows or formulas would make or break this
for real use?
```

### GitHub Discussion Prompt

```text
Which headless spreadsheet workflow should @bilig/headless prove next?

Current supported path:

- build a workbook from data
- evaluate formulas
- apply structural edits
- read ranges
- persist and restore JSON
- verify agent writes

Candidates:

- more XLSX formula-cache corpus reductions
- more finance-model examples
- a LangChain / Vercel AI SDK tool-calling example
- broader formula compatibility fixtures
- a migration guide from HyperFormula

If you have a real workflow, please describe the input shape, formulas, and the
output you need to trust.
```

## Weekly Operating Cadence

Every week until `1,000` stars:

1. Ship one public-facing example, comparison, or fixture writeup.
2. Answer five relevant external threads where the answer helps even if nobody
   clicks through.
3. Open or refresh three starter issues tied to shareable examples.
4. Publish one short evidence-led post.
5. Record stars, npm downloads, referrers, issue sources, discussion activity,
   and external PRs.
6. If a channel produces no useful feedback after two attempts, change the
   angle before posting there again.

Capture the baseline before and after each distribution push:

```sh
pnpm community:growth:snapshot
```

Set `GITHUB_TOKEN` or `GH_TOKEN` with repository traffic access to include
views, clones, popular referrers, and popular paths. Without a token, the
snapshot still records public GitHub stars, forks, open issues, package version,
and npm download windows.

## Anti-Spam Rules

- Do not buy stars, trade stars, automate follows, or ask unrelated audiences
  for stars.
- Do not claim full Excel compatibility.
- Do not claim blanket performance wins; keep the p95 caveat visible.
- Do not post the same launch copy across communities.
- Do not hide that the maintainer is recommending their own project.
- Do not contact people through private channels unless there is a specific,
  relevant reason and the message would still be useful without a repo link.

## Success Metrics

Primary:

- `1,000` legitimate GitHub stars on <https://github.com/proompteng/bilig>

Leading indicators:

- npm weekly downloads
- GitHub traffic and referrers
- external issues with real workbook examples
- external PRs
- GitHub Discussions from non-maintainers
- stars per public post
- stars per new runnable example
- conversion from npm README visits to GitHub stars

If stars rise but no external users ask questions or open issues, the growth is
weak. If external users bring real workflows and compatibility gaps, the flywheel
is working even before the star count catches up.
