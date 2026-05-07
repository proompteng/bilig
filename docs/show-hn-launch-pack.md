# Show HN Launch Pack

Status: ready-to-review launch checklist for a `bilig` Hacker News submission.

Use this only when a maintainer is available to answer comments for the next few
hours. Do not ask friends, coworkers, followers, or other accounts to upvote or
comment. Treat the launch as feedback collection first and star growth second.

Official Show HN fit:

- The project is something people can try: the `@bilig/headless` package is on
  npm and the repository has a runnable Node example.
- The submission should point at the repository or package, not a blog post,
  because Show HN is for things people can use.
- The copy should be factual, technical, and direct. Avoid launch-day marketing
  language.

## Submission

Title:

```text
Show HN: Bilig - a headless spreadsheet engine for Node services and agents
```

URL:

```text
https://github.com/proompteng/bilig
```

Use the GitHub repository URL because it exposes code, examples, issues,
benchmarks, caveats, and contribution paths in one place.

## First Comment Draft

Post this as the maintainer's own comment only after reviewing it in your own
voice:

```text
i built bilig because a lot of useful business logic still lives in spreadsheet-shaped models, but automation usually either screen-scrapes a browser grid or rewrites formulas in ad hoc code.

the current public package is @bilig/headless. it runs workbook creation, formula evaluation, structural edits, persistence round trips, and readback from node without opening a browser ui.

quick try path:

npm install @bilig/headless

or run the external example:

git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm start

there is also an agent writeback proof in the same example:

npm run agent:verify

the repo also includes benchmark evidence against hyperformula-style workloads. the current checked-in claim is 46/46 mean wins on scorecard-eligible comparable workloads, with the p95 caveat left attached. it is not a finished excel clone; the compatibility boundaries doc calls out macro, formula, xlsx, and ui gaps.

i'm especially interested in feedback from people who have built spreadsheet-backed services, formula engines, xlsx import/export, or agent workflows that need reliable workbook state.
```

## Links To Have Open

- repository: <https://github.com/proompteng/bilig>
- npm package: <https://www.npmjs.com/package/@bilig/headless>
- runnable example:
  <https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper>
- package README:
  <https://github.com/proompteng/bilig/tree/main/packages/headless#readme>
- benchmark explainer:
  <https://github.com/proompteng/bilig/blob/main/docs/what-workpaper-benchmark-proves.md>
- compatibility boundaries:
  <https://github.com/proompteng/bilig/blob/main/docs/where-bilig-is-not-excel-compatible-yet.md>
- starter issues:
  <https://github.com/proompteng/bilig/blob/main/docs/starter-issues.md>

## Preflight

Before posting:

1. Confirm the repository star count, default branch, and latest pushed SHA.
2. Confirm GitHub Actions are green for `main`.
3. Confirm the package README still shows the install command and quickstart.
4. Run the external example locally or confirm the latest smoke command was
   green.
5. Re-read the benchmark and compatibility caveats so replies do not overclaim.
6. Be available to answer comments for at least `2` hours.

Useful commands:

```sh
gh api repos/proompteng/bilig --jq '{stars:.stargazers_count,pushed_at:.pushed_at,default_branch:.default_branch}'
gh run list --repo proompteng/bilig --limit 5 --json workflowName,headSha,status,conclusion,url
pnpm workpaper:smoke:external
```

## Reply Principles

- Answer technical questions directly.
- Admit gaps quickly and link to the issue or compatibility note.
- Convert repeated objections into issues or docs.
- Keep replies short enough for the thread to stay readable.
- Do not use generated comments unedited; HN is a human discussion space.
- Do not ask people to star the repo.

## Likely Questions

### Why not HyperFormula?

`@bilig/headless` is aimed at WorkPaper-style service and agent workflows:
formula-backed business logic, structural edits, persistence, validation, and
engine-level evidence in one TypeScript repo. HyperFormula is the obvious
comparison point, so the benchmark docs keep that comparison explicit and
caveated.

### Is this Excel compatible?

Not fully. The honest claim is fixture- and corpus-scoped compatibility. The
compatibility boundaries note names the current macro, formula, XLSX corpus, and
UI gaps. The formula fixture notes show how specific claims should map to
fixtures and verifier commands.

### Can I use it without the browser app?

Yes. The `@bilig/headless` package runs from Node. The external example builds a
small revenue workbook, evaluates formulas, applies an agent-style edit,
serializes the document, restores it, and verifies summary values.

### What feedback is most useful?

Useful feedback includes:

- real workbook automation use cases
- formula parity bugs with reduced expected values
- import/export gaps with small fixture files
- benchmark workloads that match production spreadsheets
- API friction in the headless package

## Follow-Up

Within `24` hours of posting:

1. Save the HN URL in the launch discussion.
2. Create issues for concrete bugs or compatibility gaps.
3. Add docs for repeated questions.
4. Post one follow-up in the GitHub discussion with links to resulting artifacts.
5. Update `docs/github-stars-growth-plan.md` with the star count, traffic notes,
   and what changed.

## Sources

- Official Show HN guidelines: <https://news.ycombinator.com/showhn.html>
- Hacker News guidelines: <https://news.ycombinator.com/newsguidelines.html>
