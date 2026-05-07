# GitHub Stars Growth Plan

Status: researched growth plan for taking `proompteng/bilig` from early public
visibility to the first `1000` legitimate GitHub stars.

Research date: 2026-05-07.

## Objective

Reach `1000` real GitHub stars for <https://github.com/proompteng/bilig> by
turning `@bilig/headless` into an easy-to-evaluate, easy-to-share developer
tool. Stars should come from developers who understand what the project does,
not from paid, automated, reciprocal, or misleading campaigns.

## Current Public State

Verified on 2026-05-07:

- GitHub stars: `2`
- Community profile: `100%`
- Public package: `@bilig/headless`
- GitHub repo description: headless spreadsheet engine and local-first workbook
  runtime for Node services, AI agents, collaborative sync, and
  WASM-accelerated formulas
- Topics: `19`, including `spreadsheet`, `spreadsheet-engine`,
  `headless-spreadsheet`, `ai-agents`, `formula-engine`, `workbook`,
  `typescript`, and `wasm`
- Public launch assets already present:
  - root `README.md`
  - `packages/headless/README.md`
  - `docs/public-adoption-kit.md`
  - `docs/assets/github-social-preview.png`
  - `CONTRIBUTING.md`
  - `CODE_OF_CONDUCT.md`
  - `SECURITY.md`
  - `SUPPORT.md`
  - issue and pull request templates
  - `good first issue` and `help wanted` labels

## Research Findings

GitHub stars are a bookmark and appreciation signal, and GitHub explicitly says
many repository rankings and Explore surfaces depend on stars. The practical
meaning is simple: stars are not a usage metric, but they amplify discovery once
real developers begin saving the project.

GitHub's official discoverability levers are already aligned with the work done
so far:

- the README is usually the first thing a visitor sees and should explain why
  the project is useful, what people can do with it, and how to use it
- topics classify the repository by purpose, community, language, and subject
  so people can find it through topic pages and searches
- community health files reduce friction for contributors and establish
  expectations
- a social preview image makes shared repository links identifiable across
  social platforms

Open source community guidance is consistent with that: documentation is the
main conversion surface, responsiveness matters, most contributors are casual,
and early projects need to meet users in the places where they already talk.

Case studies and founder writeups are also consistent:

- Hacker News can create major spikes, but it is hit-or-miss and should be part
  of repeated content distribution rather than a one-shot launch
- Reddit, Product Hunt, Twitter/X, Indie Hackers, dev.to, community Slacks, and
  niche forums can work when the post speaks to a specific developer pain
- content based on actual user questions performs better than generic release
  announcements
- a repo must deserve the traffic before marketing: clear README, demo,
  install path, proof, and contribution path first

Do not buy stars or run fake-star exchanges. Recent research found increasing
fake-star campaigns, and GitHub Trending appears to filter out most superficial
fake-star activity. Fake growth creates trust and supply-chain risk for a
developer-tool project.

## Positioning

Lead with the smallest valuable product:

> `@bilig/headless` is a headless spreadsheet engine for agents and Node
> services.

Second sentence:

> It gives you formulas, structural edits, persistence, validation, and
> benchmark evidence without opening a browser grid.

Primary audience:

- agent developers who need a reliable workbook API instead of screen scraping
- Node service developers who need formulas and workbook persistence
- spreadsheet-engine developers comparing formula and structural-edit behavior
- TypeScript developers interested in local-first workbook infrastructure

Avoid leading with the full monorepo architecture. The browser shell, renderer,
sync, protocol, and WASM work are supporting proof, but the adoption wedge is
`@bilig/headless`.

## Star Funnel

### 0 To 50 Stars

Goal: make every visitor understand the project in under one minute and make the
first star feel like a useful bookmark.

Actions:

- Upload `docs/assets/github-social-preview.png` as the custom GitHub social
  preview image through repository settings. It is `1280x640`, under `1 MB`,
  and uses a solid background.
- Pin or surface the external example in all launch copy:
  `examples/headless-workpaper`.
- Add a Star History chart only after the repo has enough organic movement to
  avoid making `2` stars the visual focus.
- Seed `3` to `5` real `good first issue` items:
  - add a formula parity fixture
  - improve a WorkPaper recipe
  - add an example for named expressions
  - add a persistence validation example
  - improve benchmark evidence wording
- Ask existing users, collaborators, and adjacent project maintainers for
  feedback, not stars. Stars should be the natural bookmark after they evaluate.

### 50 To 250 Stars

Goal: earn attention from targeted developer communities.

Actions:

- Publish one technical launch post:
  - title angle: `A headless spreadsheet engine for AI agents and Node services`
  - include the npm install path, the maintained example, and the benchmark
    evidence
  - keep benchmark claims precise: `46/46 mean wins`, with the p95 caveat
- Submit adapted posts to:
  - Hacker News: `Show HN: bilig - a headless spreadsheet engine for agents`
  - Reddit communities where the angle fits:
    `r/javascript`, `r/typescript`, `r/node`, `r/opensource`,
    `r/coolgithubprojects`, `r/SideProject`, and only `r/programming` when the
    post is deeply technical
  - Product Hunt after the social preview, demo, and example are polished
  - dev.to or a personal/company blog with the benchmark story
- Track which posts create GitHub visitors, npm downloads, issues, and stars.
  Do not judge by likes alone.
- Answer comments quickly. If the same question appears twice, convert it into
  docs or an example.

### 250 To 1000 Stars

Goal: turn a one-time launch into a repeatable content and proof loop.

Actions:

- Publish one focused proof article per week for `6` to `8` weeks:
  - `Why agents need workbook APIs, not spreadsheet screenshots`
  - `Persisting formula-backed WorkPaper documents in Node`
  - `What our HyperFormula-style benchmark does and does not prove`
  - `How structural edits affect spreadsheet dependency graphs`
  - `Building a revenue model with @bilig/headless`
  - `Where bilig is not Excel-compatible yet`
- Create one small runnable example for each article.
- Keep monthly GitHub releases readable by humans, not just generated changelog
  entries.
- Watch mentions with alerts or search, then join conversations where they
  already happen instead of forcing everyone into a new community.
- Open GitHub Discussions only when there is enough inbound discussion to keep
  it alive.

## Launch Copy

Use this for the first broad post:

> I built `@bilig/headless`, a TypeScript spreadsheet engine for agents and Node
> services. It runs formulas, structural edits, persistence round trips, and
> validation without a browser grid. The repo includes a runnable npm example
> and checked-in benchmark evidence against HyperFormula-style workloads.

Use this for a benchmark-focused post:

> The current WorkPaper benchmark artifact records `46/46` mean wins on
> scorecard-eligible comparable workloads against HyperFormula-style workloads
> (`38/38` public, `8/8` holdout). The p95 story has nuance, so the evidence doc
> spells out exactly what is measured and what is not.

Use this for contributor outreach:

> If you like spreadsheet engines, formula semantics, local-first software, or
> agent tools, `bilig` has useful first contributions: formula parity fixtures,
> WorkPaper examples, benchmark scenarios, accessibility fixes, and docs that
> turn architecture notes into runnable code.

## Metrics

Track weekly:

- GitHub stars
- GitHub unique visitors and referring sites
- npm downloads for `@bilig/headless`
- external example successful installs or issues
- issue count and issue response time
- contributor count
- posts shipped
- questions converted into docs/examples

Target pace:

- week 1: `10` to `50` stars from direct network and first launch post
- weeks 2-4: `50` to `250` stars from targeted communities and examples
- weeks 5-10: `250` to `1000` stars from repeated proof posts, feedback loops,
  and community resharing

The target pace is not guaranteed. Treat misses as signal about positioning,
audience, or product friction.

## Immediate Next Actions

1. Upload `docs/assets/github-social-preview.png` as the GitHub social preview
   image.
2. Create `3` to `5` scoped public `good first issue` tickets.
3. Publish the first launch post using the copy above.
4. Share the post in one high-fit channel at a time and respond to every serious
   comment.
5. Convert repeated questions into README, package README, or example updates.
6. After the first organic star movement, add a Star History chart near the end
   of the README.

## Sources

- GitHub Docs, stars and Explore rankings:
  <https://docs.github.com/en/enterprise-server@3.18/get-started/exploring-projects-on-github/saving-repositories-with-stars>
- GitHub Docs, README purpose:
  <https://docs.github.com/articles/about-readmes/>
- GitHub Docs, repository topics:
  <https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/classifying-your-repository-with-topics>
- GitHub Docs, social preview:
  <https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/customizing-your-repositorys-social-media-preview>
- GitHub Docs, healthy contributions:
  <https://docs.github.com/en/communities/setting-up-your-project-for-healthy-contributions>
- Open Source Guides, building community:
  <https://opensource.guide/building-community/>
- GitHub Blog, building an open source community:
  <https://github.blog/open-source/maintainers/four-steps-toward-building-an-open-source-community/>
- Lago first-1000-stars case study:
  <https://getlago.com/blog/how-we-got-our-first-1000-github-stars>
- Indie Hackers first-1000-stars writeup:
  <https://www.indiehackers.com/post/0-to-1000-github-stars-for-your-open-source-dev-tools-db2efb62f1>
- "What's in a GitHub Star?", arXiv:
  <https://arxiv.org/abs/1811.07643>
- Hacker News launch diffusion study, arXiv:
  <https://arxiv.org/abs/2511.04453>
- Fake-star campaign study:
  <https://cmustrudel.github.io/papers/icse2026fakestars.pdf>
