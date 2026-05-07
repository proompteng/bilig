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

Verified on 2026-05-07 after the Show HN launch and topic refresh:

- GitHub stars: `6`
- GitHub open issues: `75`
- Community profile: `100%`
- Public package: `@bilig/headless`
- GitHub repo description: headless spreadsheet engine for agents and Node
  services
- Topics: `20`, including `spreadsheet`, `spreadsheet-engine`,
  `headless-spreadsheet`, `ai-agents`, `ai-spreadsheet`, `formula-engine`,
  `spreadsheet-api`, `workbook`, `workbook-api`, `excel-automation`,
  `typescript`, and `wasm`
- Public launch assets already present:
  - root `README.md`
  - `packages/headless/README.md`
  - `docs/public-adoption-kit.md`
  - `docs/launch-post-headless-workpaper.md`
  - `docs/why-agents-need-workbook-apis.md`
  - `docs/persisting-formula-backed-workpaper-documents-in-node.md`
  - `docs/what-workpaper-benchmark-proves.md`
  - `docs/building-a-revenue-model-with-headless-workpaper.md`
  - `docs/where-bilig-is-not-excel-compatible-yet.md`
  - `docs/xlsx-corpus-verifier-walkthrough.md`
  - `docs/show-hn-launch-pack.md`
  - `docs/community-launch-pack.md`
  - `docs/dev-to-workbook-apis-post.md`
  - `examples/headless-workpaper/revenue-scenarios.mjs`
  - `docs/starter-issues.md`
  - `docs/assets/github-social-preview.png`
  - GitHub repository social preview configured from
    `docs/assets/github-social-preview.png`
  - `CONTRIBUTING.md`
  - `CODE_OF_CONDUCT.md`
  - `SECURITY.md`
  - `SUPPORT.md`
  - issue and pull request templates
  - `good first issue` and `help wanted` labels

## Execution Log

2026-05-07:

- Published the DEV article:
  <https://dev.to/gregkonush/why-agents-need-workbook-apis-instead-of-spreadsheet-screenshots-3d61>.
- Added Product Hunt thumbnail, gallery images, and short WebM demo to
  [`docs/product-hunt-launch-assets.md`](product-hunt-launch-assets.md).
- Verified GitHub Pages is serving the launch asset page and demo video.
- Attempted the Product Hunt submit flow in Atlas. Posting is blocked by the
  personal-account access gate, so the safe next step is waiting for access
  before scheduling a launch.
- Posted one manual X reply from `@GregKonush` in a high-fit spreadsheet
  automation subthread:
  <https://x.com/tulexaicom/status/2052288937717063977>.
- Re-verified the GitHub mirror, public workflows, and star count after the
  demo commit. Stars remain `5`.
- Refreshed the starter issue queue to `5` open `good first issue` tickets by
  promoting scoped issues `#32` and `#49`, adding starter-scope comments, and
  removing closed issue `#102` from
  [`docs/starter-issues.md`](starter-issues.md).
- Submitted Show HN:
  <https://news.ycombinator.com/item?id=48052832>. Preflight included green
  GitHub workflow pages, public site and repository checks, and a passing
  `pnpm workpaper:smoke:external`. Stars remained `5` at launch.
- Rechecked the Show HN thread through Hacker News and Algolia. The item was
  live with no external comments to answer yet, and public GitHub stars had
  moved to `6`.
- Refreshed GitHub discovery topics by replacing broad `agents` and `headless`
  tags with exact-fit `ai-spreadsheet` and `spreadsheet-api` tags.
- Posted one adapted comment in the `r/github` self-promotion megathread under
  the logged-in maintainer account, linking the npm package, repository, and
  runnable example without asking for votes or stars:
  <https://www.reddit.com/r/github/comments/1jy8rea/promote_your_projects_here_selfpromotion/okhx8b5/>.
- Posted one adapted link submission to `r/coolgithubprojects`, using the
  repository as the primary link and a short feedback-focused body:
  <https://www.reddit.com/r/coolgithubprojects/comments/1t6jo3s/i_built_a_headless_spreadsheet_engine_for_node/>.
- Opened a targeted `jsgrids` directory contribution for `bilig` as a
  headless spreadsheet library:
  <https://github.com/statico/jsgrids/pull/92>. This is a discovery-list
  placement, not a social repost.
- Posted one no-link follow-up reply to Mert Deveci after he suggested
  IronCalc as an adjacent open-source spreadsheet engine:
  <https://x.com/GregKonush/status/2052471748826996909>. The reply
  acknowledged IronCalc as a strong Rust/WASM project and positioned `bilig`
  narrowly around Node/service WorkPaper state, mutation receipts, formula
  readback, and persistence checks for agents.

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
- curated comparison directories can create slower but more durable discovery
  when the project is a real fit and the contribution uses conservative feature
  claims
- content based on actual user questions performs better than generic release
  announcements
- a repo must deserve the traffic before marketing: clear README, demo,
  install path, proof, and contribution path first
- X replies should be manual, low-volume, and tied to the specific post.
  Official X guidance treats duplicated unsolicited replies as spam behavior,
  and its automation rules and developer guidelines prohibit using non-API
  website scripting or keyword searches to spray automated replies.

GitHub's 2026 maintainer guidance adds one important constraint: growth is only
useful if the project can absorb it. AI-assisted contributions can increase
issue and pull-request volume without increasing quality, so the growth loop
needs stronger docs, scoped issues, runnable examples, and fast maintainer
triage before a major launch spike.

The 2025 Hacker News launch-diffusion study is directionally useful for launch
planning. It analyzed 138 AI and LLM-tool repository launches from 2024-2025
and reported average post-HN gains of `121` stars within 24 hours, `189` within
48 hours, and `289` within a week. It also says timing and launch fit matter.
Treat HN as one distribution event in a repeatable proof loop, not as the whole
plan.

The most useful 2026 founder playbooks repeat the same lesson: large star
growth came from a specific, tryable product plus repeated distribution across
HN, Reddit, Product Hunt, X, and community channels. AFFiNE's public case study
claims `1,000` stars in 72 hours and `6,000` in seven days, but the useful
takeaway is not the spike; it is the clean positioning, global distribution,
and immediate user conversations after the traffic arrived.

Official HN guidance makes the same launch constraint concrete: a Show HN should
be something people can try directly, preferably without signups or email gates,
and it should be work the poster personally built and can discuss in the thread.
For `bilig`, that means the launch link should point at the repository or public
docs with the npm-only quickstart visible immediately, not a generic landing
page.

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

- Keep `docs/assets/github-social-preview.png` uploaded as the custom GitHub
  social preview image through repository settings. Verified in GitHub settings
  on 2026-05-07. It is `1280x640`, under `1 MB`, and uses a solid background.
- Regenerate the checked-in preview image with
  `pnpm docs:social-preview:generate` when the package positioning or proof
  points change, then run `pnpm docs:social-preview:check` before sharing fresh
  links.
- Pin or surface the external example in all launch copy:
  `examples/headless-workpaper`.
- Keep the website and README first-run path npm-only, copy-pasteable, and tied
  to persistence/readback output. A cold visitor should see a self-verifying
  npm smoke test before any clone-first path and should not need monorepo
  knowledge to evaluate the package.
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
  - use `docs/launch-post-headless-workpaper.md` as the starting draft
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
  - Product Hunt after the social preview, demo, and example are polished.
    Product Hunt's current help docs say posting requires a personal account,
    new accounts may need a one-week warmup before posting, and the launch flow
    now supports drafts instead of an immediate "launch now" path. Build the
    draft early, add the maintainer as a Maker, and do normal product comments
    before launch day.
  - dev.to or a personal/company blog with the benchmark story
  - use `docs/community-launch-pack.md` for platform-specific drafts and
    anti-spam guardrails before posting outside HN
  - use `docs/dev-to-workbook-apis-post.md` as the first DEV article draft,
    then adapt it in the composer instead of posting a thin repo link
- Contribute to legitimate comparison lists where the project clearly fits:
  - `jsgrids` for JavaScript spreadsheet/data-grid libraries:
    <https://github.com/statico/jsgrids/pull/92>
  - use conservative feature metadata and mark UI-only behavior as false unless
    `@bilig/headless` exposes it directly
  - do not open list PRs that stretch the category fit just to gain backlinks
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

### Continuous X Reply Loop

Goal: earn profile visits from people already discussing spreadsheets, agents,
Excel automation, workbook persistence, and formula reliability.

Operating rule: reply only when the post has a clear technical connection to
`@bilig/headless`. Do not use duplicated templates, engagement bait, automated
mentions, or bare links.

Daily loop:

1. Search X for `excel ai`, `spreadsheet agents`, `workbook api`,
   `google sheets ai`, `formula engine`, and `hyperformula`.
2. Pick at most `2` to `3` posts where a bilig maintainer can add a concrete
   technical point.
3. Reply in lowercase, plain human tone, and lead with the idea rather than the
   repo.
4. Link only when it genuinely helps the conversation; otherwise let the
   profile and previous launch posts carry the repository.
5. If the same question appears twice, convert the answer into a doc, example,
   issue, or benchmark note before replying again.

Tone rule:

- lowercase, direct, and slightly informal
- one concrete idea per reply
- no "check out my repo" unless the link answers the exact post
- no duplicated reply body across multiple posts
- no pretending to be neutral when the reply is from the maintainer

Sam Altman-style lower-case tone works because it reads casual and compressed:
short sentences, minimal punctuation, and little launch-copy polish. Use that
shape, but keep the substance specific to `bilig`; do not impersonate anyone or
turn replies into vague hype.

Good reply shapes:

> the hard part is not generating a formula once. it is preserving workbook
> state, formulas, provenance, and writeback so an agent can be checked after it
> acts.

> this is where headless workbook apis matter. screenshots are fine for demos,
> but agents need ranges, formulas, structural edits, persistence, and readback
> tests.

> the benchmark story only matters if it is auditable. for bilig i’m keeping the
> artifact, verify command, and p95 caveat in public instead of turning it into
> a vague "faster than x" claim.

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

1. Post one lower-case X update for the compatibility-boundaries article and
   XLSX verifier walkthrough.
2. Manually reply to at most `2` high-fit X posts per day. Start with the idea;
   add a link only when it materially helps the thread. Use
   [`docs/x-reply-growth-playbook.md`](x-reply-growth-playbook.md) for the
   reply budget, tone, examples, and platform-rule boundaries.
3. Create a Product Hunt draft once the launch image and demo are final. Do not
   schedule the launch until the personal account is eligible to post and the
   first maker comment is ready.
4. Respond to serious Hacker News and GitHub Discussion comments within the same
   day, then convert repeated feedback into docs or examples.
   For a concrete Hacker News launch checklist, use
   [`docs/show-hn-launch-pack.md`](show-hn-launch-pack.md).
5. Keep shipping compact formula-edge fixture articles that follow the XLOOKUP,
   SUMIFS, and GROUPBY walkthrough pattern: one real formula family, exact
   fixture inputs, expected result or spill output, and verifier command.
6. Share the HyperFormula comparison only in threads where someone is already
   evaluating headless spreadsheet engines; keep the caveats visible and do not
   frame it as a blanket replacement claim.
7. Track GitHub stars, npm downloads, GitHub traffic referrers, and issue
   quality every week.
8. Add a Star History chart only after there is enough organic movement for the
   graph to communicate momentum instead of early-stage emptiness.

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
- GitHub Blog, open source in 2026:
  <https://github.blog/open-source/maintainers/what-to-expect-for-open-source-in-2026/>
- Hacker News, Show HN guidelines:
  <https://news.ycombinator.com/showhn.html>
- Hacker News, site guidelines:
  <https://news.ycombinator.com/newsguidelines.html>
- X rules and best practices for replies:
  <https://help.x.com/en/rules-and-policies/x-rules-and-best-practices>
- X automation rules:
  <https://help.x.com/en/rules-and-policies/x-automation>
- X developer guidelines:
  <https://docs.x.com/developer-guidelines>
- Product Hunt launch guide:
  <https://www.producthunt.com/launch/>
- Product Hunt getting started:
  <https://help.producthunt.com/en/articles/2305333-getting-started>
- Product Hunt posting access:
  <https://help.producthunt.com/en/articles/481909-how-can-i-get-access-to-post>
- Product Hunt launch drafts:
  <https://help.producthunt.com/en/articles/9823193-where-did-launch-now-go>
- Product Hunt hunter and maker roles:
  <https://help.producthunt.com/en/articles/10082986-hunter-vs-makers-and-how-to-change-them>
- Lago first-1000-stars case study:
  <https://getlago.com/blog/how-we-got-our-first-1000-github-stars>
- Indie Hackers first-1000-stars writeup:
  <https://www.indiehackers.com/post/0-to-1000-github-stars-for-your-open-source-dev-tools-db2efb62f1>
- "What's in a GitHub Star?", arXiv:
  <https://arxiv.org/abs/1811.07643>
- Hacker News launch diffusion study, arXiv:
  <https://arxiv.org/abs/2511.04453>
- AFFiNE open-source growth case study:
  <https://gingiris.github.io/growth-tools/blog/2026/03/07/i-led-affine-from-0-to-60k-github-stars-here-are-my-open-source-growth-playbooks/>
- Fake-star campaign study:
  <https://cmustrudel.github.io/papers/icse2026fakestars.pdf>
- Fortune coverage of Sam Altman's lower-case social style:
  <https://fortune.com/2026/01/29/openai-ceo-sam-altman-types-all-lowercase-like-gen-z-but-could-be-career-chatgpt-boss/>
- IronCalc official site, used for the adjacent-engine positioning reply:
  <https://www.ironcalc.com/>
