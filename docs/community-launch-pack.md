# Community Launch Pack

Status: ready-to-adapt community outreach pack for `bilig`.

Use this after the Show HN pack is reviewed or when a maintainer wants a slower
community-by-community distribution path. The rule is the same across every
channel: adapt to the specific community, disclose the maintainer relationship,
and do not ask for votes, stars, or artificial engagement.

Current launch state:

- Show HN is live:
  <https://news.ycombinator.com/item?id=48052832>.
- One adapted `r/github` self-promotion megathread comment is live:
  <https://www.reddit.com/r/github/comments/1jy8rea/promote_your_projects_here_selfpromotion/okhx8b5/>.
- Do not repost the same launch angle while that thread is active.
- Use new community posts only when there is time to answer replies and adapt
  the copy to the specific community.

## Operating Rules

- Post manually.
- Post in one community at a time.
- Do not paste the same text across communities.
- Do not use multiple accounts.
- Do not ask friends or followers to upvote.
- Do not frame a launch post as a fake neutral discovery.
- Reply to comments before posting anywhere else.
- Convert repeated objections into issues, docs, or examples.

The goal is to earn useful technical feedback and legitimate bookmarks from
developers who care about spreadsheet automation, formula engines, Node services,
local-first workbooks, and agent tooling.

## Reddit

Use Reddit only when the subreddit rules explicitly allow projects, open-source
tools, Show-off posts, feedback posts, or link submissions. Reddit's spam policy
defines spam broadly as repeated or unsolicited actions that harm communities,
including repetitive mass posting and off-topic link sharing.

Do:

- Read the target subreddit rules before drafting.
- Search the subreddit for recent "spreadsheet", "excel", "formula engine",
  "node", and "open source" threads.
- Prefer a text post with the technical context first and the link later.
- Message moderators first when the rules are ambiguous.
- Post only once, then participate in replies.

Do not:

- Drop the GitHub link across multiple subreddits in one day.
- Reuse the same title and body across subreddits.
- Use a brand-new account whose only history is promotion.
- Ask for stars, upvotes, or "support".

Candidate subreddits to evaluate:

- `r/opensource`: only if the post is framed around open-source implementation,
  contribution paths, and feedback.
- `r/typescript`: only if the post emphasizes TypeScript package design and
  engine architecture.
- `r/node`: only if the post emphasizes the Node service package and runnable
  example.
- `r/javascript`: only if the post is technical enough for the community and
  not just a release announcement.
- `r/coolgithubprojects`: likely the most direct GitHub-project fit, but still
  check current rules first.

Draft for a permissive open-source or GitHub-project subreddit:

```text
I built a headless spreadsheet engine for Node services and agents

I maintain an open-source TypeScript spreadsheet engine called bilig. The public wedge is @bilig/headless: a Node package for workbook creation, formula evaluation, structural edits, persistence round trips, and readback without opening a browser grid.

The problem I am trying to solve is that a lot of business logic still lives in spreadsheet-shaped models, but service automation and coding agents usually either screen-scrape Excel/Sheets or rewrite formulas in ad hoc code.

Try path:

npm install @bilig/headless

or:

git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm start

Repo: https://github.com/proompteng/bilig

The repo includes benchmark evidence and compatibility caveats. It is not a finished Excel clone. The most useful feedback would be API friction, missing formula semantics, import/export expectations, or real workbook automation cases you would want reduced into fixtures.

Feedback thread: https://github.com/proompteng/bilig/discussions/115
```

## DEV Community

DEV's terms say posts should be on-topic, high quality, and not primarily for
promotion or backlink creation. A DEV post should therefore be a real technical
article, not a thin link wrapper.

Published post:

```text
https://dev.to/gregkonush/why-agents-need-workbook-apis-instead-of-spreadsheet-screenshots-3d61
```

Repo source mirror:

```text
docs/dev-to-workbook-apis-post.md
```

It includes DEV-style front matter, a canonical URL, tags, the current social
preview image, a concrete Node example, benchmark caveats, and feedback prompts.
Use that draft as the source of truth, then edit in the DEV composer for
community-specific polish.

Article title:

```text
Why agents need workbook APIs instead of spreadsheet screenshots
```

Tags to consider:

```text
typescript node opensource ai
```

Article outline:

1. The automation problem: business logic still lives in spreadsheet-shaped
   models.
2. Why browser-grid automation is brittle for agents.
3. What a headless workbook API needs: formulas, structural edits, persistence,
   validation, range reads, and readback.
4. A compact `@bilig/headless` Node example.
5. What is proven by the current benchmark and what is not.
6. Compatibility caveats and useful contribution areas.

Opening draft:

```text
Spreadsheets are still one of the most common ways teams encode business logic. That creates an awkward problem for automation: the logic is structured like a workbook, but most programmatic workflows either drive a browser grid or rewrite the model in application code.

That is especially fragile for coding agents. A screenshot can show a grid, but it cannot give an agent a stable contract for formulas, structural edits, persistence, validation, or post-write readback.

I maintain an open-source TypeScript project called bilig. The public package is @bilig/headless, a Node-facing WorkPaper API for programmatic workbook automation. This post explains the API shape and the current caveats, not a claim that the project is a finished Excel clone.
```

Close with:

```text
The repo is here if you want to inspect the package, examples, benchmark evidence, or starter issues:

https://github.com/proompteng/bilig
```

## Lobsters

Lobsters is a computing-focused discussion community with explicit
self-promotion guidance. It is not a good first launch surface unless the
maintainer is already an invited participant and the submission is a technical
article rather than a product announcement.

Use only if:

- the submitter is already a normal participant
- the link is a substantial technical article
- the article improves the reader's next program or deepens their understanding
- the submitter participates in comments

Better candidate link:

```text
https://github.com/proompteng/bilig/blob/main/docs/why-agents-need-workbook-apis.md
```

Possible title:

```text
why agents need workbook apis instead of spreadsheet screenshots
```

Likely tags to evaluate in the UI:

```text
programming, typescript, practices
```

Do not submit the repository homepage as a product announcement.

## Product Hunt

Product Hunt is a distribution surface for packaged products. The thumbnail and
gallery images now exist in
[`docs/product-hunt-launch-assets.md`](product-hunt-launch-assets.md), so the
next safe step is a draft, not a scheduled launch. Do not launch until there is
a support window for comments and one more concrete product proof item to point
people at.

Draft listing:

```text
Name: bilig
Tagline: headless spreadsheet engine for services and agents
URL: https://github.com/proompteng/bilig
```

Maker comment draft:

```text
I built Bilig Headless because a lot of business logic still lives in spreadsheet-shaped models, but service automation and coding agents usually have to choose between screen-scraping a browser grid or rewriting formulas in application code.

The current package, @bilig/headless, gives Node services a WorkPaper API for formulas, structural edits, persistence round trips, validation, and readback. The repo includes a runnable external example, benchmark evidence, and compatibility caveats.

It is early infrastructure, not a finished Excel clone. I would most like feedback from people building spreadsheet-backed services, formula engines, import/export pipelines, or agent workflows that need reliable workbook state.
```

Do not launch on Product Hunt until these are ready:

- small thumbnail or logo: ready in `docs/assets/product-hunt-thumbnail.png`
- gallery images showing code and output: ready in `docs/assets/product-hunt-gallery-*.png`
- one short demo recording: ready in `docs/assets/product-hunt-demo.webm`
- maintainer comment
- support window for launch-day comments
- follow-up plan for GitHub issues and docs

## Tracking

Track every community post in the GitHub feedback discussion:

```text
https://github.com/proompteng/bilig/discussions/115
```

- platform
- URL
- date
- starting star count
- ending star count after `24` and `72` hours
- useful questions
- issues/docs created from feedback

The real metric is not raw post score. It is whether the post drives qualified
visitors, stars, npm installs, issues, reduced fixtures, or concrete maintainer
feedback.

## Continuous Community Loop

Use this when there is no obvious launch window:

1. Pick one proof artifact from the repo.
2. Pick one community where that artifact answers an existing technical pain.
3. Rewrite the post for that community's norms instead of reusing launch copy.
4. Disclose the maintainer relationship in the first paragraph.
5. Stay in the comments until the thread goes quiet.
6. Convert the best question into a GitHub issue, doc patch, fixture, or
   example before starting the next community post.

Cadence: one community post per week is enough while the project is early. More
volume only helps after comments are being answered quickly and the repo is
turning feedback into visible improvements.

## Sources

- Reddit spam policy:
  <https://support.reddithelp.com/hc/en-us/articles/360043504051-Spam>
- DEV Community terms:
  <https://dev.to/terms>
- DEV Code of Conduct:
  <https://dev.to/code-of-conduct>
- Lobsters guidelines:
  <https://lobste.rs/about>
- Product Hunt launch guide:
  <https://www.producthunt.com/launch/>
- DEV editor guide:
  <https://dev.to/p/editor_guide/>
