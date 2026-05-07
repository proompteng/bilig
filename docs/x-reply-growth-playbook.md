# X Reply Growth Playbook

Status: public, account-safe outreach playbook for growing `proompteng/bilig`.

Goal: earn legitimate attention from developers who already care about
spreadsheet engines, workbook automation, local-first software, coding agents,
formula semantics, or open-source infrastructure. Do not optimize for evading
platform enforcement. Optimize for useful replies that would still make sense if
the link were removed.

## Ground Rules

- Reply manually.
- Do not use browser automation or bots to post replies.
- Do not send duplicated replies across many accounts.
- Do not reply from keyword search alone.
- Do not mass mention people.
- Do not ask for stars as the first interaction.
- Do not post unrelated links into trends.
- Link only when the thread is already about spreadsheets, workbook automation,
  formula engines, agent tooling, or open-source implementation evidence.

X's own rules say unsolicited automated replies based only on keyword searches
are not permitted, and its behavior guidance treats repeated duplicated
unsolicited replies as spam. That makes the correct strategy simple: fewer
replies, higher fit, and no automation.

## Daily Reply Budget

Use a hard cap until the account has real inbound discussion:

- `0-2` high-fit replies per day
- `0-1` link-bearing replies per day
- no more than `1` reply in a single thread unless someone asks a follow-up

This is not about hiding from filters. It is about keeping the account useful
enough that a person reading the reply sees a relevant contribution, not a
growth tactic.

## Thread Fit

Good targets:

- someone asking how to run spreadsheet logic outside Excel or Google Sheets
- a maintainer discussing formula parity, import/export fidelity, or XLSX
  compatibility
- a developer comparing HyperFormula, DuckDB, local-first data tools, or browser
  workbook UIs
- an agent-tooling thread where the missing primitive is reliable workbook state
- an open-source thread asking for good first issues or implementation-heavy
  projects

Bad targets:

- generic ai trend posts
- launch posts from unrelated products
- outrage threads
- posts where a link would feel inserted instead of requested by the topic
- celebrity threads unless the content is directly about developer tooling

## Voice

Keep the tone close to the common builder style on X: lowercase, direct, short,
and specific. Avoid polished marketing language.

Useful shape:

1. agree or add one concrete distinction
2. name the implementation evidence
3. link only if it clarifies the claim

Avoid:

- "we're revolutionizing spreadsheets"
- "please star us"
- "check out my project" without context
- emojis as decoration
- repeated sentence templates

## Current Style Scan

Checked on 2026-05-07 in Atlas on the live `@sama` X profile. Recent
high-engagement posts and replies tend to be compressed, casual, and
low-friction:

- mostly lowercase, except when announcing a product milestone
- one idea per post
- plain verbs instead of launch-copy adjectives
- little punctuation
- concrete belief or observation first, explanation second
- no decorative hashtags
- short replies can be a single sentence when the thread already has context
- stronger standalone posts add one specific reason after the claim

Use that shape only as a tone reference. Do not impersonate anyone, copy
phrasing, or turn posts into vague ai hype. For `bilig`, the useful version is
lowercase and human but still specific: workbook state, formula parity,
readback, fixtures, examples, and measured caveats.

## Live Reply Queue - 2026-05-07

These are hand-picked targets from the logged-in Atlas X session. Do not treat
this as a scraping queue. Use it as a small manual queue and stop after the first
reply unless a real follow-up appears.

### ChatGPT spreadsheet add-on post

Target:
<https://x.com/ChatGPTapp/status/2051776032127238266>

Why it fits:

- The post is directly about ChatGPT inside Excel and Google Sheets.
- The thread is already about spreadsheet automation, formulas, messy data, and
  explaining workbook changes.
- `bilig` adds the adjacent developer angle: typed workbook operations and
  verification for agents outside the browser grid.
- The post has broad reach, but the reply still needs to stand alone as a useful
  technical distinction.

Draft reply:

```text
this is exactly the direction. one thing we keep hitting while building bilig is
that agents need typed workbook operations and verification hooks, not
screenshots of grids.

open-source node api if useful for anyone experimenting:
https://github.com/proompteng/bilig
```

Lower-promotion variant:

```text
this is exactly the direction. the thing i keep wanting for agents is typed
workbook operations + verification hooks, not screenshots of grids.

the grid is the ui. workbook state is the api.
```

Use the linked version only if the maintainer account is comfortable clearly
owning the project in the reply. Use the no-link version when the thread feels
too crowded or the account needs more normal participation before linking.

Follow-up artifact if anyone engages:

- link `docs/why-agents-need-workbook-apis.md` for the conceptual argument
- link `examples/headless-workpaper` for the runnable Node example
- if someone asks about Excel parity, link the fixture-scoped caveats instead
  of making a broad compatibility claim

### AI Excel agent startups thread

Target:
<https://x.com/IM_Aeneas/status/2050729841947709822>

Why it fits:

- The post asks directly what Microsoft's Excel agent work means for AI Excel
  agent startups.
- The thread is not asking for a generic repo link; the useful contribution is
  a technical distinction about where standalone infrastructure still matters.
- `bilig` can add a maintainer-level point about workbook operations,
  recalculation correctness, persistence, and readback.

Draft reply:

```text
probably means the serious ones have to go deeper than chat around cells.

the hard part is typed workbook ops, import/export fidelity, recalculation
correctness, and verification after edits.
```

Only add a repo link if someone asks what an implementation of that boundary
looks like. If that happens, link the website or adoption kit rather than
dropping the repository into the first reply.

### PopSheets AI spreadsheet mention

Target:
<https://x.com/AudreyLimsAi/status/2052019390225555807>

Why it fits:

- The post is specifically about an AI spreadsheet that researches, analyzes,
  charts, and reports from data.
- A useful reply can add the developer infrastructure angle without competing
  with the product mention.
- Do not link on the first reply unless someone asks about implementation.

Draft reply:

```text
this is the right product surface.

the infrastructure layer i keep watching is whether the agent can prove what it
changed: ranges, formulas, recalc, readback, and export fidelity.
```

### ChatGPT Excel add-in personalization reply

Target:
<https://x.com/KieranJame86217/status/2051998803771949424>

Why it fits:

- The post is about skills syncing and personalized instructions for an Excel
  add-in.
- The useful angle is that personalization needs reliable workbook operations,
  not just prompt memory.
- Keep it as a normal technical reply with no repo link.

Draft reply:

```text
yeah, the interesting bit is making the instructions operational.

not just "remember how i like models", but stable workbook ops the add-in can
run and then verify after edits.
```

### Native Excel agent launch reply

Target:
<https://x.com/gardnersmitha/status/2051456316942458898>

Why it fits:

- The post mentions a native Excel add-in, direct data access, hot-swappable
  models, and domain skills.
- A useful reply can focus on the verification/writeback boundary that matters
  after an agent mutates a workbook.
- Use no link first; add `bilig` only if asked for open-source infrastructure
  examples.

Draft reply:

```text
direct data access is a big deal.

the other half i care about for spreadsheet agents is writeback verification:
after the model changes a workbook, can you inspect formulas/ranges and prove
what changed.
```

## Reply Templates

Use these as starting points, not copy/paste automation.

### agents and spreadsheets

```text
yeah, screenshots are the weak primitive here.

the useful thing is a workbook api the agent can mutate and verify. i wrote up
the shape we use in bilig:
https://github.com/proompteng/bilig/blob/main/docs/why-agents-need-workbook-apis.md
```

### formula parity

```text
i think the honest version is fixture-scoped parity, not "excel compatible".

for bilig i'm trying to make each claim point at the exact fixture + verifier
command, like this xlookup exact case:
https://github.com/proompteng/bilig/blob/main/docs/formula-edge-xlookup-exact-fixture.md
```

### xlsx compatibility

```text
cached xlsx parity is a useful test, but it should be described as corpus parity.

this is the report shape we use: matching formulas, mismatches, skipped formulas,
and why they were skipped:
https://github.com/proompteng/bilig/blob/main/docs/xlsx-corpus-verifier-walkthrough.md
```

### local-first workbooks

```text
the part that gets interesting is when the workbook document can round-trip
through json and still preserve formula-backed state.

this is the small node example i keep pointing people at:
https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper
```

### benchmarks

```text
the benchmark claim has to stay narrow or it becomes noise.

for bilig the public claim is 46/46 mean wins on scorecard-eligible comparable
workloads, with the p95 caveat left attached:
https://github.com/proompteng/bilig/blob/main/docs/what-workpaper-benchmark-proves.md
```

## No-Link Replies

Use these when a link would feel premature:

```text
this is where i think spreadsheet engines need more boring audit trails: exact
fixture, expected value, verifier command, and explicit gaps.
```

```text
the thing i would separate is "can import the file" vs "can prove the formulas
match cached results for this corpus".
```

```text
for agent workflows, i think the grid is the ui, not the api. the api needs
stable workbook state and readback.
```

## Follow-Up Loop

For every useful reply:

1. Save the thread URL.
2. Note the actual question or objection.
3. Convert repeated questions into a doc, example, test, fixture, or issue.
4. Reply once with the new artifact only if it directly answers the thread.

This compounds better than raw posting volume because it turns market feedback
into repository evidence.

## Continuous Growth Cadence

Run this as a weekly loop:

1. Ship one small proof artifact in the repo: fixture walkthrough, runnable
   example, benchmark note, compatibility caveat, or starter issue.
2. Publish one maintainer post that explains the proof in lowercase, direct
   language.
3. Spend `3` days watching related X discussions and add at most `2`
   high-context replies per day.
4. Log every serious objection and convert repeated questions into docs,
   issues, fixtures, or examples before posting the next link.
5. At the end of the week, compare GitHub stars, npm downloads, GitHub traffic
   referrers, issue quality, and repeat questions.

Do not optimize for reply count. The compounding unit is a public artifact that
is good enough to link when the same question appears again.

## Sources

- X automation rules and best practices:
  <https://help.x.com/en/rules-and-policies/x-automation>
- X account behavior best practices:
  <https://help.x.com/en/rules-and-policies/x-rules-and-best-practices>
- Sam Altman Codex milestone post, tone reference:
  <https://x.com/sama/status/2041658719839383945>
- Sam Altman voice-model post, tone reference:
  <https://di.gg/ai/25eceafd-986f-45ed-8b10-6a35eac12d30>
- Sam Altman AI-access reply, tone reference:
  <https://di.gg/ai/c7d99e5e-4e8b-456f-bc3d-788f8c775fb8>
- GitHub repository topics:
  <https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/classifying-your-repository-with-topics>
- Open Source Guides, building community:
  <https://opensource.guide/building-community/>
- GitHub Blog, building an open source community:
  <https://github.blog/open-source/maintainers/four-steps-toward-building-an-open-source-community/>
