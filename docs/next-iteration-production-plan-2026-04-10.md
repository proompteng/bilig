# bilig Next Iteration Production Plan

## Status
Proposed, repo-grounded plan for the next production iteration.

## Executive summary

The next iteration should **not** be another topology rewrite. The repo already has the right backbone: a shipped `apps/bilig` monolith, a worker-first `apps/web` shell, OPFS-backed local persistence, Zero/Postgres-backed authoritative sync, and an embedded Codex-based workbook assistant.

The highest-leverage move now is to turn those pieces into one coherent product definition:

> **A multiplayer, highly performant, local-first spreadsheet product where a Codex app-server agent runs inside a document chat, plans spreadsheet-native workflows from prompts, previews them safely, and commits them through the same authoritative mutation stream as human edits.**

That means the next iteration should focus on four things:

1. **Make chat and agent execution durable and multiplayer**, not session-ephemeral.
2. **Make every supported spreadsheet workflow callable from chat**, using semantic workbook tools rather than UI automation.
3. **Preserve the worker-first, local-first performance model** while scaling collaboration and agent traffic.
4. **Keep the product operationally boring**: one product backend runtime, no correctness-path microservice sprawl, no second execution model.

---

## 1. Repo-grounded starting point

The current repo already contains the core building blocks we should preserve and harden:

- `apps/bilig` is the active monolith runtime and serves the web shell, sync routes, agent routes, and Zero ingress.
- `apps/web` is already worker-first by default and mounts the worker-owned runtime path.
- `packages/storage-browser` provides OPFS-backed SQLite browser persistence for local-first state, authoritative base tiles, projection overlays, and the pending mutation journal.
- `apps/web/src/worker-runtime.ts` already supports local replay, authoritative reconcile, preview generation for agent command bundles, and crash-safe pending mutation persistence.
- `apps/bilig/src/codex-app/workbook-agent-service.ts` already embeds a Codex app-server client in the monolith and stages workbook command bundles from chat tool calls.
- `apps/web/src/use-workbook-agent-pane.tsx` and `apps/web/src/WorkbookAgentPanel.tsx` already ship a document-side chat panel with prompt submission, streaming updates, preview/apply, and execution replay.
- `apps/bilig/src/zero/service.ts` already validates preview/apply against authoritative workbook revision and persists execution records.

### What is strong today

- Worker-first browser runtime is real, not aspirational.
- Local-first persistence is real, not a thin cache.
- The authoritative path is server-ordered and durable.
- The agent already uses **semantic workbook tools** instead of DOM automation.
- Preview/apply, undo, and execution records already exist as product concepts.

### What is still missing for the next product-grade iteration

- Chat/session state is still largely **in-memory and per-user**, which prevents true shared multiplayer chat and makes resilience weaker than the workbook runtime itself.
- Pending bundles and timeline items are not yet a fully durable collaboration surface.
- The Codex runtime is effectively a singleton monolith-owned transport, which is fine for current scope but not for broader multiplayer adoption.
- The supported agent tool surface is good, but it is not yet framed as a complete **workflow runtime**.
- Some correctness/performance seams remain high risk:
  - projected viewport authority still needs further narrowing
  - formula parity / WASM production routing is not fully closed on the full desired surface
  - multi-tab/browser-lock edge cases still need explicit product behavior
  - typed binary agent payloads are not fully closed end to end

---

## 2. Product definition for the next iteration

### Product thesis

`bilig` should become a **local-first workbook operating system** for serious spreadsheet work.

The defining experience is not just “edit cells in a browser.” The defining experience is:

- open a workbook and start working immediately from local state
- collaborate with other users on the same document with fast convergence and explicit conflict handling
- ask for spreadsheet work in a chat using natural language
- watch the agent inspect the workbook, plan changes, and stage a preview
- approve or auto-apply bounded changes through the same authoritative mutation path as direct edits
- keep every accepted change undoable, replayable, and auditable

### The release-level outcome

At the end of this iteration, the product should feel like:

- **multiplayer by default**
- **local-first by default**
- **chat-native for spreadsheet workflows**
- **server-authoritative for correctness**
- **worker-fast for interaction**
- **monolithic in deployment shape, not microservice-fragmented**

---

## 3. Hard decisions

### 3.1 Keep the monolith

`apps/bilig` stays the only product backend runtime.

The monolith owns:

- auth/session resolution
- document control
- authoritative workbook mutation ordering
- Zero query/mutate ingress
- checkpoint/replay
- recalc and runtime warming
- workbook chat and workflow orchestration
- Codex app-server session management
- workflow execution
- static asset serving

**No new product microservices** should be introduced for chat, agent execution, workflow orchestration, or workbook mutation handling.

Infrastructure dependencies remain acceptable:

- Postgres for durable state
- Zero data plane / cache where already required by the runtime model

But there should be **one product backend runtime** and **one correctness path**.

### 3.2 Keep the server authoritative

The product remains optimistic and local-first in the browser, but authoritative mutation order remains server-owned.

That means:

- local edits are previewed/applied instantly in the worker runtime
- all accepted edits still commit through the authoritative mutation path
- agent bundles still validate against base revision and authoritative preview parity before commit
- no peer-to-peer multiwriter design is introduced in this tranche

### 3.3 Keep the agent semantic

The agent must continue to act through **spreadsheet-native tools and command bundles**, not browser clicks or DOM scraping.

The Codex app server can reason freely inside the document chat, but it can only change workbook state by emitting supported semantic operations.

### 3.4 Keep the browser worker-owned

The worker runtime remains the browser truth for hot-path interaction.

The UI thread stays thin:

- input plumbing
- rendering orchestration
- focus/accessibility behavior
- side rails and chrome

The worker keeps ownership of:

- workbook preview engine
- local mutation journal
- local projection state
- authoritative ingest/reconcile
- preview generation
- viewport patch publication
- local persistence coordination

---

## 4. What this iteration will deliver

### 4.1 Multiplayer workbook collaboration as a first-class product surface

The next iteration must make collaborative editing feel native rather than “network-shaped.”

Core deliverables:

- shared same-document editing with durable presence and selection state
- convergence through ordered authoritative events
- explicit conflict surfacing for same-cell divergence and stale preview applies
- shared change timeline and revertable change bundles
- collaborator-visible agent executions when a thread is shared

### 4.2 Chat-native spreadsheet workflows

The document chat becomes a first-class execution surface, not a side experiment.

A user must be able to type prompts like:

- “Summarize this workbook and explain the important sheets.”
- “Find broken formulas and fix the visible range.”
- “Normalize this imported sheet, standardize the dates, and create a summary tab.”
- “Build a monthly rollup sheet from the regional tabs.”
- “Highlight outliers in revenue and add a MoM growth formula column.”

And the system must:

1. inspect relevant workbook context
2. plan through semantic tools
3. stage a command bundle or workflow run
4. preview the impact locally
5. apply authoritatively with replayable history

### 4.3 Local-first behavior that survives real-world failure

The workbook already has a strong local-first base. The next iteration extends that standard to collaborative and chat flows.

That means:

- workbook edits survive refresh, crash, reconnect, and offline periods
- chat drafts and recent chat context survive refresh
- accepted agent runs remain visible after reconnect and restart
- the product stays useful when network or agent availability is degraded

### 4.4 Performance that does not regress under collaboration and AI load

Agent and multiplayer features must fit inside the same performance contract as direct editing.

No “smart” feature is allowed to reintroduce:

- request-path memory explosions
- heavy point-query churn on selection changes
- frame hitching during viewport interaction
- broad full-workbook invalidations for narrow edits

---

## 5. Non-goals for this iteration

This iteration should explicitly **not** try to do the following:

- split the monolith into multiple product services
- introduce a second workbook execution engine
- rely on DOM or browser automation for spreadsheet actions
- build peer-to-peer offline multiwriter replication
- launch an open plugin marketplace
- add arbitrary shell/file/web capabilities to the agent runtime
- become a notebook or BI product before the spreadsheet workflow core is sharp

---

## 6. Target architecture

```text
Browser UI shell
  -> Worker runtime
      -> OPFS SQLite local store
      -> preview engine + mutation journal + viewport tile store
      -> local agent preview generation
  -> Monolith (apps/bilig)
      -> auth/session
      -> authoritative workbook runtime
      -> Zero ingress and fanout
      -> Codex app-server pool
      -> chat/workflow orchestration
      -> checkpoint/replay + recalc
  -> Postgres
      -> workbook durability
      -> chat/workflow durability
      -> agent execution records
  -> Zero data plane
      -> narrow workbook/collaboration projections
```

### 6.1 Monolith responsibilities

The monolith should own these domains directly:

- **Workbook authority**: ordered mutation apply, snapshot/replay, revision tracking
- **Collaboration**: presence, change feeds, shared thread metadata, shared approvals
- **Agent orchestration**: Codex session lifecycle, tool dispatch, approval policy, preview verification
- **Workflow execution**: long-running spreadsheet-native operations executed against authoritative state
- **Observability**: latency, memory, queue depth, preview mismatch, reconnect health

### 6.2 Browser runtime responsibilities

The browser worker should own:

- local workbook warm start
- durable pending op journal
- projection overlay and viewport patches
- local preview rendering for agent bundles
- reconnect/rebase behavior
- local persistence of lightweight chat state and session references

### 6.3 Data model additions

The next iteration should make chat and workflows durable with explicit schema, instead of keeping them mostly as in-memory session objects.

Add durable domain tables for:

- `workbook_chat_thread`
- `workbook_chat_item`
- `workbook_chat_tool_call`
- `workbook_pending_bundle`
- `workbook_workflow_run`
- `workbook_workflow_step`
- `workbook_workflow_artifact`

Existing `workbook_agent_run` can remain as the accepted-run audit table, but the live chat and pending bundle state must become durable and replayable.

### 6.4 Streaming model

Use a two-lane model:

1. **live lane** for turn deltas and in-progress chat streaming
2. **durable lane** for reconstructable thread state and collaborator fanout

Recommended behavior:

- live token/tool deltas continue streaming from the monolith for immediacy
- every completed item, tool call, pending bundle, workflow status change, and accepted execution is persisted to Postgres
- reconnecting clients rebuild from durable state first, then resume live deltas
- Zero exposes narrow thread/run projections so collaborators can follow shared chat state without depending on one browser session’s SSE connection

### 6.5 Codex app-server execution model

The Codex app server should remain **embedded inside the monolith boundary** as a child-process pool managed by `apps/bilig`.

The next iteration should move from a mostly singleton transport model to a bounded pool with:

- concurrency caps
- idle eviction
- backpressure and queueing
- per-document / per-user quotas
- process health checks
- structured metrics per turn and per tool call

This preserves the monolith shape while making chat scale operationally.

---

## 7. Chat-native workflow execution contract

This is the most important product contract in the plan.

### 7.1 Every workflow starts in chat

A user prompt inside a document-scoped chat is the entry point.

The chat may be:

- **private**: visible only to the initiating user
- **shared**: visible to collaborators on the workbook

### 7.2 Every mutating workflow is semantic

The agent may inspect with tools freely, but every mutating result must compile to one of these bounded outputs:

- a **command bundle** for immediate preview/apply
- a **workflow run spec** for heavier multi-step execution in the monolith
- a **read/report result** with no workbook mutation

### 7.3 Every workflow is revision-aware

Every mutating plan must include:

- document id
- thread id / turn id
- base revision
- bounded affected ranges or workbook scope
- risk class
- approval mode
- preview summary

### 7.4 Every accepted workflow is undoable and auditable

When a workflow applies:

- it commits through the authoritative mutation stream
- it produces an execution record
- it has an undo path or revert bundle where semantically valid
- it is attached to the originating chat thread and workbook revision lineage

### 7.5 Every supported spreadsheet workflow must be callable from chat

“Any spreadsheet engine workflow” should mean:

- any workflow represented by the supported semantic operation surface
- any document analysis workflow that can be expressed through workbook reads and inspection tools
- any heavier transformation workflow that can be executed as a workbook-native monolith job

It should **not** mean unconstrained arbitrary code execution.

---

## 8. Day-1 workflow families

The first production workflow set should be deliberately narrow, highly reliable, and clearly useful.

### 8.1 Read and explain

- summarize workbook structure
- explain current cell / formula
- search workbook for concept, formula, or value
- trace precedents / dependents
- describe recent changes

### 8.2 Formula diagnostics and repair

- find error cells
- detect cycles
- identify JS-only fallback formulas
- propose formula repairs in bounded ranges
- stage formula rewrites as preview bundles

### 8.3 Import cleanup and normalization

- normalize headers
- infer and standardize date / currency / percentage columns
- split or combine fields
- fill formulas down imported ranges
- create clean output sheets from messy imports

### 8.4 Reshape and summarize

- create new summary sheets
- consolidate multiple sheets into one rollup
- move/copy/fill ranges semantically
- create workbook-ready reporting tabs

### 8.5 Formatting and presentation cleanup

- normalize number formats
- apply consistent styles
- highlight exceptions and outliers
- prepare review-ready sheets

### 8.6 Safe structural edits

- create sheet
- rename sheet
- bounded move/copy/fill operations
- row/column visibility and sizing operations once authoritative parity is closed on those surfaces

---

## 9. Approval and risk policy

The agent should not have one global write mode. It needs policy tied to change shape.

| Risk class | Typical scope | Default policy | Examples |
| --- | --- | --- | --- |
| Low | selection / small visible-range changes | auto or preview | format cells, write bounded formulas, normalize selected range |
| Medium | sheet-scoped edits | preview required | repair formula region, build summary block, clean imported sheet |
| High | workbook-wide or structural edits | explicit apply | create/rename sheets, large restructures, cross-sheet rewrites |

### Policy rules

- low-risk changes may auto-apply only when the affected scope is clearly bounded and the user has editor rights
- medium-risk changes always require preview
- high-risk changes always require explicit apply and should be shareable for collaborator review on shared documents
- stale or mismatched previews always fail closed and require refresh / replay

---

## 10. Core workstreams

## Workstream A — Durable multiplayer chat and agent state

### Why

The workbook runtime is already more durable than the chat runtime. That mismatch will become a product-quality problem as soon as the chat is used seriously by multiple collaborators.

### Deliverables

- persist thread metadata and timeline items
- persist pending bundles, not just accepted execution records
- support private and shared thread scope
- reconstruct sessions after monolith restart and browser reconnect
- broadcast collaborator-visible thread/run state through durable projections
- keep live delta streaming for immediacy

### Exit gate

A shared document chat can survive browser refresh, monolith restart, reconnect, and collaborator join without losing accepted thread history or pending bundle state.

---

## Workstream B — Workflow-native Codex chat

### Why

The current assistant is already useful, but it still looks like an embedded agent session. The next iteration needs it to feel like a first-class workflow runtime.

### Deliverables

- formal workflow result types: read/report, command bundle, workflow run
- expand semantic tool surface where needed for supported workflow families
- attach range, sheet, and revision citations to plans and results
- add workflow templates / intents for repeated spreadsheet jobs
- support cancellation, replay, and partial acceptance where semantically valid
- surface structured progress in chat for longer-running workflows

### Exit gate

A user can start a workbook workflow from chat, watch it reason over workbook state, preview the proposed changes, and apply the result without leaving the document.

---

## Workstream C — Authoritative semantic parity

### Why

The product only stays trustworthy if every local previewed action has a matching authoritative meaning.

### Deliverables

- audit every supported direct-edit and agent-edit operation against the authoritative mutation path
- close any remaining gaps between local engine capability and durable authoritative mutation representation
- ensure all workflow-generated edits compile to the same semantic operation family used by direct editing
- keep undo/revert support coherent across direct edits and workflow applies

### Exit gate

Every supported agent mutation and direct workbook action has one authoritative semantic representation, one replay story, and one correctness harness.

---

## Workstream D — Performance and scale preservation

### Why

Multiplayer and AI features are only acceptable if they fit inside the existing performance posture.

### Deliverables

- preserve worker-first hot path and viewport patch model
- keep Zero narrow and tile-shaped
- reduce or fully retire temporary local authority layers that can drift under load
- expand WASM production routing only for proven formula families
- add codex pool backpressure and queue metrics
- add collaboration and chat load tests to the existing performance contract

### Exit gate

The product remains within the published performance budgets while collaboration and chat are active.

---

## Workstream E — Local-first hardening and multi-tab behavior

### Why

The OPFS-backed local store is a strength, but multi-tab and lease behavior must be a deliberate product surface rather than an incidental lock error.

### Deliverables

- document writer lease / ownership model across tabs
- follower-tab behavior when a document local store is already locked
- clean lease transfer path
- offline banner and degraded-mode messaging that do not block editing
- local persistence of recent chat state and drafts

### Exit gate

The user can understand what happens when the same workbook is open in multiple tabs, and no committed work is silently lost.

---

## Workstream F — Reliability, correctness, and rollout

### Why

This product is already sophisticated enough that correctness regressions and memory failures are product failures, not engineering details.

### Deliverables

- replay tests for preview/apply mismatch and stale bundle scenarios
- property tests for workflow replay and partial acceptance invariants
- deterministic integration harnesses for chat, tool streaming, and reconnect
- production dashboards and alerts for codex pool, Zero lag, memory, and preview mismatch rate
- feature flags and canary rollout for shared chat, auto-apply, and workflow families

### Exit gate

Collaboration and chat features are behind explicit launch gates and observable with production dashboards before broad rollout.

---

## 11. Implementation sequence

## Tranche 1 — Make chat durable and document-scoped

### Build

- introduce durable chat/thread tables
- persist thread items, tool calls, pending bundles, and workflow metadata
- keep current session API surface compatible while migrating the web shell to durable threads
- reconstruct thread state from durable storage on reconnect
- expose thread list / unread / shared thread summaries through narrow projections

### Result

Chat stops being “session memory attached to a document” and becomes a durable document feature.

### Exit gate

A monolith restart no longer wipes active workbook chat history or pending bundle state.

---

## Tranche 2 — Shared multiplayer agent experience

### Build

- add shared vs private thread scope
- add collaborator-visible thread and run summaries
- allow collaborators to inspect pending bundles and approved executions on shared threads
- add approvals for medium/high-risk shared changes
- attach cell/range references to thread items and results

### Result

The agent becomes multiplayer-aware instead of purely personal.

### Exit gate

Two collaborators on the same workbook can follow a shared agent thread and understand what changed, why, and at which revision.

---

## Tranche 3 — Workflow runtime on top of Codex chat

### Build

- formalize workflow run types
- expand tool surface for the Day-1 workflow families
- add long-running monolith workflow runner for document-wide analysis/transforms
- stream workflow progress into chat
- return preview bundles or artifacts from workflow runs

### Result

Chat becomes a reliable front door for spreadsheet-native workflows.

### Exit gate

The Day-1 workflow families work end-to-end from chat prompt to applied workbook result.

---

## Tranche 4 — Performance and correctness hardening

### Build

- add perf/load contracts for collaboration + chat
- close highest-risk local authority seams
- improve codex pool management and memory isolation
- expand preview/apply correctness harnesses
- improve multi-tab writer lease behavior

### Result

The new product surface is safe to roll out broadly.

### Exit gate

Performance, correctness, and resilience gates are green for workbook + collaboration + chat together.

---

## 12. API and protocol direction

### Keep these principles

- binary sync framing remains the long-term canonical transport direction
- the agent protocol should move from JSON bodies inside frames toward typed payload codecs
- the browser should not need a second unofficial path for agent state or workbook mutations

### Recommended API shape

Keep the current endpoints working, but evolve toward durable thread-centric routes:

- `POST /v2/documents/:documentId/chat/threads`
- `POST /v2/documents/:documentId/chat/threads/:threadId/turns`
- `GET /v2/documents/:documentId/chat/threads/:threadId/events`
- `POST /v2/documents/:documentId/chat/threads/:threadId/bundles/:bundleId/apply`
- `POST /v2/documents/:documentId/chat/threads/:threadId/bundles/:bundleId/dismiss`
- `POST /v2/documents/:documentId/chat/threads/:threadId/workflows`
- `POST /v2/documents/:documentId/chat/threads/:threadId/workflows/:runId/cancel`

The important shift is not path naming. The important shift is that the backing state becomes durable and multiplayer-safe.

---

## 13. Performance, scale, and SLO targets

These targets should become explicit release gates for the iteration.

| Metric | Target |
| --- | --- |
| Local visible edit p95 | `<16ms` |
| Selection paint p95 | `<8ms` |
| 100-cell paste first local paint | `<40ms` |
| 100-cell paste first authoritative diff p95 | `<100ms` |
| Collaborator visible update p95 | `<150ms` |
| Warm reopen last workbook p95 | `<500ms` |
| Warm-start first useful paint, 100k workbook | `<250ms` |
| Warm-start first useful paint, 250k workbook | `<700ms` |
| Reconnect after offline period with 100 pending ops | `<2s` |
| Pending-op loss | `0` |
| First assistant delta p95 | `<700ms` |
| First preview highlight p95 | `<1000ms` |
| Accepted agent mutation visible locally p95 | `<1500ms` |
| Preview/apply mismatch rate | `<0.5%` |
| JS/WASM mismatch rate on promoted families | `<0.1%` |

### Suggested initial scale target

For the next iteration, the explicit same-document collaboration target should be:

- **25 active editors** on one workbook
- **100 passive viewers/followers** on the same workbook
- active chat/workflow usage without breaking the hot path

If the product exceeds that later, the architecture can be tuned from a stable baseline instead of guessed prematurely.

---

## 14. Operational plan

### 14.1 Feature flags

Roll out behind separate flags for:

- durable chat threads
- shared chat threads
- auto-apply low-risk bundles
- long-running workflow runner
- each Day-1 workflow family

### 14.2 Observability

Add dashboards and alerts for:

- monolith RSS / heap / GC pause
- codex pool busy count and queue depth
- first-delta latency and tool-call latency
- preview generation latency
- apply conflict / stale bundle / mismatch rate
- Zero lag and authoritative event catch-up time
- OPFS local store open failure / lock contention rate
- reconnect/rebase latency

### 14.3 Rollout shape

- internal dogfood
- allowlisted documents / users
- low-risk workflow families first
- shared threads after private-thread durability is proven
- auto-apply only after preview mismatch and revert confidence are inside budget

---

## 15. Acceptance matrix for the release

The release is ready only when all of the following are true:

### Product readiness

- a workbook can be opened, edited, and recovered from local state without network dependency
- same-document collaboration works with presence and ordered convergence
- a user can initiate supported spreadsheet workflows from chat prompts
- shared chat threads are durable and reconnect-safe
- accepted agent executions are replayable and auditable

### Correctness readiness

- stale bundle and preview mismatch paths fail closed and are test-covered
- workflow-generated edits use the authoritative semantic op model
- preview/apply parity tests are green
- projection parity and reconnect/rebase tests are green

### Performance readiness

- published SLO targets are green on release hardware profiles
- collaboration + chat do not regress direct-edit hot-path budgets
- no request-path memory spikes or GC stalls reintroduce prior failure modes

### Operational readiness

- dashboards and alerts exist for the new surfaces
- feature flags can disable shared chat, workflow runner, and auto-apply independently
- migration and recovery playbooks are documented
- canary rollout is clean before broad exposure

---

## 16. Biggest risks and mitigations

### Risk 1 — Codex runtime becomes a bottleneck inside the monolith

**Mitigation**

- bounded child-process pool
- concurrency caps and queueing
- idle eviction
- per-user and per-document quotas
- hard metrics and kill/restart policies

### Risk 2 — Chat becomes durable but loses streaming responsiveness

**Mitigation**

- separate live delta lane from durable lane
- stream deltas immediately, persist completed items and state transitions asynchronously but quickly
- rebuild from durable state on reconnect instead of trying to make the live lane do both jobs

### Risk 3 — Preview/apply divergence under multiplayer edits hurts trust

**Mitigation**

- keep base revision checks
- keep authoritative preview verification before apply
- attach range/revision citations to plans
- give the user one-click replay/rebuild of stale bundles

### Risk 4 — Multi-tab OPFS locking becomes confusing

**Mitigation**

- explicit document writer lease model
- follower-tab UX instead of generic lock errors
- deliberate lease transfer flow

### Risk 5 — Workflow surface grows faster than correctness coverage

**Mitigation**

- launch only narrow Day-1 workflow families
- require replay harnesses and acceptance cases before adding workflow families
- keep semantic tool surface tighter than the model’s general capability

---

## 17. Final recommendation

The right next iteration is:

> **Ship a durable, multiplayer, chat-native workflow layer on top of the existing worker-first local-first spreadsheet monolith.**

Do **not** spend this iteration inventing a new topology.
Do **not** add a second execution model.
Do **not** let the agent escape into UI automation.

Instead:

- preserve the monolith
- preserve the worker-owned local-first runtime
- make chat durable and multiplayer
- make Codex chat the front door to spreadsheet-native workflows
- keep every accepted change authoritative, previewed, undoable, and observable

That is the fastest path from the current repo to a genuinely differentiated product.
