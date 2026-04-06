## 1. Executive Summary

`bilig` today is a promising server-authoritative spreadsheet with a native grid, a warm server runtime, typed workbook operations, and a clear architectural direction around worker-first UI, narrow Zero sync, and projected viewport patches. It is **not** yet a real local-first workbook system: the mounted browser path still revolves around `ZeroWorkbookBridge` plus an in-memory viewport cache, offline writes are gated by connection state, and browser durability is still JSON snapshot persistence rather than a real local database. The attached repo brief and docs point in the right direction, but the hot path is not there yet.

The right next move is **not** “more spreadsheet features.” It is to turn `bilig` into a **local-first workbook operating system**: a browser-resident workbook runtime that opens instantly, writes locally first, stays durable offline, reconciles cleanly with a server-authoritative model, and keeps giant workbooks fluid. The product should feel native because the visible state comes from a worker-owned local database and local projection engine, not from network round-trips or React-held sheet arrays.

Direct answer to the SQLite question: **yes, use browser SQLite/WASM**, but refine the storage choice. For a giant-data workbook engine, the default should be **SQLite/WASM on OPFS in a worker**, with IndexedDB kept for fallback/bootstrap metadata rather than as the primary page store. MDN describes OPFS as origin-private storage optimized for in-place writes with synchronous worker access, and the official SQLite WASM docs expose `OpfsDb` as the browser-persistent path while the lighter JS storage path is aimed at local/session storage. That makes OPFS the right default for the serious path. ([MDN Web Docs][1])

The next major plan should do five things in order:

1. make the existing worker runtime the **real** mounted path
2. add a worker-owned **local SQLite + op journal + local projection** layer
3. narrow Zero into an **authoritative delta and collaboration feed**, not the hot UI model
4. harden the grid for giant-data UX
5. layer AI on **semantic workbook operations** with preview/apply, not on DOM automation

---

## 2. Current State

### What is already strong

The repo already has the right backbone.

- The architecture direction is solid: worker-first browser shell, server-authoritative ordering, Zero as narrow relational sync, projected viewport patches rather than raw engine state, and Postgres as the durable shared store. The repo docs are aligned on that, which matters.
- The product owns a native grid/runtime instead of hiding behind Glide. That is the only path that can produce spreadsheet-grade interaction quality at scale. The renderer docs explicitly treat the native path as the active one now.
- `packages/workbook-domain` is a real semantic operation layer, not UI event soup. That is exactly what a local op journal, replay, preview/apply AI, and authoritative reconciliation need.
- `apps/bilig` already has a warm runtime manager and authoritative mutation path. That is a major advantage: you do not need to invent the server model from scratch.
- `packages/zero-sync` is already tile/range-shaped enough to be useful. The current repo direction to keep Zero narrow and reduce projection churn is correct.
- The formula plan is disciplined: Excel semantics as canonical baseline, JS first, WASM later after parity gates. That sequencing is right.

### What is structurally weak

The browser-side product architecture is still the weak link.

- The mounted runtime path is now worker-owned, the browser store is now OPFS-backed SQLite through `@bilig/storage-browser`, and the old `viewport-cache.ts` layer has been deleted in favor of a narrower projected viewport store. The browser DB now has normalized authoritative base tables plus normalized projection overlay tables for cells, axis metadata, and styles, and the worker can serve full viewport patches from merged local base+overlay reads. Narrow authoritative event batches now ingest directly into those normalized local tables instead of forcing a full base repersist. The runtime is still only partway through the intended migration because stable `sheet_id` and a fully worker-owned tile store are not built yet.
- Offline/local-first credibility is materially better now because writes are no longer gated by Zero connection state, the worker keeps a crash-safe pending-op journal in SQLite, submitted ops stay durable until the authoritative revision feed absorbs them, reconnect now prefers authoritative event ingest before snapshot fallback, restored pending overlays survive reload, restored local sessions no longer block on a cold snapshot fetch before showing accepted local state, and absorbed authoritative events now update the local DB through a direct delta path. The product is still not fully local-first because stable `sheet_id` and the worker-owned tile store are not done.
- Browser durability is materially better now because accepted local state and pending ops live in SQLite/OPFS rather than IndexedDB/localStorage JSON. The remaining weakness is the data model inside that DB, not the browser persistence substrate.
- The original hot-path file-size debt has been materially reduced, but the runtime is still split across cache/session/runtime layers that need a cleaner local-first decomposition.
  - `packages/grid/src/WorkbookGridSurface.tsx` is now 175 lines with the interaction and render logic extracted.
  - `apps/web/src/WorkerWorkbookApp.tsx` is now 119 lines with state/sync/toolbar logic extracted.
  - `apps/web/src/worker-runtime.ts` is now 776 lines with persistence and viewport support split out.
- The selected-cell synthetic live-sync path is gone. Selected-cell state now comes through the same viewport/tile subscription path as the rest of the grid, which removes a stale-path class that used to exist.
- The grid still uses DOM text rendering and browser text measurement for autofit. The repo docs already call that out. It limits both polish and giant-workbook consistency.

### What most affects UX quality

Three things are doing the most damage:

1. **No single browser truth store.**
   Visible state, selected-cell state, editor state, and optimistic state are still too spread out.

2. **The hot path is too UI-thread-centric.**
   The repo says worker-first; the mounted product still behaves more like “UI thread shell + bridge + in-memory cache.”

3. **The renderer is not yet text-native.**
   DOM text plus browser-only measurement is good enough for bring-up, not for a 10x-feels-native product.

### What most affects speed

- no worker-owned local database
- no local op journal
- no local projection tables
- in-memory cache budget that is too small and too arbitrary for giant workbooks
- per-feature optimistic cache mutation from `WorkerWorkbookApp.tsx`
- bridge patterns that still rebuild or reproject too much state

### What most affects local-first credibility

Right now the product cannot honestly claim local-first because:

- accepted local edits are not durably written to a real local store first
- offline writes are blocked
- reconnect semantics are not built around a durable local outbox
- conflict handling is not built around rebase/replay on top of authoritative revisions
- startup does not come from a proper local workbook database

### What most affects long-term architecture quality

The biggest architectural issue is that the repo already contains the pieces of a better browser runtime, but the product is still not using them as the primary state model. That causes a bad middle state:

- a real worker runtime exists
- the UI still behaves like a bridge-and-cache app
- the browser persistence layer is snapshot-oriented
- the selection path still has special handling
- the grid shell still owns too much behavior

Bluntly: `bilig` is a **good server-authoritative spreadsheet core** with a **browser runtime that still needs to grow up**.

---

## 3. Winning Product Definition

Google Sheets already sets a strong baseline for online collaboration, comments, task assignment, filter views, structured tables/chips, and Gemini-assisted sheet creation and analysis. Excel sets the semantics bar, and Microsoft is now pushing Copilot plus Python in Excel for advanced analysis and workbook editing. Airtable and Smartsheet both push above the grid into workflow, automation, and trusted execution. `bilig` only wins if it feels materially faster and sharper at the workbook core while also making multi-step workbook work easier and safer. ([Google Workspace][2])

### Product definition by dimension

- **Input latency**
  Baseline: web spreadsheets are good enough for ordinary edits but still feel network-shaped under stress.
  `bilig`: the UI commits locally before the network matters.
  Success: local visible response p95 `<16ms`; selection paint p95 `<8ms`.

- **Visual smoothness**
  Baseline: smooth enough on normal sheets; frame pacing degrades on huge or collaboration-heavy sheets.
  `bilig`: 60fps-feeling pan/scroll/select on active viewports, even while sync is happening.
  Success: steady 60fps on modern laptop for active viewport operations; no visible hitch on same-sheet collaborator updates in viewport.

- **Edit responsiveness**
  Baseline: edits are fast when network and workbook are friendly.
  `bilig`: local edit is instant, authoritative confirm is a second step, not the first visible one.
  Success: single-cell local apply `<16ms`; single-cell authoritative ack p95 `<60ms`; 100-cell paste first local paint `<40ms`; first authoritative diff p95 `<100ms`.

- **Offline durability**
  Baseline: some offline behavior exists in incumbent tools, but the experience is not the center of the product.
  `bilig`: once a local edit is accepted, it survives refresh, crash, tab recycle, and offline periods.
  Success: 100% locally accepted ops survive reload/crash; warm-start last workbook p95 `<500ms`.

- **Sync behavior**
  Baseline: continuous sync exists, but network hiccups still leak into perceived product quality.
  `bilig`: sync is background plumbing; the sheet stays usable.
  Success: reconnect after 5 minutes offline with 100 pending ops catches up p95 `<2s`; no pending-op loss.

- **Conflict handling**
  Baseline: same-cell and same-view conflicts are still frustrating.
  `bilig`: local draft never disappears, remote authority reconciles cleanly, and the user sees what happened.
  Success: zero lost in-progress editor drafts; explicit conflict UI for same-cell divergence; rebase failure rate `<0.5%`.

- **Collaboration ergonomics**
  Baseline: Sheets is strong at basic collaboration; filter/view isolation and change review still leave room.
  `bilig`: collaborators can work in parallel without stepping on each other’s view state or destroying someone’s flow.
  Success: private views + named views + presence + changes pane live; >50% of shared workbooks use at least one of those features.

- **Formula experience**
  Baseline: Excel still owns user trust on semantics; Copilot/Gemini help, but trust comes from the engine.
  `bilig`: formula entry, autocomplete, arg help, precedents/dependents, and parity are all fast and dependable.
  Success: formula parse/preview `<50ms` for single-cell edits; targeted parity suite green on committed surface; JS/WASM mismatch rate `<0.1%` on proven subset.

- **Workbook comprehension**
  Baseline: users still spend too much time figuring out what a workbook is doing.
  `bilig`: users can ask “what changed, what drives this, what depends on this, what is broken” and get immediate answers.
  Success: click-to-dependents/precedents; recent changes indexed; broken-formula scan for active workbook `<1s` on warm runtime.

- **Workflow execution**
  Baseline: spreadsheets are still flexible but messy execution tools.
  `bilig`: recurring workbook work becomes guided and reviewable without losing grid flexibility.
  Success: private views, saved views, change bundles, import-cleanup flows, and review-ready undo bundles.

- **AI-assisted building/editing**
  Baseline: Gemini and Copilot can help create formulas, tables, pivots, charts, and analyses.
  `bilig`: AI should plan over workbook operations, preview locally, and apply as undoable bundles.
  Success: first plan token p95 `<700ms`; first preview highlight p95 `<1000ms`; first visible accepted mutation p95 `<1500ms`. ([Google Help][3])

- **Extensibility**
  Baseline: incumbents support scripting/add-ins, but the execution model is often inconsistent or brittle.
  `bilig`: the extension surface is semantic command APIs and typed workbook tools, not DOM automation.
  Success: >95% of supported automations use semantic ops; zero production DOM-driven automations.

- **Large-workbook scale**
  Baseline: incumbents can open large workbooks, but tail latency and UI feel degrade fast.
  `bilig`: giant workbooks are a normal use case.
  Success: warm-start first useful paint: 100k workbook `<250ms`, 250k workbook `<700ms`; 5k-cell paste first visible authoritative diff `<250ms`; sort/filter on 100k rows remains interactive.

### KPI / SLO recommendations

| Metric                                       |                   Target |
| -------------------------------------------- | -----------------------: |
| Local visible input p95                      |                  `<16ms` |
| Selection paint p95                          |                   `<8ms` |
| Single-cell authoritative ack p95            |                  `<60ms` |
| Collaborator visible update p95              |                 `<150ms` |
| 100-cell paste first authoritative diff p95  |                 `<100ms` |
| 5k-cell paste first authoritative diff p95   |                 `<250ms` |
| Warm-start first useful paint, 100k workbook |                 `<250ms` |
| Warm-start first useful paint, 250k workbook |                 `<700ms` |
| Last-workbook reopen p95                     |                 `<500ms` |
| Pending-op loss                              |                      `0` |
| JS/WASM verified mismatch rate               | `<0.1%` on proven subset |
| Preview/apply mismatch rate                  |                  `<0.5%` |

---

## 4. Product Thesis

**Thesis: `bilig` should become the local-first workbook operating system for serious work.**

### Why this wins

Sheets wins on casual collaboration. Excel wins on semantics. Airtable and Smartsheet win when teams want structure, workflow, and trusted automation above the grid. The gap is obvious: there is still no workbook product whose defining trait is **native-feeling local interaction plus serious spreadsheet depth plus operation-native automation**. That is the opening. ([Google Workspace][2])

### Who it fits first

Start with users whose work is workbook-heavy, collaboration-heavy, and too large or too repetitive for a “just share a sheet” model:

- finance and ops teams living in giant imported workbooks
- analysts and data-heavy PM/ops teams doing recurring reporting
- support/revenue/people/program teams running structured trackers with formulas and views
- anyone doing repeated CSV/XLSX cleanup, merge, summarize, and review work

### What use cases it should dominate first

- giant CSV/XLSX import and cleanup
- collaborative operational trackers with private views
- monthly/weekly reporting workbooks
- formula-heavy review and repair
- workbook restructuring and summary building
- large-sheet editing under shaky network conditions

### What it should not try to dominate first

- consumer note-taking spreadsheets
- presentation/dashboard-first BI
- full peer-to-peer CRDT offline collaboration
- open plugin marketplace
- generic project management with a grid skin
- Python-notebook replacement

### Incumbent weakness it exploits

- Sheets is collaborative, but its product center of gravity is still online-first.
- Excel is deep, but the web product does not feel like a browser-native local runtime.
- Copilot/Gemini help at the side; they do not fundamentally change the execution model.
- Airtable/Smartsheet add workflow and trust layers, but they are not trying to win on giant-data spreadsheet feel. ([Google Help][3])

---

## 5. Target Architecture

### The architecture in one line

**UI shell → runtime worker → local SQLite/OPFS + local engine + tile store → sync bridge → monolith → Postgres → Zero → sync bridge → runtime worker**

### Browser worker runtime

The worker must become the real browser runtime, not an optional side path.

It should own:

- local SQLite database
- local workbook engine / preview engine
- local pending-op journal
- local overlay state
- local tile cache and damage tracking
- selection/editor draft model
- warm-start/checkpoint restore

The UI thread should become a thin shell: input, chrome, accessible focus plumbing, and rendering commands.

### Native grid decomposition

The grid needs a strict split:

- **scene layer**: GPU visuals only
- **text layer**: text layout model, render backend abstracted
- **interaction layer**: pointer/keyboard/fill/drag/resize
- **editor layer**: in-cell edit, formula bar integration
- **a11y layer**: ARIA grid semantics, announcements, focus behavior
- **presentation layer**: metrics, density, tokens

WAI’s grid pattern explicitly includes spreadsheet-style applications, and MDN’s pointer capture and ResizeObserver guidance match the current direction for robust drag and precise sizing. `bilig` should keep using those primitives, then move the grid internals to cleaner boundaries. ([W3C][4])

### Local SQLite layer in browser

Refine the proposed architecture this way:

- **Primary local DB**: SQLite/WASM on **OPFS** in a worker
- **Fallback/bootstrap**: IndexedDB for manifest, recent-workbook list, tiny settings, and fallback when OPFS path is unavailable
- **Not allowed**: syncing raw SQLite pages/files through Zero

Why:

- OPFS is optimized for persistent file-like writes and worker access. ([MDN Web Docs][1])
- The official SQLite WASM path has a first-class OPFS-backed database object, while its lighter browser storage mode targets local/session storage. ([SQLite][5])
- A giant workbook engine needs file-style durability and fast local transactions, not JSON blobs or ad hoc browser object stores.

### IndexedDB persistence model

IndexedDB stays, but it becomes supporting infrastructure:

- bootstrap manifest
- recent workbook registry
- feature flags / session settings
- fallback persistence on non-OPFS path
- small crash-recovery breadcrumbs

It stops being the main runtime database.

### Workbook engine placement

Use **two runtimes** with different jobs:

- **Browser worker runtime**: local preview, local apply, local formula help, local warm-start, local overlay
- **Monolith warm runtime**: shared authoritative apply, revision ordering, heavy recalc, cross-user coordination, import/export finalization

That is the correct split. Browser speed and server trust stop fighting each other.

### Projection/cache model

Replace the current model with:

- **authoritative base tables** in local SQLite
- **pending overlay tables** in local SQLite
- **hot tile cache** in worker memory
- **damage tracker** in worker memory
- **renderer patch emitter** from worker to UI

The grid never reads raw Zero rows directly.
The grid never reads React-held sheet arrays.
The formula bar and selected cell read from the same local store, not from a side channel.

### Zero sync role

Keep Zero. Narrow it further.

Zero becomes:

- authoritative downstream sync for `cell_render`, `cell_input`, rows/cols, styles, views, presence, change feed
- collaborative metadata fan-out
- resumable catch-up feed

Zero is **not**:

- the browser’s primary local database
- the browser’s outbox
- the place where you replicate SQLite internals
- a second UI state model

### Monolith role

Keep the monolith as the authoritative coordinator.

It owns:

- command validation
- authoritative ordering and revision assignment
- warm shared runtime
- heavy recalc scheduling
- import/export finalization
- change bundles
- sync cursors
- AI apply/approval path
- durability guarantees

### Postgres role

Postgres stays the shared durable system of record for:

- workbook source rows
- authoritative render rows
- workbook events / revisions
- change bundles
- collaboration metadata
- checkpoints

### Formula runtime / WASM role

Respect the repo’s current rule:

- JS remains the oracle first
- WASM follows behind parity gates
- production authoritative flip happens only after parity closes

Refinement for local-first UX:

- browser worker uses the same JS oracle path first
- WASM is introduced as an acceleration engine only for verified-equal workloads
- browser and server must share the same command semantics and formula validation rules

### Collaboration model

Make collaboration **server-authoritative, locally buffered**.

- local user sees immediate local result
- server assigns authoritative revision order
- remote diffs arrive through Zero
- local pending ops rebase over new authoritative base
- private views isolate filter/sort state
- selected-cell editor draft is protected from unrelated remote edits

### Operation log / event model

Use two logs:

- **local `pending_op` journal** in browser SQLite
- **authoritative revision log** on server

Upstream sync format should be **semantic command bundles**, not raw row diffs and never SQLite pages.

Downstream sync format should be **authoritative relational deltas**.

That is the right asymmetry.

### Import/export model

- **Local import staging** for CSV/XLSX: parse locally, preview locally, map locally
- **Authoritative finalize** on server: persist canonical source/render, produce change bundle
- **Export** from authoritative server state with compatibility reporting
- For giant local files, parsing should happen in worker threads or dedicated browser workers before server finalization

### Observability model

Add both local and server observability.

Local:

- startup source (`warm-local`, `cold-local`, `network-hydrate`)
- tile hit/miss
- local DB query time
- pending-op queue depth
- rebase count
- conflict count
- patch size
- renderer frame cost

Server:

- authoritative apply latency
- recalc latency by class
- diff size
- Zero publish latency
- reconnect catch-up latency
- import/export throughput

### Rollout model

Roll out by **session mode**, not by mixed paths inside one session.

Bad rollout:

- some interactions use local DB
- some interactions still mutate the old cache path

Good rollout:

- at workbook-open, choose `legacy-session` or `local-runtime-session`
- keep one visible-state pipeline per session
- migrate feature by feature inside the new session architecture, not across parallel browser truths

### What remains

- monolith
- Postgres
- Zero
- typed workbook ops
- native grid
- worker-first direction
- projected rendering model

### What gets refactored

- `runtime-session.ts`
- `runtime-machine.ts`
- `WorkerWorkbookApp.tsx`
- `viewport-cache.ts`
- `ZeroWorkbookBridge.ts`
- `WorkbookGridSurface.tsx`
- selected-cell/edit model
- browser persistence strategy

### What gets removed

- offline-write gating tied to connection state
- JSON snapshot persistence as the primary browser runtime store
- selected-cell synthetic reproject path
- UI-thread optimistic mutation as the main state mechanism
- any path where the grid renders from something other than the worker-owned projected state

### Explicit tradeoffs

- **Server-authoritative local-first**, not CRDT-native peer merge
  Better trust and replay, less magical offline merge.

- **OPFS primary, IndexedDB secondary**
  Better giant-data performance, slightly more browser-path complexity.

- **Command sync upward, relational diffs downward**
  Cleaner reconciliation, less opaque write amplification.

- **One session mode at a time**
  Slower migration, fewer impossible bugs.

- **Incremental renderer cleanup before full GPU text rewrite**
  Faster delivery, less all-or-nothing rewrite risk.

### Assumptions and uncertainties

- **Assumption:** target browsers support OPFS worker path on the serious supported matrix.
- **Uncertainty:** Zero client compatibility inside workers may force a main-thread sync bridge in the first cut.
- **Uncertainty:** local engine checkpoint memory cost on 250k+ workbooks needs bench validation.
- **Uncertainty:** multi-tab write coordination should start with one active writer per workbook/profile until the lock model is hardened.

---

## 6. Local-First Runtime Model

### Storage layout

Use:

- `catalog.sqlite`
  recent workbooks, last-open viewport, sync cursors, settings, feature flags

- `workbook_<id>.sqlite`
  actual workbook-local state

Suggested local tables:

- `sheet`
- `sheet_row`
- `sheet_col`
- `cell_input_base`
- `cell_render_base`
- `cell_input_overlay`
- `cell_render_overlay`
- `cell_style`
- `number_format`
- `defined_name`
- `table_def`
- `pivot_def`
- `spill_owner`
- `pending_op`
- `pending_op_effect`
- `sync_cursor`
- `checkpoint`
- `local_change_bundle`

Keep `sheet_id` stable in local and server schemas.

### How reads work

1. UI asks the runtime worker for visible tiles and selected-cell state.
2. Worker checks hot in-memory tile cache.
3. Cache miss reads from local SQLite base + overlay.
4. Worker composes visible projection and emits a minimal renderer patch.
5. Grid paints from that patch only.

Selected cell works the same way:

- grid render reads local `cell_render`
- formula bar reads local `cell_input`
- editor draft overlays that local state
- no selected-cell query is allowed to bypass the worker-owned local model for rendering

That eliminates the current special-case selected-cell bridge problem.

### How writes work

1. User action becomes a **semantic command bundle**.
2. Worker opens a local SQLite transaction.
3. Bundle is appended to `pending_op` with:
   - `op_id`
   - `client_id`
   - `local_seq`
   - `base_revision`
   - `scope`
   - `payload`
   - `status`

4. Worker applies the bundle locally through the local engine / deterministic op executor.
5. Worker writes overlay rows and emits damage to the renderer immediately.
6. UI shows the result now, not after network.

A local write is considered “accepted” only after local SQLite commit succeeds.

### How sync works

Use a split-direction model.

**Upstream**

- worker outbox drains `pending_op` in order
- bundles go to a monolith command-apply endpoint
- each bundle carries `base_revision` and idempotency key
- successful submit marks the local op as `submitted`; it is removed from the journal only when the authoritative revision feed absorbs that client mutation

**Downstream**

- authoritative revision/event feed is the thing that clears submitted local ops
- Zero carries authoritative row deltas, change bundles, presence, and views back to all clients
- worker ingests authoritative revision events first and falls back to full snapshot hydrate only when the event feed cannot close the gap

The browser does **not** sync SQLite contents into Zero.
It syncs semantic intent upward and consumes authoritative deltas downward.

### How local projections work

The local read model is:

`visible_state = authoritative_base + pending_overlay + editor_draft_overlay`

Where:

- **authoritative_base** = last confirmed server state in local SQLite
- **pending_overlay** = local accepted but not yet authoritative ops
- **editor_draft_overlay** = text currently being typed and not yet committed

That gives the product three things incumbents still struggle with:

- no lost typing
- instant local visibility
- clean authoritative reconciliation

### Operation ordering

- local order: `local_seq`
- shared order: `revision`
- a pending op is always tied to the `base_revision` it was created against
- when authoritative revisions arrive, pending ops are replayed or rebased on top of the new base

### Conflict handling

Start with a simple, strict model.

- **Non-overlapping edits**: apply authoritative delta, replay pending local ops, no drama.
- **Remote edit to visible cell outside active editor**: apply authoritative delta, subtle reconcile animation.
- **Remote edit to same cell while user is typing**: keep draft local; show remote-change badge; do not clobber the editor.
- **User submits a draft on stale base revision**:
  - attempt trivial rebase first
  - if non-trivial, show compare affordance with explicit choice:
    - apply my value/formula
    - keep authoritative
    - duplicate mine elsewhere / copy to formula bar buffer

First version should optimize for zero lost work, not magical merge heroics.

### Recovery after reconnect

On reconnect:

1. fetch latest authoritative cursor
2. ingest missing authoritative deltas through Zero
3. rebase pending local ops
4. continue draining outbox
5. refresh hot visible tiles

Target outcome:

- same viewport
- same draft
- same local pending changes
- no full-sheet flash

### Durability model

- pending ops are written to SQLite before the UI considers them committed locally
- checkpoints are written periodically for warm-start
- crash recovery loads:
  - last authoritative base
  - pending op journal
  - optional engine checkpoint
  - last viewport/selection

This is the difference between “optimistic UI” and “local-first product.”

### Large workbook paging / tiling

Keep stable tiles.
The repo already uses `128 x 32`; that is a sane starting point.

Rules:

- tile-shaped local queries only
- hot tile LRU in worker memory
- authoritative full workbook lives on disk, not in React memory
- neighbor preload ring of 1 tile
- last-view heat-based retention
- row/column metadata separate from cell tables
- style ids/materials cached separately

Do **not** keep full sheet render arrays in React state.

### Startup / warm-start / hydration

Startup should look like this:

1. open local DB
2. render last known workbook chrome + first visible tile from local DB
3. restore selection and scroll position
4. start sync catch-up in background
5. refresh visible region when authoritative deltas land

That gives instant feel even with weak network.

### Why this feels native

Because every critical user loop becomes local:

- selection is local
- scrolling is local
- typing is local
- visible cell projection is local
- draft durability is local
- startup is local

Network becomes shared truth plumbing, not the source of immediacy.

---

## 7. Top Strategic Bets

| Bet                                              | Why it matters                                                                         | User outcome                                    | Depends on                              | Codebase changes                                                  | Risk   | Payoff    |
| ------------------------------------------------ | -------------------------------------------------------------------------------------- | ----------------------------------------------- | --------------------------------------- | ----------------------------------------------------------------- | ------ | --------- |
| 1. Make the worker runtime the real mounted path | Current product cannot feel native while hot state stays outside a real runtime worker | Immediate feel, lower UI-thread contention      | runtime session rewrite                 | `runtime-session.ts`, `workbook.worker.ts`, `worker-runtime.ts`   | Medium | Very high |
| 2. OPFS-backed SQLite local DB                   | Giant-data local durability and warm-start need a real DB                              | Offline-safe, instant reopen                    | browser storage layer                   | new browser-sqlite package/module, de-emphasize `storage-browser` | Medium | Very high |
| 3. Local semantic op journal + rebase            | Local-first without durable pending ops is fake                                        | Survives crash/offline, clean reconcile         | typed ops                               | `packages/workbook-domain`, worker runtime, server apply endpoint | High   | Very high |
| 4. Single browser truth store                    | Fixes selected-cell drift, editor loss, stale-path bugs                                | Sharper editing and fewer weird bugs            | local DB + overlay model                | `ZeroWorkbookBridge`, `viewport-cache`, editor state              | Medium | Very high |
| 5. Grid shell decomposition                      | Giant hot-path files will keep slowing product quality                                 | Faster iteration, fewer regressions             | none                                    | `WorkbookGridSurface.tsx`, `WorkerWorkbookApp.tsx`                | Low    | High      |
| 6. Text/layout cleanup path                      | DOM text and browser-only autofit cap polish                                           | Crisper, more consistent grid                   | grid interfaces                         | `gridTextScene.ts`, `GridTextOverlay.tsx`, autofit service        | Medium | High      |
| 7. Giant-workbook warm-start path                | Large data is the wedge                                                                | Opens feel instant instead of heavy             | local DB + tile model                   | worker tile store, startup path, checkpoints                      | Medium | Very high |
| 8. Private views + changes + presence            | Core workflow differentiation without bloating the product                             | Collaboration stops feeling destructive         | stable identities + change feed         | Zero schema/queries, UI rails                                     | Medium | High      |
| 9. Local-first import and cleanup                | Huge CSV/XLSX workflows are where current tools feel clumsy                            | Immediate preview, faster cleanup               | local DB + worker parse                 | browser workers, import pipeline, server finalize                 | High   | High      |
| 10. Plan/preview/apply AI over workbook ops      | Beats chat-sidecars by turning AI into a real execution model                          | Faster workbook building and repair, with trust | typed ops, preview engine, undo bundles | AI tools, command preview, changes pane                           | High   | Very high |

---

## 8. Execution Roadmap

### Status as of 2026-04-06

**Completed from the original roadmap**

- the worker runtime is the mounted browser session path
- writes are no longer gated on simple Zero connection state
- the selected-cell special reproject path has been deleted
- `WorkerWorkbookApp.tsx`, `WorkbookGridSurface.tsx`, and `worker-runtime.ts` have been split below the hot-path ceiling
- OPFS SQLite is now the primary browser workbook store
- warm-start now comes from the local SQLite state path rather than JSON persistence
- the worker has a persisted pending-mutation queue with crash-safe journal replay
- `viewport-cache.ts` has been deleted and replaced with `projected-viewport-store.ts`

**Still not completed**

- stable `sheet_id` across browser/server/local layers
- a fully worker-owned tile store backed by normalized local DB tables for both base and overlay reads
- collaboration/product layers in Phases 2 through 4

**What this roadmap now needs to do**

- focus only on the still-missing production work
- treat worker cutover and hot-path file splits as done
- stop referring to deleted bridge behavior as active architecture

### Prioritized initiative table

| Priority | Initiative                                                  | Why now                                                            |
| -------- | ----------------------------------------------------------- | ------------------------------------------------------------------ |
| 1        | Add stable `sheet_id` across local/server/browser layers    | Needed for views, changes, comments, tasks later                   |
| 2        | Move projected tiles fully behind worker-owned local tables | Giant-data warm-start and ingest still depend on in-memory patches |
| 3        | Add storage and reconnect failure harnesses                 | The local-first path now exists; it needs production-grade failure proofing |
| 4        | Add private views, changes pane, collaborator jump          | Best near-term workflow differentiation                            |
| 5        | Build plan/preview/apply AI on semantic bundles             | Biggest differentiated UX after local-first core                   |

### Dependency list

- real worker runtime → local SQLite → normalized local base → overlay/rebase
- stable `sheet_id` → views, change bundles, anchored metadata
- local projection model → selected-cell cleanup → reliable editor model
- normalized local base tables → overlay tables → giant workbook warm-start → smooth collaboration ingest
- typed command bundles → preview/apply AI → undo bundles → changes pane

### Phase 0: architecture hardening

**Objective**
Stop fighting the current browser architecture.

**User-visible outcomes**

- cleaner editing behavior
- fewer selection/editor glitches
- lower interaction jank

**Engineering outcomes**

- actual worker runtime mounted
- hot-path files carved apart
- selected-cell side path isolated for deletion
- storage benchmark harness in place

**Key epics**

- wire `workbook.worker.ts` / `worker-runtime.ts` into `runtime-session.ts`
- split `WorkerWorkbookApp.tsx`
- split `WorkbookGridSurface.tsx`
- add `sheet_id`
- add storage harness comparing current JSON persistence vs OPFS SQLite
- remove write gating from simple connection state

**Dependencies**

- none

**Risks**

- migration regressions in selection/editing
- worker transport edges

**Exit criteria**

- worker runtime is the active session path
- no hot-path file over ~900 LoC
- selected-cell render no longer depends on synthetic one-cell patching
- offline-safe Class A edits work on the new path in test mode

### Phase 1: local-first foundation

**Objective**
Make `bilig` genuinely local-first.

**User-visible outcomes**

- local durable edits
- instant reopen of recent workbooks
- usable workbook interaction under poor network or offline

**Engineering outcomes**

- OPFS SQLite local DB
- local base/overlay tables
- pending-op journal with submitted-state durability
- reconnect/rebase path
- authoritative ingest into local base tables

**Key epics**

- implement local SQLite storage layer
- add worker tile store + damage tracker
- implement `pending_op` journal
- build upstream command sync
- build downstream authoritative delta ingest
- warm-start from local DB and checkpoint

**Dependencies**

- Phase 0 worker path
- stable `sheet_id`

**Risks**

- browser storage edge cases
- local/server divergence bugs
- checkpoint size

**Exit criteria**

- 100% locally accepted ops survive refresh/crash
- warm-start last workbook p95 `<500ms`
- local visible input p95 `<16ms`
- reconnect with 100 pending ops p95 `<2s`
- no pending-op loss in failure harness

### Phase 2: collaboration and workflow superiority

**Objective**
Make shared workbook work smoother than Sheets for serious day-to-day use.

**User-visible outcomes**

- private views
- named views
- presence and jump-to-collaborator
- show changes
- conflict UX that preserves drafts

**Engineering outcomes**

- change bundle model
- presence feed
- view model
- editor/remote reconcile guarantees

**Key epics**

- `workbook_change` UI and pipeline
- private/saved view state
- collaborator presence and location jump
- same-cell conflict compare flow
- named versions / undo bundles

**Dependencies**

- Phase 1 local rebase model
- stable sheet/object anchors

**Risks**

- collaboration UI bloat
- reconcile edge cases

**Exit criteria**

- no lost typing during remote edits
- private views do not disturb others
- change bundle jump/revert works
- > 50% of shared pilot workbooks use views, changes, or collaborator jump

### Phase 3: agent-native workbook platform

**Objective**
Turn AI into a real workbook execution tool.

**User-visible outcomes**

- plan → preview → apply
- formula repair and workbook cleanup
- structured workbook edits with undo bundles
- local preview before commit

**Engineering outcomes**

- preview engine
- AI tool surface over workbook ops
- approval plumbing
- replayable agent execution records

**Key epics**

- command preview/diff
- MCP/internal tool layer
- AI pane with scope control
- approval models by risk class
- undoable command bundles

**Dependencies**

- local projection model
- typed ops
- change bundles

**Risks**

- model quality disappointment
- preview/apply mismatch
- scope creep into generic chat product

**Exit criteria**

- first plan token p95 `<700ms`
- first preview highlight p95 `<1000ms`
- first accepted mutation p95 `<1500ms`
- 100% agent writes use semantic command bundles

### Phase 4: category-defining capabilities

**Objective**
Make giant-data workbook work obviously better than incumbents.

**User-visible outcomes**

- near-instant local staging of large imports
- workbook comprehension tools
- scenario/scratchpad flows for heavy analysis
- high-confidence large-model cleanup and restructure

**Engineering outcomes**

- local import staging engine
- workbook semantic index
- scenario/checkpoint tools
- larger-scale runtime budgets

**Key epics**

- local CSV/XLSX staging + preview
- broken-formula and dependency inspector
- semantic workbook search/index
- scenario scratchpads / temporary branches
- giant-workbook benchmark corpus

**Dependencies**

- all prior phases

**Risks**

- performance budget overruns
- product sprawl

**Exit criteria**

- 250k workbook warm-start first useful paint `<700ms`
- staged import preview for large files in seconds, not minutes
- workbook comprehension actions return near-instant on warm runtime

### Risk matrix

| Risk                                 | Likelihood | Impact | Mitigation                                                              |
| ------------------------------------ | ---------- | ------ | ----------------------------------------------------------------------- |
| OPFS/browser-path quirks             | Medium     | High   | explicit fallback path, browser matrix harness, feature flag            |
| Worker runtime migration regressions | High       | High   | session-mode rollout, golden workbook tests, shadow telemetry           |
| Local/server divergence bugs         | Medium     | High   | deterministic command model, replay tests, rebase harness               |
| Conflict UX becomes annoying         | Medium     | Medium | protect draft first, keep first model simple, instrument conflict rates |
| Renderer work expands too early      | High       | Medium | interface-first cleanup, defer full GPU text until local runtime stable |
| AI scope creep                       | High       | Medium | semantic tools only, preview/apply discipline, keep chat narrow         |

---

## 9. What To Postpone

Postpone these until the local-first core is solid:

- **full GPU text atlas rewrite**
  Do the scene/text interface cleanup now, but land the full text backend after the worker runtime and local DB are stable.

- **long-tail Excel parity**
  Keep parity discipline on the targeted worksheet surface. Avoid spending months on obscure function tails while the product still lacks a great local runtime.

- **large chart/dashboard push**
  Core grid, import, formula, views, and changes matter more than beautiful secondary visualization.

- **plugin marketplace**
  Build semantic internal tool surfaces first. External platformization can wait.

- **Python-in-grid / arbitrary code cells**
  That is the wrong complexity class for the next 12 months.

- **peer-to-peer merge or CRDT-native shared editing**
  The right model here is server-authoritative local-first, not distributed merge-first.

- **multi-tab concurrent writers for the same workbook**
  Start with one active writer per workbook/profile and harden the coordination model before expanding.

- **generic AI chat breadth**
  Q&A is fine. The priority is semantic plan/preview/apply over workbook ops.

- **heavy comments/tasks bureaucracy**
  Views, changes, presence, and conflict-safe editing add more value earlier.

- **visual novelty work**
  Quiet dense tool-like UI is the right design language for this product. The repo’s own UI rules are right about that.

---

## 10. AI/Agent Strategy

Sheets and Excel already offer AI-assisted sheet creation, formula help, analysis, charts, pivots, and editing. Excel also layers Python into the workbook surface. `bilig` only wins if AI is **faster, safer, and more operationally real**, not just another side panel. ([Google Help][3])

### Core rule

The model plans.
The workbook engine verifies.
The runtime previews.
The server authoritatively applies.

### What agents should do

- create formulas
- repair broken formulas
- clean imported data
- restructure sheets and tables
- create pivots/charts/views
- batch-format and relabel
- explain workbook logic
- generate summary sheets
- convert messy data into structured workbook flows

### What agents should not do

- drive the DOM
- mutate unseen workbook scope silently
- bypass parser/binder/validator
- rewrite large workbook areas without preview
- overwrite in-progress user drafts
- create external side effects without explicit approval

### Execution model

Every mutating AI action is a **command bundle**:

- scope
- requested intent
- concrete workbook ops
- affected sheets/ranges
- estimated cost
- local preview diff
- undo bundle

### Verification loop

1. model reads workbook context
2. model proposes command bundle
3. deterministic validation runs
4. runtime worker previews locally
5. UI shows range highlights and diff summary
6. user accepts whole plan or scoped subset
7. monolith applies authoritatively
8. authoritative diff lands back through sync
9. change bundle is recorded and undoable

### Approval model

Use risk classes.

- **low-risk**: local formatting inside current selection
  can allow one-click auto-apply inside strict scope

- **medium-risk**: formula generation in selected table, cleanup transforms
  preview required

- **high-risk**: sheet structure changes, workbook-wide transforms, external system actions
  explicit approval required

### Replayability

Persist:

- goal text
- workbook context references
- plan
- preview
- accepted scope
- applied revision
- undo bundle id
- result summary

That turns AI from “chat transcript” into “replayable workbook operation.”

### Deterministic boundaries

These stay deterministic:

- formula parsing
- reference binding
- validation
- local preview
- apply
- undo
- rebase
- export
- sync reconciliation

The model never becomes execution truth.

### Why this beats sidebar copilots

Because `bilig` can preview locally in milliseconds, apply as semantic bundles, survive offline planning, and reconcile against the same operation model the workbook already uses. That is a fundamentally better product surface than “assistant suggests something, then the product sort of does it.”

---

## 11. Repo Execution Map

### `apps/web`

#### `/Users/gregkonush/github.com/bilig/apps/web/src/WorkerWorkbookApp.tsx`

Current problem: it owns too much.

Change it into:

- `WorkbookShell.tsx`
- `useWorkbookSession.ts`
- `useWorkbookCommands.ts`
- `useWorkbookEditor.ts`
- `useWorkbookSelection.ts`
- `useWorkbookViews.ts`
- `useWorkbookChangesRail.ts`
- `useWorkbookAiRail.ts`

Remove from the component:

- direct optimistic cache mutation
- selection-to-data lookup logic
- bridge-specific state wiring
- editing/reconcile rules

The component becomes composition only.

#### `/Users/gregkonush/github.com/bilig/apps/web/src/runtime-session.ts`

Make this the composition root for the **real worker runtime**.

New responsibilities:

- boot worker runtime
- mount local DB session
- bridge UI ↔ worker
- bridge sync feed ↔ worker
- own bootstrap mode (`legacy` or `local-runtime`)

It should stop constructing only a thin projected viewport store plus live Zero subscriptions as the session substrate.

#### `/Users/gregkonush/github.com/bilig/apps/web/src/runtime-machine.ts`

Expand the state machine to reflect the real lifecycle:

- `booting`
- `hydratingLocal`
- `localReady`
- `syncing`
- `live`
- `offline`
- `reconciling`
- `recovering`
- `failed`

That gives the shell correct save/sync/status behavior.

#### `/Users/gregkonush/github.com/bilig/apps/web/src/projected-viewport-store.ts`

The old `viewport-cache.ts` file is now deleted. The replacement projected store should stay narrow and temporary rather than quietly becoming the next browser truth store.

Split/replace it with:

- `ViewportTileStore` (worker-side)
- `OverlayStore` (worker-side)
- `DamageTracker`
- `RendererPatchEmitter`

The current `MAX_CACHED_CELLS_PER_SHEET = 6000` model is still only a stopgap.
Hot visible cells belong in worker memory; full authoritative state belongs in local SQLite.

#### `/Users/gregkonush/github.com/bilig/apps/web/src/zero/ZeroWorkbookBridge.ts`

Refactor into an **AuthoritativeSyncBridge**.

Keep:

- tile subscriptions
- workbook chrome metadata subscriptions
- presence / change feed subscriptions

Delete:

- selected-cell synthetic reproject path
- any render-path reliance on special selected-cell queries

New behavior:

- ingest authoritative deltas into worker local DB
- notify worker of damaged tiles
- keep grid render fully local

#### `apps/web/src/workbook.worker.ts` and `apps/web/src/worker-runtime.ts`

These become the center of gravity.

Add:

- SQLite/WASM bootstrap
- OPFS-backed persistence
- `pending_op` journal
- checkpointing
- local preview/apply
- overlay/base table composition
- warm-start logic
- tile query APIs for UI
- selected cell/formula bar APIs
- rebase/reconnect APIs

Demote current JSON snapshot persistence to fallback/bootstrap use only.

### `packages/grid`

#### `/Users/gregkonush/github.com/bilig/packages/grid/src/WorkbookGridSurface.tsx`

Turn it into a composition root only.

Carve out:

- `GridViewportController`
- `GridSelectionController`
- `GridEditorOverlay`
- `GridFillHandleController`
- `GridResizeController`
- `GridA11yLayer`
- `GridAutofitService`

#### `/Users/gregkonush/github.com/bilig/packages/grid/src/WorkbookView.tsx`

Keep as shell composition seam.
Add explicit rails/slots for:

- changes
- views
- collaborator jump
- AI preview highlights

#### `/Users/gregkonush/github.com/bilig/packages/grid/src/gridGpuScene.ts`

Keep visuals only:

- fills
- borders
- selection
- frozen lines
- diff/highlight overlays

No workbook business logic.

#### `/Users/gregkonush/github.com/bilig/packages/grid/src/gridTextScene.ts`

Convert this into a pure text layout model.
It should output layout objects that can be rendered by:

- current DOM text overlay
- later GPU text backend

#### DOM text path

Keep `GridTextOverlay.tsx` as the temporary renderer.
Add a stable text-measurement interface so autofit stops depending directly on ad hoc browser measurement calls.

#### `/Users/gregkonush/github.com/bilig/packages/grid/src/gridInteractionController.ts`

Narrow it to pointer/keyboard state transitions only.
Keep pointer capture semantics; that part is aligned with platform guidance already. ([MDN Web Docs][6])

#### `/Users/gregkonush/github.com/bilig/packages/grid/src/gridPresentation.ts`

#### `/Users/gregkonush/github.com/bilig/packages/grid/src/gridMetrics.ts`

Move hardcoded values into tokenized density/metrics contracts.
The repo’s UI philosophy is right here: state, behavior, and rendering should stay separate.

### `packages/core`

Use `packages/core` for deterministic local/server workbook mechanics.

Add:

- local checkpoint import/export helpers
- preview/apply effects summary
- rebase helpers for pending ops
- deterministic diff generation
- selected-cell formula/value resolution helpers
- verified JS/WASM comparison hooks

The browser and server need the same deterministic workbook semantics.

### `packages/zero-sync`

Keep it narrow and tile-first.

Add / change:

- stable `sheet_id`
- tile queries that stay stable under scroll
- `workbook_change`
- `presence_coarse`
- `sheet_view`

Remove client dependence on:

- whole-sheet aggregation in React state
- selected-cell render special casing
- anything that assumes the client render model is Zero-first instead of local-first

### `apps/bilig`

Keep the monolith and strengthen it.

Add:

- command-apply endpoint for browser outbox
- revision ack + authoritative diff response
- preview endpoint for AI and heavy local/server compare
- change bundle builder
- better checkpoint APIs
- import finalize pipeline
- rebase-aware apply errors
- stable `sheet_id` persistence

Strengthen `workbook-runtime/runtime-manager.ts` with:

- preview sessions
- lightweight checkpoint export for browser warm-start
- explicit diff payload helpers
- memory budgeting by workbook size class

### Docs / RFCs

Rewrite or add:

- `docs/architecture.md`
- new `docs/local-first-runtime.md`
- new `docs/browser-sqlite-opfs-rfc.md`
- new `docs/command-journal-rebase-rfc.md`
- new `docs/grid-renderer-v3.md`
- update `docs/excel-parity-program.md` with browser/server shared execution plan
- update `docs/react-spectrum-ui-philosophy.md` with grid-specific a11y/runtime rules
- new `docs/session-rollout-mode-rfc.md`

### Tests / benchmarks / acceptance gates

Add:

- worker runtime boot/recover tests
- local DB crash/reload durability tests
- offline edit + reconnect + rebase tests
- same-cell conflict/editor-draft tests
- multi-user authoritative ordering tests
- giant workbook warm-start benchmarks
- storage-path matrix tests (`OPFS`, fallback)
- JS vs WASM differential tests
- tile-cache hit/miss benchmarks
- change bundle correctness tests
- AI preview/apply/undo consistency tests

Acceptance gates:

- no hot-path browser state path outside worker runtime for local-runtime sessions
- no selected-cell synthetic render patch path
- no whole-sheet arrays in React state
- warm-start and input SLOs enforced in CI perf harness
- feature rollout only by session mode, never mixed in-session state paths

---

## 12. Executive Decision Memo

**One-sentence thesis**
`bilig` should become the **local-first workbook operating system**: a browser-resident spreadsheet runtime that feels native because edits, projections, startup, and AI previews all happen locally first, while the server stays authoritative for shared truth.

**Top 5 strategic bets**

1. make the worker runtime the real mounted path
2. move workbook state into OPFS-backed SQLite in the browser
3. build a durable local op journal with authoritative rebase
4. unify visible state so grid, selection, and formula bar read from one local model
5. make AI operate on previewable command bundles

**Top 5 engineering priorities**

1. refactor `runtime-session.ts` around the actual worker runtime
2. replace `viewport-cache.ts` with worker-owned tile/overlay stores
3. remove the selected-cell synthetic path from `ZeroWorkbookBridge.ts`
4. split `WorkbookGridSurface.tsx` and `WorkerWorkbookApp.tsx` into stable layers
5. add stable `sheet_id` and local/server checkpoint + rebase plumbing

**Top 5 UX upgrades**

1. instant local edits that survive offline and refresh
2. warm-start last workbook from local DB
3. conflict-safe editing that never loses a draft
4. private views + changes pane + collaborator jump
5. formula and selected-cell behavior sourced from one truth

**Top 5 architecture risks**

1. browser storage/OPFS edge cases
2. local/server divergence during rebase
3. migration pain while switching to real worker-owned state
4. renderer rewrite expanding too early
5. multi-tab coordination complexity

**Top 5 product risks**

1. shipping “local-first” before durability is truly real
2. keeping dual browser state paths too long
3. letting AI become chat garnish instead of workbook execution
4. spending too much time on parity tail before local runtime feels superb
5. diluting the roadmap with secondary surfaces before giant-data UX is solved

**Recommended next move**
Take one workbook session path and make it real: wire the existing runtime worker into `apps/web`, back it with OPFS SQLite, add a local pending-op journal plus authoritative rebase, and route the grid to a worker-owned projected state model. That is the shortest path to a workbook product that actually feels 10x faster, sharper, and more trustworthy than Sheets or Excel web.

[1]: https://developer.mozilla.org/en-US/docs/Web/API/File_System_API/Origin_private_file_system "https://developer.mozilla.org/en-US/docs/Web/API/File_System_API/Origin_private_file_system"
[2]: https://workspace.google.com/products/sheets/ "https://workspace.google.com/products/sheets/"
[3]: https://support.google.com/docs/answer/14356410?hl=en "https://support.google.com/docs/answer/14356410?hl=en"
[4]: https://www.w3.org/WAI/ARIA/apg/patterns/grid/ "https://www.w3.org/WAI/ARIA/apg/patterns/grid/"
[5]: https://sqlite.org/wasm/doc/trunk/api-oo1.md "https://sqlite.org/wasm/doc/trunk/api-oo1.md"
[6]: https://developer.mozilla.org/en-US/docs/Web/API/Element/setPointerCapture "https://developer.mozilla.org/en-US/docs/Web/API/Element/setPointerCapture"
