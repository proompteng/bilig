# bilig × Zero v4
## Hard-cut monolith plan with AX and enterprise UX

## 1. What is wrong in the current repo

The current repo is already close enough that the bottlenecks are visible.

### 1.1 Current shape in the repo

Relevant code already exists:

- `packages/zero-sync`
- `apps/web/src/zero/ZeroWorkbookBridge.ts`
- `apps/web/src/zero/tile-subscriptions.ts`
- `apps/sync-server/src/zero/service.ts`
- `apps/sync-server/src/zero/server-mutators.ts`
- `apps/sync-server/src/zero/store.ts`
- `apps/local-server/src/local-workbook-session-manager.ts`
- `packages/workbook-domain`
- `packages/agent-api`

The repo already contains the right raw ingredients, but the data flow is wrong for performance.

### 1.2 The four current performance killers

#### A. Hot persistence is still snapshot-centric

`packages/zero-sync/src/schema.ts` keeps `snapshot` and `replicaSnapshot` on `workbooks`.
`apps/sync-server/src/zero/store.ts` reads and writes those fields directly.

That keeps the authoritative write path anchored to full workbook snapshots instead of to a strict source/render relational model.

#### B. Server mutators still cold-boot the engine per mutation

`apps/sync-server/src/zero/server-mutators.ts` loads workbook state, creates a fresh `SpreadsheetEngine`, imports the snapshot, applies the mutation, and exports a new snapshot.

That makes hot edit latency scale with engine restore cost and workbook size.

#### C. Persistence still diffs the entire workbook projection on every mutation

`persistWorkbookMutation(...)` in `apps/sync-server/src/zero/store.ts` builds `previousProjection` and `nextProjection` from whole-workbook snapshots and then runs `diffProjectionRows(...)` across sheets, cells, metadata, styles, formats, and render rows.

This is better than full table truncation, but it is still whole-workbook work per edit.

#### D. Client Zero consumption still allocates too much

`apps/web/src/zero/tile-subscriptions.ts` creates multiple `zero.materialize()` views per tile and concatenates arrays from every tile on every update.
`packages/zero-sync/src/queries.ts` also keeps workbook-wide style/format registry queries and duplicate `computedCells` / `cellEval` query shapes.

The UI query surface is much narrower than older drafts, but it still does more allocation and invalidation work than necessary.

### 1.3 Current web bridge still does too much array rebuilding

`apps/web/src/zero/tile-subscriptions.ts` still rebuilds aggregate arrays per tile update, and it does it on top of workbook-wide style/format registry reads.

That is acceptable for bring-up. It is not the architecture for Excel-class polish.

---

## 2. Non-negotiable target architecture

### 2.1 Product truth

This product should be described internally as:

**server-authoritative multiplayer spreadsheet with local-first feel**

Not:

- true offline-write local-first
- CRDT-native merge-first spreadsheet
- workbook-snapshot-over-sync spreadsheet

### 2.2 Deployment topology

#### Target production topology

- `bilig-app` deployment
- `bilig-zero` deployment
- `bilig-db` Postgres cluster
- optional object storage only if checkpoint retention or export volume justifies it
- no Redis

#### Routing

Single public origin:

- `/` -> web shell from `bilig-app`
- `/api/zero/query` -> `bilig-app`
- `/api/zero/mutate` -> `bilig-app`
- `/api/agent/*` -> `bilig-app`
- `/zero` -> `bilig-zero`
- `/v1/frames` -> branch bring-up only; removed from production routing before merge

#### What gets removed from production

- separate `bilig-sync` deployment
- separate `local-server` deployment model
- remote `@bilig/agent-api` network surface
- Redis-based coordination/presence path

### 2.3 Monolith module layout

Inside `bilig-app`, use modules, not services:

- `src/http` – Fastify app, auth, static assets, public API
- `src/zero` – query registry, mutators, endpoint handlers, publication config
- `src/workbook-runtime` – warm workbook sessions, engine lifecycle, session eviction
- `src/recalc` – incremental calc scheduling, dirty-set tracking, render materialization
- `src/domain` – canonical workbook command model extracted from `packages/workbook-domain`
- `src/import-export` – XLSX import/export, compatibility transforms
- `src/ax` – Codex app-server orchestration, AX state, approvals, plan/apply
- `src/skills` – agent skill definitions over optimized internal workbook APIs
- `src/background` – PG-backed work loops (`SKIP LOCKED`, advisory locks, heartbeats)
- `src/presence` – coarse presence, collaborator navigation, session expiration
- `src/telemetry` – metrics, tracing, query analysis, SLO alerts

This gives one deployable app while preserving clean boundaries.

---

## 3. Required architectural change: strict source/render split

This is the core design rule.

### 3.1 Source state

Source state is what users edit semantically:

- workbook/sheet metadata
- raw cell input values
- raw formula text
- styles / number formats
- row/column structure
- names / tables / pivots / validations / protected ranges
- durable workbook change log

### 3.2 Render state

Render state is what the viewport actually needs:

- authoritative visible computed values
- cell flags / error state / spill flags / formula display mode
- effective style reference
- visible row/col metadata
- sheet tabs and workbook chrome
- coarse collaborator presence

### 3.3 Rule

Never sync these through Zero hot tables:

- workbook snapshot blobs
- replica snapshots
- CRDT clocks or transport artifacts
- compiled ASTs
- dependency graph internals
- WASM heap/runtime state
- calc scheduler internals

Persist them only as private server/runtime artifacts when needed.

---

## 4. Target relational model

Entity names in this section describe the target data model. During implementation, keep current pluralized SQL/Zero naming where that avoids pointless churn. This migration is about fixing hot-path behavior and ownership boundaries, not spending time on cosmetic renames.

## 4.1 Public / synced tables

### `workbook`

- `id uuid pk`
- `name text`
- `owner_user_id uuid`
- `head_revision bigint`
- `calculated_revision bigint`
- `calc_mode text`
- `compatibility_mode text`
- `recalc_epoch bigint`
- `created_at timestamptz`
- `updated_at timestamptz`

### `workbook_member`

- `workbook_id uuid`
- `user_id uuid`
- `role text`
- `created_at timestamptz`

### `sheet`

- `id uuid pk`
- `workbook_id uuid`
- `name text`
- `position int`
- `freeze_rows int`
- `freeze_cols int`
- `tab_color text null`
- `hidden boolean`
- `created_at timestamptz`
- `updated_at timestamptz`

### `sheet_view`
Per-user private view state, modeled after Excel Sheet Views / Google filter views.

- `id uuid pk`
- `workbook_id uuid`
- `sheet_id uuid`
- `owner_user_id uuid`
- `name text`
- `kind text` (`temporary`, `saved`)
- `filter_json jsonb`
- `sort_json jsonb`
- `slicer_json jsonb`
- `is_default boolean`
- `updated_at timestamptz`

### `sheet_row`
Sparse non-default row metadata.

- `workbook_id uuid`
- `sheet_id uuid`
- `row_num int`
- `axis_id uuid null`
- `height int null`
- `hidden boolean`
- `outline_level int`
- `source_revision bigint`
- `updated_at timestamptz`

PK `(workbook_id, sheet_id, row_num)`

### `sheet_col`
Sparse non-default column metadata.

- `workbook_id uuid`
- `sheet_id uuid`
- `col_num int`
- `axis_id uuid null`
- `width int null`
- `hidden boolean`
- `outline_level int`
- `source_revision bigint`
- `updated_at timestamptz`

PK `(workbook_id, sheet_id, col_num)`

### `cell_style`

- `workbook_id uuid`
- `style_id uuid`
- `style_json jsonb`
- `hash text`
- `created_at timestamptz`

PK `(workbook_id, style_id)`

### `number_format`

- `workbook_id uuid`
- `format_id uuid`
- `kind text`
- `code text`
- `created_at timestamptz`

PK `(workbook_id, format_id)`

### `cell_input`
Editable source.

- `workbook_id uuid`
- `sheet_id uuid`
- `row_num int`
- `col_num int`
- `input_json jsonb null`
- `formula_source text null`
- `style_id uuid null`
- `format_id uuid null`
- `editor_text text null`
- `source_revision bigint`
- `updated_by uuid`
- `updated_at timestamptz`

PK `(workbook_id, sheet_id, row_num, col_num)`

### `cell_render`
Authoritative visible render state. Sparse.

- `workbook_id uuid`
- `sheet_id uuid`
- `row_num int`
- `col_num int`
- `value_tag text`
- `number_value double precision null`
- `string_value text null`
- `boolean_value boolean null`
- `error_code text null`
- `style_id uuid null`
- `format_id uuid null`
- `flags int`
- `calc_revision bigint`
- `updated_at timestamptz`

PK `(workbook_id, sheet_id, row_num, col_num)`

### `defined_name`

- `workbook_id uuid`
- `scope_sheet_id uuid null`
- `name text`
- `normalized_name text`
- `value_json jsonb`
- `source_revision bigint`

PK `(workbook_id, scope_sheet_id, normalized_name)`

### `table_def`

- `workbook_id uuid`
- `table_id uuid`
- `sheet_id uuid`
- `name text`
- `start_row int`
- `start_col int`
- `end_row int`
- `end_col int`
- `column_names_json jsonb`
- `header_row boolean`
- `totals_row boolean`
- `source_revision bigint`

PK `(workbook_id, table_id)`

### `pivot_def`

- `workbook_id uuid`
- `pivot_id uuid`
- `sheet_id uuid`
- `name text`
- `anchor_row int`
- `anchor_col int`
- `source_json jsonb`
- `layout_json jsonb`
- `source_revision bigint`

PK `(workbook_id, pivot_id)`

### `spill_owner`

- `workbook_id uuid`
- `sheet_id uuid`
- `owner_row int`
- `owner_col int`
- `rows int`
- `cols int`
- `calc_revision bigint`

PK `(workbook_id, sheet_id, owner_row, owner_col)`

### `data_validation`

- `id uuid pk`
- `workbook_id uuid`
- `sheet_id uuid`
- `range_json jsonb`
- `rule_json jsonb`
- `message text null`
- `strict boolean`
- `source_revision bigint`

### `protected_range`

- `id uuid pk`
- `workbook_id uuid`
- `sheet_id uuid`
- `range_json jsonb`
- `warning_only boolean`
- `editors_json jsonb`
- `source_revision bigint`

### `presence_coarse`

- `workbook_id uuid`
- `user_id uuid`
- `session_id uuid`
- `sheet_id uuid`
- `active_row int`
- `active_col int`
- `selection_json jsonb`
- `color text`
- `updated_at timestamptz`

PK `(workbook_id, session_id)`

### `workbook_change`
User-facing change log for Show Changes / history UI.

- `workbook_id uuid`
- `change_id uuid`
- `revision bigint`
- `actor_user_id uuid`
- `kind text`
- `sheet_id uuid null`
- `range_json jsonb null`
- `summary_json jsonb`
- `before_json jsonb null`
- `after_json jsonb null`
- `created_at timestamptz`

PK `(workbook_id, change_id)`

## 4.2 Internal-only tables

These should not be published in Zero hot sync.

### `workbook_event`
Append-only durable command log.

- `workbook_id uuid`
- `revision bigint`
- `actor_user_id uuid`
- `op_kind text`
- `payload_json jsonb`
- `idempotency_key text null`
- `created_at timestamptz`

PK `(workbook_id, revision)`

### `workbook_snapshot`

- `workbook_id uuid`
- `revision bigint`
- `snapshot_format text`
- `snapshot_bytes bytea`
- `created_at timestamptz`

PK `(workbook_id, revision)`

### `recalc_job`

- `id uuid pk`
- `workbook_id uuid`
- `from_revision bigint`
- `to_revision bigint`
- `priority text`
- `dirty_regions_json jsonb`
- `status text`
- `attempts int`
- `lease_owner text null`
- `lease_until timestamptz null`
- `last_error text null`
- `created_at timestamptz`

### `agent_job`

- `id uuid pk`
- `workbook_id uuid`
- `thread_id text`
- `status text`
- `goal text`
- `plan_json jsonb null`
- `result_json jsonb null`
- `created_at timestamptz`
- `updated_at timestamptz`

---

## 5. Zero design that actually performs

## 5.1 Publication rules

Create a new Zero publication and a new app id for the target schema.

Do **not** try to keep the current and target clients alive against the same publication. The production move is a hard cut:

- backfill the target tables from current authoritative workbook state
- start the new monolith and the new Zero publication
- move the browser client to the new query surface
- validate against real workbooks
- delete the old publication and old app id

The current publication is retired, not extended.

### Publish:

- `workbook`
- `sheet`
- `sheet_row`
- `sheet_col`
- `cell_style`
- `number_format`
- `cell_input` (limited queries only)
- `cell_render`
- `defined_name`
- `table_def`
- `pivot_def`
- `spill_owner`
- `data_validation`
- `protected_range`
- `presence_coarse`
- `workbook_change`
- `sheet_view`

### Do not publish:

- `workbook_snapshot`
- `workbook_event`
- `recalc_job`
- engine/private staging tables

## 5.2 Query rules

### Hard rules

- no whole-workbook query
- no snapshot query
- no giant relation tree query
- no viewport query with continuously unique coordinates per scroll tick
- no sync for unused columns via fat rows

### Spreadsheet query set

#### Workbook chrome

- `workbook.get(workbookId)`
- `sheet.list(workbookId)`
- `sheetView.list(workbookId, sheetId, userId)`
- `presence.byWorkbook(workbookId)`

#### Tiled viewport

Use stable tiles.

Recommended starting tile size:

- `128 x 32` for desktop
- overscan ring of 1 tile in each direction

Queries:

- `cellRender.tile(workbookId, sheetId, tileRow, tileCol)`
- `sheetRow.tile(workbookId, sheetId, tileRow)`
- `sheetCol.tile(workbookId, sheetId, tileCol)`
- `style.tileRefs(...)` only if a style id enters view

#### Selected-cell / editing queries

- `cellInput.get(workbookId, sheetId, rowNum, colNum)`
- `definedName.byWorkbook(workbookId)`
- `tableDef.bySheet(workbookId, sheetId)`
- `pivotDef.bySheet(workbookId, sheetId)`
- `dataValidation.intersectTile(...)`
- `protectedRange.intersectTile(...)`

#### History / collaboration queries

- `workbookChange.recent(workbookId, cursor)`
- `presence.byWorkbook(workbookId)`

## 5.3 Client consumption rules

### Replace current tile aggregation pattern

Current bridge pattern:

- `materialize()` query
- replace whole tile arrays
- concatenate full arrays across tiles
- reproject entire viewport patch

Target pattern:

- use `materialize()` only where acceptable for tiny result sets
- use **custom `View`** for tile queries to receive `add` / `edit` / `remove`
- maintain per-tile mutable maps keyed by `(row,col)`
- project **only changed cells** into the viewport cache
- push minimal patch deltas to the renderer

### Browser-side rules

- do not hold full sheet render arrays in React state
- renderer owns viewport cache and dirty rectangles
- bridge owns tile lifecycle and map-level diffing
- formula bar pulls selected-cell source on demand
- use `preload()` for neighbor tiles, sheet list, recent workbook changes, and likely next sheet
- keep query identities stable

---

## 6. Monolith workbook runtime

## 6.1 Warm workbook sessions

The single biggest performance improvement is to stop rehydrating the engine on every edit.

### Introduce `WorkbookRuntimeManager`

Inside `bilig-app`:

- LRU cache keyed by `workbook_id`
- each runtime holds:
  - warm `SpreadsheetEngine`
  - loaded snapshot revision
  - event tail position
  - dependency graph state
  - dirty region tracker
  - pending presence state
  - metrics counters

### Session budgets

- keep hottest 50–200 workbooks warm per pod, depending on memory profile
- TTL-based idle eviction
- memory-pressure eviction hook
- snapshot checkpoint every N revisions or M seconds

## 6.2 Command application model

Do not apply database row mutations directly as business semantics.

Instead:

1. UI / agent / import layer produces canonical domain commands.
2. Command planner normalizes them.
3. Monolith mutator writes source rows and `workbook_event`.
4. Warm runtime applies the same canonical commands.
5. Runtime calculates dirty closure and render diffs.
6. Only changed `cell_render` and impacted metadata rows are upserted/deleted.

The canonical command model should live in `packages/workbook-domain`, extracted from the old `packages/crdt` op vocabulary and expanded where needed.

## 6.3 Edit classes

### Class A — hot inline edits

Examples:

- set value in single cell
- set formula in single cell
- short fill/paste
- style patch on small range
- rename sheet
- resize single column/row

Flow:

- acquire workbook advisory lock
- update source rows
- apply command on warm runtime synchronously
- compute dirty closure
- upsert changed render rows synchronously
- commit

Target: collaborator sees authoritative update in `< 150ms p95` inside active region.

### Class B — medium batch edits

Examples:

- paste thousands of cells
- autofill large range
- structural insert/delete across moderate region
- bulk formatting
- sort/filter materialization affecting many rows

Flow:

- transaction writes source rows and outbox/recalc job
- UI uses speculative local preview via worker
- background loop on same monolith pod processes job immediately if runtime warm and queue depth low
- authoritative render deltas stream back in chunks

### Class C — heavy jobs

Examples:

- large CSV/XLSX import
- pivot rebuild on huge source
- workbook-wide name/table rebinding
- agent jobs that touch many sheets and charts

Flow:

- asynchronous staged job
- progress events + partial visible render updates
- explicit change bundle / undo bundle

## 6.4 Dirty-set and calc strategy

Adopt Excel’s calc discipline, not brute-force recompute:

- dependency tree
- calculation chain
- dirty marking of direct + indirect dependents
- structural edits rebuild only affected dependency segments
- volatile functions tracked separately
- names/tables/pivots participate in dependency invalidation
- spill owner tracks spill region occupancy and invalidation

This must be bilig’s server runtime strategy, with WASM in the hot path.

## 6.5 Projection writes

### Replace current whole-workbook diff path

`apps/sync-server/src/zero/store.ts` currently uses `persistWorkbookMutation(...)` to build full `previousProjection` and `nextProjection` objects and then runs `diffProjectionRows(...)` across the workbook.

Delete that algorithm from the hot write path and replace it with incremental persistence.

New persistence contract:

- `upsertCellInputDelta(...)`
- `applyRenderDelta(...)`
- `applyRowDelta(...)`
- `applyColDelta(...)`
- `applyTableDelta(...)`
- `applyPivotDelta(...)`
- `appendWorkbookChanges(...)`
- `checkpointSnapshotIfNeeded(...)`

Never delete/reinsert the workbook projection wholesale.

---

## 7. Exact database/indexing plan

## 7.1 Hot indexes

### `cell_render`

- primary key `(workbook_id, sheet_id, row_num, col_num)`
- btree `(workbook_id, sheet_id, row_num, col_num)`
- optional partial index for `error_code is not null`
- optional partial index for `flags <> 0`

### `cell_input`

- primary key `(workbook_id, sheet_id, row_num, col_num)`

### `sheet_row`

- `(workbook_id, sheet_id, row_num)`

### `sheet_col`

- `(workbook_id, sheet_id, col_num)`

### `sheet`

- unique `(workbook_id, position)`
- unique `(workbook_id, name)`

### `presence_coarse`

- `(workbook_id, updated_at)`

### `workbook_change`

- `(workbook_id, created_at desc)`
- `(workbook_id, revision desc)`

### `recalc_job`

- `(status, lease_until, created_at)`
- `(workbook_id, status, created_at)`

## 7.2 Query plan discipline

Every hot Zero query must have:

- explicit order matching upstream index order
- no `TEMP B-TREE`
- bounded result size
- stable query identity
- tile-aligned shapes only

Before rollout, run Zero slow-query analysis on:

- `cellRender.tile`
- `sheetRow.tile`
- `sheetCol.tile`
- `presence.byWorkbook`
- `workbookChange.recent`

---

## 8. AX

## 8.1 What should be removed

`@bilig/agent-api` exits the merged production design.

It can exist only as temporary branch scaffolding while code is being extracted. It is not part of the final runtime, network surface, or operator story.

## 8.2 Required architecture

Use:

- **Codex app-server protocol** for the rich agent runtime
- **agent-optimized internal APIs** for workbook reads, previews, and mutations
- **versioned bilig skill packs** that bind Codex to those APIs
- **side-by-side AX pane** in the web app

This is the correct separation:

- Codex handles conversation, planning, streaming, approvals, and skill orchestration.
- bilig exposes fast workbook-specific skills backed by the monolith runtime.

## 8.3 Repo leverage from `lab`

Reuse the existing internal wrapper:

- `lab/packages/codex/src/app-server-client.ts`
- `lab/services/jangar/src/server/codex-client.ts`

Create in bilig:

- `packages/agent-runtime-codex`
- or import the codex package directly if repo policy allows

## 8.4 Production rule for Codex app-server

The OpenAI docs currently describe `codex app-server` as experimental / primarily for development and debugging.

Therefore production engineering must be:

- pin exact Codex CLI version
- generate and vendor the protocol bindings for that version
- use only stable API surface by default
- keep `experimentalApi` **off** unless a specific feature is accepted behind a feature flag
- use **stdio transport**, not websocket
- wrap all Codex traffic behind `AgentRuntime` interface inside the monolith
- support process restart and thread recovery

That is not a workaround. That is the correct way to use an evolving protocol safely.

## 8.5 Skill surface

Expose workbook capabilities to Codex through versioned bilig skill packs, not through a generic remote protocol.

### Context and read skills

- `get_context`
- `get_selection`
- `read_range`
- `read_table`
- `read_named_range`
- `get_formula`
- `get_dependents`
- `get_precedents`
- `get_sheet_schema`
- `list_sheets`
- `get_recent_changes`
- `get_visible_viewport`
- `find_relevant_regions`

### Write skills

- `batch_edit_cells`
- `batch_set_formulas`
- `batch_style`
- `insert_rows`
- `insert_cols`
- `delete_rows`
- `delete_cols`
- `sort_range`
- `filter_view_create`
- `filter_view_apply`
- `create_table`
- `create_pivot`
- `create_chart`
- `rename_sheet`
- `add_sheet`
- `delete_sheet`
- `undo_change_bundle`

### Preview and planning skills

- `preview_batch`
- `diff_preview`
- `validate_formula_batch`
- `estimate_cost`

## 8.6 Agent execution model

### AX flow

- agent sees active workbook, active sheet, current selection, visible viewport, recent changes, sheet list, and named objects
- agent produces **plan → preview → apply**
- user can accept whole plan or per-step chunks
- for small deterministic edits, optionally allow “auto-apply” mode inside selected range or workbook scope
- agent appears as a live participant while it is running, with visible working range, current step label, and authored changes flowing through normal presence/history UI

### Background workflow

For large workflows:

1. user prompts in the AX pane
2. agent gathers context through bilig skills
3. agent returns streamed plan + preview highlights immediately
4. user accepts or auto-apply policy accepts
5. bilig applies as one or more command bundles
6. visible screen updates incrementally via Zero almost immediately
7. change bundle appears in history and can be undone as one action

### Tool design rule

The agent should never automate the DOM to use the spreadsheet.

It should call semantic workbook skills backed by optimized internal APIs.

That is the only way to get reliable, ultrafast, auditable, batch-safe spreadsheet automation.

## 8.7 Agent performance path

Agent write tools must call high-performance internal APIs, not generic public REST.

Implement local server-side interfaces such as:

- `applyCommandBundle(bundle)`
- `previewCommandBundle(bundle)`
- `planRangeTransforms(spec)`
- `executeBatchMutation(batch)`
- `streamRenderDiffs(revision)`

These operate inside the monolith process and can directly touch the warm runtime and Postgres transaction path.

---

## 9. Exact UX blueprint

The product must not be “spreadsheet with chat bolted on.”

The baseline experience should feel comparable to Excel Online / Google Sheets before AX is considered.

The AX bar is not parity with current spreadsheet copilots. It must beat them on latency, preview clarity, reversibility, collaboration visibility, and trust.

## 9.1 Main shell layout

### Top bar

- workbook title with inline rename
- save/sync state chip
- share button
- collaborator avatars
- undo / redo
- command palette
- search / find
- comments / notifications entry
- AX entrypoint

### Formula row

- name box
- function `fx` affordance
- formula bar with rich text tokens for ranges/functions/names
- autocomplete dropdown
- function signature/help strip
- precedents/dependents quick affordance

### Grid area

- frozen panes
- row/column headers
- smooth 60fps scroll
- exact hit-testing
- crisp selection visuals
- multi-range selection
- autofill handle
- drag-reorder tabs and sheet objects
- inline error and spill visualization
- polished context menus

### Bottom bar

- sheet tabs
- tab scroll
- add sheet
- sheet menu
- status bar with:
  - sum/count/avg
  - mode/status
  - calc indicator
  - AX activity indicator when relevant

### Right rail

Tabbed:

- AX
- Comments
- Changes
- Named objects / inspector

The AX pane should be collapsible and resizable.

## 9.2 Multiplayer UX

### Presence

- colored collaborator cursor/selection outline
- collaborator initials/avatar near active range
- hover to reveal identity
- click avatar to jump to collaborator location
- subtle pulse when remote edit lands in your current viewport
- AX appears as a first-class collaborator with a distinct avatar, color, and live working range while it is applying changes

### Conflict UX

- authoritative value wins
- if local preview diverges before server response, show a one-frame reconcile animation, not a jarring snap
- selection never teleports unexpectedly
- formula bar never loses in-progress typing due to remote edits elsewhere

### Show Changes

Implement a first-class changes pane:

- who changed what, where, when
- previous value and new value
- grouped batch edits
- range filter / sheet filter
- click entry to jump to range
- revert bundle where legal
- agent-authored changes carry step summaries so the user can see which skill or action changed which ranges

### Sheet Views / Filter Views

Provide both:

- per-user private view state
- named saved views

Behavior:

- user filters/sorts without disrupting others
- can save a view, rename it, share linkable view state
- temporary views auto-promote to saved when named

## 9.3 Editing UX baseline

These are mandatory:

- single-click select
- double-click edit
- type-to-replace
- F2 edit
- Enter / Shift+Enter / Tab / Shift+Tab semantics
- Home / End / Ctrl+Arrow / PageUp / PageDown parity
- drag selection and autofill
- clipboard copy/cut/paste, including multi-cell paste
- formula autocomplete with function arg hints
- merged cells behavior correctness
- frozen pane correctness
- row resize
- column resize and autofit
- hide/unhide rows/cols
- right-click structural actions
- keyboard-accessible menus and ribbon actions

## 9.4 AX

### AX pane behavior

The AX pane should operate in two modes:

#### Chat mode

- explanations
- formula help
- workbook Q&A
- non-mutating analysis

#### Edit mode

- natural language edit request
- streamed plan
- live preview highlight in sheet
- per-step or full apply
- one-click undo bundle

### Required AX elements

- active workbook context summary at top of pane
- visible selection chip
- plan list with statuses (`planned`, `running`, `applied`, `failed`)
- preview overlay in grid before apply
- side-by-side rendered result explanation
- citation of sheets/ranges it used
- “apply to current selection / current sheet / whole workbook” scope control
- approval prompt for destructive or wide-scope edits
- active AX step pinned to the affected range in-grid while it runs
- agent presence visible in the collaborator strip and changes pane, not hidden inside the side panel
- step-level undo as well as bundle-level undo for long multi-step plans
- interrupt, pause, and resume controls without losing workbook context

### AX quality bar

- first useful preview appears fast enough that the user does not wonder whether the system understood the request
- every action is spatially anchored in the grid, with visible range ownership before and during apply
- the user can understand scope, risk, and rollback path at a glance without reading a wall of text
- long workflows stay interruptible, reversible, and legible instead of becoming a hidden background job
- agent work feels like collaborating with a very fast operator inside the sheet, not sending a request into a black box

### Gold standard interaction examples

- “clean this imported CSV and turn it into a monthly report”
- “merge these three sheets into a single summary table and make a pivot with slicers”
- “find broken formulas and fix them”
- “recreate this analysis in a more readable dashboard layout”
- “convert this workbook to match our finance template”

The user should see a plan and visible progress almost immediately, and sheet state should change incrementally as the agent applies semantic bundles.

---

## 10. Excel / Google Sheets feature parity map

## 10.1 Must-have collaboration features

Modeled after Excel / Sheets:

- live coauthoring with collaborator jump-to-location
- Show Changes pane with previous value
- version history with named versions
- sheet views / filter views
- protected sheets and protected ranges
- named ranges and named functions
- tables and pivots
- slicers / saved filter state

## 10.2 Must-have AX editing features

Modeled after Excel Copilot editing:

- add/rename/delete sheets
- insert/modify cell values and ranges
- apply conditional formatting / styles / borders / validation
- create tables
- create pivots
- perform cross-sheet formulas and transformations
- update workbook layout

## 10.3 Compatibility requirements

### Open XLSX

Supported:

- values
- formulas
- styles
- number formats
- merges
- row/column sizes
- hidden rows/cols
- freeze panes
- sheet names/order
- tables
- pivots where feasible
- validations
- named ranges

### Save XLSX

Supported:

- round-trip as much as possible without semantic loss
- preserve unsupported features where possible via opaque part pass-through
- emit compatibility warnings when semantic downgrade occurs

### Compatibility subsystem

Add:

- `CompatibilityReport`
- `ImportWarnings`
- `ExportWarnings`
- workbook-level “Compatibility” drawer with actionable issues

---

## 11. Performance targets

The repo already has budgets. Keep them and add server/collaboration/agent budgets.

## 11.1 Browser targets

- visible local input response p95 `< 16ms`
- viewport scroll on active workbook: maintain 60fps on modern laptop
- selection paint p95 `< 8ms`
- remote visible patch apply p95 `< 16ms`

## 11.2 Server targets

### Hot inline edit

- single-cell authoritative commit p95 `< 60ms`
- collaborator visible update in same region p95 `< 150ms`

### Small-range edit

- 100-cell paste authoritative commit p95 `< 100ms`

### Medium-range edit

- 5k-cell paste first visible authoritative diff `< 250ms`
- full completion depending on formula graph, streamed progressively

### Workbook restore

Keep repo budgets:

- `100k` restore p95 `< 500ms`
- `250k` restore p95 `< 1500ms`

## 11.3 Agent targets

- first streamed plan token `< 700ms p95`
- first preview highlight `< 1000ms p95`
- first visible sheet mutation for accepted plan `< 1500ms p95`
- batch apply throughput should exceed manual UX by at least an order of magnitude on multi-step workflows

## 11.4 Zero targets

Single-node first:

- tile query warm hit `< 10ms` inside `zero-cache`
- reconnects should reuse warm query state where possible
- no hot query with `TEMP B-TREE`

---

## 12. Exact implementation plan

This lands as one migration branch and one production cut. Use as many commits as needed, but do not merge or deploy an intermediate hybrid state.

## 12.1 Hard-cut rules

- `apps/sync-server` is the monolith root for this migration. Do not create `apps/app`.
- `packages/zero-sync` is rewritten in place. Do not create `packages/zero-sync-v2`.
- `apps/web` remains the browser build target. The monolith serves its built assets.
- `apps/local-server` is an extraction source only. Any behavior worth keeping moves into shared monolith modules; do not preserve a second authority implementation.
- old snapshot-backed hot writes, whole-workbook query surfaces, and workbook-wide style/format subscriptions are removed in the migration branch before merge
- the production cut creates a new Zero app id/publication and retires the old one immediately after validation

## 12.2 Reshape the monolith in the current repo

Expand `apps/sync-server` from “sync server” into `bilig-app` in place:

- web asset serving
- Zero query/mutate handlers
- workbook runtime manager
- recalc/outbox loops
- AX runtime
- skill registry and execution adapters
- operational and health endpoints

Extract reusable engine/session behavior from `apps/local-server/src/local-workbook-session-manager.ts` into modules under `apps/sync-server/src/`:

- warm engine cache
- per-workbook lifecycle management
- idle eviction
- checkpoint scheduling
- batch/event fanout only where the monolith still needs it

Delete `apps/local-server` from the product runtime after extraction. It must not ship, start, or own workbook authority in the merged system.

Also remove Redis assumptions from compose, manifests, and runtime docs.

## 12.3 Do the one-time data migration from current state

Current authoritative state is still anchored to `workbooks.snapshot`, `workbooks.replicaSnapshot`, and the projection tables built around them.

Build one migration script that:

1. reads each workbook snapshot from the current schema
2. restores the engine once
3. emits the new source/render rows from that authority state
4. writes private checkpoint rows for recovery
5. sets `head_revision` and `calculated_revision` in the target schema

Use the current projection code only as migration scaffolding where it helps, especially `apps/sync-server/src/zero/projection.ts`. Do not keep those whole-workbook projection builders in the hot path after the migration lands.

Production sequence:

1. apply the schema migration
2. run the one-time backfill
3. deploy the monolith and the new Zero publication/app id
4. move traffic to the new client/runtime path
5. validate on representative workbooks
6. delete the old publication and old production deployment shape

## 12.4 Rewrite `packages/zero-sync` in place

Update `packages/zero-sync/src/schema.ts` from the current snapshot-centric model (`workbooks`, `cells`, `cell_eval`, `row_metadata`, `column_metadata`, `sheet_style_ranges`, `sheet_format_ranges`, and related tables) to the target model in section 4.

Rules for that rewrite:

- keep naming conservative where it reduces churn; do not rename tables or fields just to make the plan prettier
- remove synced snapshot/blob fields and any Zero-facing table that only mirrors private runtime state
- keep source rows, render rows, row/column metadata, presence, and recent change history as the public sync surface

Update `packages/zero-sync/src/queries.ts` so the public query surface is only:

- workbook chrome queries
- tile-bounded render/source queries
- selected-cell source queries
- bounded collaboration/history queries

Delete the `computedCells` alias and remove workbook-wide `styles.byWorkbook` / `numberFormats.byWorkbook` reads from the hot client path.

Update mutate arg schemas so browser and agent writes both flow through canonical command bundles rather than ad-hoc table-shaped mutations.

## 12.5 Replace snapshot-backed hot writes in `apps/sync-server`

`apps/sync-server/src/zero/server-mutators.ts` must stop calling `createWorkbookEngine(...)` per request.

Introduce `WorkbookRuntimeManager` and per-workbook runtime handles that:

- acquire the advisory lock
- load or create the warm runtime
- apply the canonical command bundle
- compute dirty closure and render deltas
- persist only affected source/render/metadata rows
- append change/event records
- checkpoint only when thresholds are crossed

Split `apps/sync-server/src/zero/store.ts` into clear responsibilities:

- checkpoint load/save
- source row persistence
- render delta persistence
- revision/change append
- recalc job leasing

Delete whole-workbook `previousProjection` vs `nextProjection` diffing from hot mutations.

Keep full snapshot export only as recovery/checkpoint material, not as the write record replicated to clients.

Fold `apps/sync-server/src/zero/recalc-worker.ts` into the same monolith job loops first. Split it later only if measured CPU isolation is needed.

## 12.6 Replace browser Zero consumption in `apps/web`

`apps/web/src/zero/ZeroWorkbookBridge.ts` must stop materializing workbook-wide style and number-format registries on startup.

`apps/web/src/zero/tile-subscriptions.ts` must stop building aggregate arrays by concatenating per-tile query results on every listener update.

The new client bridge should:

- keep per-tile maps keyed by stable cell/row/col identities
- apply `add` / `edit` / `remove` deltas into those maps
- project only changed cells into `viewport-projector`
- fetch selected-cell source separately for the formula bar/editor
- preload neighboring tiles and adjacent sheet metadata with stable query identities

The renderer or worker owns the viewport cache and dirty rectangles. React state does not own full render arrays.

Remove any remaining production use of whole-workbook render/state subscriptions.

## 12.7 Collaboration, history, and AX path

Finish the migration on the same runtime path:

- presence uses coarse synced rows plus monolith-owned ephemeral fanout where needed
- `workbook_change` powers Show Changes and undoable change bundles
- `sheet_view` stays private per-user state
- the AX pane uses the same `plan -> preview -> apply` bundle path as normal edits
- the agent is visible as a live collaborator while it is mutating the workbook

Remove production dependence on `@bilig/agent-api`. Add an internal `AgentRuntime` inside the monolith, backed by pinned `codex app-server` stdio sessions and bilig skill packs that call the same command planner/runtime manager as user edits.

## 12.8 Delete list for the migration branch

Delete or retire these before merge:

- old synced snapshot columns and Zero queries that expose them
- `computedCells` duplicate query surface
- workbook-wide style/format hot subscriptions
- cold-boot-per-mutation engine creation path
- whole-workbook projection diffing in hot writes
- Redis-based presence/coordination assumptions
- separate production deployment model for `local-server`
- standalone production use of `@bilig/agent-api`
- any generic tool/protocol layer that bypasses the bilig skill packs

## 12.9 Merge gate and production validation

Merge only when all of these are true:

- representative existing workbooks were backfilled from current snapshots into the new schema
- open workbook, edit cell, paste range, rename sheet, recalc, refresh, and collaborator update all work against the new monolith
- selected-cell formula/source reads are correct and stay in sync with rendered cell values
- Zero slow-query checks show no hot query using temp b-tree or unbounded result shapes
- monolith restart recovers from checkpoints without data loss
- AX preview/apply/undo works through skill-backed command bundles, and the agent is visible in-grid while it runs
- import/export still round-trips representative fixtures
- the old publication/app id and old runtime surface can be removed without breaking traffic

---

## 13. Operational design

## 13.1 Keep ops simple

### Initial production shape

- 2–4 replicas of `bilig-app`
- 1 replica of `bilig-zero` initially
- 1 Postgres primary

No Redis.
No NATS.
No separate agent service.
No separate recalc service.
No sidecar complexity.

### Scale policy

Scale in this order:

1. optimize queries/indexes
2. warm runtime tuning
3. app replica count
4. Zero single-node resource tuning
5. move Zero to replication-manager + view-syncers only when single-node is saturated
6. split heavy recalc into dedicated deployment only when monolith CPU isolation becomes necessary

## 13.2 Monolith internals for concurrency

Use:

- advisory locks per workbook for write serialization
- PG outbox / `SKIP LOCKED` for background jobs
- worker threads or child pools for CPU-heavy XLSX transforms if necessary
- bounded in-process queue per workbook to coalesce burst edits

This preserves operational simplicity while still achieving high throughput.

---

## 14. Testing plan

## 14.1 Correctness suites

- formula parity corpus against bilig engine
- JS vs WASM differential tests
- XLSX import/export round-trip fixtures
- multiplayer order/race tests
- selected-cell/edit-mode remote update tests
- sheet view isolation tests
- change history correctness tests

## 14.2 Performance suites

- hot single-cell edit
- 100-cell paste
- 5k-cell paste
- sort 100k rows
- open 100k workbook
- open 250k workbook
- remote collaborator edit burst
- agent batch edit workflow

## 14.3 Failure-mode tests

- `zero-cache` restart during editing
- auth expiry -> `needs-auth`
- long `connecting` grace period
- monolith pod recycle with snapshot restore
- Codex child process crash and recovery
- partial apply rollback on failed bundle

---

## 15. Final recommendation

The implementation-ready recommendation is:

- **one product monolith** for everything bilig-specific
- **one Zero deployment** because Zero itself is a required adjacent runtime
- **one Postgres cluster**
- **no CRDT sync server**
- **no workbook snapshot syncing**
- **no whole-workbook Zero query**
- **no full projection rebuilds**
- **warm WASM workbook sessions** in the monolith
- **incremental render diffs** only
- **Codex app-server + bilig skills** for AX
- **UX and AX quality that is materially better than current spreadsheet copilots on speed, clarity, and trust**

Execute that as one hard cut using the current repo surfaces in place:

- `apps/sync-server` becomes the monolith root
- `packages/zero-sync` is rewritten in place
- `apps/web` keeps the browser build while the monolith takes over serving it
- current snapshots are backfilled once into the new source/render model and then removed from the hot path

If only three work items matter before anything else, they are:

1. replace `apps/sync-server/src/zero/server-mutators.ts` and `persistWorkbookMutation(...)` with a warm runtime manager plus targeted source/render writes
2. rewrite `packages/zero-sync` around tile-bounded source/render queries and selected-cell source reads
3. rebuild `apps/web/src/zero/ZeroWorkbookBridge.ts` and `apps/web/src/zero/tile-subscriptions.ts` around per-tile delta application instead of workbook-wide registries and array aggregation

Those are not follow-up optimizations. They are the migration.
