
# bilig → Zero 1.0 production implementation plan
## Repo-grounded, CRDT/sync-server deprecation plan
## Status: implementation-ready design
## Last revalidated against local repos: 2026-03-28

## 0. Scope and baseline used for this plan

This plan is grounded in the current local worktrees, not in a generic spreadsheet architecture:

- **Primary product repo used**: `~/github.com/bilig`
- **Infrastructure repo used**: `~/github.com/lab`

This revision no longer depends on uploaded zip snapshots or a second bilig extract. It is grounded in the active checked-out repos listed above.

I used the **current bilig repo** as the primary baseline because it already contains the current Zero integration surfaces:
- `packages/zero-sync`
- `apps/sync-server/src/zero/*`
- `apps/web` wired to `@rocicorp/zero`
- `lab` manifests in `~/github.com/lab/argocd/applications/bilig` already deploying `rocicorp/zero:1.0.0`

This update revalidated the plan against the current code in:
- `packages/zero-sync/src/queries.ts`
- `packages/zero-sync/src/schema.ts`
- `packages/zero-sync/src/mutators.ts`
- `apps/sync-server/src/server.ts`
- `apps/sync-server/src/zero/service.ts`
- `apps/sync-server/src/zero/server-mutators.ts`
- `apps/sync-server/src/zero/store.ts`
- `apps/web/src/WorkerWorkbookApp.tsx`
- `apps/web/src/worker-runtime.ts`
- `~/github.com/lab/argocd/applications/bilig/README.md`
- `~/github.com/lab/argocd/applications/bilig/zero-deployment.yaml`
- `~/github.com/lab/argocd/applications/bilig/postgres-cluster.yaml`
- `~/github.com/lab/argocd/applications/bilig/frontend-ingressroute.yaml`

This plan also assumes **current Zero 1.0 semantics in the checked-out repos**, not earlier preview-era guidance.

---

## 1. Executive conclusion

The right architecture is:

- **Deprecate `@bilig/crdt` as the sync/replication system**
- **Deprecate `apps/sync-server` as the CRDT binary sync gateway**
- **Keep bilig’s formula/runtime engine as the only calc oracle**
- **Use Zero 1.0 as the read-sync + write-ingress + collaboration distribution layer**
- **Use Postgres as semantic source of truth**
- **Use a headless bilig workbook service to compute authoritative formula outputs and materialize them back into Postgres**
- **Keep the worker-first browser shell and viewport-patch UI contract**
- **Rename the product internally from “true local-first spreadsheet” to “server-authoritative multiplayer spreadsheet with local-first UX”**

That last point is not optional. With Zero 1.0, the target is **instant local reads plus speculative local preview**, but **not offline-write local-first**. The product should still feel instant and multiplayer, but it is a server-authoritative system.

The highest-leverage product decision is this:

> **Zero should sync a relational workbook model and materialized evaluation state, not whole workbook snapshots.**

The current repo does the opposite. That is why the current Zero path is a prototype, not a production design.

---

## 2. Current repo reality

### 2.1 Current Zero integration is prototype-grade

The current `bilig` repo already has a Zero path, but it is not the architecture you want to scale:

#### Current read path
- `packages/zero-sync/src/queries.ts`
  - `workbooks.byId({documentId})` loads the entire workbook graph:
    - workbook row
    - sheets
    - all cells
    - all computed cells
    - row metadata
    - column metadata
    - defined names
    - workbook metadata
    - calculation settings

This is whole-workbook sync, not query-shaped viewport sync.

#### Current write path
- `apps/sync-server/src/zero/server-mutators.ts`
  - loads the full snapshot
  - instantiates a `SpreadsheetEngine`
  - imports the full snapshot
  - applies a mutation
  - exports a full snapshot
  - materializes computed cells by iterating all populated cells
  - persists a complete projection

- `apps/sync-server/src/zero/store.ts`
  - `persistWorkbookProjection()` clears and reinserts entire workbook projection tables

This is effectively “store workbook snapshot + regenerate projection on every mutation”.

#### Current browser path
- `apps/web/src/WorkerWorkbookApp.tsx`
  - `useQuery(queries.workbooks.byId({documentId}))`
  - `projectWorkbookToSnapshot(remoteWorkbook, documentId)`
  - `replaceSnapshot(remoteSnapshot)` into the worker
  - after many worker mutations, the browser writes `replaceSnapshot` back to Zero again

This means the browser is treating Zero as snapshot storage plus whole-workbook replication.

### 2.2 The good news: the repo already has the right seams

The codebase already contains the right long-lived architecture seams:

- `apps/web/src/worker-runtime.ts`
  - worker-first runtime
  - viewport patch derivation
  - persistence restore
  - WASM lifecycle
  - selection/edit/runtime services

- `packages/worker-transport/src/viewport-patch.ts`
  - established UI contract for viewport patches:
    - patched cells
    - display/copy/editor text
    - row/column patches
    - styles
    - metrics

- `docs/browser-runtime.md`
  - already says the grid should consume derived viewport patches, not raw engine state

- `docs/workbook-metadata-runtime-rfc.md`
  - already treats names, tables, pivots, spills, filters, sorts, and row/column metadata as first-class workbook state

- `packages/protocol/src/types.ts`
  - already has durable types for:
    - workbook metadata
    - styles
    - number formats
    - names
    - spills
    - pivots
    - tables
    - axis entries
    - calc settings

- `packages/core/src/engine.ts`
  - already owns:
    - workbook state
    - metadata
    - formula runtime
    - snapshot import/export
    - structural row/column operations
    - tables/pivots/names
    - spill behavior

That means this migration should be a **sync/distribution architecture rewrite**, not a spreadsheet-engine rewrite.

### 2.3 `@bilig/crdt` is mixed: transport logic and domain model are fused

`packages/crdt/src/index.ts` currently holds two different things:

1. **Replication-specific concerns**
   - replica clocks
   - applied batch IDs
   - compareBatches
   - shouldApplyBatch
   - replica snapshots

2. **Useful domain concepts**
   - `WorkbookOp`
   - sheet/cell/style/metadata/structural ops
   - op families that are still valuable even without CRDT

That means you should **not delete the semantic op model when deprecating CRDT**. You should extract it into a transport-neutral package and delete only the replica/merge layer.

### 2.4 The lab repo is already close to the desired deployment shape

The current `lab` repo already has the right production shell for Zero:

- `argocd/applications/bilig/zero-deployment.yaml`
  - `rocicorp/zero:1.0.0`
  - persistent replica file
  - `/api/zero/query` and `/api/zero/mutate` wired to `bilig-sync`

- `argocd/applications/bilig/postgres-cluster.yaml`
  - Postgres 17
  - logical replication enabled
  - publication `zero_data`

- `argocd/applications/bilig/frontend-ingressroute.yaml`
  - `/api/zero` routed to `bilig-sync`
  - `/zero` routed to `bilig-zero`

This is good. The infrastructure direction is already mostly correct. The application data model is what must change.

---

## 3. Non-negotiable truths for the target architecture

### 3.1 What the target product is now

After this migration, bilig should be described as:

- **server-authoritative**
- **multiplayer**
- **query-driven partial-sync**
- **local-first-feeling**
- **speculative-preview capable**
- **read-only offline tolerant**
- **not offline-write local-first**

That wording matters because it affects product promises, UX, testing, and failure handling.

### 3.2 What Zero should own

Zero should own:

- online sync of workbook state from Postgres to clients
- local query cache on client
- realtime fanout of changed rows
- optimistic mutation ingress
- auth-scoped query/mutate APIs
- multiplayer visibility for durable collaboration state
- consistent partial replicas of active workbook data

### 3.3 What bilig should own

bilig should continue to own:

- formula parse/bind/evaluate semantics
- JS/WASM execution
- snapshot import/export
- structural rewrite rules
- copy/fill/paste translation
- tables / names / pivots / spills semantics
- format and display derivation
- viewport-patch shape and rendering contracts
- Excel parity

### 3.4 What Postgres should own

Postgres should own:

- workbook semantic source of truth
- workbook ordering/revisions
- durable audit/event stream
- recalc queue / transactional outbox
- authoritative rendered value materialization
- sharing and permissions

---

## 4. The target architecture

### 4.1 High-level shape

```text
browser UI
  └─ worker runtime
       ├─ local parse / preview / formatting / patch projection
       └─ Zero workbook bridge
            ├─ subscribes to tiled Zero queries
            ├─ maintains authoritative viewport projection
            └─ emits existing viewport patches to UI

Zero client
  └─ zero-cache
       ├─ local partial replica
       ├─ query sync
       └─ mutation transport

bilig-sync (repurposed)
  ├─ /api/zero/query
  ├─ /api/zero/mutate
  ├─ auth/context
  ├─ transactional write executor
  └─ embedded recalc worker or separate recalc worker

Postgres
  ├─ workbook semantic tables
  ├─ workbook materialized evaluation tables
  ├─ workbook revisions/events
  ├─ recalc jobs
  └─ snapshots for warm start only
```

### 4.2 Final component responsibilities

#### `apps/web`
Keeps:
- UI shell
- grid
- formula bar
- selection UX
- worker boot
- patch consumption

Changes:
- stop loading whole workbook snapshots from Zero
- stop pushing `replaceSnapshot` after edits
- stop depending on CRDT replica semantics for Zero mode
- start consuming tiled authoritative workbook queries
- start using real auth-derived `userID`

#### `apps/sync-server`
Do **not** delete this deployment name immediately because lab manifests already use it.

Instead:
- keep Kubernetes service/deployment name `bilig-sync`
- repurpose app role from “binary CRDT sync server” into:
  - Zero query endpoint
  - Zero mutate endpoint
  - workbook write coordinator
  - recalc worker host (phase 1)
  - agent API host if still needed

Later, when stable, you can rename the app/package if desired.

#### `apps/local-server`
Deprecate as a product authority. Reuse useful pieces:
- warm `SpreadsheetEngine` session manager pattern
- import/export helpers
- worksheet execution helpers

Practical recommendation:
- keep `apps/local-server` for local development and test tooling
- extract shared runtime pieces into a new internal package used by the production workbook service

#### `packages/zero-sync`
Keep the package name to minimize import churn.

Replace its responsibilities:
- from prototype snapshot projection package
- to production Zero model package:
  - schema
  - queries
  - mutators
  - client helpers
  - context types
  - tile helpers
  - auth-aware query registry

#### new package: `packages/workbook-domain`
Create this package by extracting the transport-neutral workbook mutation model out of `packages/crdt`.

It should contain:
- `WorkbookOp`
- `WorkbookTxn`
- shared mutation payload types
- planner input/output types
- no replica clocks
- no applied batch IDs
- no CRDT convergence helpers

#### `packages/crdt`
Phase out. Short-term:
- leave compatibility re-exports to avoid a giant compile break
- mark deprecated
- remove once all imports are moved

#### new package: `packages/workbook-service-runtime` (recommended)
Shared production runtime helpers:
- workbook loading from Postgres
- snapshot warm-start helpers
- event replay helpers
- recalc diff materialization helpers
- planner helpers for complex mutations
- idempotency helpers

---

## 5. The target data model

## 5.1 Design rules

1. **Do not use whole-workbook snapshot rows as the synced client model**
2. **Separate source state from evaluated/rendered state**
3. **Keep hot synced rows narrow**
4. **Use sparse tables**
5. **Store numeric coordinates, not only A1 strings**
6. **Keep style and number-format registries normalized**
7. **Store range semantics as ranges when per-cell expansion would explode**
8. **Use revision numbers, not replica clocks**
9. **Use snapshots only as warm-start accelerators**
10. **Keep internal operational tables out of the client publication when possible**

## 5.2 Table inventory

### Published to Zero clients

#### `workbook`
One row per workbook.

Columns:
- `id text primary key`
- `name text not null`
- `owner_user_id text not null`
- `head_revision bigint not null default 0`
- `calculated_revision bigint not null default 0`
- `calc_mode text not null default 'automatic'`
- `compatibility_mode text not null default 'excel-modern'`
- `recalc_epoch bigint not null default 0`
- `created_at timestamptz not null`
- `updated_at timestamptz not null`

Purpose:
- workbook identity
- revision tracking
- calc settings
- staleness detection (`head_revision` vs `calculated_revision`)

#### `workbook_member`
Sharing and authorization.

Columns:
- `workbook_id text not null`
- `user_id text not null`
- `role text not null` (`owner`, `editor`, `commenter`, `viewer`)
- `granted_by text`
- `created_at timestamptz not null`

Unique:
- `(workbook_id, user_id)`

Purpose:
- query/mutate authorization
- sharing model

#### `sheet`
Stable sheet identity and durable sheet-level semantics.

Columns:
- `id text primary key`
- `workbook_id text not null`
- `name text not null`
- `position int not null`
- `freeze_rows int not null default 0`
- `freeze_cols int not null default 0`
- `created_at timestamptz not null`
- `updated_at timestamptz not null`

Unique:
- `(workbook_id, name)`
- `(workbook_id, position)`

Purpose:
- stable sheet IDs
- visible sheet order
- rename/reorder support
- freeze panes

#### `sheet_filter`
Small sheet-scoped durable filters.

Columns:
- `id text primary key`
- `workbook_id text not null`
- `sheet_id text not null`
- `start_row int not null`
- `end_row int not null`
- `start_col int not null`
- `end_col int not null`
- `source_revision bigint not null`

Purpose:
- durable filter regions

#### `sheet_sort`
Small sheet-scoped durable sorts.

Columns:
- `id text primary key`
- `workbook_id text not null`
- `sheet_id text not null`
- `start_row int not null`
- `end_row int not null`
- `start_col int not null`
- `end_col int not null`
- `keys_json jsonb not null`
- `source_revision bigint not null`

Purpose:
- durable sort regions

#### `sheet_row`
Sparse row overrides.

Columns:
- `workbook_id text not null`
- `sheet_id text not null`
- `row_num int not null`
- `row_id text null`
- `size int null`
- `hidden boolean null`
- `source_revision bigint not null`
- `updated_at timestamptz not null`

Primary key:
- `(workbook_id, sheet_id, row_num)`

Purpose:
- row height / hide state
- optional stable row identity alignment with protocol

#### `sheet_col`
Sparse column overrides.

Columns:
- `workbook_id text not null`
- `sheet_id text not null`
- `col_num int not null`
- `col_id text null`
- `size int null`
- `hidden boolean null`
- `source_revision bigint not null`
- `updated_at timestamptz not null`

Primary key:
- `(workbook_id, sheet_id, col_num)`

Purpose:
- column width / hide state
- optional stable column identity alignment with protocol

#### `cell_source`
User-authored source cells.

Columns:
- `workbook_id text not null`
- `sheet_id text not null`
- `row_num int not null`
- `col_num int not null`
- `literal_input_json jsonb null`
- `formula_source text null`
- `explicit_format_id text null`
- `source_revision bigint not null`
- `updated_by text not null`
- `updated_at timestamptz not null`

Primary key:
- `(workbook_id, sheet_id, row_num, col_num)`

Rules:
- row exists only if the cell has authored source state
- blank/default cell = no row
- exactly one of `literal_input_json` or `formula_source` should be set, unless the cell is a format-only cell

Purpose:
- semantic source of cell content
- selected cell editor source
- formula source of truth

#### `cell_eval`
Authoritative evaluated cell outputs.

Columns:
- `workbook_id text not null`
- `sheet_id text not null`
- `row_num int not null`
- `col_num int not null`
- `value_tag smallint not null`
- `number_value double precision null`
- `boolean_value boolean null`
- `string_value text null`
- `error_code text null`
- `flags int not null default 0`
- `version bigint not null`
- `calc_revision bigint not null`
- `updated_at timestamptz not null`

Primary key:
- `(workbook_id, sheet_id, row_num, col_num)`

Rules:
- row exists only if the cell has visible non-default evaluated state
- includes spill children and pivot outputs
- missing row = empty/default evaluated cell

Purpose:
- hot synced grid model
- authoritative server-evaluated values

#### `cell_style`
Workbook style registry.

Columns:
- `workbook_id text not null`
- `style_id text not null`
- `record_json jsonb not null`
- `hash text not null`
- `created_at timestamptz not null`

Primary key:
- `(workbook_id, style_id)`

Purpose:
- deduped style definitions

#### `cell_number_format`
Workbook number-format registry.

Columns:
- `workbook_id text not null`
- `format_id text not null`
- `code text not null`
- `kind text not null`
- `created_at timestamptz not null`

Primary key:
- `(workbook_id, format_id)`

Purpose:
- deduped format registry
- range/cell explicit format references

#### `sheet_style_range`
Range-based style assignments.

Columns:
- `id text primary key`
- `workbook_id text not null`
- `sheet_id text not null`
- `start_row int not null`
- `end_row int not null`
- `start_col int not null`
- `end_col int not null`
- `style_id text not null`
- `source_revision bigint not null`
- `updated_at timestamptz not null`

Purpose:
- semantic style source
- avoids per-cell style materialization for large ranges

#### `sheet_format_range`
Range-based format assignments.

Columns:
- `id text primary key`
- `workbook_id text not null`
- `sheet_id text not null`
- `start_row int not null`
- `end_row int not null`
- `start_col int not null`
- `end_col int not null`
- `format_id text not null`
- `source_revision bigint not null`
- `updated_at timestamptz not null`

Purpose:
- semantic range-format source

#### `defined_name`
Workbook-scoped defined names.

Columns:
- `workbook_id text not null`
- `name text not null`
- `normalized_name text not null`
- `value_json jsonb not null`
- `source_revision bigint not null`
- `updated_at timestamptz not null`

Primary key:
- `(workbook_id, normalized_name)`

Purpose:
- formula binding inputs
- exact mapping to current `WorkbookDefinedNameValueSnapshot`

#### `table_def`
Workbook tables.

Columns:
- `id text primary key`
- `workbook_id text not null`
- `sheet_id text not null`
- `name text not null`
- `start_row int not null`
- `end_row int not null`
- `start_col int not null`
- `end_col int not null`
- `column_names_json jsonb not null`
- `header_row boolean not null`
- `totals_row boolean not null`
- `source_revision bigint not null`
- `updated_at timestamptz not null`

Purpose:
- structured reference source semantics

#### `pivot_def`
Workbook pivots.

Columns:
- `id text primary key`
- `workbook_id text not null`
- `sheet_id text not null`
- `name text not null`
- `anchor_row int not null`
- `anchor_col int not null`
- `source_sheet_id text not null`
- `source_start_row int not null`
- `source_end_row int not null`
- `source_start_col int not null`
- `source_end_col int not null`
- `group_by_json jsonb not null`
- `values_json jsonb not null`
- `rows int not null`
- `cols int not null`
- `source_revision bigint not null`
- `updated_at timestamptz not null`

Purpose:
- pivot definition source
- pivot output still lands in `cell_eval`

#### `spill_owner`
Dynamic-array owner metadata.

Columns:
- `workbook_id text not null`
- `sheet_id text not null`
- `owner_row int not null`
- `owner_col int not null`
- `rows int not null`
- `cols int not null`
- `calc_revision bigint not null`
- `updated_at timestamptz not null`

Primary key:
- `(workbook_id, sheet_id, owner_row, owner_col)`

Purpose:
- authoritative spill owner bounds
- child spill cells stay in `cell_eval`

#### `presence_coarse`
Low-frequency collaborative presence.

Columns:
- `workbook_id text not null`
- `user_id text not null`
- `session_id text not null`
- `sheet_id text not null`
- `active_row int null`
- `active_col int null`
- `selection_json jsonb null`
- `color text null`
- `updated_at timestamptz not null`

Primary key:
- `(workbook_id, session_id)`

Purpose:
- coarse cursor/selection presence
- not per-keystroke typing

### Internal tables (not on the primary client hot path)

#### `workbook_event`
Append-only ordered semantic transactions.

Columns:
- `workbook_id text not null`
- `revision bigint not null`
- `actor_user_id text not null`
- `client_mutation_id text null`
- `txn_json jsonb not null`
- `created_at timestamptz not null`

Primary key:
- `(workbook_id, revision)`

Purpose:
- replay
- audit
- snapshots
- background rebuilds
- shadow validation

#### `applied_client_mutation`
Idempotency receipts.

Columns:
- `workbook_id text not null`
- `user_id text not null`
- `client_mutation_id text not null`
- `revision bigint not null`
- `created_at timestamptz not null`

Primary key:
- `(workbook_id, user_id, client_mutation_id)`

Purpose:
- dedupe retries
- safe mutation replay handling

#### `recalc_job`
Transactional outbox for recalculation.

Columns:
- `id text primary key`
- `workbook_id text not null`
- `from_revision bigint not null`
- `to_revision bigint not null`
- `dirty_regions_json jsonb null`
- `status text not null` (`pending`, `leased`, `done`, `failed`)
- `attempts int not null default 0`
- `lease_until timestamptz null`
- `last_error text null`
- `created_at timestamptz not null`
- `updated_at timestamptz not null`

Purpose:
- async authoritative recalc
- exactly-once-ish background execution with leases

#### `workbook_snapshot`
Warm-start acceleration artifacts.

Columns:
- `workbook_id text not null`
- `revision bigint not null`
- `format text not null`
- `payload bytea or jsonb not null`
- `created_at timestamptz not null`

Primary key:
- `(workbook_id, revision)`

Purpose:
- fast rebuild of engine sessions
- not client sync source of truth

## 5.3 Why source/eval split is mandatory

This split is the single biggest design improvement over the current repo.

If you do not split:
- viewport queries become wider than necessary
- style/format metadata pollutes hot value sync
- formula source rides along when grid only needs values
- recomputation and user authoring state are coupled
- you cannot cheaply query just the selected cell’s source while the grid reads only evaluated values

The split gives:

- `cell_eval`: hot visible grid path
- `cell_source`: selected-cell/editor path
- styles/formats/ranges: semantic formatting path
- workbook metadata: low-frequency preload path

---

## 6. Schema details and indexing

## 6.1 Canonical upstream indexes

Create these indexes in Postgres because Zero implicitly mirrors useful upstream indexes into its replica.

### Hot value path
```sql
create index cell_eval_sheet_row_col_idx
  on cell_eval(workbook_id, sheet_id, row_num, col_num);

create index cell_source_sheet_row_col_idx
  on cell_source(workbook_id, sheet_id, row_num, col_num);

create index sheet_row_sheet_row_idx
  on sheet_row(workbook_id, sheet_id, row_num);

create index sheet_col_sheet_col_idx
  on sheet_col(workbook_id, sheet_id, col_num);
```

### Registry lookups
```sql
create index cell_style_workbook_idx
  on cell_style(workbook_id, style_id);

create index cell_number_format_workbook_idx
  on cell_number_format(workbook_id, format_id);
```

### Metadata
```sql
create index sheet_workbook_position_idx
  on sheet(workbook_id, position);

create index defined_name_workbook_norm_idx
  on defined_name(workbook_id, normalized_name);

create index table_def_workbook_sheet_idx
  on table_def(workbook_id, sheet_id);

create index pivot_def_workbook_sheet_idx
  on pivot_def(workbook_id, sheet_id);
```

### Range overlap helpers
```sql
create index sheet_style_range_workbook_sheet_rows_idx
  on sheet_style_range(workbook_id, sheet_id, start_row, end_row, start_col, end_col);

create index sheet_format_range_workbook_sheet_rows_idx
  on sheet_format_range(workbook_id, sheet_id, start_row, end_row, start_col, end_col);
```

### Queue and events
```sql
create index recalc_job_status_lease_created_idx
  on recalc_job(status, lease_until, created_at);

create index workbook_event_workbook_created_idx
  on workbook_event(workbook_id, created_at);
```

## 6.2 Coordinate rules

Store all cell/range coordinates as:
- `row_num int`
- `col_num int`

Do **not** use A1 strings in primary data tables.

A1 notation should exist only:
- at UI edges
- in formula source strings
- at import/export boundaries
- inside some JSON payloads where protocol compatibility requires it

Mutators should parse A1 once at ingress.

## 6.3 Postgres 17 reality in lab

The lab repo currently deploys Postgres 17. That means:
- do **not** depend on `generated stored` columns being synced through Zero
- if you want persisted tile helper columns, write them explicitly
- simpler initial plan: just query numeric row/col ranges directly and index them

## 6.4 Publication strategy

The current lab cluster already creates publication `zero_data FOR TABLES IN SCHEMA public`.

Recommended strategy:
- keep all published Zero tables in `public`
- keep internal tables in `public` initially if that is simplest
- but do not expose internal tables in the client Zero schema
- if later you want stronger separation, move internal tables to a non-published schema or a separate publication

---

## 7. The Zero schema package design

## 7.1 Keep `packages/zero-sync`, replace its contents

The package name is already imported widely. Keep it.

Replace:
- `src/schema.ts`
- `src/queries.ts`
- `src/mutators.ts`
- `src/snapshot.ts`

with a real production model.

Recommended package structure:

```text
packages/zero-sync/src/
  schema.ts
  relationships.ts
  queries.ts
  mutators.ts
  args.ts
  context.ts
  runtime-config.ts
  tile.ts
  selectors.ts
  presence.ts
  zql.ts
  index.ts
```

## 7.2 Schema design in Zero

Published tables in `schema.ts` should include only the client-facing model:

- `workbook`
- `workbookMember` (optional on client if you want share UI)
- `sheet`
- `sheetFilter`
- `sheetSort`
- `sheetRow`
- `sheetCol`
- `cellSource`
- `cellEval`
- `cellStyle`
- `cellNumberFormat`
- `sheetStyleRange`
- `sheetFormatRange`
- `definedName`
- `tableDef`
- `pivotDef`
- `spillOwner`
- `presenceCoarse`

Do **not** put `workbook_event`, `recalc_job`, or `workbook_snapshot` into the client schema.

## 7.3 Relationships

Relationships should be minimal and helpful, not fancy for their own sake.

Recommended:
- workbook → sheets
- sheet → workbook
- sheet → cellEval
- sheet → cellSource
- sheet → row/col overrides
- workbook → styles
- workbook → formats
- workbook → defined names
- workbook → tables
- workbook → pivots
- workbook → presence
- sheetStyleRange → style
- sheetFormatRange → numberFormat

Do not over-chain relationships on hot tile queries. Use direct flat queries where they are cheaper and more stable.

---

## 8. Query model: how Zero should actually be used here

## 8.1 First principle

The client should subscribe to **stable tile-shaped queries**, not arbitrary whole-workbook queries and not per-scroll-pixel queries.

Recommended tile sizes for v1:
- `128` rows × `32` columns for `cell_eval`
- `256` rows for `sheet_row`
- `64` columns for `sheet_col`

These are starting points. Benchmark them. Do not hard-code them forever.

## 8.2 Query families

### Workbook/sheet bootstrap
- `workbook.get({workbookID})`
- `sheet.list({workbookID})`

These are always on for an open workbook.

### Grid value tiles
- `cellEval.tile({workbookID, sheetID, rowStart, rowEnd, colStart, colEnd})`

This is the main hot path.

### Selected cell source
- `cellSource.get({workbookID, sheetID, rowNum, colNum})`

Used for:
- formula bar
- in-cell editor
- inspector
- exact formula source display

### Axis metadata
- `sheetRow.tile({workbookID, sheetID, rowStart, rowEnd})`
- `sheetCol.tile({workbookID, sheetID, colStart, colEnd})`

### Formatting semantics
- `sheetStyleRange.intersectTile(...)`
- `sheetFormatRange.intersectTile(...)`
- `cellStyle.byWorkbook({workbookID})`
- `cellNumberFormat.byWorkbook({workbookID})`

Preload style and format registries per workbook. They are usually small.

### Workbook metadata
- `definedName.byWorkbook({workbookID})`
- `tableDef.bySheet({workbookID, sheetID})`
- `pivotDef.bySheet({workbookID, sheetID})`
- `spillOwner.tile(...)` or `spillOwner.bySheet(...)` depending on cardinality

### Presence
- `presence.byWorkbook({workbookID})`

## 8.3 Query behavior in the browser

### What to preload
At workbook open:
- workbook header
- sheet list
- styles/formats registry
- active-sheet metadata
- active-sheet tile ring (visible + overscan)
- names / tables / pivots for active workbook or active sheet

### What not to preload
- whole workbook cells
- all computed cells
- all source cells
- all spills and pivots across a large workbook unless needed

## 8.4 The bridge pattern

Create a **Zero workbook bridge** that converts Zero query updates into the existing viewport-patch contract.

Recommended location:
- `apps/web/src/zero/ZeroWorkbookBridge.ts`
- or inside worker runtime if you want this off the main thread

Responsibilities:
- own current visible tile subscriptions
- own overscan preload
- listen to query changes
- normalize rows into the existing cache model
- emit `ViewportPatch` bytes compatible with the current grid
- keep selected-cell source and visible-cell eval state in sync

This lets you preserve:
- `packages/worker-transport/src/viewport-patch.ts`
- the current grid cache model
- most of the UI rendering stack

## 8.5 How granular updates should work

Use Zero’s `materialize()` and a custom `View` implementation for hot tile queries.

Why:
- `useQuery()` is fine for simple React consumption
- but the grid already wants a patch stream
- the bridge should consume **add/edit/remove** changes, not re-render huge arrays

Recommended use:
- `materialize()` for each active tile query
- custom `View` to get add/edit/remove events
- turn those events into cell/row/column/style cache updates
- emit minimal viewport patches

This is the cleanest way to align Zero with the existing patch-driven UI.

## 8.6 Avoid query churn

Do not register a new query every few pixels of scroll.

Instead:
- snap viewport to tiles
- maintain visible tile set + overscan ring
- reuse subscriptions while viewport stays within a tile
- preload adjacent tiles when scroll velocity suggests the next ring

---

## 9. Mutation model

## 9.1 Replace CRDT batches with ordered workbook transactions

The new write model is:

- every workbook mutation is an ordered transaction
- transactions are serialized per workbook
- each committed transaction advances `workbook.head_revision`
- each committed transaction appends one `workbook_event`
- async recalc jobs derive `cell_eval` updates and advance `workbook.calculated_revision`

No replica clocks.
No batch merge logic.
No convergence algorithm.
Just one durable server order.

## 9.2 Mutation classes

### Class A — direct SQL mutations
Safe to execute directly in the mutator transaction without a planner:

- set literal cell value
- set literal formula source
- clear cell
- update row height / hidden
- update column width / hidden
- set freeze panes
- update coarse presence
- set workbook properties
- set calc mode / compatibility mode

These mutate source tables directly and enqueue recalc if needed.

### Class B — engine-planned semantic mutations
Need bilig semantic planning because they rewrite many dependent structures:

- fill range
- copy range
- paste large ranges with translation
- insert rows / delete rows / move rows
- insert columns / delete columns / move columns
- rename sheet
- reorder sheets
- delete sheet
- upsert/delete table
- upsert/delete pivot
- defined names with formula/reference semantics
- style/format range normalization when overlapping existing ranges
- future structured-reference-aware operations

These should run through a planner API backed by the bilig engine.

## 9.3 Mutation API shape

### Required common mutation envelope
Every mutator should receive:
- `workbookID`
- `clientMutationID`
- mutation-specific payload

Example:
```ts
{
  workbookID: string;
  clientMutationID: string;
  ...
}
```

This is required for idempotency and retry safety.

## 9.4 Recommended mutator surface

### Workbook
- `workbook.rename`
- `workbook.setProperties`
- `workbook.setCalculationSettings`
- `workbook.forceRecalc`

### Sheet
- `sheet.create`
- `sheet.rename`
- `sheet.reorder`
- `sheet.delete`
- `sheet.setFreezePane`
- `sheet.clearFreezePane`
- `sheet.setFilter`
- `sheet.clearFilter`
- `sheet.setSort`
- `sheet.clearSort`

### Cells
- `cell.editOne`
- `cell.editRange`
- `cell.clearRange`
- `cell.fillRange`
- `cell.copyRange`
- `cell.applyPatches` (normalized batch of cell source edits)

### Formatting
- `format.setCellFormat`
- `format.setRangeFormat`
- `format.clearRangeFormat`
- `style.setRangeStyle`
- `style.clearRangeStyleFields`

### Structure
- `row.insert`
- `row.delete`
- `row.move`
- `row.resize`
- `row.hide`
- `row.unhide`
- `col.insert`
- `col.delete`
- `col.move`
- `col.resize`
- `col.hide`
- `col.unhide`

### Metadata
- `definedName.upsert`
- `definedName.delete`
- `table.upsert`
- `table.delete`
- `pivot.upsert`
- `pivot.delete`

### Presence
- `presence.update`

## 9.5 The server write transaction contract

Every mutator should do the following in one DB transaction:

1. authenticate user
2. authorize workbook + role
3. dedupe `clientMutationID`
4. acquire per-workbook serialization lock
5. read current workbook header / relevant rows
6. either:
   - apply direct SQL source changes, or
   - call planner and apply planner-produced deltas
7. increment `workbook.head_revision`
8. write `workbook_event`
9. insert `recalc_job`
10. persist mutation receipt
11. return revision info

### Locking recommendation
Use one of:
- `select ... for update` on the `workbook` row
- plus `pg_advisory_xact_lock(hashtext(workbook_id))`

The simplest robust approach is:
- lock workbook row
- advisory lock by workbook ID for explicit per-workbook serialization

## 9.6 Planner output contract

Planner-backed mutations should produce:

```ts
interface PlannedWorkbookDelta {
  sourceUpserts: ...
  sourceDeletes: ...
  rowUpserts: ...
  rowDeletes: ...
  colUpserts: ...
  colDeletes: ...
  styleRegistryUpserts: ...
  numberFormatRegistryUpserts: ...
  styleRangeUpserts: ...
  styleRangeDeletes: ...
  formatRangeUpserts: ...
  formatRangeDeletes: ...
  sheetUpserts: ...
  sheetDeletes: ...
  tableUpserts: ...
  tableDeletes: ...
  pivotUpserts: ...
  pivotDeletes: ...
  definedNameUpserts: ...
  definedNameDeletes: ...
  dirtyRegions: ...
  txn: WorkbookTxn;
}
```

Mutators apply this delta transactionally.

This avoids putting full-engine recomputation inside the mutator itself.

---

## 10. Formula architecture with bilig WASM

## 10.1 The hard rule

The **only formula source of truth** is:
- `cell_source.formula_source`

Everything else is derived.

Do not persist:
- lowered ASTs
- compiled WASM programs
- dependency graph internals
- string pool internals
- engine-private memoized structures
- worker-runtime caches

## 10.2 The authoritative calc pipeline

### Commit path
1. browser commits a mutator
2. mutator writes semantic source state and revision/event/outbox row
3. UI shows optimistic local preview where possible
4. recalc worker consumes `recalc_job`
5. worker materializes authoritative `cell_eval`, `spill_owner`, and any changed derived metadata
6. worker updates `workbook.calculated_revision`
7. Zero syncs changed rows to all interested clients

### Why this is the right split
- mutators stay fast
- calc cost is moved out of request latency
- authoritative values are still derived quickly
- the bilig engine stays the semantic oracle
- clients can remain instant via local Zero data + preview overlay

## 10.3 Workbook service design

Create a workbook service runtime that hosts **warm `SpreadsheetEngine` sessions** keyed by workbook ID.

Recommended reuse source:
- patterns from `apps/local-server/src/local-workbook-session-manager.ts`

### Engine session responsibilities
- warm `SpreadsheetEngine` pool
- load snapshot + tail events
- apply planned or committed workbook txns
- run recalc
- export snapshots periodically
- diff outputs vs existing materialization
- write `cell_eval` updates
- write spill metadata
- maintain eviction TTL / LRU

### Warm-start strategy
For a recalc job:
1. look for in-memory warm session
2. if absent, load latest `workbook_snapshot`
3. replay `workbook_event` revisions after snapshot
4. apply target transaction(s)
5. recalc
6. diff + write materialization

Snapshot cadence:
- every `N` revisions (recommend 100–250)
- and on idle shutdown
- keep last few snapshots per workbook
- periodically compact old ones

## 10.4 Diff materialization rules

The recalc worker should not blindly rewrite all `cell_eval` rows.

It should diff:
- previous evaluated rows in affected regions
- new evaluated rows in affected regions

Then:
- upsert changed rows
- delete rows that became default/blank
- leave untouched rows untouched

This is crucial for Zero efficiency and query churn.

## 10.5 Dirty region strategy

### Minimum viable region strategy
Mutators enqueue dirty regions conservatively:
- edited cells
- pasted ranges
- impacted source/target ranges for fill/copy
- whole sheet for structural ops if needed

The recalc worker may widen dirty regions using engine semantics.

### Advanced strategy
Later, if needed:
- engine can expose more exact impacted regions
- or a dependency-aware invalidation summary
- but do not block v1 on perfect invalidation metadata

## 10.6 Formula preview in the browser

Keep local bilig parsing and formatting in the worker.

Recommended browser preview behavior:
- literal edits: always optimistic
- simple formulas with hydrated dependencies: optimistic preview allowed
- formulas with missing dependencies / cross-sheet context not in memory: show pending state
- authoritative `cell_eval` always wins

This preserves “instant feel” without lying about offline-write semantics.

## 10.7 Names, tables, pivots, spills

### Defined names
- stored in `defined_name.value_json`
- bilig binder resolves them
- planner/recalc worker loads them into engine before evaluating formulas

### Tables
- `table_def` stores durable table source semantics
- structured references bind through bilig
- table edits must run through planner-backed mutations

### Pivots
- `pivot_def` stores pivot definition
- pivot outputs land in `cell_eval`
- pivot lifecycle is planner-backed

### Spills
- owner formula source stays in `cell_source`
- owner bounds live in `spill_owner`
- child rendered cells live in `cell_eval`
- blocked spill = owner error in `cell_eval`, not fake child semantics

---

## 11. Browser runtime plan

## 11.1 Preserve the worker-first UI contract

Do **not** rewrite the product shell into direct React state from Zero queries.

Keep:
- worker runtime as a dedicated runtime boundary
- viewport patches as the grid input
- main thread focused on rendering and input

## 11.2 Introduce `ZeroWorkbookBridge`

Recommended new browser components:

```text
apps/web/src/zero/
  ZeroWorkbookBridge.ts
  tile-subscriptions.ts
  viewport-projector.ts
  presence.ts
  query-hooks.ts
```

### Responsibilities
- convert active grid viewport into tile query subscriptions
- preload overscan
- consume `cell_eval`, formatting, row/col metadata, and selected `cell_source`
- build/update `WorkerViewportCache`-compatible projection state
- emit current `ViewportPatch` payloads

## 11.3 Remove whole-snapshot Zero path

Delete/replace these patterns in `apps/web/src/WorkerWorkbookApp.tsx`:

Current pattern:
- `useQuery(queries.workbooks.byId({documentId}))`
- `projectWorkbookToSnapshot(...)`
- `replaceSnapshot(...)`
- `mutators.workbook.replaceSnapshot(...)`

Replace with:
- workbook bootstrap query
- sheet list query
- tile subscriptions
- selected-cell source query
- semantic mutators only

## 11.4 Keep these worker responsibilities

The worker should continue to own:
- formula bar helpers
- parse services
- display/copy/editor text formatting helpers
- optimistic preview overlays
- viewport patch encoding
- selected cell inspection
- maybe lightweight local workbook projection state

The worker should **not** remain the durable sync authority for Zero mode.

## 11.5 Selected-cell edit flow

### On selection change
1. selected cell changes
2. bridge ensures `cell_source.get` query for that cell
3. worker/editor state updates formula bar with exact source

### On edit commit
1. worker computes best-effort local preview
2. browser fires semantic Zero mutator
3. bridge marks cell as pending if needed
4. authoritative `cell_eval` arrives
5. pending preview cleared

## 11.6 Offline / disconnected UX

Zero-mode UX rules:
- `connecting`: allow editing, but show reconnect/pending badge
- `connected`: normal
- `disconnected`: switch workbook to read-only
- `error`: read-only + retry affordance
- `needs-auth`: read-only + refresh auth flow

The UI must stop pretending it can safely commit edits once Zero rejects writes.

Recommended implementation:
- expose connection state, not just boolean online/offline
- show a visible mode banner
- disable formula bar and cell input when writes are unsafe

---

## 12. Auth and permissions

## 12.1 Current repo problem

`apps/web/src/main.tsx` currently mounts `ZeroProvider` with:

```tsx
userID="anon"
```

This must not survive into production.

## 12.2 Production auth model

### Browser
- obtain real authenticated user session
- mount `ZeroProvider` with real `userID`
- ensure query/mutate endpoints receive auth cookie or token

### Server
- parse auth on `/api/zero/query` and `/api/zero/mutate`
- build Zero context:
  - `userID`
  - `roles`
  - maybe org/account context later

### Authorization
- all query and mutate paths validate against `workbook_member`

## 12.3 Permission rules

At minimum:
- viewer: read workbook
- commenter: read + presence + comments if added later
- editor: read + write
- owner: full rights + sharing

### Query enforcement
Queries should filter by `ctx.userID` through membership.

### Mutator enforcement
Mutators should require role checks before any write.

## 12.4 Multi-tenant hardening

If this becomes multi-tenant SaaS:
- add `account_id` / `org_id` to workbook and member tables
- make it part of all hot indexes
- ensure auth context includes it
- prefer compound uniqueness scoped to tenant

---

## 13. Presence and collaboration state

## 13.1 What should go through Zero

Good for Zero:
- active sheet
- selected cell/range
- coarse cursor/selection updates
- collaborator colors
- “currently viewing sheet X”

Recommended update frequency:
- 2–4 Hz max
- debounced/throttled

## 13.2 What should not go through Zero initially

Do not put these on the main correctness path yet:
- every keystroke in a text editor
- sub-100ms typing presence
- transient IME composition state
- cursor animation trails

If needed later:
- separate ephemeral websocket channel
- or keep Redis for non-durable presence only

## 13.3 Redis recommendation

The current lab repo deploys Redis for the old sync/presence model.

Recommended production path:
- remove Redis from correctness path
- keep it only if/when you add high-frequency ephemeral awareness later
- after Zero migration stabilizes, Redis can likely be removed entirely unless agent or presence features still need it

---

## 14. Service and codebase refactor plan

## 14.1 Package/file changes by area

### A. Extract domain model from CRDT

Create:
- `packages/workbook-domain`

Move from `packages/crdt/src/index.ts`:
- `WorkbookOp`
- related op payload types
- new `WorkbookTxn`
- helper types for planner outputs

Leave in `packages/crdt` temporarily:
- re-exports with deprecation comments

Delete later:
- replica clocks
- batch comparison
- applied batch tracking
- replica snapshots
- convergence helpers

### B. Refactor `packages/core`

Add:
- a public transaction/planner seam
- engine-friendly apply/replay helpers that do not require CRDT replica concepts
- ability to load from relational source state or snapshot + events

Add/complete missing op families:
- `renameSheet`
- `reorderSheets`
- explicit range ops if needed
- any missing structured-ref / spill-owner ops needed for durable semantics

### C. Replace `packages/zero-sync`

Rebuild:
- schema
- queries
- mutators
- shared argument validators
- runtime config helpers
- tile helpers

Delete:
- whole-workbook snapshot projection helpers from the hot path

### D. Repurpose `apps/sync-server`

Short-term keep app/deployment name:
- because `lab` already wires ingress and service names

Internally split modules into:
```text
apps/sync-server/src/
  zero-api/
    query.ts
    mutate.ts
    auth.ts
    context.ts
  workbook-service/
    session-pool.ts
    planner.ts
    recalc-worker.ts
    diff-materializer.ts
    snapshots.ts
    jobs.ts
```

Retire:
- `/v1/frames`
- websocket browser sync
- CRDT cursor catch-up endpoints
- binary frame durability as the primary product path

Keep only if still needed temporarily:
- agent endpoints

### E. Rework `apps/web`

Replace:
- whole-workbook `useQuery`
- `replaceSnapshot` writeback loop

Add:
- tile subscriptions
- Zero workbook bridge
- selected-cell source fetch
- auth-derived user ID
- connection-state-driven read-only mode

### F. Rework `apps/local-server`

Keep for:
- local development
- engine test harnesses
- maybe import/export utilities

Extract reusable session/warm engine utilities into shared runtime package.

### G. Cleanup `packages/storage-server` and binary protocol surfaces

After CRDT sync path removal:
- reduce scope of `packages/storage-server`
- likely deprecate if only used for old sync path
- binary protocol may remain only for agent tooling, if still useful

---

## 15. Detailed implementation phases

## Phase 0 — Decision freeze and branch prep

### Deliverables
- architecture decision record:
  - Zero is the only sync layer
  - CRDT transport is deprecated
  - server-authoritative multiplayer is the product truth
- engineering owner assignment by package/app
- rollout flag strategy

### Required codebase changes
- create tracking issue set
- add deprecation notice docs for `@bilig/crdt` and old sync ingress
- freeze new feature work on old sync path unless it is migration-critical

### Exit gate
- team agrees on server-authoritative model
- no ongoing product work assumes old CRDT sync will survive

---

## Phase 1 — Extract the domain model from CRDT

### Goals
- stop coupling semantic workbook ops to replica clocks
- keep compile impact controlled

### Work
1. create `packages/workbook-domain`
2. move domain types:
   - `WorkbookOp`
   - row/column/style/name/table/pivot payload types
3. define `WorkbookTxn`
4. update imports in:
   - `packages/core`
   - `packages/worker-transport`
   - `apps/web`
   - `apps/local-server`
   - `packages/binary-protocol`
5. leave `packages/crdt` as thin deprecated compatibility layer

### Exit gate
- no core package depends on CRDT replica semantics except old sync code
- domain mutation language stands alone

---

## Phase 2 — Add the new Postgres schema behind the current app

### Goals
- create relational source/eval tables without switching read path yet
- keep old Zero snapshot path alive during migration

### Work
1. add SQL migrations for all new tables
2. add indexes
3. keep existing `workbooks/sheets/cells/computed_cells/...` tables temporarily
4. build backfill script:
   - load current workbook snapshot
   - import into engine
   - export relational source rows
   - materialize initial `cell_eval`
5. run backfill for dev/staging datasets

### Lab alignment
- apply additive schema migrations first
- keep publication stable
- verify `zero-cache` backfill before client switch

### Exit gate
- new tables populated for test workbooks
- old UI still works
- new tables can fully reconstruct workbook semantics

---

## Phase 3 — Build the workbook service and recalc pipeline

### Goals
- stop doing full-engine recalc inside request path
- stop rewriting full workbook projections

### Work
1. create workbook session pool
2. implement snapshot load/replay logic
3. implement `recalc_job` lease worker
4. implement source→engine loader
5. implement `cell_eval` diff materializer
6. write snapshot cadence logic
7. add idempotent mutation receipts
8. instrument recalc metrics

### Temporary strategy
- run recalc worker inside `bilig-sync` process first
- later split into dedicated worker deployment only if needed

### Exit gate
- a committed source mutation updates `cell_eval` through background recalc
- no request-path mutation rewrites entire workbook projections

---

## Phase 4 — Replace `packages/zero-sync` with real queries/mutators

### Goals
- make Zero read/write the real relational workbook model
- keep browser still using current UI shell

### Work
1. rewrite `schema.ts`
2. rewrite `queries.ts` to tile-based queries
3. rewrite `mutators.ts` to semantic mutators
4. rewrite server query registration
5. rewrite server mutate dispatch
6. remove `replaceSnapshot` mutator from production path
7. keep old snapshot helpers only for migration tooling if needed

### Exit gate
- Zero API surfaces map to relational workbook model
- no whole-workbook query remains on hot path

---

## Phase 5 — Introduce `ZeroWorkbookBridge` and shadow-read mode in web

### Goals
- preserve existing grid/UI runtime while moving authority to Zero tiles

### Work
1. add tile calculation helpers
2. add workbook bootstrap query path
3. add tile subscription manager
4. add selected-cell source query
5. add style/format range resolver
6. add patch emitter compatible with existing `ViewportPatch`
7. add feature flag:
   - `ZERO_VIEWPORT_BRIDGE=on/off`
8. run shadow mode:
   - old worker state vs new Zero bridge state
   - compare visible patches/cell values

### Exit gate
- grid renders from Zero-backed viewport patches
- differences are measured and explainable
- whole snapshot `replaceSnapshot` loop removed from normal flow

---

## Phase 6 — Move simple edits to semantic Zero mutators

### Goals
- use real write path for common actions

### Work
Move first:
- edit one cell
- clear one cell
- edit range paste
- resize column
- row/col hide/unhide
- freeze panes
- presence update

Keep behind planner later:
- fill/copy
- structural ops
- rename/reorder sheets
- tables/pivots

### Exit gate
- normal spreadsheet editing works through relational source/eval path
- old snapshot writeback is not used for common flows

---

## Phase 7 — Move complex/planner-backed operations

### Goals
- close feature parity with existing engine semantics

### Work
Implement planner-backed mutators for:
- fill range
- copy range
- insert/delete/move rows
- insert/delete/move columns
- rename sheet
- reorder sheets
- delete sheet
- defined name reference semantics
- table lifecycle
- pivot lifecycle
- style/format range normalization

### Exit gate
- all durable worksheet semantics run through Zero + workbook service
- old CRDT/sync-server mutation path is no longer required

---

## Phase 8 — Remove old sync plane

### Goals
- fully deprecate CRDT and binary browser sync path

### Work
1. remove browser websocket sync dependency
2. remove `/v1/frames` and browser sync frame handling
3. remove old cursor catch-up/snapshot sync endpoints if unused
4. remove replica snapshot persistence for product path
5. remove CRDT convergence tests and replace with revision-order tests
6. keep only agent surfaces that still matter

### Exit gate
- production browser uses Zero only
- CRDT is not on the request path or browser runtime path

---

## Phase 9 — Infra hardening and scale-up

### Goals
- productionize after correctness is closed

### Work
1. zero-cache tuning
2. query analyzer on hot queries
3. recalc worker concurrency tuning
4. snapshot retention and vacuum strategy
5. optional split of recalc worker from API process
6. optional multi-node zero topology
7. optional removal of Redis if no longer needed

### Exit gate
- staging and production SLOs hold
- deployment and rollback paths are documented

---

## 16. Concrete browser implementation plan

## 16.1 `apps/web/src/main.tsx`

### Change
Replace:
```tsx
userID="anon"
```

With:
- real auth-derived user ID
- auth-aware Zero provider boot
- explicit connection-state handling

### Add
- auth bootstrapping hook
- error/loading shell if auth or schema update needed

## 16.2 `apps/web/src/WorkerWorkbookApp.tsx`

### Remove
- `useQuery(queries.workbooks.byId({documentId}))`
- `projectWorkbookToSnapshot(...)`
- `replaceSnapshot(remoteSnapshot)`
- `mutators.workbook.replaceSnapshot(...)`

### Add
- workbook header hook
- sheet list hook
- selected-sheet ID map
- bridge lifecycle
- explicit connection-state banner/read-only mode
- selected-cell source hook

### Keep
- formula bar
- selection model
- optimistic local editor behavior
- styling ribbon behavior
- grid integration

## 16.3 `apps/web/src/worker-runtime.ts`

### Keep
- viewport patch encoder/output
- formatting helpers
- selected cell services
- parse/preview helpers
- clipboard helpers

### Change
- stop assuming authoritative engine state always comes from snapshot/CRDT path
- add authoritative overlay mode backed by Zero bridge
- keep `replaceSnapshot` only for local/dev or migration debug

### Recommended new responsibility split
- worker runtime = local UX runtime
- Zero bridge = authoritative remote state projector

## 16.4 `packages/worker-transport`

Keep the current viewport patch payload stable for the first migration.
Do not combine this migration with a patch codec redesign.

Typed binary patch codecs can come later. The sync architecture rewrite is already large enough.

---

## 17. Concrete server implementation plan

## 17.1 `apps/sync-server/src/server.ts`

### Keep
- `/healthz`
- `/api/zero/query`
- `/api/zero/mutate`
- agent endpoints if still needed

### Remove later
- `/v1/frames`
- websocket browser sync gateway
- old sync browser attach/detach flow

## 17.2 `apps/sync-server/src/zero/service.ts`

Replace current prototype service with:
- real auth/context extraction
- real query registry
- real mutator dispatch
- real DB transaction helpers
- shared workbook service runtime access

## 17.3 New internal modules

Recommended:

```text
apps/sync-server/src/workbook-service/
  context.ts
  load-workbook.ts
  apply-delta.ts
  planner.ts
  session-pool.ts
  materialize-cell-eval.ts
  reconcile-style-format.ts
  snapshots.ts
  jobs.ts
  idempotency.ts
```

## 17.4 Query endpoint implementation

The query endpoint should:
- parse auth
- build Zero context
- dispatch only allowed queries
- avoid exposing internal tables
- keep query identities stable

## 17.5 Mutator endpoint implementation

The mutate endpoint should:
- parse auth
- dispatch mutator by name
- wrap each call in transaction
- dedupe `clientMutationID`
- serialize writes per workbook
- emit outbox/recalc job
- return revision info

---

## 18. Planner design for complex operations

## 18.1 Why planner-backed mutations exist

Some spreadsheet operations are not plain row updates:
- row insert/delete/move rewrites formulas, tables, pivots, spills, named ranges, filters, sorts
- fill/copy needs translation semantics
- rename sheet rewrites sheet references
- reorder sheets affects UI semantics and maybe external assumptions

These must be planned with bilig’s engine semantics.

## 18.2 Planner contract

Planner input:
```ts
interface PlannerRequest {
  workbookID: string;
  baseRevision: number;
  action: WorkbookAction;
}
```

Planner output:
```ts
interface PlannerResponse {
  delta: PlannedWorkbookDelta;
  diagnostics?: unknown[];
}
```

## 18.3 How planner should execute

Recommended v1:
- planner runs inside the workbook service process
- uses a warm `SpreadsheetEngine`
- applies the user’s requested operation to the warm engine
- computes normalized source/metadata delta
- returns delta to the mutator layer

This is implementation-ready with the current codebase.

Do **not** try to implement structural rewrite semantics directly in SQL first.

---

## 19. Snapshot strategy

## 19.1 Snapshots should remain, but only as accelerators

Keep snapshots because they are useful for:
- fast warm start
- disaster rebuild
- import/export
- debugging

Do not use them as:
- hot client sync payload
- semantic source of truth
- normal edit roundtrip payload

## 19.2 Snapshot storage format

Recommended:
- compressed JSONB or compressed bytes
- keyed by `(workbook_id, revision)`
- store every 100–250 revisions and on shutdown/idle
- retain last 3–5 snapshots per active workbook plus periodic long-lived checkpoints

## 19.3 Snapshot build path

The workbook service should be able to:
- export snapshot from warm engine
- write to `workbook_snapshot`
- optionally mirror to object storage later if snapshot size or retention needs grow

The current lab manifests already mention snapshot bucket config as optional. That can remain deferred.

---

## 20. Migration and data backfill

## 20.1 Expand / migrate / contract sequence

### Expand
- add new relational tables
- keep old snapshot/projection tables
- deploy server capable of populating both if needed

### Migrate
- backfill relational tables from existing snapshots
- enable new query bridge in shadow mode
- switch mutations gradually

### Contract
- stop writing old snapshot/projection tables
- remove old browser sync path
- delete obsolete tables/endpoints later

## 20.2 Backfill algorithm

For each workbook:
1. load current snapshot from `workbooks.snapshot`
2. import into `SpreadsheetEngine`
3. read:
   - sheets
   - source cells
   - row/col metadata
   - styles/formats/ranges
   - names
   - tables
   - pivots
   - spills
4. materialize `cell_eval`
5. write `workbook_snapshot` checkpoint
6. mark workbook migrated

## 20.3 Dual-write recommendation

Short-term during migration:
- allow mutators to dual-write old snapshot/projection tables and new tables
- only until shadow verification is green
- then remove dual write quickly

Do not leave dual write around for long. It becomes a correctness trap.

---

## 21. Testing and validation plan

## 21.1 Replace old sync correctness tests with new correctness tests

Remove emphasis on:
- CRDT convergence
- duplicate batch replay
- replica clock comparison

Add emphasis on:
- revision-ordered workbook transaction correctness
- source/eval materialization correctness
- planner delta correctness
- browser bridge correctness

## 21.2 Required test layers

### Unit
- source row mappers
- cell_eval row encoders/decoders
- style/format range intersection logic
- planner delta application
- idempotency receipts
- auth permission filters

### Integration
- mutator transaction tests against Postgres
- recalc worker end-to-end
- snapshot warm-start replay
- revision ordering under concurrency
- tile query correctness

### Browser / Playwright
Must cover:
- single-cell edit
- formula edit
- paste large range
- fill handle
- row/column resize
- hide/unhide
- reconnect/read-only mode
- selection presence
- million-row navigation
- workbook reopen
- collaborative second-tab propagation

### Shadow equivalence tests
For the same workbook and edit stream:
- old engine/local path vs new Zero path
- compare visible value
- compare flags
- compare display text
- compare style/format results for visible cells
- compare spills/tables/pivots/names behavior

This shadow layer is the most important confidence mechanism.

## 21.3 Performance tests

Retain current repo budgets and add service budgets.

### Existing bilig UI/runtime budgets to preserve
- local visible edit response p95 `< 16ms`
- `10k` downstream recalc p95 `< 25ms`
- `100k` workbook restore p95 `< 500ms`
- `250k` preset restore p95 `< 1500ms`

### New service budgets to add
Recommended starting SLOs:
- mutator DB commit p95 `< 60ms` for simple cell edits
- authoritative same-region cell propagation p95 `< 250ms`
- tile query materialization p95 `< 30ms` on warm cache
- recalc queue lag p95 `< 100ms` under normal load
- workbook warm-start from snapshot p95 `< 300ms` for common active workbooks

Use these as starting operational targets, then tune with real workloads.

---

## 22. Observability plan aligned to lab

The lab repo already defines rollout and observability contracts. Extend them instead of inventing a separate system.

## 22.1 Required new metrics

### Mutation path
- mutation count by name
- mutation latency
- mutation error count
- idempotency-hit count
- planner latency

### Recalc path
- recalc job enqueue count
- queue depth
- lease timeouts
- recalc duration
- changed cell count per job
- rows written / deleted in `cell_eval`
- snapshot restore time
- warm-session hit rate

### Zero path
- query count by query name
- query latency
- sync lag (`head_revision - calculated_revision`)
- active tile subscription count
- client connection states

### Browser path
- optimistic preview apply time
- authoritative convergence time
- bridge patch generation time
- patch size
- visible tile cache hit rate

## 22.2 Critical dashboards

Operators must be able to answer:
- which workbooks are behind on authoritative recalc?
- which mutators are slow or erroring?
- are tile queries causing temp sorts?
- are large style/format ranges hurting query latency?
- is the bridge emitting too many full patches?
- did a rollout regress formula latency or visible edit latency?

## 22.3 Alerts

Recommended alerts:
- recalc queue depth too high
- recalc job stuck leased
- `workbook.head_revision - calculated_revision` exceeds threshold
- Zero query endpoint 5xx rate
- Zero mutate endpoint 5xx rate
- p95 tile query latency regression
- snapshot restore failures
- auth failure spike

---

## 23. Infrastructure plan aligned to the current lab manifests

## 23.1 Keep the existing public routing shape

Current routing in lab is good:
- `/api/zero` → `bilig-sync`
- `/zero` → `bilig-zero`
- main site → `bilig-web`

Keep that.

## 23.2 Keep the existing Zero deployment first

Current:
- one `bilig-zero` deployment
- persistent volume for replica
- direct DB URI

That is a good starting production topology.

Do not split into replication-manager/view-syncer until:
- you have demonstrated need
- single-node zero-cache is actually the bottleneck

## 23.3 Keep Postgres as is

The existing CNPG cluster is already suitable:
- Postgres 17
- logical replication on
- publication present

Recommended DB additions:
- careful index creation
- autovacuum tuning for `cell_eval` and event tables
- partitioning only later if table sizes demand it

## 23.4 `bilig-sync` deployment evolution

Phase 1:
- keep deployment name
- increase memory for workbook sessions if needed
- run API + recalc worker in same pod process or same binary

Phase 2:
- split recalc worker into separate deployment if queue latency or noisy-neighbor issues appear

## 23.5 Redis

Keep deployment for now only if:
- agent API still depends on it
- or you want future ephemeral presence

Otherwise remove it after migration.

## 23.6 Rolling updates

Use additive migration order:
1. DB expand migration
2. deploy `bilig-sync` with new server logic
3. wait for zero-cache backfill
4. deploy web using new schema/queries
5. remove obsolete tables and endpoints later

This is mandatory to avoid schema/version reload loops.

---

## 24. Risk register and mitigation

## Risk 1 — Structural ops are more complex than expected
**Why:** row/col insert/delete/move affect formulas, tables, pivots, filters, sorts, spills, and names.

**Mitigation:**
- do not implement them directly in SQL
- use planner-backed engine semantics
- ship simple mutations first
- gate structural operations behind shadow tests

## Risk 2 — Large formatting ranges create bad overlap-query performance
**Why:** 2D intersection queries can degrade.

**Mitigation:**
- start simple
- benchmark real style cardinality
- if needed, add precomputed tile coverage side tables later
- do not prematurely materialize all styles per cell

## Risk 3 — Browser bridge produces too many full patches
**Why:** naive React query consumption is too coarse.

**Mitigation:**
- use `materialize()` + custom `View`
- diff into patch cache incrementally
- measure full-patch frequency

## Risk 4 — Snapshot and event replay drift
**Why:** two restore paths can diverge.

**Mitigation:**
- one authoritative transaction language
- regular rebuild tests from snapshot + tail events
- shadow validation jobs

## Risk 5 — Teams keep calling the product “local-first”
**Why:** product expectations become wrong.

**Mitigation:**
- explicitly rename architecture internally
- implement read-only disconnected UX
- document Zero semantics in product/engineering docs

---

## 25. Exact code changes recommended first

If I were assigning the first engineering week, I would start here.

### 25.1 Week 1–2 foundational changes
1. create `packages/workbook-domain`
2. move semantic op types out of `packages/crdt`
3. add missing op placeholders for `renameSheet` / `reorderSheets`
4. add new SQL migrations for relational source/eval model
5. scaffold new `packages/zero-sync` schema and query registry
6. scaffold workbook service runtime and `recalc_job` worker
7. add `clientMutationID` to planned mutator interfaces

### 25.2 Week 2–3 read path changes
1. implement workbook/sheet/tile queries
2. build `ZeroWorkbookBridge`
3. preload styles/formats registry
4. build selected-cell source fetch
5. render active sheet from Zero tiles under feature flag
6. add shadow comparison logging

### 25.3 Week 3–4 write path changes
1. implement simple cell edit mutators
2. implement recalc materialization
3. wire worker optimistic preview to new mutators
4. add read-only disconnected mode
5. remove snapshot writeback from common cell edit path

### 25.4 Week 4+ structural parity
1. planner-backed fill/copy
2. row/col structural ops
3. sheet rename/reorder
4. table/pivot/name flows
5. remove old sync plane

---

## 26. Final recommendation

This is the plan I would actually implement.

It is the highest-probability path to a strong production result because it:

- keeps the worker-first browser shell that already exists
- keeps bilig as the semantic spreadsheet engine
- uses Zero exactly where Zero is strong
- deletes the custom sync responsibilities that are least differentiated
- aligns with the current lab deployment manifests instead of fighting them
- avoids pretending Zero is an offline-write CRDT system
- fixes the current prototype flaw of whole-workbook snapshot sync
- leaves room for later optimization without requiring a speculative rewrite now

The most important architectural choices in this plan are:

1. **Deprecate CRDT transport, not the semantic workbook mutation language**
2. **Split source state from evaluated state**
3. **Use tiled queries and a bridge into the existing viewport patch contract**
4. **Move authoritative recalc out of the request path**
5. **Describe the product truthfully as server-authoritative multiplayer with local-first UX**

If you implement those five decisions cleanly, the rest of the migration becomes engineering work rather than architecture risk.

---

## 27. Acceptance checklist

The migration is complete when all of the following are true:

### Product/runtime
- browser reads live workbook state only through Zero in production mode
- browser no longer writes `replaceSnapshot`
- grid still renders via viewport patches
- disconnected/error/needs-auth states are read-only

### Server/data
- no request-path mutation rewrites entire workbook projections
- `cell_eval` is updated incrementally through recalc jobs
- `workbook_event` and `workbook_snapshot` rebuild the same state
- planner-backed structural operations produce correct source deltas

### Collaboration
- multi-tab and multi-user edits converge through ordered revisions
- coarse presence is visible
- no CRDT replica state is required for the main product path

### Infra
- `bilig-zero` serves production traffic
- `bilig-sync` serves `/api/zero/*`
- Postgres publication and backfill are healthy
- rollout/rollback are documented

### Cleanup
- old binary browser sync path is removed or permanently disabled
- `@bilig/crdt` is removed or reduced to legacy-only compatibility
- obsolete snapshot projection tables are dropped in contract migration



---

## 28. Appendix A — concrete initial SQL migration skeleton

The following is the recommended **first additive migration**. It is intentionally explicit and close to what should actually be checked in.

```sql
create table if not exists workbook (
  id text primary key,
  name text not null,
  owner_user_id text not null,
  head_revision bigint not null default 0,
  calculated_revision bigint not null default 0,
  calc_mode text not null default 'automatic',
  compatibility_mode text not null default 'excel-modern',
  recalc_epoch bigint not null default 0,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists workbook_member (
  workbook_id text not null references workbook(id) on delete cascade,
  user_id text not null,
  role text not null,
  granted_by text,
  created_at timestamptz not null default now(),
  primary key (workbook_id, user_id)
);

create table if not exists sheet (
  id text primary key,
  workbook_id text not null references workbook(id) on delete cascade,
  name text not null,
  position int not null,
  freeze_rows int not null default 0,
  freeze_cols int not null default 0,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (workbook_id, name),
  unique (workbook_id, position)
);

create table if not exists sheet_filter (
  id text primary key,
  workbook_id text not null references workbook(id) on delete cascade,
  sheet_id text not null references sheet(id) on delete cascade,
  start_row int not null,
  end_row int not null,
  start_col int not null,
  end_col int not null,
  source_revision bigint not null
);

create table if not exists sheet_sort (
  id text primary key,
  workbook_id text not null references workbook(id) on delete cascade,
  sheet_id text not null references sheet(id) on delete cascade,
  start_row int not null,
  end_row int not null,
  start_col int not null,
  end_col int not null,
  keys_json jsonb not null,
  source_revision bigint not null
);

create table if not exists sheet_row (
  workbook_id text not null references workbook(id) on delete cascade,
  sheet_id text not null references sheet(id) on delete cascade,
  row_num int not null,
  row_id text,
  size int,
  hidden boolean,
  source_revision bigint not null,
  updated_at timestamptz not null default now(),
  primary key (workbook_id, sheet_id, row_num)
);

create table if not exists sheet_col (
  workbook_id text not null references workbook(id) on delete cascade,
  sheet_id text not null references sheet(id) on delete cascade,
  col_num int not null,
  col_id text,
  size int,
  hidden boolean,
  source_revision bigint not null,
  updated_at timestamptz not null default now(),
  primary key (workbook_id, sheet_id, col_num)
);

create table if not exists cell_number_format (
  workbook_id text not null references workbook(id) on delete cascade,
  format_id text not null,
  code text not null,
  kind text not null,
  created_at timestamptz not null default now(),
  primary key (workbook_id, format_id)
);

create table if not exists cell_style (
  workbook_id text not null references workbook(id) on delete cascade,
  style_id text not null,
  record_json jsonb not null,
  hash text not null,
  created_at timestamptz not null default now(),
  primary key (workbook_id, style_id)
);

create table if not exists cell_source (
  workbook_id text not null references workbook(id) on delete cascade,
  sheet_id text not null references sheet(id) on delete cascade,
  row_num int not null,
  col_num int not null,
  literal_input_json jsonb,
  formula_source text,
  explicit_format_id text,
  source_revision bigint not null,
  updated_by text not null,
  updated_at timestamptz not null default now(),
  primary key (workbook_id, sheet_id, row_num, col_num),
  foreign key (workbook_id, explicit_format_id)
    references cell_number_format(workbook_id, format_id)
    on delete set null
);

create table if not exists cell_eval (
  workbook_id text not null references workbook(id) on delete cascade,
  sheet_id text not null references sheet(id) on delete cascade,
  row_num int not null,
  col_num int not null,
  value_tag smallint not null,
  number_value double precision,
  boolean_value boolean,
  string_value text,
  error_code text,
  flags int not null default 0,
  version bigint not null,
  calc_revision bigint not null,
  updated_at timestamptz not null default now(),
  primary key (workbook_id, sheet_id, row_num, col_num)
);

create table if not exists sheet_style_range (
  id text primary key,
  workbook_id text not null references workbook(id) on delete cascade,
  sheet_id text not null references sheet(id) on delete cascade,
  start_row int not null,
  end_row int not null,
  start_col int not null,
  end_col int not null,
  style_id text not null,
  source_revision bigint not null,
  updated_at timestamptz not null default now(),
  foreign key (workbook_id, style_id)
    references cell_style(workbook_id, style_id)
    on delete cascade
);

create table if not exists sheet_format_range (
  id text primary key,
  workbook_id text not null references workbook(id) on delete cascade,
  sheet_id text not null references sheet(id) on delete cascade,
  start_row int not null,
  end_row int not null,
  start_col int not null,
  end_col int not null,
  format_id text not null,
  source_revision bigint not null,
  updated_at timestamptz not null default now(),
  foreign key (workbook_id, format_id)
    references cell_number_format(workbook_id, format_id)
    on delete cascade
);

create table if not exists defined_name (
  workbook_id text not null references workbook(id) on delete cascade,
  name text not null,
  normalized_name text not null,
  value_json jsonb not null,
  source_revision bigint not null,
  updated_at timestamptz not null default now(),
  primary key (workbook_id, normalized_name)
);

create table if not exists table_def (
  id text primary key,
  workbook_id text not null references workbook(id) on delete cascade,
  sheet_id text not null references sheet(id) on delete cascade,
  name text not null,
  start_row int not null,
  end_row int not null,
  start_col int not null,
  end_col int not null,
  column_names_json jsonb not null,
  header_row boolean not null,
  totals_row boolean not null,
  source_revision bigint not null,
  updated_at timestamptz not null default now(),
  unique (workbook_id, name)
);

create table if not exists pivot_def (
  id text primary key,
  workbook_id text not null references workbook(id) on delete cascade,
  sheet_id text not null references sheet(id) on delete cascade,
  name text not null,
  anchor_row int not null,
  anchor_col int not null,
  source_sheet_id text not null references sheet(id) on delete cascade,
  source_start_row int not null,
  source_end_row int not null,
  source_start_col int not null,
  source_end_col int not null,
  group_by_json jsonb not null,
  values_json jsonb not null,
  rows int not null,
  cols int not null,
  source_revision bigint not null,
  updated_at timestamptz not null default now(),
  unique (workbook_id, name)
);

create table if not exists spill_owner (
  workbook_id text not null references workbook(id) on delete cascade,
  sheet_id text not null references sheet(id) on delete cascade,
  owner_row int not null,
  owner_col int not null,
  rows int not null,
  cols int not null,
  calc_revision bigint not null,
  updated_at timestamptz not null default now(),
  primary key (workbook_id, sheet_id, owner_row, owner_col)
);

create table if not exists presence_coarse (
  workbook_id text not null references workbook(id) on delete cascade,
  user_id text not null,
  session_id text not null,
  sheet_id text not null references sheet(id) on delete cascade,
  active_row int,
  active_col int,
  selection_json jsonb,
  color text,
  updated_at timestamptz not null default now(),
  primary key (workbook_id, session_id)
);

create table if not exists workbook_event (
  workbook_id text not null references workbook(id) on delete cascade,
  revision bigint not null,
  actor_user_id text not null,
  client_mutation_id text,
  txn_json jsonb not null,
  created_at timestamptz not null default now(),
  primary key (workbook_id, revision)
);

create table if not exists applied_client_mutation (
  workbook_id text not null references workbook(id) on delete cascade,
  user_id text not null,
  client_mutation_id text not null,
  revision bigint not null,
  created_at timestamptz not null default now(),
  primary key (workbook_id, user_id, client_mutation_id)
);

create table if not exists recalc_job (
  id text primary key,
  workbook_id text not null references workbook(id) on delete cascade,
  from_revision bigint not null,
  to_revision bigint not null,
  dirty_regions_json jsonb,
  status text not null,
  attempts int not null default 0,
  lease_until timestamptz,
  last_error text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists workbook_snapshot (
  workbook_id text not null references workbook(id) on delete cascade,
  revision bigint not null,
  format text not null,
  payload jsonb not null,
  created_at timestamptz not null default now(),
  primary key (workbook_id, revision)
);

create index if not exists sheet_workbook_position_idx
  on sheet(workbook_id, position);

create index if not exists cell_source_sheet_row_col_idx
  on cell_source(workbook_id, sheet_id, row_num, col_num);

create index if not exists cell_eval_sheet_row_col_idx
  on cell_eval(workbook_id, sheet_id, row_num, col_num);

create index if not exists sheet_row_sheet_row_idx
  on sheet_row(workbook_id, sheet_id, row_num);

create index if not exists sheet_col_sheet_col_idx
  on sheet_col(workbook_id, sheet_id, col_num);

create index if not exists defined_name_workbook_norm_idx
  on defined_name(workbook_id, normalized_name);

create index if not exists table_def_workbook_sheet_idx
  on table_def(workbook_id, sheet_id);

create index if not exists pivot_def_workbook_sheet_idx
  on pivot_def(workbook_id, sheet_id);

create index if not exists sheet_style_range_workbook_sheet_rows_idx
  on sheet_style_range(workbook_id, sheet_id, start_row, end_row, start_col, end_col);

create index if not exists sheet_format_range_workbook_sheet_rows_idx
  on sheet_format_range(workbook_id, sheet_id, start_row, end_row, start_col, end_col);

create index if not exists recalc_job_status_lease_created_idx
  on recalc_job(status, lease_until, created_at);
```

---

## 29. Appendix B — concrete `packages/zero-sync/src/schema.ts` skeleton

```ts
import {
  createSchema,
  relationships,
  table,
  string,
  number,
  boolean,
  json,
} from "@rocicorp/zero";

const workbook = table("workbook").columns({
  id: string(),
  name: string(),
  ownerUserID: string().from("owner_user_id"),
  headRevision: number().from("head_revision"),
  calculatedRevision: number().from("calculated_revision"),
  calcMode: string().from("calc_mode"),
  compatibilityMode: string().from("compatibility_mode"),
  recalcEpoch: number().from("recalc_epoch"),
  createdAt: string().from("created_at"),
  updatedAt: string().from("updated_at"),
}).primaryKey("id");

const sheet = table("sheet").columns({
  id: string(),
  workbookID: string().from("workbook_id"),
  name: string(),
  position: number(),
  freezeRows: number().from("freeze_rows"),
  freezeCols: number().from("freeze_cols"),
  createdAt: string().from("created_at"),
  updatedAt: string().from("updated_at"),
}).primaryKey("id");

const cellSource = table("cell_source").columns({
  workbookID: string().from("workbook_id"),
  sheetID: string().from("sheet_id"),
  rowNum: number().from("row_num"),
  colNum: number().from("col_num"),
  literalInputJSON: json().from("literal_input_json").optional(),
  formulaSource: string().from("formula_source").optional(),
  explicitFormatID: string().from("explicit_format_id").optional(),
  sourceRevision: number().from("source_revision"),
  updatedBy: string().from("updated_by"),
  updatedAt: string().from("updated_at"),
}).primaryKey("workbookID", "sheetID", "rowNum", "colNum");

const cellEval = table("cell_eval").columns({
  workbookID: string().from("workbook_id"),
  sheetID: string().from("sheet_id"),
  rowNum: number().from("row_num"),
  colNum: number().from("col_num"),
  valueTag: number().from("value_tag"),
  numberValue: number().from("number_value").optional(),
  booleanValue: boolean().from("boolean_value").optional(),
  stringValue: string().from("string_value").optional(),
  errorCode: string().from("error_code").optional(),
  flags: number(),
  version: number(),
  calcRevision: number().from("calc_revision"),
  updatedAt: string().from("updated_at"),
}).primaryKey("workbookID", "sheetID", "rowNum", "colNum");

const sheetRow = table("sheet_row").columns({
  workbookID: string().from("workbook_id"),
  sheetID: string().from("sheet_id"),
  rowNum: number().from("row_num"),
  rowID: string().from("row_id").optional(),
  size: number().optional(),
  hidden: boolean().optional(),
  sourceRevision: number().from("source_revision"),
  updatedAt: string().from("updated_at"),
}).primaryKey("workbookID", "sheetID", "rowNum");

const sheetCol = table("sheet_col").columns({
  workbookID: string().from("workbook_id"),
  sheetID: string().from("sheet_id"),
  colNum: number().from("col_num"),
  colID: string().from("col_id").optional(),
  size: number().optional(),
  hidden: boolean().optional(),
  sourceRevision: number().from("source_revision"),
  updatedAt: string().from("updated_at"),
}).primaryKey("workbookID", "sheetID", "colNum");

const cellStyle = table("cell_style").columns({
  workbookID: string().from("workbook_id"),
  styleID: string().from("style_id"),
  recordJSON: json().from("record_json"),
  hash: string(),
  createdAt: string().from("created_at"),
}).primaryKey("workbookID", "styleID");

const cellNumberFormat = table("cell_number_format").columns({
  workbookID: string().from("workbook_id"),
  formatID: string().from("format_id"),
  code: string(),
  kind: string(),
  createdAt: string().from("created_at"),
}).primaryKey("workbookID", "formatID");

const sheetStyleRange = table("sheet_style_range").columns({
  id: string(),
  workbookID: string().from("workbook_id"),
  sheetID: string().from("sheet_id"),
  startRow: number().from("start_row"),
  endRow: number().from("end_row"),
  startCol: number().from("start_col"),
  endCol: number().from("end_col"),
  styleID: string().from("style_id"),
  sourceRevision: number().from("source_revision"),
  updatedAt: string().from("updated_at"),
}).primaryKey("id");

const sheetFormatRange = table("sheet_format_range").columns({
  id: string(),
  workbookID: string().from("workbook_id"),
  sheetID: string().from("sheet_id"),
  startRow: number().from("start_row"),
  endRow: number().from("end_row"),
  startCol: number().from("start_col"),
  endCol: number().from("end_col"),
  formatID: string().from("format_id"),
  sourceRevision: number().from("source_revision"),
  updatedAt: string().from("updated_at"),
}).primaryKey("id");

// ... definedName, tableDef, pivotDef, spillOwner, presenceCoarse, workbookMember

export const schema = createSchema({
  tables: [
    workbook,
    sheet,
    cellSource,
    cellEval,
    sheetRow,
    sheetCol,
    cellStyle,
    cellNumberFormat,
    sheetStyleRange,
    sheetFormatRange,
    // ...
  ],
  relationships: [
    relationships(workbook, ({many}) => ({
      sheets: many({
        sourceField: ["id"],
        destField: ["workbookID"],
        destSchema: sheet,
      }),
      styles: many({
        sourceField: ["id"],
        destField: ["workbookID"],
        destSchema: cellStyle,
      }),
      formats: many({
        sourceField: ["id"],
        destField: ["workbookID"],
        destSchema: cellNumberFormat,
      }),
    })),
    relationships(sheet, ({one, many}) => ({
      workbook: one({
        sourceField: ["workbookID"],
        destField: ["id"],
        destSchema: workbook,
      }),
      evalCells: many({
        sourceField: ["workbookID", "id"],
        destField: ["workbookID", "sheetID"],
        destSchema: cellEval,
      }),
      sourceCells: many({
        sourceField: ["workbookID", "id"],
        destField: ["workbookID", "sheetID"],
        destSchema: cellSource,
      }),
      rows: many({
        sourceField: ["workbookID", "id"],
        destField: ["workbookID", "sheetID"],
        destSchema: sheetRow,
      }),
      cols: many({
        sourceField: ["workbookID", "id"],
        destField: ["workbookID", "sheetID"],
        destSchema: sheetCol,
      }),
    })),
  ],
});

declare module "@rocicorp/zero" {
  interface DefaultTypes {
    schema: typeof schema;
  }
}
```

---

## 30. Appendix C — concrete `queries.ts` skeleton

```ts
import {defineQueriesWithType, defineQuery} from "@rocicorp/zero";
import {z} from "zod";
import {schema} from "./schema.js";
import {zql} from "./zql.js";

const defineQueries = defineQueriesWithType<typeof schema>();

const tileArgs = z.object({
  workbookID: z.string().min(1),
  sheetID: z.string().min(1),
  rowStart: z.number().int().nonnegative(),
  rowEnd: z.number().int().nonnegative(),
  colStart: z.number().int().nonnegative(),
  colEnd: z.number().int().nonnegative(),
});

export const queries = defineQueries({
  workbook: {
    get: defineQuery(
      z.object({workbookID: z.string().min(1)}),
      ({args, ctx}) =>
        zql.workbook
          .where("id", args.workbookID)
          // optionally filter with ctx.userID membership
          .one(),
    ),
  },

  sheet: {
    list: defineQuery(
      z.object({workbookID: z.string().min(1)}),
      ({args}) =>
        zql.sheet
          .where("workbookID", args.workbookID)
          .orderBy("position", "asc"),
    ),
  },

  cellEval: {
    tile: defineQuery(
      tileArgs,
      ({args}) =>
        zql.cellEval
          .where("workbookID", args.workbookID)
          .where("sheetID", args.sheetID)
          .where("rowNum", ">=", args.rowStart)
          .where("rowNum", "<=", args.rowEnd)
          .where("colNum", ">=", args.colStart)
          .where("colNum", "<=", args.colEnd)
          .orderBy("rowNum", "asc")
          .orderBy("colNum", "asc"),
    ),
  },

  cellSource: {
    get: defineQuery(
      z.object({
        workbookID: z.string().min(1),
        sheetID: z.string().min(1),
        rowNum: z.number().int().nonnegative(),
        colNum: z.number().int().nonnegative(),
      }),
      ({args}) =>
        zql.cellSource
          .where("workbookID", args.workbookID)
          .where("sheetID", args.sheetID)
          .where("rowNum", args.rowNum)
          .where("colNum", args.colNum)
          .one(),
    ),
  },

  sheetRow: {
    tile: defineQuery(
      z.object({
        workbookID: z.string().min(1),
        sheetID: z.string().min(1),
        rowStart: z.number().int().nonnegative(),
        rowEnd: z.number().int().nonnegative(),
      }),
      ({args}) =>
        zql.sheetRow
          .where("workbookID", args.workbookID)
          .where("sheetID", args.sheetID)
          .where("rowNum", ">=", args.rowStart)
          .where("rowNum", "<=", args.rowEnd)
          .orderBy("rowNum", "asc"),
    ),
  },

  sheetCol: {
    tile: defineQuery(
      z.object({
        workbookID: z.string().min(1),
        sheetID: z.string().min(1),
        colStart: z.number().int().nonnegative(),
        colEnd: z.number().int().nonnegative(),
      }),
      ({args}) =>
        zql.sheetCol
          .where("workbookID", args.workbookID)
          .where("sheetID", args.sheetID)
          .where("colNum", ">=", args.colStart)
          .where("colNum", "<=", args.colEnd)
          .orderBy("colNum", "asc"),
    ),
  },

  styleRange: {
    intersectTile: defineQuery(
      tileArgs,
      ({args}) =>
        zql.sheetStyleRange
          .where("workbookID", args.workbookID)
          .where("sheetID", args.sheetID)
          .where("endRow", ">=", args.rowStart)
          .where("startRow", "<=", args.rowEnd)
          .where("endCol", ">=", args.colStart)
          .where("startCol", "<=", args.colEnd),
    ),
  },

  formatRange: {
    intersectTile: defineQuery(
      tileArgs,
      ({args}) =>
        zql.sheetFormatRange
          .where("workbookID", args.workbookID)
          .where("sheetID", args.sheetID)
          .where("endRow", ">=", args.rowStart)
          .where("startRow", "<=", args.rowEnd)
          .where("endCol", ">=", args.colStart)
          .where("startCol", "<=", args.colEnd),
    ),
  },

  style: {
    byWorkbook: defineQuery(
      z.object({workbookID: z.string().min(1)}),
      ({args}) => zql.cellStyle.where("workbookID", args.workbookID),
    ),
  },

  numberFormat: {
    byWorkbook: defineQuery(
      z.object({workbookID: z.string().min(1)}),
      ({args}) => zql.cellNumberFormat.where("workbookID", args.workbookID),
    ),
  },

  presence: {
    byWorkbook: defineQuery(
      z.object({workbookID: z.string().min(1)}),
      ({args}) => zql.presenceCoarse.where("workbookID", args.workbookID),
    ),
  },
});
```

---

## 31. Appendix D — concrete `mutators.ts` skeleton

```ts
import {defineMutator, defineMutatorsWithType} from "@rocicorp/zero";
import {z} from "zod";
import {schema} from "./schema.js";

const defineMutators = defineMutatorsWithType<typeof schema>();

const baseMutation = z.object({
  workbookID: z.string().min(1),
  clientMutationID: z.string().min(1),
});

const editOneArgs = baseMutation.extend({
  sheetID: z.string().min(1),
  rowNum: z.number().int().nonnegative(),
  colNum: z.number().int().nonnegative(),
  literalInput: z.union([z.string(), z.number(), z.boolean(), z.null()]).optional(),
  formulaSource: z.string().optional(),
});

const clearRangeArgs = baseMutation.extend({
  sheetID: z.string().min(1),
  startRow: z.number().int().nonnegative(),
  endRow: z.number().int().nonnegative(),
  startCol: z.number().int().nonnegative(),
  endCol: z.number().int().nonnegative(),
});

const resizeColumnArgs = baseMutation.extend({
  sheetID: z.string().min(1),
  colNum: z.number().int().nonnegative(),
  size: z.number().int().positive(),
});

async function serverOnly(): Promise<void> {
  // no-op in client mutator typing side
}

export const mutators = defineMutators({
  cell: {
    editOne: defineMutator(editOneArgs, serverOnly),
    clearRange: defineMutator(clearRangeArgs, serverOnly),
  },
  col: {
    resize: defineMutator(resizeColumnArgs, serverOnly),
  },
});
```

### Concrete server dispatch contract

Inside the server mutate handler, dispatch to functions that follow this pattern:

```ts
await tx.dbTransaction.query(
  `select pg_advisory_xact_lock(hashtext($1))`,
  [workbookID],
);

// 1. authorize
// 2. idempotency check
// 3. planner/direct-SQL mutation
// 4. bump workbook.head_revision
// 5. append workbook_event
// 6. insert recalc_job
// 7. persist applied_client_mutation
```

---

## 32. Appendix E — first 10 implementation PRs

If this plan is executed as a real program, these are the first 10 PRs I would open.

1. **Create `packages/workbook-domain` and move semantic op types out of `packages/crdt`**
2. **Add additive Postgres migration for relational workbook model**
3. **Scaffold workbook service runtime and `recalc_job` worker**
4. **Rewrite `packages/zero-sync` schema and query registry**
5. **Add real auth/context to `/api/zero/query` and `/api/zero/mutate`**
6. **Implement `cell.editOne`, `cell.clearRange`, `col.resize`, `presence.update` mutators**
7. **Implement `cell_eval` materializer and snapshot warm-start**
8. **Add `ZeroWorkbookBridge` and render active sheet from tiled queries behind flag**
9. **Cut browser snapshot `replaceSnapshot` loop from the normal edit path**
10. **Add shadow-equivalence tests and remove old websocket browser sync from prod path**
