# bilig

`bilig` is a production spreadsheet engine monorepo with a worker-first web shell, a package-based custom workbook reconciler, a framework-agnostic core engine, replication-ready mutation pipelines, and an AssemblyScript/WASM numeric fast path.

It already has the foundations of a serious spreadsheet/runtime stack: a real engine, a real local session loop, a real binary sync layer, a real reconciler, and a reasonably mature grid shell. The biggest remaining gap is not basic spreadsheet arithmetic; it is the seam between what the local engine can represent and what the authoritative replicated model can express.

> Current topology note:
> `apps/bilig` is the only shipped product runtime.
> `apps/web` remains the browser source tree, but the built assets now ship from the monolith image.
> Any remaining references below to `apps/local-server` or `apps/sync-server` are historical architecture context, not the current supported deployment shape.

## bilig — what this project does

### Overview

**bilig** is a local-first spreadsheet system built as a monorepo. Its purpose is to provide an Excel-like worksheet engine that can run in the browser, execute formulas deterministically, synchronize edits in realtime, expose spreadsheet state to agents, and progressively move hot calculation paths into WebAssembly for higher performance.

This project is not just a spreadsheet UI. It combines a workbook engine, formula language stack, synchronization model, binary transport, browser storage, a fullstack monolith runtime, React-based workbook authoring layer, and a reusable grid UI.

At a high level, bilig is trying to become a full spreadsheet runtime platform with these properties:

- **local-first**: the workbook can restore and continue from local state
- **realtime**: edits can be mirrored and replayed through ordered mutation streams
- **browser-native**: the main product surface is a web spreadsheet shell
- **WASM-accelerated**: proven formula families can run on a WebAssembly fast path
- **AI-native**: agents can read, mutate, and subscribe to worksheet state through a shared API
- **package-oriented**: core responsibilities are split into reusable packages instead of buried in one app

### What problem it solves

Most spreadsheet products mix several concerns together: UI, workbook model, formula semantics, persistence, collaboration, and automation. bilig separates those concerns into packages so the same workbook model can be used in different runtimes:

- inside a browser UI
- inside a local machine server
- inside a future durable remote sync service
- from CLI or agent tooling
- from a React-authored workbook description

That makes bilig closer to a **spreadsheet engine platform** than a single spreadsheet application.

### What the system does today

From the current repository structure and docs, bilig already provides a meaningful spreadsheet foundation.

#### 1. It runs a real workbook engine

The core engine supports workbook and sheet mutation, cell values, formulas, formatting, range operations, history, selection, dependency inspection, snapshots, and replica synchronization hooks.

The documented engine surface includes operations such as:

- create and delete sheets
- set cell values, formulas, and formats
- clear cells and ranges
- set values and formulas across ranges
- fill, copy, and paste ranges
- manage selection state
- undo and redo
- inspect dependencies and dependents
- explain a cell's computed state
- import and export workbook snapshots
- import and export replica snapshots
- apply remote mutation batches
- subscribe to engine changes and outbound batches

This means bilig already behaves like a real worksheet runtime, not just a table editor.

#### 2. It parses and evaluates spreadsheet formulas

The formula subsystem owns:

- A1 addressing
- lexing and parsing
- binding and compilation
- JS evaluation as the semantic oracle
- compatibility tracking for formula families
- oracle fixture integration for parity testing

The project's current formula milestone is a **100-case canonical Excel worksheet formula corpus**. The repo is explicitly structured so formula support lands in JavaScript first, is verified against checked-in fixtures, and only then gets promoted into WASM production routing.

#### 3. It has a WebAssembly acceleration layer

The repo includes an AssemblyScript-based WASM kernel. The current design is conservative and deliberate:

- JavaScript remains the semantic source of truth while parity is being closed
- WASM handles closed and proven formula families
- production routing flips to WASM only after differential parity is green

This is the right model for a spreadsheet engine, because performance improvements do not come at the cost of silent semantic drift.

#### 4. It supports local-first persistence and replica replay

The browser side already has a persistence layer and snapshot flow. The current design is meant to restore:

1. workbook snapshot
2. replica snapshot and queued outbound edits
3. engine runtime
4. sync connectivity

The browser shell already demonstrates local persistence and replica mirroring behavior, which means bilig already operates as a local-first product surface, even though the final worker-first production path is still in progress.

#### 5. It has a collaboration and sync model

The repository includes:

- a CRDT-ready mutation and replica package
- a binary sync protocol package
- a local server that hosts live workbook sessions
- a remote sync server scaffold for durability and cross-device fanout

The local-first collaboration loop envisioned by the repo is:

1. browser restores local workbook state
2. browser initializes runtime and WASM
3. browser connects to a local app server when available
4. edits commit through one ordered workbook stream
5. the browser renders committed mutations immediately
6. the same committed mutations relay to a remote sync backend for durability and multiplayer convergence

So bilig is already architected for multiplayer collaboration, even though the remote durable service is not fully closed yet.

#### 6. It exposes spreadsheet state to agents

One of the strongest characteristics of this codebase is that it treats spreadsheet automation as a first-class runtime concern.

The repo includes an `agent-api` package and a local server that can execute canonical worksheet requests against live workbook sessions. The intended scope includes:

- open and close worksheet sessions
- read cells and ranges
- write values and formulas
- clear, fill, copy, and paste ranges
- inspect precedents and dependents
- subscribe to range changes
- import and export snapshots
- query metrics and traces
- execute batched mutations with idempotency keys

This makes bilig more than a spreadsheet app. It is being designed as a **spreadsheet engine that both people and agents can drive**.

#### 7. It has a React-based workbook authoring model

The repo contains a custom reconciler package and a workbook DSL with shapes such as:

- `<Workbook>`
- `<Sheet name="...">`
- `<Cell addr="A1" value={...} />`
- `<Cell addr="A1" formula="..." />`
- `<Cell addr="A1" format="..." />`

That means React is used as a declarative authoring and commit layer for workbook structure, while the core engine remains framework-agnostic.

This is useful for:

- deterministic workbook construction
- programmatic workbook generation
- test scenarios
- UI-to-engine consistency
- internal tooling

#### 8. It ships with a reusable spreadsheet UI

The monorepo also includes a reusable grid package and browser shell. The current product surface includes:

- a virtualized spreadsheet grid
- sheet tabs
- keyboard navigation
- dependency inspection
- recalc metrics and inspection panels
- a worker-first web shell for exercising and shipping the engine

So bilig already contains both the engine and the interactive product shell that sits on top of it.

### How the architecture fits together

A useful way to understand bilig is to think of it as five connected layers.

#### Layer 1: workbook semantics

`@bilig/core` is the engine. It owns workbook state, mutation application, recalculation orchestration, snapshots, selection, and introspection.

#### Layer 2: formula language and execution

`@bilig/formula` handles parsing, binding, compilation, and JS evaluation. `@bilig/wasm-kernel` provides the WASM execution path for formula families that are already proven safe to promote.

#### Layer 3: replication and transport

`@bilig/core` owns replica bookkeeping, local replay semantics, and ordered workbook mutation streams. `@bilig/zero-sync` defines the shared Zero schema, projection helpers, and workbook event payloads. `@bilig/binary-protocol` defines how sync frames move over the wire, and `@bilig/worker-transport` gives the repo a path to move the engine off the main browser thread.

#### Layer 4: UI and declarative authoring

`@bilig/renderer` is the custom workbook reconciler and DSL. `@bilig/grid` is the reusable spreadsheet UI layer.

#### Layer 5: runtimes and storage

- `apps/web` is the browser source tree compiled into the monolith
- `apps/bilig` is the only supported product runtime and serves the built web shell
- `@bilig/storage-browser` and `@bilig/storage-server` handle browser and server persistence concerns

This split is important because it allows the project to evolve into a spreadsheet platform rather than a single-page app with tightly coupled logic.

### Package map

#### Applications

| Path | Role |
| --- | --- |
| `apps/web` | Vite/React browser source compiled into the monolith |
| `apps/bilig` | Fullstack monolith runtime, API surface, and static asset server |

#### Packages

| Package | Role |
| --- | --- |
| `@bilig/protocol` | Shared enums, constants, and protocol types |
| `@bilig/formula` | Formula grammar, binding, compilation, JS oracle evaluation |
| `@bilig/core` | Workbook engine, recalc, snapshots, selection, sync hooks |
| `@bilig/zero-sync` | Zero schema, workbook projection, shared event payloads |
| `@bilig/binary-protocol` | Wire format for sync frames |
| `@bilig/agent-api` | Agent request/response/event model and framing |
| `@bilig/worker-transport` | Engine host/client bridge for worker execution |
| `@bilig/renderer` | Custom workbook reconciler and workbook DSL |
| `@bilig/grid` | Reusable spreadsheet UI components and hooks |
| `@bilig/wasm-kernel` | AssemblyScript/WASM compute fast path |
| `@bilig/storage-browser` | Browser-side persistence |
| `@bilig/storage-server` | Server-side storage integration points |
| `@bilig/excel-fixtures` | Checked-in formula parity fixtures |
| `@bilig/benchmarks` | Benchmark harness and performance contracts |

### What makes bilig different

The project has several distinguishing design choices.

#### Local-first is not an add-on

Local persistence, snapshots, relay queues, and local-session hosting are treated as core architecture, not a later enhancement.

#### Agents are part of the system design

The repo is explicitly preparing for spreadsheet state to be read and mutated by local and remote agents through a shared API, instead of forcing automation to screen-scrape the UI.

#### Performance work is tied to semantic proof

The WASM fast path is governed by formula parity contracts, fixture checks, and differential testing, which is a much stronger approach than rewriting computation paths and hoping they match Excel semantics.

#### React is used as an authoring layer, not as the engine itself

The core engine remains reusable and framework-agnostic. React is used to describe workbook structure and reconcile that description into engine mutations.

#### Collaboration is modeled as ordered mutation streams

The repo already defines outbound batch APIs, binary frame families, replica snapshots, and cursor/watermark concepts. That gives the project a strong foundation for durable collaboration.

### Current limits and open work

The repo is ambitious, but it is not finished. The main open areas are clear from the current docs.

#### Excel parity is incomplete

The canonical corpus exists, but full parity is not yet closed. The open areas explicitly called out in the repo include:

- defined names
- tables
- structured references
- dynamic arrays and spill behavior
- `LET` / `LAMBDA`
- some metadata-aware and volatile semantics

#### Worker-first browser runtime is not the default yet

The production browser shell still runs the engine in-process today. The worker-backed runtime exists as a direction and partial implementation, but it is not yet the default boot path.

#### The remote sync backend is not final

The remote sync server exists, but the docs still describe it as incomplete relative to the final durable collaboration service.

#### Agent transport is in an interim state

Agent frames currently use JSON payload bodies inside a binary envelope. The target design is typed binary frames end to end.

### The project's intended end state

If bilig reaches its architectural target, it becomes a system where:

- a browser spreadsheet shell can restore instantly from local state
- formulas execute with Excel-like semantics
- closed formula families run in WASM for speed
- local user edits and local agent edits flow through the same ordered stream
- committed updates can be relayed to a durable remote backend
- multiple clients can converge on the same workbook state
- agents can understand and modify the spreadsheet using stable programmatic APIs
- the engine can be reused outside the browser shell

In other words, bilig is aiming to be a **high-performance, local-first, collaborative, agent-addressable spreadsheet engine and runtime platform**, not just a spreadsheet UI.

### Plain-English summary

If this project had to be explained in one paragraph:

> bilig is a browser-native spreadsheet engine platform. It combines a real workbook engine, a formula parser/compiler, a WebAssembly acceleration layer, a collaboration pipeline, a local-first persistence model, a React-based workbook reconciler, reusable grid UI, and agent APIs. The system is designed so spreadsheets can be edited by people or agents, run locally or with sync, restore quickly, and progressively reach Excel-level formula compatibility while keeping high-performance execution paths in WASM.

## Workspace layout

- `apps/web`: Vite 8 React app shell that composes the packages
- `packages/protocol`: shared enums, opcodes, constants, and types
- `packages/formula`: A1 addressing, lexer, parser, binder, compiler, JS evaluator
- `packages/core`: spreadsheet engine, storage, scheduler, snapshots, selectors, sync ownership, WASM facade
- `packages/zero-sync`: shared Zero schema, workbook queries, projection, and event payload helpers
- `packages/renderer`: custom workbook reconciler and workbook DSL
- `packages/grid`: reusable React spreadsheet UI, hooks, selection, metrics, and inspectors
- `packages/wasm-kernel`: AssemblyScript VM and numeric kernels
- `packages/benchmarks`: benchmark harness
- `docs`: architecture, API, reconciler layering, sync ownership, formula language

## Quickstart

```bash
pnpm install
pnpm wasm:build
pnpm typecheck
pnpm test
pnpm dev
```

## Local Docker Compose

```bash
docker compose up --build
```

This brings up the full local stack:

- `http://localhost:3000` for the monolith web shell with `/v2`, `/api/zero/v2`, and `/zero`
- `http://localhost:4321/healthz` for the monolith app runtime
- `http://localhost:4848/keepalive` for Zero cache
- `postgresql://bilig:bilig@localhost:5432/bilig` for Postgres

The compose stack uses `docker/runtime-config.local.json` for browser runtime wiring and creates the `zero_data` publication during Postgres init so Zero can replicate from the local database.

To reset local state:

```bash
docker compose down -v
```

## Commands

```bash
pnpm dev
pnpm protocol:generate
pnpm build
pnpm typecheck
pnpm test
pnpm bench
pnpm bench:contracts
pnpm release:check
pnpm run ci
```

## Notes

- Reusable React code now lives in `packages/renderer` and `packages/grid`; `apps/web` is the thin shell.
- The spreadsheet engine remains usable without React.
- The custom reconciler lives under `packages/renderer`.
- The public cell model supports `format` as a persisted attribute alongside `addr`, `value`, and `formula`.
- The WASM kernel is a custom AssemblyScript fast path, not an embedded proprietary spreadsheet runtime.
- The TS protocol enums/opcodes and AssemblyScript protocol mirror are generated from `scripts/gen-protocol.ts` so JS/WASM ABI drift fails fast in CI.
- The web shell includes a scroll-windowed sheet surface, sheet tabs, keyboard cell navigation, dependency inspection, and recalc metrics.
- The web shell operator surface now spans a 100k-row by 256-column virtualized window while keeping the engine hard limits at 1,048,576 rows by 16,384 columns.
- The demo workbook now exercises JS row/column range formulas and a WASM-backed branch formula in the visible UI so browser smoke covers both paths.
- The cell inspector now exposes formula mode, topo rank, versioning, and dependency edges from the core engine.
- The web shell also demonstrates local-first replica mirroring through the engine’s outbound and inbound batch APIs.
- The web shell persists workbook and replica snapshots in local storage so the demo survives reloads as a local-first app surface.
- The web shell relay queue now persists paused replica traffic across reloads, so offline-style catch-up is visible instead of being memory-only.
- The paused relay queue is compacted with the CRDT entity-order rules, so repeated offline edits do not grow an unbounded replay backlog for the same cell or sheet entity.
- The imperative engine now includes a single-sheet CSV bridge for import/export without pulling React into shared packages.
- CI now enforces performance contracts for 100k snapshot load, 10k-downstream edits, and 10k-cell render commits instead of relying on a loose smoke check alone.
- `pnpm release:check` enforces the documented production budgets for the built app JS and bundled WASM asset.
- The next highest-leverage architecture work is to make the authoritative workbook op model exhaustive enough to match the local engine surface, then build worker-first runtime, durable multiplayer, and typed binary agent work on top of that seam.

## CI

- Forgejo Actions is the primary CI surface for this repo via `.forgejo/workflows/forgejo-ci.yml`.
- GitHub Actions mirrors the verification contract in `.github/workflows/ci.yml`.
- Image publishing does not run on GitHub Actions for this repo.
- Forgejo image publishing runs via `.forgejo/workflows/release-images.yml` on the private `docker` and `docker-arm64` runner labels and publishes multi-arch `linux/amd64,linux/arm64` tags to `registry.ide-newton.ts.net`.
- The workflow is strict: frozen lockfile install, full `pnpm run ci`, artifact budget checks, browser smoke, and a tracked-file cleanliness check.
- Forgejo verification jobs run on the private `docker` runner label and pin their own Node `24.x` container, so they do not depend on host-level language tooling.
- GitHub Actions runs the same repository contract on Node 22 and Node 24.11.1 so compatibility drift is visible before release.
