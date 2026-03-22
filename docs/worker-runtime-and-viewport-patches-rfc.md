# Worker Runtime and Viewport Patches RFC

## Problem

The browser runtime still runs the engine in-process by default, and the grid does not yet consume a dedicated derived viewport patch stream. That makes the runtime/UI seam weaker than the local engine seam and keeps the main thread closer to workbook semantics than it should be.

## Goals

- make worker-first browser execution the default runtime shape
- keep the UI thread focused on paint, input, clipboard, overlays, and formula-bar UX
- introduce a derived viewport patch stream for visible workbook state
- replace address-by-address subscription patterns with region-based viewport subscriptions

## Non-goals

- move workbook semantics into the UI layer
- make React or `@bilig/grid` the canonical runtime state
- depend on worker-first execution before the authoritative op model exists

## Runtime responsibilities

### UI thread

- shell composition
- user input
- clipboard
- overlays
- formula bar
- viewport subscription requests
- rendering derived patches

### Worker

- workbook engine
- formula execution
- authoritative workbook transactions
- persistence restore and replay
- sync connectivity
- WASM lifecycle
- viewport patch derivation

## Streams

The runtime should separate:

1. authoritative workbook op stream
2. derived viewport patch stream
3. ephemeral presence and awareness stream

The grid should consume the second stream, not raw engine state or raw transaction replay.

## Proposed viewport model

### Region subscriptions

Replace cell-by-cell or narrow ad hoc subscriptions with region-based subscriptions:

- visible sheet region
- overscan region
- frozen panes region if applicable
- auxiliary subscriptions for inspectors or formula bar when needed

### Derived patch payloads

A patch stream should be optimized for rendering, not for workbook semantics. Candidate fields:

- changed visible cells
- display strings
- error tags
- style or format IDs
- row heights
- column widths
- visible selection overlays if needed
- patch cursor or version

The worker can derive these from authoritative workbook state without making the grid responsible for replay logic.

## Proposed package direction

- `@bilig/worker-transport`
  - request/response and subscription transport between UI and worker
- a future render-patch or viewport-runtime package
  - viewport subscriptions
  - patch derivation
  - patch codecs
- `@bilig/grid`
  - patch consumer, not workbook-state owner
- `apps/web`
  - worker-first shell boot

## Boot sequence target

1. restore workbook snapshot and local queue
2. boot worker transport
3. initialize WASM in worker
4. connect local daemon or fallback local-only runtime
5. replay authoritative transaction tail
6. subscribe visible regions
7. stream viewport patches to UI
8. surface sync state and reconnect status in shell

## Preconditions

This RFC depends on earlier architecture work:

- the authoritative workbook op model must exist
- metadata needed for names, tables, spills, and structure must already be first-class

Without those, worker-first execution just moves incomplete semantics to another thread.

## Migration order

1. formalize worker boot contract
2. introduce viewport subscription model and derived patch types
3. add region-based subscriptions for the visible grid
4. move `apps/web` to worker-first default boot
5. keep in-process execution only as an explicit fallback during rollout
6. remove address-by-address hot paths from the browser rendering loop

## Suggested PR breakdown

1. worker boot and runtime contract cleanup
2. viewport subscription and patch type introduction
3. worker-derived patch implementation for visible cells and metrics
4. `@bilig/grid` adaptation to patch consumption
5. `apps/web` worker-first default boot rollout
6. reconnect/restore/browser test hardening

## Exit gate

- `apps/web` boots through a worker-backed engine by default
- the grid renders from worker-derived viewport patches
- visible updates no longer depend on in-process engine coupling
- reconnect, restore, and local-daemon catch-up work through the worker runtime
- region-based subscriptions replace address-by-address hot paths for the primary grid surface
