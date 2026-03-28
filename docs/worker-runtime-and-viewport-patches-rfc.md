# Worker Runtime and Viewport Patches RFC

## Current state

The current browser runtime implements these RFC items:

- `apps/web` boots a worker-backed engine by default
- the UI consumes viewport patches rather than raw engine state
- the worker owns engine boot, persistence restore, sync connectivity, WASM lifecycle, and patch derivation

The remaining runtime items are:

- retire the deprecated `apps/playground` package from active product flows
- stronger region-oriented patch subscriptions and rollout hardening
- typed patch codecs instead of JSON payloads inside a byte envelope
- full reconnect and remote-catch-up behavior once sync-server closes more of the multiplayer path

## Goals

- keep worker-first browser execution as the default runtime shape
- keep the UI thread focused on paint, input, clipboard, overlays, and formula-bar UX
- keep a derived viewport patch stream for visible workbook state
- continue replacing direct engine coupling and narrow ad hoc subscription patterns with region-based viewport subscriptions

## Non-goals

- move workbook semantics into the UI layer
- make React or `@bilig/grid` the canonical runtime state
- reintroduce in-process engine ownership into the product shell

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

The runtime should continue to separate:

1. authoritative workbook op stream
2. derived viewport patch stream
3. ephemeral presence and awareness stream

The grid should consume the second stream, not raw engine state or raw transaction replay.

## Viewport model

### Region subscriptions

The target subscription model is:

- visible sheet region
- overscan region
- frozen panes region if applicable
- auxiliary subscriptions for inspectors or formula bar when needed

### Derived patch payloads

A patch stream should be optimized for rendering, not for workbook semantics. Current payloads carry:

- changed visible cells
- display strings
- copy and editor strings
- format and style ids
- row heights
- column widths
- patch version

The remaining improvement is codec quality and region-subscription hardening.

## Proposed package direction

- `@bilig/worker-transport`
  - request, response, and subscription transport between UI and worker
- a future render-patch or viewport-runtime package
  - viewport subscriptions
  - patch derivation
  - patch codecs
- `@bilig/grid`
  - patch consumer, not workbook-state owner
- `apps/web`
  - worker-first shell boot

## Remaining migration order

1. finish worker boot and reconnect contract cleanup
2. harden region-based subscriptions for the visible grid
3. replace JSON patch payloads with typed codecs
4. retire the deprecated playground shell and keep `apps/web` as the product surface
5. harden reconnect, restore, and browser test coverage

## Exit gate

- `apps/web` boots through a worker-backed engine by default
- the grid renders from worker-derived viewport patches
- visible updates no longer depend on in-process engine coupling
- reconnect, restore, and local-daemon catch-up work through the worker runtime
- region-based subscriptions back the primary grid surface
- viewport patches use typed codecs rather than JSON payload wrappers
