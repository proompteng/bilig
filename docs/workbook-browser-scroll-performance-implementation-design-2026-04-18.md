# Workbook Browser Scroll Performance Implementation Design

Date: `2026-04-18`

Status: `design capture, not implemented`

Grounding:

- current mounted workbook surface in `packages/grid`
- current worker-first viewport/runtime path in `apps/web`
- current product performance goals in `docs/05-06-next-phase.md`

Related documents:

- `docs/design.md`
- `docs/glide-removal-renderer-v2.md`
- `docs/05-06-next-phase.md`
- `docs/performance-budgets.md`
- `docs/next-iteration-production-plan-2026-04-10.md`

## Purpose

This document defines the production implementation design for eliminating browser jank
when horizontally browsing large or wide workbook surfaces.

It is not a generic optimization checklist and it is not a request for local patches.

It is a design for the mounted browser workbook runtime that must produce:

- steady high-FPS horizontal and vertical browsing
- no frame-dropping or hitchy scroll behavior on large sheets
- no DOM-node-per-cell rendering ceiling
- no correctness regressions in text, selection, freeze panes, or collaboration

The target is a workbook surface that feels smooth under sustained use, not a sheet that
passes isolated microbenchmarks while still feeling bad in the browser.

## Executive Verdict

The current architecture cannot reliably deliver a no-jank wide-sheet browsing experience.

The main reasons are structural:

- viewport identity changes on scroll currently drive subscription churn
- worker viewport subscriptions emit full initialization patches on attach
- React-visible workbook state is still fanned out too broadly from one viewport store
- the grid rebuilds full visible scenes on viewport movement instead of treating scroll as
  mostly local transform work
- visible cell text is still rendered as one absolutely positioned DOM node per visible text
  item

That combination means large horizontal scrolls pay work in the wrong places:

- worker patch setup work
- patch decode and cache mutation work
- React subscription wakeups
- full visible scene recomputation
- DOM layout and paint cost for hundreds or thousands of text nodes

The correct fix is not a tuning pass.

The correct fix is to separate:

- scroll position from data residency
- viewport residency from patch session lifecycle
- scene caching from React render
- text rendering from DOM node creation

## Product-Level Performance Contract

This design is only considered successful if it meets the following browser-path contract.

### Scroll smoothness

For a warmed active workbook viewport:

- sustained horizontal browse frame time p95 must stay below `16.7ms`
- sustained horizontal browse frame time p99 must stay below `25ms`
- sustained horizontal browse must produce zero long tasks over `50ms`
- pure scroll inside an already resident tile window must not trigger worker resubscription
- pure scroll inside an already resident tile window must not trigger React shell rerenders
  outside the grid renderer boundary

### Interaction quality

- freeze panes must remain visually locked with no lag against the scrolling body pane
- selection outline, hovered cell chrome, fill handle, and editor overlay must remain
  attached to the correct cells while scrolling
- same-sheet collaborator updates must not visibly hitch the active viewport
- scrolling must never enter a reduced-fidelity mode that hides text, borders, fills, or
  formatting

### Data residency and patch behavior

- local pan within resident tiles: `0` full viewport patches
- local pan within resident tiles: `0` worker subscription attach or detach operations
- tile-boundary crossing: one stable session region update, not per-frame subscribe and
  unsubscribe churn
- visible patch application must remain damage-based, not full-viewport replacement, once a
  session is active

### Browser workloads that must pass

The harness must measure at least these workloads:

1. `wide-250k-main-body`
   - wide worksheet, no freeze panes, sustained horizontal browse across multiple screens
2. `wide-250k-frozen-columns`
   - same workbook, frozen leading columns, sustained horizontal browse
3. `wide-250k-frozen-rows-and-columns`
   - same workbook, frozen top rows and leading columns
4. `wide-250k-collaboration-visible`
   - same workbook, active horizontal browse while same-sheet collaborator patches land inside
     and outside the visible area
5. `wide-250k-variable-width`
   - same workbook shape with non-uniform column widths so scroll math cannot cheat on fixed
     widths

The target browser for the primary gate is Chromium on a modern laptop because that matches the
current product shell and Playwright path, but the design must not depend on Chromium-only
behavior to be correct.

### Exact browser harness mechanics

The browser-path contract above must be enforced by the existing Playwright stack, not by an
invented benchmark runner.

Primary harness files:

- `playwright.config.ts`
- `scripts/run-browser-tests.ts`
- `e2e/tests/web-shell-scroll-performance.pw.ts`
- `e2e/tests/web-shell-helpers.ts`
- `apps/web/src/perf/workbook-perf.ts`
- `apps/web/src/perf/workbook-scroll-perf.ts`

Required harness behavior:

- the perf suite runs inside the same Chromium Playwright project already used by `pnpm
  test:browser`
- the test opens a deterministic wide-sheet workbook fixture through the normal browser stack
- the fixture comes from the checked-in benchmark corpus family in `packages/benchmarks`, not from
  an ad hoc browser-only builder
- the test drives horizontal browse using the real scroll viewport element, not synthetic direct
  state injection
- the page installs a `PerformanceObserver` for `longtask`
- the page installs a `requestAnimationFrame` sampler that records frame deltas during the browse
  interval
- the page records workbook-specific counters through `performance.mark` plus a dedicated browser
  perf collector
- the test exports one JSON artifact per workload through `testInfo.outputPath(...)`

Required metrics per workload artifact:

- workload name
- browser name and viewport
- workbook fixture id
- frame-time sample list
- frame-time min, median, p95, p99, and max
- long-task count and worst long-task duration
- viewport session attach count
- viewport session region-update count
- full patch count
- damage patch count
- React commit count for the workbook shell boundary
- visible tile residency hit rate
- mounted grid text DOM node count

Required run shape:

- each workload runs `5` times in CI
- the gate uses p95 for frame time and the max for long-task duration
- the artifact preserves all raw samples so regressions can be diagnosed after failure

Required CI policy:

- phase `0`: perf suite runs in CI and publishes artifacts, but does not yet block merge
- phase `1` completion gate: session churn thresholds become merge-blocking
- phase `4` completion gate: full scroll-frame thresholds become merge-blocking
- final state: the perf suite is a required `pnpm run ci` stage for workbook browser changes

This is the only acceptable rollout order because the suite must be trustworthy before it becomes
hard-gating.

### Deterministic fixture and sampling rules

The perf suite must not benchmark arbitrary user workbooks.

It must use checked-in deterministic fixtures with:

- stable row and column counts
- stable width metadata
- stable freeze-pane metadata where relevant
- stable visible seed viewport
- stable text distributions that exercise spill, wrap, and numeric rendering

The primary wide-sheet fixture family should live beside the existing benchmark corpus definitions
and expose browser-oriented cases for:

- wide dense body
- wide variable-width body
- wide frozen leading columns
- wide frozen rows plus columns

Sampling rules:

- run each scenario at a fixed browser viewport size
- use a fixed device scale factor for CI
- reset to the same starting viewport before each iteration
- sample only after the workbook reaches warm ready state
- record for a fixed browse duration with constant commanded velocity
- exclude initial navigation and workbook bootstrap from the steady-state browse window

## Non-Negotiable Constraints

The implementation may not rely on any of the following:

- scroll debouncing or throttling as the primary fix
- temporary “fast mode while scrolling”
- hiding text until scroll settles
- dropping borders, fills, wrap, spill, or alignment fidelity during movement
- broad `useMemo` and `useDeferredValue` tuning as a substitute for architectural separation
- benchmark or acceptance-budget softening
- a browser-only second truth model that diverges from the worker/runtime contract

This is a quality rewrite of the hot path, not a trick.

## Explicit Decisions

This document makes the following hard choices so execution does not drift.

### Decision 1: viewport sessions replace exact viewport subscriptions

The mounted grid hot path will not continue to use exact-viewport subscription identity.

`subscribeViewport(sheetName, viewport, listener)` can remain as a compatibility substrate, but
the mounted renderer must move to a session-oriented API that preserves residency and patch state
across scroll movement.

### Decision 2: Canvas 2D is the production text backend

The production text backend for the workbook grid hot path is retained pane-aligned `Canvas 2D`,
not DOM text and not an immediate WebGPU glyph-atlas rewrite.

This is a final architectural decision for the current redesign, not a bridge tactic.

### Decision 3: frozen panes become physical pane boundaries

Freeze panes will be represented as separate rendered panes with distinct text and rect surfaces.
They will not remain a purely mathematical partition inside one monolithic scene.

### Decision 4: React owns shell chrome, not cell rendering

React remains the owner of workbook shell composition and overlays, but it stops owning visible
cell text and stops coordinating per-scroll scene rebuild work.

### Decision 5: performance counters are first-class product instrumentation

The grid browser path will expose structured performance counters for tests and diagnostics.

These counters are not debug-only logs. They are part of the maintained runtime surface used to
prove:

- session churn is gone
- patch churn stays bounded
- React commits stay outside the hot path
- DOM text is no longer mounted

## Current Architecture and Concrete Failure Points

### 1. Scroll changes viewport identity too aggressively

Current path:

- `packages/grid/src/useWorkbookGridRenderState.ts`

The mounted grid computes a new exact viewport from scroll position and re-runs the viewport
subscription effect whenever that viewport changes.

Current consequence:

- horizontal movement turns into subscription lifecycle churn
- viewport movement is treated like data-source movement instead of local browsing inside a
  resident region

### 2. Worker viewport attach currently emits a full patch

Current path:

- `apps/web/src/worker-runtime-viewport-publisher.ts`
- `apps/web/src/projected-viewport-patch-coordinator.ts`

Current consequence:

- subscription churn is expensive even when the underlying data is already locally available
- large horizontal browsing can repeatedly pay initialization work

### 3. Browser state fanout is too broad

Current path:

- `apps/web/src/use-worker-workbook-app-state.tsx`
- `apps/web/src/projected-viewport-store.ts`
- `apps/web/src/projected-viewport-cell-cache.ts`

Current consequence:

- cell, axis, freeze, and selected-cell consumers all hang off broad store subscription paths
- patch application can wake too much React even when the visible requirement is only grid
  surface repaint

### 4. The grid rebuilds scenes instead of retaining them

Current path:

- `packages/grid/src/useWorkbookGridRenderState.ts`
- `packages/grid/src/gridGpuScene.ts`
- `packages/grid/src/gridTextScene.ts`
- `packages/grid/src/visibleGridAxes.ts`

Current consequence:

- viewport movement rebuilds visible item lists
- GPU scene arrays are rebuilt
- text scene arrays are rebuilt
- axis bounds are re-derived from visible cells
- hot path allocations and computation scale with the full visible scene every time

### 5. Text rendering still uses DOM nodes per visible item

Current path:

- `packages/grid/src/GridTextOverlay.tsx`

Current consequence:

- the current text plane has an unavoidable DOM layout and paint ceiling
- wide-sheet horizontal browse pays browser DOM cost exactly where the product most needs a
  retained renderer

### 6. Freeze pane handling is logically correct but not renderer-native

Current path:

- `packages/grid/src/workbookGridViewport.ts`
- `packages/grid/src/gridGpuScene.ts`
- `packages/grid/src/gridTextScene.ts`

Current consequence:

- freeze behavior is represented through scene math inside one logical surface instead of a
  compositor-friendly multi-pane renderer model
- this makes smooth body scrolling with pinned panes harder than it needs to be

## Required End-State Architecture

The mounted browser workbook must move to a renderer architecture with four properties:

1. scroll is local and mostly transform-driven
2. viewport data residency is stable and tile-based
3. renderer scenes are retained and damage-based
4. text is rendered by the renderer, not by React DOM cell nodes

### 1. Stable viewport residency sessions

The product must stop expressing scroll as exact viewport subscription churn.

Replace the current exact-viewport subscription model with a stable session model:

- the grid opens one viewport residency session per mounted sheet surface
- the session has a resident region expressed in tile-aligned bounds with overscan
- local scroll updates move the visible window inside the resident region without changing the
  underlying session identity
- worker/runtime traffic happens only when the resident region must expand or shift across tile
  boundaries

Required protocol behavior:

- session attach: initialize resident region and patch state once
- session move within resident region: no attach, no full patch
- session expansion or boundary crossing: region update against the same session identity
- session teardown: one explicit disposal path

This requires a session-oriented API instead of the current
`subscribeViewport(sheetName, viewport, listener)` shape as the mounted hot path.

### 2. Tile-stable worker/browser contract

The worker/runtime path must become stable under scroll.

Required changes:

- the worker publisher must preserve patch state per viewport session rather than per exact
  subscription rectangle
- the local tile store must become the primary source for already-resident browse movement
- region updates must compute entering and leaving tile bands instead of replacing the full
  viewport identity
- the projected viewport store must treat session movement as resident-window evolution, not a
  fresh subscription

Design requirement:

- if the user scrolls horizontally across cells that are already covered by resident tiles, the
  worker path should be silent

That is the minimum standard for a smooth browser browsing experience.

### 3. Quadrant-based renderer surfaces

Freeze panes should become first-class renderer structure, not just math.

The mounted grid should be composed into four panes:

- top-left frozen corner
- top scrolling header band
- left scrolling row band
- main scrolling body pane

Each pane owns:

- a rect surface
- a text surface
- hit-testing geometry aligned to that pane

Body scroll then becomes:

- translate the main body pane horizontally and or vertically
- leave frozen panes still
- repaint only newly exposed strips

This turns freeze panes from a global scene complication into a compositor-aligned layout model.

### 4. Retained scene caches instead of full scene rebuilds

The grid renderer must stop rebuilding the whole visible scene on movement.

Introduce retained caches for:

- visible column bounds by pane
- visible row bounds by pane
- cell background and border geometry by tile or stripe
- text layout records by cell signature
- text paint strips by pane and row-band or column-band

Damage should be represented as:

- entering columns
- leaving columns
- entering rows
- leaving rows
- cell-content changes
- axis-size changes
- selection or hover overlay changes

Pure horizontal browse in the main body should do only this:

- update body pane transform
- compute newly exposed column strip if needed
- paint the new strip into retained surfaces
- update selection and hover overlay positions if they remain visible

It must not rebuild the entire body scene.

### 5. Renderer-owned text backend

The current DOM text overlay must be replaced.

The chosen end-state backend for this design is:

- a retained canvas text renderer
- mounted as pane-aligned text surfaces
- fed by cached text layout records and strip-level paint invalidation

This design explicitly chooses a renderer-owned Canvas 2D text surface instead of a DOM text
overlay and instead of a speculative full WebGPU glyph-atlas rewrite.

Reasons:

- it removes the DOM-node-per-cell ceiling immediately
- it preserves full browser text raster quality and font shaping behavior
- it is compatible with retained strip repaint and pane composition
- it is substantially lower risk than attempting a full custom GPU text stack in the same
  project phase

This is not a temporary fallback.
It is the intended production text renderer for the workbook grid hot path.

#### Text renderer requirements

The text backend must preserve:

- left, center, and right alignment
- wrapping
- spill into adjacent empty cells
- clipping against frozen boundaries and pane edges
- underline and text decorations already supported by product styling
- number, date, string, and error display parity

The text backend must not:

- create one React element per visible text item
- rely on React reconciliation for scroll-time text movement
- measure and recompute the full visible text plane on every viewport change

#### Exact text parity matrix

The text backend migration is incomplete unless all of the following cases are covered by named
tests and pass before the DOM text layer is deleted.

Alignment and formatting:

- left-aligned plain strings
- centered headers
- right-aligned numbers
- right-aligned date displays
- underline
- mixed font size and family from style records already supported by the product

Wrapping and clipping:

- single-line unwrapped text clipped to cell bounds
- wrapped multi-line text within one cell
- wrapped text clipped by viewport bottom edge
- wrapped text clipped by frozen pane boundaries
- wrapped text inside resized rows and columns

Spill behavior:

- left-aligned text spilling through contiguous empty cells
- spill blocked by a non-empty adjacent cell
- spill blocked by a hidden column boundary
- spill clipped at the visible pane boundary
- spill clipped at the frozen-pane split

Viewport and freeze interactions:

- text in the main body pane while horizontally browsing
- text in frozen columns while the body scrolls
- text in frozen top rows while the body scrolls vertically
- text in the frozen corner while both axes move

Mutation and collaboration:

- text repaint after local edit inside the visible pane
- text repaint after collaborator patch inside the visible pane
- no repaint of unaffected strips after a narrow visible patch
- stable text while scrolling and collaborator patches overlap

Selection and editor interactions:

- selected-cell text hidden only when the editor overlay owns that cell
- non-edited cells remain visible while the editor is open
- editor overlay anchoring stays correct while the surrounding pane scrolls

The parity suite must use two complementary mechanisms:

- deterministic unit tests for layout and clipping decisions
- browser tests for visually observable behavior under real scroll and mutation conditions

The migration may temporarily keep the old `gridTextScene.ts` logic as a test oracle, but not as a
mounted production renderer path.

### 6. Store fanout decomposition

The browser store must stop waking the workbook shell broadly during pure browse movement.

Split the current projected viewport store responsibilities into narrower observable domains:

- resident cell tile data
- axis metadata and hidden state
- freeze pane metadata
- selected cell summary
- style lookup cache

The mounted grid renderer should subscribe only to the data domains needed for:

- pane scene invalidation
- overlay invalidation
- editor overlay invalidation

The formula bar, toolbar, and selection summary should not rerender during pure body pan.

### 7. React boundary simplification

React should own:

- workbook shell composition
- formula bar
- sheet tabs
- status surfaces
- editor overlay
- menu and popover shells

React should not own:

- per-cell text nodes
- per-frame visible cell scene assembly
- per-scroll geometry mutation

The grid surface should become an imperative renderer boundary with narrow declarative inputs.

## Required File and Module Restructure

This design is not credible unless it names the ownership changes.

### Primary grid refactor

Current files to slim down:

- `packages/grid/src/useWorkbookGridRenderState.ts`
- `packages/grid/src/WorkbookGridSurface.tsx`

Create or split into focused modules:

- `packages/grid/src/useGridScrollState.ts`
  - local scroll position
  - pane transform state
  - visible window inside resident region
- `packages/grid/src/useGridViewportResidency.ts`
  - resident tile window management
  - viewport session lifecycle
  - worker session region updates
- `packages/grid/src/gridQuadrantLayout.ts`
  - four-pane geometry
  - frozen and scrolling pane boundaries
- `packages/grid/src/gridSceneCache.ts`
  - retained pane scene cache
  - dirty-strip invalidation
- `packages/grid/src/gridAxisBoundsCache.ts`
  - retained row and column bounds by pane
- `packages/grid/src/GridTextCanvasSurface.tsx`
  - pane text canvas mounting
  - imperative paint lifecycle
- `packages/grid/src/gridTextLayoutOracle.ts`
  - shared text layout semantics used by tests during migration
- `packages/grid/src/gridTextLayoutCache.ts`
  - cell-signature keyed text layout records
- `packages/grid/src/gridTextStripPainter.ts`
  - retained row-band and column-band repaint logic
- `packages/grid/src/gridOverlayLayer.ts`
  - selection, hover, resize, and fill overlays without full scene rebuild coupling

### Worker/runtime and store changes

Current files to change:

- `apps/web/src/use-worker-workbook-app-state.tsx`
- `apps/web/src/projected-viewport-store.ts`
- `apps/web/src/projected-viewport-patch-coordinator.ts`
- `apps/web/src/projected-viewport-cell-cache.ts`
- `apps/web/src/worker-runtime-viewport-publisher.ts`
- `apps/web/src/worker-runtime-support.ts`
- `apps/web/src/worker-runtime.ts`
- `apps/web/src/worker-viewport-tile-store.ts`

Create new modules:

- `apps/web/src/viewport-session.ts`
  - stable client-side viewport session abstraction
- `apps/web/src/projected-viewport-session-store.ts`
  - resident-region session state and region update logic
- `apps/web/src/projected-viewport-cell-channel.ts`
  - narrow cell tile invalidation channel for grid renderer consumers
- `apps/web/src/perf/workbook-scroll-perf.ts`
  - browser performance marks and trace helpers for scroll workloads
- `apps/web/src/perf/workbook-scroll-perf-collector.ts`
  - browser-side accumulation of frame, long-task, and patch churn samples
- `apps/web/src/perf/workbook-react-commit-counter.tsx`
  - `React.Profiler`-backed commit counting around workbook shell boundaries in perf runs

### Test and perf harness changes

Create or expand:

- `packages/grid/src/__tests__/grid-scroll-state.test.ts`
- `packages/grid/src/__tests__/grid-scene-cache.test.ts`
- `packages/grid/src/__tests__/grid-text-layout-cache.test.ts`
- `packages/grid/src/__tests__/grid-quadrant-layout.test.ts`
- `apps/web/src/__tests__/viewport-session.test.ts`
- `apps/web/src/__tests__/projected-viewport-session-store.test.ts`
- `apps/web/src/__tests__/workbook-scroll-performance.test.ts`
- `apps/web/src/__tests__/workbook-scroll-perf-collector.test.ts`
- `e2e/tests/web-shell-scroll-performance.pw.ts`

## Phased Implementation Plan

This work must be landed in strict phases.

Do not combine the phases into one giant renderer rewrite.

### Phase 0: Measurement and guardrails

Goal:

- make the current jank measurable and prevent “feels faster on my machine” work

Deliverables:

- browser scroll perf harness
- workbook trace instrumentation
- explicit wide-sheet workloads
- patch attach and session churn counters
- React commit counters around the grid shell
- artifact JSON schema and helper utilities for perf result capture
- deterministic benchmark-corpus-backed browser fixtures for wide-sheet browsing

Exit criteria:

- current failure is reproducible in CI and local harnesses
- steady-state horizontal browse metrics are captured in artifacts
- no implementation work starts without these baselines
- artifact shape is stable enough that later phases can gate on it without rewriting the harness

### Phase 1: Stable viewport sessions and resident tile windows

Goal:

- remove per-scroll viewport resubscription churn

Deliverables:

- session-based viewport API
- tile-aligned resident window with overscan
- region update semantics instead of exact-viewport attach semantics
- worker publisher state preserved across movement
- session counters exported into the perf collector

Exit criteria:

- pure scroll inside a resident tile window causes zero session attach and zero full patch
- worker patch counts scale with tile-boundary crossings, not scroll frames
- these thresholds are enforced by CI as a merge-blocking perf check

### Phase 2: Store fanout reduction and imperative renderer boundary

Goal:

- ensure pure browse movement does not wake the workbook shell

Deliverables:

- narrow store channels
- grid renderer subscriptions isolated from formula bar, toolbar, and status surfaces
- `useWorkbookGridRenderState` decomposed into controller modules
- explicit React commit instrumentation around the workbook shell root

Exit criteria:

- formula bar and toolbar do not rerender during pure horizontal browse
- React commit counts during pure pan are limited to grid-boundary control work only

### Phase 3: Quadrant renderer and retained scene caches

Goal:

- make scroll mostly transform work plus strip invalidation

Deliverables:

- pane-based layout
- retained axis bounds
- retained GPU rect scene caches
- dirty-strip scene updates for entering and leaving rows and columns
- no full visible rect-scene rebuilds on resident horizontal pan

Exit criteria:

- freeze panes remain fixed with no perceptible lag
- visible rect scene is not rebuilt wholesale during steady-state pan

### Phase 4: Renderer-owned text backend

Goal:

- eliminate DOM text as the scroll bottleneck

Deliverables:

- pane-aligned canvas text surfaces
- cached text layout records keyed by cell signature
- strip repaint logic for newly exposed rows and columns
- spill, wrap, and clipping parity against current product behavior
- full text parity matrix wired into unit and browser suites

Exit criteria:

- no per-cell DOM text nodes remain in the mounted grid hot path
- sustained wide-sheet browse meets frame and long-task budgets
- text fidelity matches product requirements under scroll, freeze, and collaboration
- the perf suite is merge-blocking for browser workbook changes

### Phase 5: Hardening under collaboration and mutation load

Goal:

- keep the browsing path smooth even while patches and edits land

Deliverables:

- damage coalescing rules for visible collaborator patches
- patch prioritization for active viewport correctness
- explicit tests for browse plus collaborator update overlap

Exit criteria:

- same-sheet collaborator updates do not hitch the active scrolling pane
- visible patch application remains damage-based and bounded

## Verification Strategy

This design is incomplete unless the verification path is equally strict.

### Browser instrumentation

Capture:

- frame times from `requestAnimationFrame`
- main-thread long tasks
- React commit counts inside the workbook shell
- viewport session attach, move, and dispose counts
- full patch count versus damage patch count
- strip repaint counts
- visible tile residency hit rate
- DOM node count for mounted grid text surfaces, to prove the DOM text plane is gone
- explicit shell-boundary commit counts from a profiler-backed counter, not inferred heuristics

### Artifact schema

Each browser perf artifact must be machine-readable JSON with one top-level object:

- `workload`
- `fixture`
- `browser`
- `samples`
- `summary`
- `counters`
- `thresholds`
- `pass`

Where:

- `samples` contains raw frame deltas and long-task durations
- `summary` contains min, median, p95, p99, and max rollups
- `counters` contains session, patch, repaint, and React commit counts
- `thresholds` records the exact threshold values used by the gate

The artifact format must be stable across phases so results are comparable over time.

### Traceable browser scenarios

The perf suite must run:

- cold first browse
- warm steady-state browse
- browse with frozen panes
- browse with variable-width columns
- browse while collaborator patches land
- browse while selection remains visible

The perf suite must also record a deterministic scripted browse path:

- open the workbook fixture
- wait for warm ready state
- locate the real scroll viewport
- execute a fixed-duration horizontal browse with constant scroll velocity
- stop sampling
- collect browser metrics and attach the JSON artifact

No workload may collect results from ad hoc mouse movement or uncontrolled manual scrolling.

The page-side collector should be exposed in a narrow structured form such as
`window.__biligScrollPerf` during perf runs so Playwright can retrieve counters without scraping
console output.

### Correctness and parity coverage

The renderer rewrite must ship with parity tests for:

- wrapped text
- spilled text
- numeric and date alignment
- hidden rows and columns
- freeze panes
- selection and hover overlays
- editor overlay anchoring
- collaborator patch visibility

The parity suite must be named and versioned so regressions are visible as renderer regressions,
not as generic browser failures.

The text parity suite must also pin:

- CI font family
- device scale factor
- viewport size

so canvas text results are deterministic enough for meaningful browser assertions without relying
on brittle whole-screen pixel snapshots.

### Acceptance criteria for rollout

The work is not production-ready until all of the following hold on the same committed tree:

- scroll performance harnesses pass
- functional browser tests pass
- no parity regressions remain open for visible workbook behavior
- no design step depends on “temporary degraded mode”

## Migration Risks and Mitigations

### Risk: text fidelity regressions during DOM text removal

Mitigation:

- treat text parity as a first-class test matrix
- migrate pane by pane behind a renderer boundary rather than interleaving DOM and canvas text
- compare layout behavior against the current mounted surface before removal

### Risk: viewport sessions become a second truth model

Mitigation:

- keep the worker/runtime authoritative
- use sessions only as a residency and patch-delivery protocol, not as a semantic workbook state

### Risk: oversized rewrite diff

Mitigation:

- land by phase
- require each phase to have its own tests and acceptance gates
- keep existing product behavior visible while internal ownership changes
- cap each phase to one architectural axis of change so regressions are attributable

### Risk: freeze panes regress during compositor-oriented rewrite

Mitigation:

- make quadrant layout explicit before the text backend swap
- test frozen and unfrozen workloads separately

## What This Design Deliberately Does Not Do

- It does not defer the hard part by hiding content during scroll.
- It does not push more state into React to “make it simpler.”
- It does not require a second spreadsheet execution engine.
- It does not assume a full custom GPU text system is necessary for the first successful
  production result.
- It does not soften the browser experience target.
- It does not allow the final architecture to keep exact-viewport resubscribe churn in the mounted
  scroll path.

## Final Design Statement

The production path to a genuinely smooth browser workbook is:

- stable viewport sessions
- resident tile browsing
- narrow browser-store fanout
- quadrant-based renderer layout
- retained scene caches
- renderer-owned canvas text

Anything short of that may improve metrics in spots, but it will not produce the browsing
experience this product needs.

That is the implementation design this repo should execute if the requirement is:

- high FPS
- no frame dropping
- no jank
- no hacks
- no cheating
- and no degraded “perf mode” compromises.
