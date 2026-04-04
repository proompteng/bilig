# Glide Removal And Native Grid Renderer V2

## Status

`bilig` now mounts the native
[packages/grid/src/WorkbookGridSurface.tsx](/Users/gregkonush/github.com/bilig/packages/grid/src/WorkbookGridSurface.tsx)
runtime as the primary workbook grid surface. The current surface owns the visible
sheet plane directly:

- body fills
- grid lines
- custom borders
- selection outline and fill
- fill preview and fill handle
- boolean cell visuals
- header backgrounds
- row marker backgrounds
- visible cell text
- visible row numbers
- visible column labels
- hover state
- resize guide
- header drag guide
- select-all corner control

The current repo also routes renderer-facing text/autofit work through native
render-cell snapshots, uses local selection containers instead of Glide selection
objects, and falls back to Canvas2D when WebGPU is unavailable.

This document now serves as the implementation record and cleanup plan for the
native workbook surface rather than as a future-state proposal.

## Problem

The current architecture cannot reach the desired product bar.

### Product problems

- The workbook still feels like a product shell wrapped around a third-party grid.
- Interaction polish is inconsistent because paint ownership and behavior ownership
  are split across two systems.
- Google Sheets-level quality requires one coherent rendering and interaction model,
  not a hybrid stack.

### Performance problems

- Canvas2D callback rendering through Glide imposes an architectural ceiling.
- WebGPU cannot fully own the frame while Glide still drives the sheet plane.
- The current hot path still pays React and Glide integration overhead that a native
  render loop would avoid.

### Maintainability problems

- The renderer surface is constrained by Glide’s component and callback model.
- Hit-testing and interaction policy are harder to reason about because geometry and
  selection state are partly ours and partly Glide’s.
- Every major polish change requires working around another system instead of owning
  the behavior directly.

## Goals

- Remove Glide from the workbook surface completely.
- Make `bilig` own rendering, hit-testing, scrolling, selection, resize, fill,
  editing activation, and header behavior.
- Use WebGPU as the primary visual plane for the sheet.
- Keep workbook semantics and engine behavior independent from the renderer.
- Maintain or improve current functionality while increasing consistency and speed.

## Non-goals

- This is not a formula-engine rewrite.
- This is not a file-import/export rewrite.
- This is not a general design-system overhaul outside the workbook surface.
- This is not a “ship a parallel experimental grid forever” plan.

## Success Criteria

The renderer replacement is complete when all of the following are true:

- `SheetGridView` no longer imports or mounts `DataEditor`.
- The main workbook surface renders through `WorkbookGridSurface`.
- WebGPU owns the body plane and selection plane.
- Header/body hit-testing is fully local to `@bilig/grid`.
- Column resize, row/column selection, fill drag, and select-all all run through
  local geometry and interaction controllers.
- Editing is activated and positioned through local hit-testing and editor geometry,
  not Glide callbacks.
- Browser perf is stable at 60fps for common navigation and editing flows on a
  representative workbook.

## Design Principles

### One renderer, one interaction model

The grid surface must be a single owned system. Paint and behavior must not be
split between a GPU layer and a foreign selection engine.

### Headless workbook semantics

The workbook engine remains the semantic source of truth. The renderer consumes
snapshots, viewport deltas, styles, and metrics. It does not become the model.

### Immediate-mode rendering, retained interaction state

The frame should be drawn from scene data each tick or damage pass. Interaction
state should be retained in small explicit controller state, not inferred from DOM.

### Typed geometry everywhere

All hit-testing, bounds resolution, drag math, and damage computation should use
explicit row/column geometry utilities, never reparsed DOM assumptions.

### Minimal React in the hot path

React should coordinate the shell and editor overlays. It should not be the
frame-by-frame rendering engine for the sheet.

## Target Architecture

### New primary surface

Add a native surface component:

- `packages/grid/src/WorkbookGridSurface.tsx`

Responsibilities:

- own the scroll container
- own the GPU canvases and text overlay
- own hit-testing for cells, headers, row markers, resize handles, and fill handle
- own pointer/keyboard interaction controllers
- expose semantic callbacks to the workbook shell

This replaces Glide as the mounted sheet body.

### Renderer modules

The renderer should be broken into explicit layers:

- `grid-scene-body.ts`
  - fills, grid lines, custom borders, boolean visuals
- `grid-scene-selection.ts`
  - selection fill, outline, hover, drag guides, resize guides
- `grid-scene-header.ts`
  - header backgrounds, row marker surfaces, corner affordance surfaces
- `grid-scene-text.ts`
  - visible cell text, row numbers, column labels
- `grid-scene-editor.ts`
  - editor target geometry and overlay anchoring

The current scene modules can be evolved into this split; they should not stay as
one ever-growing `gridGpuScene.ts`.

### Interaction modules

The renderer should own these behaviors explicitly:

- `grid-hit-testing.ts`
  - resolve cell, header, row marker, corner, resize handle, fill handle
- `grid-pointer-controller.ts`
  - pointer down/move/up state machine
- `grid-keyboard-controller.ts`
  - navigation, range extension, selection modes, clipboard shortcuts
- `grid-scroll-controller.ts`
  - viewport movement, wheel/touch handling, inertial behavior policy
- `grid-ime-editor-controller.ts`
  - edit activation, overlay seed, composition handling, commit/cancel

Some current controller code already exists. It should migrate here and become the
sole interaction implementation rather than an adapter around Glide behavior.

### Viewport model

The renderer should consume a typed viewport model:

- visible row range
- visible column bounds
- scroll offsets
- frozen regions, when added
- damage regions

This should be updated through a local scroll controller rather than Glide’s
`onVisibleRegionChanged`.

### Text rendering strategy

Text should remain separate from rect rendering:

- short term: canvas text overlay driven by local scene data
- medium term: glyph atlas or SDF text path when the current canvas text overlay
  becomes the bottleneck

The key requirement is that text layout and visibility are owned locally, not by
Glide.

## Proposed Package Shape

### Keep

- `@bilig/core`
- `@bilig/formula`
- `@bilig/protocol`
- `@bilig/grid`

### Add or formalize inside `@bilig/grid`

- `render/`
  - scene builders
  - TypeGPU pipeline setup
  - text overlay primitives
- `interaction/`
  - hit-testing
  - pointer controller
  - keyboard controller
  - fill/resize drag controllers
- `viewport/`
  - metrics
  - visible ranges
  - damage computation
- `surface/`
  - `WorkbookGridSurface.tsx`

## Migration Plan

### Phase 0: Lock the contract

Before more renderer work, freeze the `SheetGridView` contract that the workbook
shell expects:

- selection change
- begin edit
- commit/cancel edit
- paste/copy/fill/copy-range
- column width updates
- visible viewport updates

This allows a renderer swap without rewriting the app shell at the same time.

### Phase 1: Introduce `WorkbookGridSurface`

Create the new component alongside Glide.

Initial scope:

- local scroll container
- local viewport math
- GPU body and header surfaces
- current text overlay reused

At this phase, the surface may still forward some semantic callbacks to existing
controllers, but it must not mount Glide.

### Phase 2: Own hit-testing

Move these fully local:

- body cell hit-testing
- row header hit-testing
- column header hit-testing
- select-all corner hit-testing
- resize handle hit-testing
- fill handle hit-testing

Glide should no longer be the thing telling us what the pointer hit.

### Phase 3: Own selection and drag lifecycle

Replace Glide selection semantics with local controllers:

- click selection
- shift-click range expansion
- drag selection
- row slice selection
- column slice selection
- select all
- fill drag
- resize drag

This is the point where the selection model becomes truly coherent.

### Phase 4: Own editing activation

Replace Glide edit lifecycle with local activation:

- double click to edit
- type-to-edit
- keyboard activation
- IME composition
- overlay geometry
- commit/cancel movement semantics

The existing `CellEditorOverlay` can remain, but its trigger and target bounds
must come from local geometry.

### Phase 5: Delete Glide

Delete:

- `DataEditor` mount usage
- Glide selection reconciliation code
- Glide-specific grid cell conversion code
- Glide theme workarounds that only exist to hide its paint

At the end of this phase, [packages/grid/README.md](/Users/gregkonush/github.com/bilig/packages/grid/README.md)
must stop describing the grid as “built on Glide Data Grid.”

## Required Refactors Before Or During Migration

### 1. Split `SheetGridView`

`SheetGridView` still carries too much orchestration. It should become a thin
wrapper or disappear into:

- `WorkbookGridSurface`
- `GridEditorLayer`
- `GridClipboardController`
- `GridSelectionController`

### 2. Replace `cellToGridCell`

Current cell conversion still thinks in terms of Glide cell kinds. That API must be
replaced with a renderer-native snapshot-to-scene model.

Recommended replacement:

- `snapshotToRenderCell(...)`

This should output only what the renderer needs:

- display text
- semantic kind
- style refs
- alignment
- interaction flags

### 3. Replace Glide selection objects

Current selection logic leans on Glide `GridSelection` and `CompactSelection`.
That should be replaced with local selection structures:

- active cell
- rectangular selection
- row slice
- column slice
- whole sheet

with explicit helpers rather than imported third-party container types.

### 4. Move viewport ownership local

Visible ranges should come from local scroll position and local metrics, not
Glide callbacks.

### 5. Remove theme suppression hacks

`gridPresentation.ts` currently contains several transparency hacks whose only job
is to make Glide invisible. Those should disappear as Glide leaves.

## Performance Plan

### Rendering targets

- steady 60fps for scroll on representative workbooks
- no long main-thread tasks above 16ms in standard navigation flows
- no React-driven rerender of the sheet body on every pointer move

### Data targets

- O(visible cells) scene build cost
- O(damaged cells) incremental update path where feasible
- no address stringify/parse loops in the hot render path

### Memory targets

- bounded scene allocations per frame
- reusable typed arrays or pooled rect buffers for GPU uploads
- no unbounded off-viewport visual cache growth

## Verification Plan

### Unit tests

Add or expand tests for:

- hit-testing
- drag state machines
- resize math
- viewport calculations
- scene generation for selection and hover states

### Browser tests

Add browser automation for:

- click selection
- shift-click range expansion
- row and column header drags
- select-all corner
- fill drag
- double click edit
- type-to-edit
- resize and autofit

### Performance tests

Add repeatable performance scripts for:

- large scroll sweep
- repeated arrow navigation
- repeated fill drag
- large paste
- multi-column resize

Metrics should be captured before and after the Glide removal.

### Visual regression

Add targeted screenshots for:

- default empty sheet
- selected cell
- range selection
- row slice
- column slice
- fill preview
- boolean cells
- hovered header
- resize guide
- drag guides

## Risks

### IME and accessibility regression

Glide currently gives us some behavior for free. Removing it means we must own:

- keyboard focus
- edit activation semantics
- composition events
- screen-reader strategy for the surface

This must be planned, not patched later.

### Browser-specific rendering behavior

Canvas text and WebGPU composition differ across browsers. The renderer needs
explicit verification on Chromium, Safari, and Firefox where applicable.

### Scope creep

Do not mix this with formula, codec, or collaboration rewrites. Renderer removal
is already a large project.

## Milestones

### Milestone A

`WorkbookGridSurface` exists and renders a scrollable sheet without Glide.

### Milestone B

All pointer hit-testing and selection semantics are local.

### Milestone C

Editing activation and overlay positioning are local.

### Milestone D

Glide is removed from `@bilig/grid`.

### Milestone E

Perf and visual acceptance gates are green.

## Recommended Immediate Next Steps

1. Create `WorkbookGridSurface.tsx` with local scroll ownership.
2. Introduce renderer-native selection types to replace Glide selection objects.
3. Replace `cellToGridCell(...)` with `snapshotToRenderCell(...)`.
4. Port hit-testing and visible-range updates to the new surface before adding more
   hybrid polish.
5. Add a dedicated browser perf harness for scroll, selection, resize, and fill.

## Decision

`bilig` should stop treating Glide as the workbook runtime and move to a fully
owned native renderer. The current hybrid work was useful as an incremental bridge,
but it is not the end-state. The end-state is one renderer, one interaction model,
and one coherent workbook surface owned by `@bilig/grid`.
