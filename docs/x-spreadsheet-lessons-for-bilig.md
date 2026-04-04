# x-spreadsheet Lessons For `bilig`

This document captures what `bilig` should learn from the local x-spreadsheet repo at `/Users/gregkonush/github.com/x-data-spreadsheet`.

Reviewed source:

- `/Users/gregkonush/github.com/x-data-spreadsheet/README.md`
- `/Users/gregkonush/github.com/x-data-spreadsheet/src/core/data_proxy.js`
- `/Users/gregkonush/github.com/x-data-spreadsheet/src/component/sheet.js`
- `/Users/gregkonush/github.com/x-data-spreadsheet/src/component/table.js`
- `/Users/gregkonush/github.com/x-data-spreadsheet/src/canvas/draw.js`
- `/Users/gregkonush/github.com/x-data-spreadsheet/src/core/history.js`

## Executive Summary

x-spreadsheet is a useful small-to-medium spreadsheet UI reference, not a world-class workbook engine reference.

It is worth copying for:

- pragmatic component decomposition
- direct canvas drawing utilities
- compact data model ideas
- simple event-driven sheet interaction structure

It is not worth copying for:

- history model design
- formula engine depth
- long-term scalability assumptions
- deep workbook semantics

The best lesson from x-spreadsheet is that a spreadsheet product can stay understandable when the pieces are small and boring.

## What x-spreadsheet Does Well

### 1. The codebase is decomposed into understandable pieces

Compared with Luckysheet, x-spreadsheet is much easier to reason about.

There is a visible split between:

- core data structures in `src/core/*`
- UI components in `src/component/*`
- canvas primitives in `src/canvas/*`

That is valuable even if the implementation is simpler.

For `bilig`, the lesson is:

- keep module responsibilities obvious
- prefer small understandable surfaces over giant “do everything” components

### 2. `data_proxy.js` is a good example of compact workbook state ownership

`src/core/data_proxy.js` is not a full engine, but it is a decent example of a central sheet data model that still has a recognizably bounded job.

It owns:

- rows and columns
- merges
- clipboard
- validations
- selection-related sheet operations

That is much healthier than a giant ambient global store.

For `bilig`, the lesson is not to copy the exact object model.

The lesson is:

- state owners should have a readable scope
- sheet operations should hang off a coherent data model instead of scattering everywhere

### 3. The canvas utility layer is pragmatic

`src/canvas/draw.js` is worth reading because it stays direct:

- DPR helpers
- text positioning
- line drawing
- border drawing
- draw-box geometry

This is the same category of code a serious spreadsheet renderer needs.

For `bilig`, the lesson is:

- rendering helpers should be small and geometry-focused
- not overloaded with product logic

### 4. Table rendering is organized around visible work

`src/component/table.js` separates:

- cell rendering
- fixed header rendering
- autofilter rendering
- merge rendering

That is not a complete render engine, but it is a useful pattern.

For `bilig`, this reinforces:

- render passes should be explicit
- different visual layers should not be tangled together accidentally

### 5. The sheet interaction surface is readable

`src/component/sheet.js` is still imperative, but the interaction model is understandable:

- selection movement
- scroll syncing
- resize affordances
- overlay interaction

That clarity is valuable.

For `bilig`, it is a reminder that interaction code should stay locally understandable, especially for:

- selection movement
- scroll coupling
- resize behavior
- overlay/editor coordination

## What x-spreadsheet Does Poorly

### 1. History is too naive

`src/core/history.js` stores undo/redo snapshots as `JSON.stringify(data)`.

That is fine for a smaller widget.

It is not a good model for `bilig`.

Problems:

- high allocation cost
- poor scalability
- no semantic operation modeling
- no incremental patching

### 2. Formula and engine depth are limited

x-spreadsheet is not where the advanced calc-engine lessons are.

It is mainly a product-shell and UI interaction reference.

### 3. Some architecture is still imperative UI code rather than explicit runtime design

This is acceptable for a smaller library, but `bilig` needs stronger runtime boundaries than this codebase provides.

## What `bilig` Should Copy

### A. Keep modules small and comprehensible

The strongest x-spreadsheet lesson is stylistic:

- do not let every subsystem become huge
- keep render helpers focused
- keep sheet interaction paths readable

### B. Explicit render passes

`bilig` should continue moving toward distinct render responsibilities for:

- body cells
- headers
- merges
- selection chrome
- overlays
- filter or decoration layers

### C. Geometry-first canvas or GPU primitives

Even though `bilig` is moving toward GPU-backed rendering, the same rule applies:

- geometry primitives should remain tiny and explicit

## What `bilig` Should Not Copy

### 1. Snapshot-by-JSON undo/redo

This is far too naive for a serious workbook engine.

### 2. Treating UI state as the engine

x-spreadsheet is better than Luckysheet structurally, but it is still primarily a UI library, not a deep workbook runtime.

### 3. Limited engine semantics

Do not let a simpler widget architecture become the standard for workbook correctness, recalculation, or collaborative behavior.

## Recommended Actions For `bilig`

### Near term

- keep extracting interaction and render helpers into small modules
- keep render layers explicit and separate

### Medium term

- keep geometry helpers small as the GPU renderer grows
- make selection, scrolling, and overlay behavior locally understandable

### Long term

- preserve x-spreadsheet’s readability without inheriting its simplistic history and engine model

## Bottom Line

x-spreadsheet is worth studying because it stays relatively understandable.

Copy:

- compact decomposition
- explicit render passes
- small geometry primitives

Reject:

- snapshot-by-JSON history
- shallow engine assumptions
- using a UI-library architecture as the whole workbook-runtime model
