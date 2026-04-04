# Univer Lessons For `bilig`

This document captures what `bilig` should learn from the local Univer repo at `/Users/gregkonush/github.com/univer`.

Reviewed source:

- `/Users/gregkonush/github.com/univer/README.md`
- `/Users/gregkonush/github.com/univer/pnpm-workspace.yaml`
- `/Users/gregkonush/github.com/univer/docs/FIX_MEMORY_LEAK.md`
- `/Users/gregkonush/github.com/univer/packages`
- `/Users/gregkonush/github.com/univer/e2e/perf/scroll.spec.ts`
- `/Users/gregkonush/github.com/univer/e2e/visual-comparison/sheets/sheets-visual-comparison.spec.ts`
- `/Users/gregkonush/github.com/univer/e2e/visual-comparison/sheets/sheets-scroll.spec.ts`

## Executive Summary

Univer is the strongest full-product architecture reference in this competitor set.

It is worth copying for:

- package decomposition
- plugin-oriented product architecture
- isomorphic runtime thinking
- dedicated render engine boundaries
- explicit performance, memory, and visual regression testing

It is not worth copying for:

- feature sprawl by default
- package explosion without discipline
- assuming `bilig` needs the same document/slides scope

The best Univer lesson is not “build every feature they have”.

The best lesson is:

- split core, render, domain, and feature plugins cleanly
- treat memory and visual regressions as first-class engineering concerns
- make the spreadsheet product extensible without making the core incoherent

## What Univer Does Well

### 1. It has a real package architecture, not a giant app blob

The `packages/` tree is the first important signal:

- `packages/core`
- `packages/engine-render`
- `packages/engine-formula`
- `packages/sheets`
- `packages/sheets-ui`
- many feature packages such as `sheets-filter`, `sheets-sort`, `sheets-table`, `sheets-data-validation`, `sheets-formula-ui`

This is worth studying because it shows discipline around responsibility boundaries.

For `bilig`, the lesson is:

- keep core model/runtime separate from renderer
- keep feature modules separate from the base sheet product
- avoid one package or app owning all spreadsheet behavior

### 2. Plugin and feature modularity are treated as architecture, not an afterthought

Univer clearly thinks in terms of:

- base capabilities
- feature packages
- UI packages
- adapter layers

That is a mature product move.

For `bilig`, this means:

- sort/filter/data validation/comments/charts should not all land in the same core module
- feature growth should happen through deliberate module seams
- the shell should compose capabilities instead of accumulating one giant runtime file

### 3. They take render architecture seriously

The repo has an explicit `engine-render` package, and the README is direct about the rendering engine being a core capability.

That is important even if `bilig` does not copy Univer’s renderer.

The valuable lesson is:

- rendering deserves its own architecture boundary
- it should not be an incidental byproduct of the sheet UI layer

For `bilig`, this strongly supports the direction of:

- owning more of the render plane directly
- not leaving the renderer as a generic third-party grid plus product shell

### 4. They treat visual and performance verification as normal

The repo includes:

- `e2e/perf/scroll.spec.ts`
- `e2e/memory/memory.spec.ts`
- `e2e/visual-comparison/sheets/*`

That is one of the most valuable operational lessons in the repo.

For `bilig`, it reinforces:

- scroll performance should be tested, not just discussed
- rendering correctness should have visual regression coverage
- UI regressions in grids often need screenshot-based tests

### 5. They have explicit memory-leak engineering guidance

`docs/FIX_MEMORY_LEAK.md` is worth copying in spirit.

It calls out concrete leak sources:

- undisposed subscriptions
- singleton services holding current-unit references
- big objects in React dependency arrays

This is mature engineering.

For `bilig`, this means:

- memory-leak prevention should be documented as a repo practice
- render/session/runtime lifecycle should be reviewed with disposal discipline
- “it works after mount” is not enough for long-lived workbook sessions

### 6. Isomorphic thinking is useful

Univer positions itself as browser and server capable, with formula work that can run off the main UI surface.

Whether or not `bilig` copies their exact model, the principle is valuable:

- engine capabilities should not depend on one UI environment
- compute should be able to move between browser, worker, and server boundaries as needed

That aligns with `bilig`’s own workbook-runtime ambitions.

## What `bilig` Should Copy

### A. Stronger package boundaries around features

`bilig` should continue splitting workbook capabilities into explicit modules instead of accreting them into a few large files or packages.

Especially for:

- sorting and filtering
- data validation
- number formatting UI vs number formatting engine
- comments/notes
- tables and pivots
- render-only concerns vs workbook-domain concerns

### B. Treat rendering as a first-class engine

The important Univer lesson is not “use canvas because they do”.

It is:

- rendering should have a dedicated architecture
- performance work belongs in that subsystem
- product shell code should not own core render behavior

For `bilig`, that supports further movement toward a real grid render surface rather than layered UI tweaks.

### C. Add more visual and performance regression coverage

`bilig` should grow:

- scroll perf tests
- screenshot regression tests for key sheet states
- memory/regression scenarios for workbook mount/unmount

### D. Document lifecycle and leak rules

`bilig` should add explicit guidance for:

- subscription disposal
- render-unit ownership
- avoiding large object captures in React dependency arrays
- workbook/session teardown invariants

### E. Modularize product features without overcoupling the shell

Univer is a reminder that mature spreadsheet products need modular feature ownership.

`bilig` should keep pushing features out of giant shell and grid files into dedicated modules with clearer seams.

## What `bilig` Should Not Copy

### 1. Do not copy the entire feature surface

Univer supports spreadsheets, docs, and slides.

`bilig` does not need to inherit that scope creep.

### 2. Do not create package sprawl without discipline

Univer’s decomposition is useful because it is systematic.

Blindly making many packages without clear contracts would just create a worse monorepo.

### 3. Do not assume all plugin systems are good by default

Feature modularity is valuable.

An unprincipled plugin system can become indirection and startup overhead without real payoff.

## Recommended Actions For `bilig`

### Near term

- keep decomposing oversized workbook and grid files
- strengthen screenshot and perf testing around the sheet surface
- add memory/leak guidance to engineering docs

### Medium term

- keep splitting core runtime, render runtime, and feature modules
- make render ownership clearer and less dependent on third-party grid assumptions
- give each major feature a cleaner package or module boundary

### Long term

- maintain a spreadsheet architecture that can support workers/server execution cleanly
- formalize render-engine boundaries the way Univer formalizes `engine-render`

## Bottom Line

Univer is the best reference here for spreadsheet product architecture at scale.

The main things worth copying are:

- modular package boundaries
- render-engine explicitness
- plugin-style feature decomposition
- visual/perf/memory verification discipline

The main thing not to copy is the temptation to chase their scope instead of their engineering discipline.
