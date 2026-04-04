# HyperFormula Lessons For `bilig`

This document captures what `bilig` should learn from the local HyperFormula repo at `/Users/gregkonush/github.com/hyperformula`.

Reviewed source:

- `/Users/gregkonush/github.com/hyperformula/README.md`
- `/Users/gregkonush/github.com/hyperformula/src/BuildEngineFactory.ts`
- `/Users/gregkonush/github.com/hyperformula/src/GraphBuilder.ts`
- `/Users/gregkonush/github.com/hyperformula/src/DependencyGraph/DependencyGraph.ts`
- `/Users/gregkonush/github.com/hyperformula/src/DependencyGraph/AddressMapping/ChooseAddressMappingPolicy.ts`
- `/Users/gregkonush/github.com/hyperformula/src/parser/ParserWithCaching.ts`

## Executive Summary

HyperFormula is the strongest direct engine reference in this competitor set.

It is worth copying for:

- headless engine boundaries
- explicit engine composition
- dependency graph design
- adaptive dense vs sparse address storage
- parser caching
- formula-centric CRUD and undo/redo thinking

It is not worth copying for:

- product shell or workbook UI
- collaboration architecture
- renderer strategy
- file import/export breadth

The main lesson for `bilig` is that a serious spreadsheet engine is not a giant workbook object with helper methods.

It is a deliberately composed runtime with:

- parser
- dependency graph
- evaluator
- operation layer
- serialization/export layer
- configurable storage/index strategies

## What HyperFormula Does Well

### 1. It is unapologetically headless

HyperFormula keeps the workbook engine separate from any visual grid.

The repo structure reinforces that:

- `src/BuildEngineFactory.ts`
- `src/DependencyGraph/*`
- `src/interpreter/*`
- `src/parser/*`
- `src/CrudOperations.ts`
- `src/UndoRedo.ts`

This is the correct shape for engine code.

For `bilig`, the lesson is:

- the calc/runtime core must remain independent from the grid renderer
- UI concerns should not leak into dependency evaluation, graph updates, or parser behavior
- workbook semantics should be testable without React, DOM, or browser layout

### 2. Engine construction is explicit and compositional

`BuildEngineFactory.ts` is one of the most useful files in the repo.

It wires together:

- config
- statistics
- dependency graph
- parser
- interpreter
- cell content parser
- operations
- undo/redo
- clipboard operations
- exporter
- serialization

That is valuable because the engine is assembled from subsystems instead of hidden inside one class.

For `bilig`, this suggests:

- make engine composition explicit
- keep subsystem seams visible
- make it obvious where parsing, graph mutation, evaluation, and export live

### 3. Adaptive dense vs sparse storage is a real performance idea

`ChooseAddressMappingPolicy.ts` is small, but it exposes a strong design choice:

- use dense storage when sheet fill is high
- use sparse storage when sheet fill is low

That is exactly the kind of practical engine optimization that matters.

For `bilig`, this is worth copying conceptually:

- do not assume one cell storage shape is ideal for every workbook
- allow sparse and dense strategies, especially for large sparse sheets
- make the switch a policy decision rather than hardcoded behavior

### 4. The dependency graph is the center of the engine

`DependencyGraph.ts` and `GraphBuilder.ts` make the design very clear:

- cells become vertices
- ranges become vertices
- formulas attach dependencies explicitly
- dirty/volatile/structural-change semantics are first-class

This is worth copying more than any one algorithm detail.

For `bilig`, the main lesson is:

- the dependency graph should be the semantic truth for recalculation
- recalculation should not be an ad hoc cascade over workbook cells
- structure-change awareness needs first-class representation

### 5. Parser caching is treated as normal infrastructure

`ParserWithCaching.ts` is valuable because it treats parser caching as part of the normal engine path, not an optional micro-optimization.

It does several useful things:

- token-based hashing
- AST reuse
- dependency recollection from cached ASTs
- normalization of reversed ranges

For `bilig`, this reinforces:

- parse caching should be built in, not bolted on
- formulas should normalize into stable forms early
- caching should happen at AST and dependency levels, not only display text levels

### 6. CRUD, clipboard, undo/redo, and graph semantics are tied together

HyperFormula does not treat workbook edits as loose mutations.

The engine has dedicated layers for:

- CRUD operations
- clipboard operations
- undo/redo
- serialization/export

That matters because spreadsheet correctness depends on edits, formula translation, dependency updates, and recomputation staying coherent.

For `bilig`, the lesson is:

- workbook operations should be semantic operations
- not just arbitrary state patches
- edits must flow through translation, invalidation, and graph update rules in one disciplined path

## What `bilig` Should Copy

### A. A more explicit engine assembly surface

`bilig` should continue moving toward an explicit engine composition model with clearly owned subsystems for:

- parser
- graph
- evaluator
- operation layer
- export/serialization
- performance statistics

### B. Dense vs sparse sheet storage policies

This is one of the highest-value concrete ideas in HyperFormula.

`bilig` should support:

- sparse storage for lightly filled sheets
- dense storage for heavily filled sheets
- policy-driven selection rather than one universal storage format

### C. First-class structural invalidation semantics

HyperFormula models:

- volatile behavior
- structure-dependent formulas
- range vertices

`bilig` should keep strengthening this area instead of treating structural edits as oversized generic invalidations.

### D. Parser caching and normalized AST handling

`bilig` should keep pushing parser work toward:

- stable AST normalization
- cached parse results
- dependency reuse
- lower-cost rebuilds after formula-preserving edits

### E. Engine statistics and profiling hooks

HyperFormula’s explicit `Statistics` wiring is a good habit.

`bilig` should keep or expand:

- parser timing
- graph-build timing
- recomputation timing
- operation timing

If performance matters, the engine should measure itself.

## What `bilig` Should Not Copy

### 1. HyperFormula is not a product shell

It does not solve:

- polished workbook UI
- selection and editing UX
- renderer architecture
- collaborative sync

Do not treat it as a product blueprint.

### 2. Do not blindly copy object names or class boundaries

The useful part is the architecture and performance mindset, not literal file names or class names.

### 3. Do not import licensing assumptions accidentally

HyperFormula is GPLv3 or commercial.

Treat it as an architectural reference, not a copy-paste dependency plan.

## Recommended Actions For `bilig`

### Near term

- keep the engine headless and separated from UI
- continue tightening structural invalidation semantics
- add or refine parser caching where formulas are reparsed often
- review whether sparse vs dense storage should become a first-class policy

### Medium term

- expose a clearer engine assembly surface
- make engine timing and profiling more systematic
- harden CRUD/edit operations as semantic engine operations instead of generic workbook mutations

### Long term

- build a true engine/runtime design doc for `bilig` that is as explicit as HyperFormula’s code structure
- support multiple storage/index strategies where workbook shape justifies it

## Bottom Line

HyperFormula is the best direct reference here for spreadsheet engine architecture.

If `bilig` wants to be a serious workbook engine, the main things to copy are:

- headless runtime discipline
- dependency-graph centrality
- parser caching
- adaptive storage policy
- explicit engine composition

Those lessons matter more than any single function implementation.
