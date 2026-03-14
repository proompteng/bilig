# Design

This repository implements the spreadsheet monorepo described in the implementation-ready design document, with the current package-centric layout:

- reusable React code lives in `packages/renderer` and `packages/grid`
- `apps/playground` is a thin app shell that composes those packages
- the engine is CRDT-ready and local-first by design
- the AssemblyScript package provides a browser-embedded numeric fast path
- the core runtime now includes extracted `RangeRegistry`, `EdgeArena`, `RecalcScheduler`, and SCC-based cycle detection instead of keeping the graph inline in the engine
- cell `format` is a real persisted attribute in the public model

The rest of the docs in this folder define the current codebase shape and public boundaries.
