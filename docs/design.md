# Design

This repository implements the spreadsheet monorepo described in the implementation-ready design document, with these enforced overrides:

- React exists only in `apps/playground`
- shared packages remain React-free
- the custom workbook reconciler lives in `apps/playground/src/reconciler`
- the engine is CRDT-ready and local-first by design
- the AssemblyScript package provides a browser-embedded numeric fast path

The rest of the docs in this folder define the current codebase shape and public boundaries.
