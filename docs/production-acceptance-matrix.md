# Production Acceptance Matrix

## Core contract

- documented `@bilig/core` APIs exist in code
- range mutation helpers are tested
- undo/redo is tested
- selection state and sync state are tested

## Architecture

- canonical docs are updated and consistent
- worker transport package exists and is tested
- binary protocol package exists and is tested
- sync-server app exists and is tested

## Spreadsheet product

- Excel shell UX passes browser acceptance
- worker-first browser runtime is active
- offline restore and reconnect are proven

## Formula and WASM

- checked-in Excel parity corpus exists
- JS oracle matches the corpus
- WASM overlap matches JS on all claimed kernels

## Collaboration

- CRDT convergence passes under replay, reorder, duplicate delivery, and reconnect
- backend append-before-ack is proven
- cursor catch-up and snapshot restore are proven

## Operations

- GitHub green
- Forgejo green
- Argo pre-prod manifests exist in `lab`
- performance budgets are green

## First tranche status

Architecture-foundation rows are closed, and the first core-contract tranche is now partially closed. Product-parity, worker-runtime, durable-backend, and full formula-parity rows remain open until the deeper implementation tranches land.
