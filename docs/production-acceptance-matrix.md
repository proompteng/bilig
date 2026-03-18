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
- local-server app exists and is tested
- `apps/web` exists as the shipping app wrapper

## Spreadsheet product

- Excel shell UX passes browser acceptance
- worker-first browser runtime is active
- offline restore and reconnect are proven
- localhost app-server reconnect and cursor catch-up are proven

## Formula and WASM

- checked-in Excel parity corpus exists
- JS oracle matches the corpus
- WASM overlap matches JS on all claimed kernels

## Collaboration

- CRDT convergence passes under replay, reorder, duplicate delivery, and reconnect
- local agent or browser mutations are emitted through one ordered workbook commit stream
- backend append-before-ack is proven
- cursor catch-up and snapshot restore are proven

## Operations

- GitHub green
- Forgejo green
- Argo pre-prod manifests exist in `lab`
- performance budgets are green

## First tranche status

Architecture-foundation rows are closed, and the local app server plus product-app split tranche is now partially closed. Product-parity, worker-runtime, chat orchestration, durable-backend, and full formula-parity rows remain open until the deeper implementation tranches land.
