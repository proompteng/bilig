# Testing and Benchmarks

## Required gate families

- unit and integration tests for every shared package
- Playwright browser tests for Excel-like UX
- binary protocol roundtrip and malformed-frame tests
- worker transport request/subscription parity tests
- CRDT convergence and replay tests
- backend durability/session tests
- Excel parity fixture tests
- performance budget enforcement

## Browser and UX coverage

Playwright coverage must include:

- navigation shortcuts
- drag selection
- fill handle
- copy/cut/paste
- formula bar and in-cell sync
- scrollbar gutter safety
- offline reload and restore
- million-row navigation

## Realtime coverage

- append-before-ack validation
- reconnect and cursor catch-up
- snapshot restore
- duplicate delivery
- concurrent multi-replica convergence

## Performance coverage

The source of truth for SLOs is [performance-budgets.md](/Users/gregkonush/github.com/bilig/docs/performance-budgets.md). CI and release checks must gate on those budgets rather than narrative claims.

## Current tranche status

This tranche adds tests for:

- binary sync framing
- agent stdio framing
- worker transport parity
- in-memory server durability primitives
- sync-server binary ingress
