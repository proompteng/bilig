# Typed Agent Protocol RFC

## Problem

The current agent surface is a meaningful foundation, but the payload format is still `JSON.stringify(...)` inside a binary frame envelope. That is enough for the current CLI and local server path, but it is not yet the typed binary agent protocol implied by the architecture direction.

## Goals

- define a true typed binary agent protocol
- support both stdio and remote transports with the same schema
- enable streaming events and subscriptions as first-class behavior
- define safe mutation, diff preview, and rollback-aware agent surfaces

## Non-goals

- scrape UI state instead of using the workbook runtime
- make the first version cover every future agent workflow
- remove the current JSON-in-binary-envelope transport without a migration path

## Current state

Today the stack already has:

- typed TypeScript request/response/event unions
- binary framing for stdio and network payloads
- local workbook session execution in `apps/local-server`
- remote ingress shape in `apps/sync-server`

What it does not yet have is a real typed binary payload schema, a streaming-first event model, and a mature remote worksheet execution story.

## Target protocol families

### Request families

- session lifecycle
- viewport and range reads
- worksheet mutation proposals
- mutation apply/commit
- diff preview
- explain and dependency inspection
- snapshot export/import
- metrics and traces
- subscription setup and teardown

### Response families

- ok/ack
- payload responses for reads
- diff preview payloads
- structured errors
- capability negotiation

### Event families

- subscription updates
- sync state
- mutation commit notifications
- session lifecycle notifications
- remote-execution progress where relevant

## Recommended tool surfaces

The protocol should support stable agent-facing capabilities such as:

- `readViewport`
- `readRange`
- `readSchema`
- `explainCell`
- `traceImpact`
- `proposeBatch`
- `applyBatch`
- `watchRange`
- `exportSnapshot`

The point is to give agents semantic workbook access rather than UI scraping surfaces.

## Safety model

- mutation proposals should be previewable before commit
- commits should expose stable IDs and rollback hooks where appropriate
- remote requests should be authenticated and tenant-scoped
- replay and idempotency rules should be explicit for mutating calls

## Transport model

### Stdio

- low-overhead local control
- identical schema to network transport

### Remote

- HTTP for request/response ingress where needed
- websocket or equivalent stream transport for events and subscriptions

The schema should be shared even when transport semantics differ.

## Compatibility strategy

1. keep the current binary envelope and typed TS unions working
2. introduce a binary schema versioned alongside the current protocol
3. dual-run encoding/decoding paths during migration where practical
4. retire JSON payload bodies only after parity and tooling coverage are proven

## Package responsibilities

- `@bilig/agent-api`
  - schema definitions, codecs, versioning, conformance fixtures
- `apps/local-server`
  - live worksheet execution and session authority
- `apps/sync-server`
  - authenticated remote ingress and long-lived remote agent execution path
- CLI and scripts
  - protocol consumers, not bespoke protocol definitions

## Migration order

1. define binary schema families and versioning
2. add codec layer alongside current JSON payload path
3. migrate stdio
4. migrate local network ingress
5. migrate remote ingress and streaming subscriptions
6. add richer diff-preview and watch surfaces

## Suggested PR breakdown

1. schema and codec introduction
2. stdio parity path
3. local server parity path
4. remote ingress parity path
5. subscription and streaming event rollout
6. diff-preview and safe-mutation surface expansion

## Exit gate

- agent payloads are typed binary end to end rather than JSON inside a binary envelope
- stdio and remote transports produce identical semantic results for shared operations
- agents can read, explain, preview, commit, and subscribe to worksheet changes through protocol surfaces alone
- the protocol supports long-lived subscriptions and safe mutation flows without UI scraping
