# Binary Protocol

## Current state

- the typed sync codec exists in code and has roundtrip coverage
- sync traffic uses real binary frame encoding
- the remote agent interface is still transitioning from JSON-in-binary-envelope to fully typed binary frames

## Purpose

The binary protocol is the canonical hot-path transport for:

- browser sync
- backend relay
- snapshot chunking
- remote agent frames

## Versioning

- protocol magic: fixed binary header
- protocol version: `1`
- incompatible changes require a version bump

## Current sync frame families

- `hello`
- `appendBatch`
- `ack`
- `snapshotChunk`
- `cursorWatermark`
- `heartbeat`
- `error`

## Encoding model

- little-endian framing
- typed binary header
- explicit string and byte lengths
- explicit op tags for CRDT batches

## Target state

- browser sync, backend relay, snapshot transport, stdio agent traffic, and remote agent traffic all use typed binary frames
- websocket fanout and remote agent control run on the same canonical protocol family

## Exit gate

- no canonical hot path still relies on JSON payload bodies
- browser, backend, stdio, and remote agent transports interoperate on the same typed frame definitions
- roundtrip and malformed-frame tests cover every public frame family
