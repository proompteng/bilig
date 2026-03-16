# Binary Protocol

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

## Current tranche

The repo now contains a real typed binary codec package and roundtrip tests. Websocket transport and server fanout are follow-up tranches on top of this wire contract.
