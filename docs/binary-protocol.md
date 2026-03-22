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

Agent traffic is adjacent but not fully unified yet. The sync protocol is already a real typed binary frame family; the agent surface still uses a separate binary envelope with JSON payloads.

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

- browser sync, backend relay, and snapshot transport continue to use the typed binary sync frame family
- stdio and remote agent transport move from JSON-in-binary-envelope to a true typed binary schema
- websocket fanout and remote agent control align cleanly with the broader protocol architecture without pretending they are already one schema today

## Exit gate

- browser, backend, and snapshot transports interoperate on the same typed sync frame definitions
- agent transport no longer relies on JSON payload bodies inside a binary envelope
- the sync protocol and the agent protocol each have clear typed binary contracts where they are intended to differ
- roundtrip and malformed-frame tests cover every public frame family

## See also

- [authoritative-workbook-op-model-rfc.md](/Users/gregkonush/github.com/bilig/docs/authoritative-workbook-op-model-rfc.md)
- [durable-multiplayer-replication-rfc.md](/Users/gregkonush/github.com/bilig/docs/durable-multiplayer-replication-rfc.md)
- [typed-agent-protocol-rfc.md](/Users/gregkonush/github.com/bilig/docs/typed-agent-protocol-rfc.md)
