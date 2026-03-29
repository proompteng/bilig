# Runtime Kernel

## Responsibility

`@bilig/runtime-kernel` owns the Effect-side infrastructure used by every application surface:

- typed error algebra
- runtime entrypoints
- config loading
- structured logging
- tracing/span helpers
- resource acquisition and cleanup
- injected clock, random, ID, fetch, and websocket services

## Error Model

All expected failures are tagged domain failures:

- `TransportError`
- `HttpError`
- `DecodeError`
- `ValidationError`
- `SessionError`
- `SnapshotError`
- `DocumentStateError`
- `AgentRequestError`

Rules:

- expected failures return tagged errors in the Effect error channel
- route handlers map tagged errors to schema-defined error envelopes
- thrown generic `Error` is allowed only inside low-level adapters and must be normalized immediately

## Service Rules

- global `fetch` is wrapped behind an injected service
- websocket construction is wrapped behind an injected service
- runtime config is loaded through Effect config/services, not ambient reads in application code
- retry policies are explicit and owned by the caller
- resource lifetimes use scoped acquisition and release

## Runtime Entrypoints

Expose a small stable surface:

- `runPromise`
- `runPromiseExit`
- `withSpan`
- `decodeWithSchema`
- `encodeWithSchema`
- live layers for browser and node fetch services

Do not export a catch-all bag of Effect re-exports from this package.
