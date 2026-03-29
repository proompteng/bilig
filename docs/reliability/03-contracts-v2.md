# Contracts v2

## Scope

`@bilig/contracts` defines all `v2` payloads used at product control-plane boundaries:

- `/v2/session`
- `/v2/documents/:documentId/state`
- `/v2/documents/:documentId/snapshot/latest`
- `/v2/documents/:documentId/frames`
- `/v2/documents/:documentId/ws`
- `/v2/agent/frames`
- `/api/zero/v2/query`
- `/api/zero/v2/mutate`

## Rules

- all payloads are Effect Schema definitions
- all success and failure payloads are exported from the package
- no route defines its own runtime validation inline
- all route handlers decode request input and encode response output through schemas

## Error Envelope

Every `v2` HTTP error returns:

- `error` tagged code
- `message`
- `retryable`
- optional `details`

Websocket and worker control-plane failures use the same tagged error vocabulary, adapted to their transport.

## Current First-Wave Contracts

First implementation wave covers:

- runtime session response
- document state summary
- snapshot response metadata
- generic `v2` error envelope

Subsequent waves add Zero query/mutator argument/result schemas and worker transport RPC schemas.
