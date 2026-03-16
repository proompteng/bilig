# Public APIs

## Stable packages

- `@bilig/core`
- `@bilig/formula`
- `@bilig/wasm-kernel`
- `@bilig/crdt`
- `@bilig/renderer`
- `@bilig/grid`
- `@bilig/binary-protocol`
- `@bilig/worker-transport`
- `@bilig/agent-api`
- `@bilig/storage-browser`
- `@bilig/storage-server`
- `@bilig/excel-fixtures`

## Workbook DSL

`@bilig/renderer` keeps the declarative workbook DSL unchanged:

- `<Workbook>`
- `<Sheet name="...">`
- `<Cell addr="..." value={...} />`
- `<Cell addr="..." formula="..." />`
- `<Cell addr="..." format="..." />`

## Core engine surface

The canonical engine surface includes:

- `createSheet`
- `deleteSheet`
- `setCellValue`
- `setCellFormula`
- `setCellFormat`
- `clearCell`
- `setRangeValues`
- `setRangeFormulas`
- `clearRange`
- `fillRange`
- `copyRange`
- `pasteRange`
- `setSelection`
- `undo`
- `redo`
- `getCell`
- `getDependencies`
- `getDependents`
- `explainCell`
- `exportSnapshot`
- `importSnapshot`
- `exportReplicaSnapshot`
- `importReplicaSnapshot`
- `applyRemoteBatch`
- `subscribe`
- `subscribeBatches`
- `connectSyncClient`
- `disconnectSyncClient`
- `getSyncState`

## Binary protocol

`@bilig/binary-protocol` exposes:

- `PROTOCOL_VERSION`
- `encodeFrame(frame): Uint8Array`
- `decodeFrame(bytes): ProtocolFrame`

The current frame families are:

- `hello`
- `appendBatch`
- `ack`
- `snapshotChunk`
- `cursorWatermark`
- `heartbeat`
- `error`

## Worker transport

`@bilig/worker-transport` exposes:

- `createWorkerEngineHost(engine, port)`
- `createWorkerEngineClient({ port })`

The first tranche already supports:

- method invocation
- engine events
- outbound batch subscriptions

## Agent API

`@bilig/agent-api` exposes:

- `AgentRequest`
- `AgentResponse`
- `AgentEvent`
- `encodeAgentFrame`
- `decodeAgentFrame`
- `encodeStdioMessage`
- `decodeStdioMessages`

The canonical transport surface covers both stdio and remote network usage with the same typed request/response/event model.
