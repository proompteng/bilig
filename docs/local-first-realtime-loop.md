# Local-First Realtime Loop

## Current state

- the repo now has all three runtime building blocks: browser shell, local agent server, and remote sync server.
- the local agent server can already host live workbook sessions and emit committed binary batch frames over websocket.
- the browser worker and local-server loop are not yet unified into the final default boot path.

## Target state

1. browser restores local snapshot and queue
2. browser initializes worker and WASM
3. browser connects to the local agent server over localhost websocket
4. browser catches up by cursor and subscribes viewport/ranges
5. local user and local agent mutations commit through one ordered workbook stream
6. committed mutations render immediately in the browser
7. committed mutations relay upstream to the remote sync backend for durability and cross-device fanout

## Exit gate

- local edit and local agent loops both update the browser through the same binary packet family
- cursor catch-up is proven after reconnect
- remote replay converges with the same local commit stream without semantic drift
