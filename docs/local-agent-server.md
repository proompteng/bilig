# Local Agent Server

## Current state

- `apps/local-server` exists and hosts live local workbook sessions.
- browser clients can connect over binary websocket frames and receive committed batch broadcasts.
- canonical worksheet request/response operations now execute against live local workbook sessions through the local agent API ingress.
- agent chat orchestration is not wired yet; this tranche is worksheet-session first, not chat-to-action complete.

## Target state

- the local agent server is the default authoritative runtime for local workbook sessions and agent-driven work.
- chat messages enter the local agent server, run through Codex CLI plus skills, and commit worksheet mutations through the same ordered transaction stream as UI edits.
- committed mutations emit binary CRDT frames immediately to the frontend and relay upstream to the remote sync backend when connected.

## Exit gate

- local browser sessions connect to the local agent server by default
- chat-driven agent work mutates live workbook sessions and updates the frontend in near realtime
- local crash/restart recovery preserves workbook state, cursors, and pending outbound sync state
