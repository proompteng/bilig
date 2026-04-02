# Production Stability Remediation Plan

## Goal

Stabilize the fullstack `bilig` monolith in production without adding temporary hacks. The target state is:

- no origin `502` during normal editing or deployment
- no app pod OOM restarts under low-volume interactive use
- immediate workbook paint before Zero bridge convergence
- no steady-state per-selection Zero point-query churn
- durable workbook recovery from checkpoint plus event replay instead of hot-path full-state persistence

## Root cause

The production outage on `2026-04-02` had two direct causes:

1. `bilig-app` was exporting full workbook snapshots and full replica snapshots on request-path mutations.
2. The browser selection bridge was issuing new `cellInput.one` and `cellEval.one` materializations on each selected-cell change.

The request-path snapshot churn blew the Node heap. Once V8 entered long GC / OOM failure, the 1-second probes caused both pods to fall out of service and Cloudflare returned `502`.

## Design

### 1. Durable state recovery

- stop writing hot-path full snapshots into the `workbooks` row on every mutation
- keep event journaling authoritative
- keep periodic checkpoints in `workbook_snapshot`
- on cold load, rebuild state from latest checkpoint plus ordered `workbook_event` replay

This removes the need to persist giant snapshots on each edit while preserving durable recovery semantics.

### 2. Request-path memory reduction

- remove request-path replica snapshot export from mutations
- only export checkpoint snapshots / replica snapshots in the async recalc path when the revision is checkpoint-worthy

### 3. Frontend hydration

- paint the workbook shell as soon as the worker cache exists
- do not block visible grid mount on bridge readiness
- treat viewport cache as the primary selected-cell source
- only fall back to targeted Zero `.one(...)` selection queries when the selected cell is outside the hydrated cache

### 4. Production availability

- add `startupProbe`
- increase readiness / liveness time budgets beyond GC-pause scale
- enforce a safer rollout strategy and PodDisruptionBudget so one failure cannot drop the origin

### 5. Verification

- isolated-document Playwright production smoke via `?document=...`
- local compose/browser regression validation
- full repo CI
- live Argo CD + `kubectl` verification after rollout

## Execution checklist

- [x] add checkpoint-plus-event replay support to workbook runtime loading
- [x] stop hot-path `workbooks.snapshot` / `workbooks.replica_snapshot` writes on mutation
- [x] stop hot-path replica snapshot export in mutation handling
- [x] checkpoint snapshots only in async recalc when the revision warrants it
- [x] paint workbook UI before bridge readiness
- [x] remove steady-state per-selection Zero materialization for in-viewport cells
- [x] add isolated-document production Playwright smoke
- [x] harden app deployment probes / rollout safety
- [ ] run local validation
- [ ] publish repos
- [ ] sync Argo CD
- [ ] verify live smoke and stability

## Exit criteria

- `pnpm run ci` passes
- `pnpm test:browser:prod` passes against production
- `kubectl -n bilig get pods` shows no restart loop after rollout
- `argocd app get bilig` is `Synced` and `Healthy`
- public host returns `200` and interactive workbook smoke passes on an isolated document
