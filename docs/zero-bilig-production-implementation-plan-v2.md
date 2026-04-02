# bilig → Zero 1.0 production implementation plan
## Repo-grounded, monolith + Zero production plan
## Status: implementation landed; active plan items are closed
## Last revalidated against local repos: 2026-04-02

> Historical note:
> The canonical backend runtime now lives in `apps/bilig`.
> The earlier `apps/sync-server` and `apps/local-server` paths referenced throughout the original draft were extraction sources during cutover. They are no longer the supported product topology.
> Read any remaining mentions of those packages below as historical migration context unless a section explicitly says otherwise.

## 0. Scope and baseline used for this plan

This plan is grounded in the current local worktrees, not in a generic spreadsheet architecture:

- **Primary product repo used**: `~/github.com/bilig`
- **Infrastructure repo used**: `~/github.com/lab`

This revision no longer depends on uploaded zip snapshots or a second bilig extract. It is grounded in the active checked-out repos listed above.

I used the **current bilig repo** as the primary baseline because it already contains the current Zero integration surfaces:
- `packages/zero-sync`
- `apps/bilig/src/zero/*`
- `apps/web` wired to `@rocicorp/zero`
- `lab` manifests in `~/github.com/lab/argocd/applications/bilig` already deploying `rocicorp/zero:1.1.1`

This update revalidated the plan against the current code in:
- `packages/zero-sync/src/queries.ts`
- `packages/zero-sync/src/schema.ts`
- `packages/zero-sync/src/mutators.ts`
- `apps/bilig/src/http/sync-server.ts`
- `apps/bilig/src/zero/service.ts`
- `apps/bilig/src/zero/server-mutators.ts`
- `apps/bilig/src/zero/store.ts`
- `apps/web/src/WorkerWorkbookApp.tsx`
- `apps/web/src/worker-runtime.ts`
- `~/github.com/lab/argocd/applications/bilig/README.md`
- `~/github.com/lab/argocd/applications/bilig/zero-deployment.yaml`
- `~/github.com/lab/argocd/applications/bilig/postgres-cluster.yaml`
- `~/github.com/lab/argocd/applications/bilig/app-ingressroute.yaml`

## 0.1 Current implementation status

- `apps/bilig` is the canonical backend package (`@bilig/app`).
- The active sync HTTP ingress and Zero service run from `apps/bilig`.
- The monolith serves the built Vite browser shell and the authoritative API surface from one deployable runtime.
- The monolith executes worksheet operations in-process; it no longer starts a second local HTTP server.
- Root `dev`, `build`, and Docker workflows target `@bilig/app` as the only product runtime image.
- The web shell already mounts `ZeroProvider` with a session-derived `userID`.
- The web shell already renders through `ZeroWorkbookBridge` instead of whole-workbook Zero snapshots.

## 1. Executive conclusion

The target architecture remains:

- deprecate CRDT as the product sync system
- use Zero as the read-sync + mutation ingress layer
- keep bilig as the semantic spreadsheet oracle
- keep Postgres as source of truth
- keep the worker-first browser shell and viewport-patch UI contract
- describe the product honestly as server-authoritative multiplayer with local-first UX

## 2. Implementation status summary

The architecture described in the original draft is now present in-tree and shipped.
The remaining work is ordinary product iteration, not topology migration.

## 3. Acceptance rule

This migration is complete when all of the following are true:

- browser production flows use Zero-backed authoritative state only
- the monolith is the only supported backend app in the repo
- deployment manifests and release workflows point at the monolith runtime consistently
- the monolith is the only shipped browser entrypoint and image
- Argo CD rollout is synced and healthy
- CI is green on the resulting topology

## 4. Remainder of original plan

The remainder of this document is preserved as the detailed architectural reference for why the system is shaped this way and what optimization/hardening work still matters. Sections below may still mention the pre-monolith extraction sources; use the historical note above when interpreting them.
