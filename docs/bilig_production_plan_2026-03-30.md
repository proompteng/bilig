# bilig implementation execution tracker

## Purpose

This is the active execution checklist for finishing the monolith + Zero production cutover.
It replaces the earlier speculative draft with a repo-grounded todo list.

## Competitive baselines reviewed

- Microsoft Excel limits/specs and collaboration expectations
- Zero production query/mutate and materialize guidance
- public Google Sheets scale expectations for large collaborative workbooks

Design consequence:

- keep the browser worker-first
- keep Zero tile-shaped and narrow
- keep the server authoritative
- keep CI and rollout boring and deterministic

## Current decisions

- one backend runtime: `apps/bilig`
- one browser shell: `apps/web`
- Zero stays
- Postgres stays
- no separate `apps/local-server` or `apps/sync-server` product apps
- no Redis on the correctness path

## Execution checklist

### Repo cleanup

- [x] make the monolith the only app on the root TypeScript graph
- [x] fix image release workflow to use the actual Docker runtime target
- [x] align browser test compose service names with the monolith compose file
- [x] remove placeholder and duplicate monolith files with no imports
- [ ] rename remaining snapshot-era relational helpers and tables where the migration cost is justified

### App topology

- [x] remove retired `apps/local-server`
- [x] remove retired `apps/sync-server`
- [ ] remove any remaining legacy browser-sync endpoint usage from the product path

### Docs

- [x] rewrite the top-level design docs to reflect the monolith + Zero architecture
- [x] mark the old speculative plan as historical and keep the execution tracker concise
- [ ] finish sweeping deep RFCs that still speak about the removed app topology as current

### Deploy / rollout

- [x] remove Redis from the `lab` app manifests
- [ ] sync the updated `lab` app via Argo CD
- [ ] monitor rollout health with `kubectl` and `argocd`

### Validation

- [ ] run install/update if the lockfile or workspace graph changed
- [ ] run lint
- [ ] run typecheck
- [ ] run tests and browser tests
- [ ] run full `pnpm run ci`

## Exit gate

The cutover is done when:

- repo CI is green
- Forgejo image publishing is aligned with the Dockerfile
- Argo CD is synced and healthy on the updated manifests
- no duplicate backend applications remain in the repo or deployment manifests
