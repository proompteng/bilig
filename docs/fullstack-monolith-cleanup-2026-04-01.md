# Fullstack monolith cleanup checklist
## Date: 2026-04-01
## Scope: make `apps/bilig` the single shipped product runtime

## Goals

- serve the built Vite app from `apps/bilig`
- remove the separate web runtime image and deployment path
- keep local worksheet execution support inside the monolith process
- publish one Forgejo image for the product runtime
- align the lab GitOps app to one product deployment
- rerun CI and verify live Argo CD and Kubernetes state truthfully

## Execution checklist

- [x] add Fastify static and proxy support to `@bilig/app`
- [x] serve `apps/web/dist` from the monolith and emit same-origin `runtime-config.json`
- [x] remove the nginx-based `web-runtime` image from local build and release flow
- [x] collapse local compose to `bilig-app + zero-cache + postgres`
- [x] remove `bilig-web` manifests from the lab repo and replace `bilig-sync` with `bilig-app`
- [x] update current docs and operator guidance to the monolith topology
- [x] rerun the full validation chain
- [x] verify Argo CD and Kubernetes status against the updated topology, including the current live mismatch while local lab changes remain unpublished

## Done definition

The cleanup is complete when the repo ships one app image, the lab repo deploys one app workload, browser traffic lands on the monolith, Zero still converges correctly, and the substantive CI steps are green.
