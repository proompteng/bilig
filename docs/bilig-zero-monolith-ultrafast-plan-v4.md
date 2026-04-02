# bilig Zero Monolith Ultrafast Plan v4

## Status

Archived implementation draft. The monolith cutover has landed.

## Landed outcomes

- `apps/bilig` is the monolith runtime.
- `apps/web` is the only shipped browser shell.
- Zero query and mutate ingress are served by the monolith.
- Recalc and checkpoint persistence run inside the monolith runtime.
- The standalone `apps/local-server` and `apps/sync-server` product apps are removed.
- The browser no longer depends on the legacy websocket document-sync transport in the shipped product path.

## Active docs instead of this draft

- [zero-bilig-production-implementation-plan-v2.md](/Users/gregkonush/github.com/bilig/docs/zero-bilig-production-implementation-plan-v2.md)
- [bilig_production_plan_2026-03-30.md](/Users/gregkonush/github.com/bilig/docs/bilig_production_plan_2026-03-30.md)
- [production-stability-remediation-2026-04-02.md](/Users/gregkonush/github.com/bilig/docs/production-stability-remediation-2026-04-02.md)

## Why this file remains

The original draft captured design intent during the cutover. It is preserved as an archive, not as the current execution plan.
