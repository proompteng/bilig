# Renderer Oracle Capture And Execution Plan

Captured with Browser Use from <https://chatgpt.com/c/69ed3f9f-89c8-83e8-a63a-f7ad769f6aec>.

## Capture Result

The current ChatGPT Atlas conversation is titled "Oracle-level performance audit", but the extracted answer is not a `bilig` workbook renderer audit. It is a 44,747 character audit of `bilig2` WorkPaper formula-engine performance. The answer says it inspected an uploaded zip under `/mnt/data/bilig2_audit`, used benchmark evidence from `pnpm bench:workpaper:competitive`, and focused on lookup ownership, formula-family execution, dirty scheduling, range aggregation, runtime restore, and WASM upload behavior.

The captured response headings were:

- What changed since the old structural-remap diagnosis
- Top 5 bottlenecks now
- Architecture target for SOTA WorkPaper formula-engine performance
- Implementation sequence
- Tranche 0: add counters before changing behavior
- Tranche 1: split exact and approximate lookup owner summaries
- Tranche 2: coalesce batch mutation observers by owner, not by cell
- Tranche 3: make formula families executable build owners
- Tranche 4: add family-compressed dependency graph and dirty spans
- Tranche 5: patch calc chain after dynamic topo repair
- Tranche 6: replace sliding aggregate suffix-prefix updates with column aggregate state
- Tranche 7: make runtime restore genuinely warm
- Tranche 8: add WASM delta program/range upload after JS ownership is fixed
- Tranche 9: later replace AxisMap with a real piece table/B-tree
- Bench gates
- Do not do
- Open questions

Representative captured opening excerpt:

> I inspected the uploaded zip as authoritative. Line anchors below refer to the extracted source under /mnt/data/bilig2_audit. I did not rerun pnpm bench:workpaper:competitive; benchmark evidence below uses your latest local numbers, while code conclusions come from the zip.

Representative captured closing themes:

- Add counters before behavior changes.
- Confirm benchmark/product semantics before treating restore as a warm runtime-image path.
- Keep AssemblyScript/WASM as a closed fast path after JS parity and differential tests.
- Do not treat older structural-remap guidance as the current top formula-engine bottleneck.

## Applicability To This Repo

That response is useful as a discipline reminder, but not as renderer implementation truth for this checkout. It targets formula-engine internals in `bilig2`; the active work here is the `bilig` TypeGPU workbook renderer under `packages/grid/src/renderer-v2`, the resident-scene pipeline in `apps/web/src/projected-scene-store.ts`, and browser scroll gates in `e2e/tests/web-shell-scroll-performance.pw.ts`.

Accepted principles:

- Add or keep counters and tests before claiming performance wins.
- Make the hot-path data contract authoritative and packet/revision based.
- Remove compatibility paths once the migration is complete.
- Prefer structural fixes over benchmark-specific thresholds.

Rejected as direct renderer guidance:

- Lookup owner splitting, formula family execution ownership, dirty topology repair, runtime restore, and WASM formula upload work do not address the current TypeGPU renderer review findings.
- The answer's source anchors under `/mnt/data/bilig2_audit` do not map to this checkout.

## Current Renderer Findings Checked Against Code

- Stale tile reuse drawing wrong rows is already fixed in `packages/grid/src/renderer-v2/workbook-typegpu-backend.ts`: stale packets are only substituted when `viewport.rowStart` and `viewport.colStart` match the desired packet, and `packages/grid/src/__tests__/workbook-typegpu-backend.test.ts` covers the different-origin rejection.
- The V2 pane contract is already packed-only in `packages/grid/src/renderer-v2/pane-scene-types.ts`: `WorkbookPaneScenePacket` and `WorkbookRenderPaneState` carry `packedScene`, not object `gpuScene`/`textScene`.
- Resource sync is already packet-signature based in `packages/grid/src/renderer-v2/typegpu-buffer-pool.ts`: text and rect signatures are derived from `GridScenePacketV2` counts/signatures and tile key, not from full object-scene hashing.
- The live hot-path defect was `packages/grid/src/renderer-v2/tile-gpu-cache.ts`: `findStaleValid`, `get`, `markVisible`, and `evictTo` replaced cache entry objects and allocated arrays/sets/sorts in draw-adjacent paths.
- `packages/grid/src/useWorkbookGridRenderState.ts` remains too large and cross-coupled at 1438 lines. That needs a separate controller split, but it is not the first correctness/perf blocker after the stale-origin guard.
- The frozen-pane perf test still has an intentionally loose `longTaskMax: 600` gate. Tightening that should follow measured renderer stability, not precede it.

## Execution Plan

### Phase 1 - Complete In This Change

Refactor `TileGpuCache` into an in-place mutable resident tile cache:

- Keep a stable `TileGpuCacheEntry` object per tile key.
- Update newer packets in place.
- Select stale-valid overlap with a single pass and no array/sort.
- Touch recency by mutating `lastUsedSeq`, not by replacing entries.
- Mark visibility with explicit `beginVisibilityPass()` and `markVisibleKey()` calls so backend draw code does not allocate a `Set`.
- Evict least-recent non-visible tiles with bounded single-pass scans and no sort allocation.
- Add focused tests for entry identity, iterable visibility marking, and most-recent stale selection.

### Phase 2 - Next Renderer Split

Split `useWorkbookGridRenderState.ts` into behavior-owned modules without changing public behavior:

- `renderer-v2/useWorkbookGridCamera.ts`: scroll snapshot, camera sequence, and resident viewport policy.
- `renderer-v2/useWorkbookGridSceneResidency.ts`: visible/prefetch request selection, worker/local scene source, and generation guards.
- `renderer-v2/useWorkbookGridPaneAssembly.ts`: pane frame assembly, frozen/header/body pane composition, and preload pane construction.
- `renderer-v2/useWorkbookGridInteractionState.ts`: selection/edit overlays only.

Acceptance gates:

- Existing grid interaction E2E tests still pass.
- No renderer feature flag or legacy fallback path is reintroduced.
- `useWorkbookGridRenderState.ts` falls under 1000 lines.

### Phase 3 - Tighten Perf Gate After Measurement

Run the frozen-pane scroll perf gate after Phase 1 and Phase 2. If measured steady-scroll long-task max is consistently below the current loose ceiling, lower `longTaskMax` in `e2e/tests/web-shell-scroll-performance.pw.ts` with the perf report committed as evidence. Do not change the gate to hide churn; it must stay stricter, not looser.

### Phase 4 - Packet-Only Resource Upload Follow-Up

Audit remaining object-scene creation points in `scene-packet-v2.ts`, `dynamic-overlay-packet.ts`, and header scene construction. The goal is not to remove object scenes from builders prematurely; the goal is to keep the draw/runtime resource contract packet-only and prevent object identity churn from driving uploads.

