# Workbook Renderer V3 Atlas Validated Design

Primary source: ChatGPT Atlas thread `https://chatgpt.com/c/69ee5ede-d4f4-83e8-9f52-c1fbc2a99f6a`

Follow-up source: pasted Oracle response titled `Renderer V3 Architecture and Migration Plan`

Capture date: 2026-04-26

Status: validated against the current `main` checkout plus the local renderer-v3 implementation tranche in this working tree.

## 1. Validation Verdict

The Atlas response is directionally correct: the mounted grid path is no longer a legacy DOM/canvas renderer, but it is still a resident scene-packet TypeGPU renderer rather than a renderer-native tile runtime.

The current hot path still has these validated properties:

- `packages/grid/src/useWorkbookGridRenderState.ts` is too large and owns render residency, subscriptions, pane construction, headers, overlay geometry, scroll restoration, and React-facing state in one 1400+ line hook.
- `apps/web/src/projected-scene-store.ts` retains resident scene packets keyed by viewport/pane request, not durable tile handles.
- `packages/grid/src/gridResidentDataLayer.ts` still builds pane scenes through `buildGridGpuScene()` and `buildGridTextScene()`.
- `packages/grid/src/renderer-v2/tile-gpu-cache.ts` no longer uses array materialization plus `toSorted()` for stale lookup and eviction after this implementation tranche, but it is still not a true byte-budgeted tile residency system with numeric keys, compatibility buckets, and O(1) visible marking.
- `packages/grid/src/renderer-v2/workbook-typegpu-backend.ts` still resolves draw panes by consulting that cache each frame.
- `packages/grid/src/renderer-v2/typegpu-atlas-manager.ts` still uses one growing atlas canvas and redraws all glyphs on growth.
- `packages/grid/src/renderer-v2/dynamic-overlay-packet.ts` still packs interaction overlays as `GridScenePacketV2`; the visible header-index helper no longer sorts arrays after this tranche, but overlays are still data-scene-shaped packets rather than a dedicated overlay instance layer.
- Browser perf tests now prove several resident-scene wins, but they still allow bounded scene refreshes at tile-boundary/collaboration cases rather than enforcing a pure tile-delta runtime.

Important corrections to the Atlas text:

- `packages/grid/src/renderer-v2/typegpu-buffer-pool.ts` no longer hashes every text and rect item directly during resource sync. It now compares packet-level signatures.
- The hashing work still exists during packet packing in `packages/grid/src/renderer-v2/scene-packet-v2.ts`, where `resolveRectSignature()` and `resolveTextSignature()` walk the full rect/text payload.
- `scene-packet-v2.ts` already has a tile-shaped key, but it is still serialized to a string for the cache and still wraps viewport-sized resident pane packets. It is not yet a true binary tile delta contract.
- The practical next step is not to delete the whole path immediately. The correct migration is measured replacement: add counters, remove hot-path allocation/sort work, then introduce renderer-native deltas behind tested contracts.
- The follow-up Oracle response is correct that `useWorkbookGridRenderState.ts` must stop being the renderer owner, `ViewportPatch` must not remain the renderer contract, content tile identity must split from pane placement, and overlays need their own runtime. It is stale where it says the current `tile-gpu-cache.ts` still uses `filter().toSorted()` and where it implies `typegpu-buffer-pool.ts` hashes every item during resource sync.

## 2. Target Architecture

The destination is renderer-v3:

- camera-driven scroll, with the renderer receiving direct camera updates outside React state churn;
- renderer-native tile deltas, separate from app-state viewport patches;
- numeric tile identity and O(1) residency transitions;
- GPU arenas and dirty-span uploads instead of full pane resource sync;
- independent header and interaction overlay layers;
- atlas dependency tracking with no active-scroll repack;
- scene packet runtime deleted from normal scroll/edit/select paths once tile payloads are proven.

The scene packet path may remain temporarily as a correctness oracle and cold bootstrap.

## 3. Phase Plan

### Phase 1: make current hot path measurable

Files:

- `packages/grid/src/renderer-v2/grid-render-contract.ts`
- `packages/grid/src/renderer-v2/grid-render-counters.ts`
- `apps/web/src/perf/workbook-scroll-perf.ts`
- `packages/grid/src/renderer-v2/tile-gpu-cache.ts`
- `packages/grid/src/renderer-v2/typegpu-buffer-pool.ts`
- `apps/web/src/projected-tile-scene-store.ts`
- `e2e/tests/web-shell-scroll-performance.pw.ts`

Implement:

- counters for tile cache stale lookups, scanned entries, sort calls, visible markings, evictions, and buffer signature checks;
- counters for renderer delta batches, delta mutations, delta apply time, and dirty tile count;
- gates proving steady resident-window scroll does not allocate/sort/upload;
- diagnostics that distinguish exact hits, stale-compatible hits, and visible tile misses.

### Phase 2: remove sorted/string-allocation cache hot paths

Files:

- `packages/grid/src/renderer-v2/tile-gpu-cache.ts`
- `packages/grid/src/renderer-v2/workbook-typegpu-backend.ts`
- `packages/grid/src/__tests__/tile-gpu-cache.test.ts`
- `packages/grid/src/__tests__/workbook-typegpu-backend.test.ts`

Implement:

- replace `toSorted()` stale lookup with a direct newest-compatible scan;
- replace sort-based eviction with repeated oldest-non-visible selection until the cache is within budget;
- keep string cache keys only at the API boundary for this tranche, but prevent array/sort allocation during draw;
- expose counters so the browser perf harness can enforce `tileCacheSorts = 0`.

### Phase 3: split camera and render loop out of React

Files:

- `packages/grid/src/useWorkbookGridRenderState.ts`
- `packages/grid/src/WorkbookGridSurface.tsx`
- `packages/grid/src/renderer-v2/WorkbookPaneRendererV2.tsx`
- new `packages/grid/src/gridViewportController.ts`
- new `packages/grid/src/renderer-v2/render-scheduler.ts`

Implement:

- scroll updates write a camera store directly;
- renderer subscribes to camera snapshots;
- React state updates only when the resident tile window changes;
- editor and overlay anchors subscribe to the camera store instead of relying on re-rendered hook state.

### Phase 4: introduce renderer tile delta transport

Files:

- new `packages/worker-transport/src/render-tile-delta.ts`
- `packages/worker-transport/src/index.ts`
- `apps/web/src/worker-runtime-viewport-publisher.ts`
- `apps/web/src/projected-viewport-store.ts`
- new `apps/web/src/projected-tile-scene-store.ts`

Implement:

- binary `RenderTileDeltaBatch` with tile identity, version, mutation kind, and dirty spans;
- cell-run, axis, freeze, invalidate, and overlay mutation types;
- validation tests for version mismatches and bounded payloads;
- keep `ViewportPatch` as app-state transport, not renderer transport.

### Phase 5: replace routine scene packet regeneration

Files:

- `packages/grid/src/gridResidentDataLayer.ts`
- `apps/web/src/worker-runtime-render-scene.ts`
- `apps/web/src/projected-scene-store.ts`
- `packages/grid/src/renderer-v2/scene-packet-v2.ts`
- new `packages/grid/src/renderer-v2/tile-payload-builder.ts`

Implement:

- initial tile payload builder checked against current scene builders;
- delta apply into CPU tile payloads;
- dirty spans for rect/text/glyph buffers;
- full tile replacement only for cold bootstrap, sheet switch, or structural axis version change.

### Phase 6: independent header and overlay layers

Files:

- `packages/grid/src/renderer-v2/dynamic-overlay-packet.ts`
- `packages/grid/src/renderer-v2/WorkbookPaneRendererV2.tsx`
- new `packages/grid/src/renderer-v2/interaction-overlay-layer.ts`
- new `packages/grid/src/renderer-v2/header-tile-layer.ts`

Implement:

- headers as their own tile family;
- selection/fill/resize/collaboration overlay buffer independent from data tiles;
- selection drag and resize preview with zero data tile uploads.
- in the interim V2 packet path, visible header-index generation must avoid sort allocation while it is still packet-based.

### Phase 6.5: establish renderer-v3 tile primitives

Files:

- new `packages/grid/src/renderer-v3/tile-key.ts`
- new `packages/grid/src/renderer-v3/tile-residency.ts`
- new `packages/grid/src/renderer-v3/tile-damage-index.ts`
- `apps/web/src/worker-runtime-render-tile-delta.ts`
- `apps/web/src/worker-viewport-tile-store.ts`
- new `packages/worker-transport/src/workbook-delta-v3.ts`
- new `packages/worker-transport/src/tile-interest-v3.ts`
- new `packages/grid/src/runtime/gridAxisRuntime.ts`
- new `packages/grid/src/runtime/gridCameraRuntime.ts`
- new `packages/grid/src/renderer-v3/overlay-layer.ts`
- new `packages/grid/src/runtime/gridOverlayRuntime.ts`
- new `packages/grid/src/renderer-v3/gpu-buffer-arena.ts`
- new `packages/grid/src/runtime/gridRuntimeHost.ts`

Implement:

- reversible numeric `TileKey53` content tile keys using protocol tile dimensions;
- renderer-v3 dirty tile marking by fixed tile, not viewport subscription;
- residency exact lookup, compatibility buckets, LRU touch/eviction, and generation-based visible marking;
- bootstrap render-delta tile IDs generated from shared numeric content tile keys instead of ad hoc hashes;
- render-delta transport preserving the full safe-integer tile-key range;
- render-delta subscriptions materializing fixed 32x128 content tiles instead of resident pane windows.
- worker local viewport tile cache using numeric tile residency instead of serialized string keys and scan-based eviction.
- binary V3 contracts for sheet-level dirty range deltas and renderer tile interest batches.
- runtime-owned axis query primitive for offsets, spans, tile origins, visible ranges, and monotonic axis revisions.
- runtime-owned camera primitive for visible region computation and fixed visible tile-key derivation.
- dynamic overlay runtime and packed instance-buffer contract independent from data scene packets.
- backend-agnostic GPU buffer arena primitive with capacity-class free lists and explicit trim destruction.
- imperative runtime host that composes axis, camera, overlay, and visible tile-key state for the future React shell adapter.
- dirty tile index application of V3 workbook delta batches, including bounded axis dirty ranges that are consumed per visible tile.

### Phase 7: text atlas service

Files:

- `packages/grid/src/renderer-v2/typegpu-atlas-manager.ts`
- `packages/grid/src/renderer-v2/line-text-quad-buffer.ts`
- new `packages/grid/src/renderer-v2/text-atlas-service.ts`

Implement:

- glyph dependency table per tile;
- atlas dirty-rect uploads and idle repack only;
- visible/warm dependency preservation during active scroll;
- no visible tile blanking on rare glyph allocation.

### Phase 8: delete scene-packet runtime use

Delete from normal runtime path:

- `apps/web/src/projected-scene-store.ts`
- `apps/web/src/worker-runtime-render-scene.ts`
- `packages/grid/src/gridResidentDataLayer.ts` as renderer data runtime

Keep only test-oracle builders until tile payload v3 reaches parity.

## 4. Current Execution Tranche

This document starts implementation with phases 1, 2, a narrow tested slice of phase 3, and the transport/projection-contract slice of phase 4 because they are validated, low risk, and unblock stronger gates for the later rewrite.

Definition of done for this tranche:

- no `toSorted()` or array materialization in `TileGpuCache.findStaleValid()` or `TileGpuCache.evictTo()`;
- scroll perf reports include tile cache lookup/sort/eviction counters;
- `packages/worker-transport/src/render-tile-delta.ts` exists as the renderer-native binary delta contract;
- worker transport can relay `renderTileDeltas` independently from `viewportPatches`;
- app-side projection can subscribe to, decode, store, and invalidate render tile deltas;
- the worker runtime can publish full tile-replace render deltas from fixed content tile viewports while tile-payload v3 is developed;
- scroll perf reports include renderer delta batch, mutation, apply-time, and dirty-tile counters;
- viewport-window and render-scroll math is split out of `useWorkbookGridRenderState.ts` under focused tests;
- dynamic overlay visible header-index generation avoids the old `Set` plus `toSorted()` allocation path;
- renderer-v3 numeric tile keys, dirty tile index, and tile residency primitives exist under focused tests;
- bootstrap render-delta tile IDs are shared content tile keys, not ad hoc per-pane hashes;
- existing tile cache and typegpu backend tests pass;
- render tile delta codec and worker relay tests pass;
- focused browser scroll performance still passes;
- full CI can be run after the tranche if the diff remains stable.

Completed in the first implementation tranche:

- `TileGpuCache.findStaleValid()` now does a direct newest-compatible scan instead of materializing and sorting cache entries.
- `TileGpuCache.evictTo()` now selects non-visible LRU entries without `toSorted()`.
- TypeGPU scroll perf counters include tile cache stale lookups, scanned entries, stale hits, visible marks, evictions, and sort calls.
- Browser steady-scroll perf now asserts `typeGpuTileCacheSorts === 0`.
- `RenderTileDeltaBatch` is encoded/decoded as binary and exported through `@bilig/worker-transport`.
- `createWorkerEngineHost()` and `createWorkerEngineClient()` relay `renderTileDeltas` with subscription args.
- `ProjectedTileSceneStore` applies decoded render tile deltas, ignores stale batches, and invalidates sheet tiles on structural axis/freeze batches.
- `ProjectedViewportStore` exposes projected render tile subscriptions alongside existing viewport and resident-scene subscriptions.
- `WorkbookWorkerRuntime.subscribeRenderTileDeltas()` publishes binary tile deltas using the current resident-scene builder as the bootstrap oracle.
- `gridViewportController.ts` owns resident-window comparison, tile-key conversion, and render-scroll transform math outside the large React hook.
- `ProjectedTileSceneStore` reports renderer delta batches, mutations, apply duration, and dirty tile counts into the scroll perf collector.
- `dynamic-overlay-packet.ts` no longer sorts visible row/column header indexes while building overlay packets.
- `renderer-v3/tile-key.ts` packs/unpacks fixed content tile coordinates into safe numeric keys.
- `renderer-v3/tile-damage-index.ts` maps cell-range damage to touched fixed content tiles.
- `renderer-v3/tile-residency.ts` provides exact lookup, compatibility-bucket stale lookup, generation-based visible marking, pinning, and byte-budget eviction primitives.
- `worker-runtime-render-tile-delta.ts` now emits tile IDs from `packTileKey53()` so body/frozen pane bootstrap deltas share content tile identity.
- `packages/worker-transport/src/render-tile-delta.ts` now round-trips 53-bit tile keys instead of truncating tile IDs to `u32`.
- `worker-runtime-render-tile-delta.ts` now splits render subscriptions into protocol-sized fixed content tile replacements, avoiding frozen/body pane duplicate materialization in the render-delta path.
- `worker-viewport-tile-store.ts` now keys cached projection tiles with `packTileKey53()` and evicts through `TileResidencyV3` rather than scanning serialized string-key maps.
- `packages/worker-transport/src/workbook-delta-v3.ts` defines and round-trips the first sheet-level dirty-range delta batch.
- `packages/worker-transport/src/tile-interest-v3.ts` defines and round-trips visible/warm/pinned tile interest batches using safe-integer tile keys.
- `packages/grid/src/runtime/gridAxisRuntime.ts` starts the axis runtime split with update-owned prefix indexes and tile-origin queries outside React render state.
- `packages/grid/src/runtime/gridCameraRuntime.ts` starts the camera runtime split with scroll-to-visible-region math and content tile-interest key derivation outside React render state.
- `packages/grid/src/renderer-v3/overlay-layer.ts` and `packages/grid/src/runtime/gridOverlayRuntime.ts` add small packed overlay batches for selection/resize/hover/presence-style visuals without data tile invalidation.
- `packages/grid/src/renderer-v3/gpu-buffer-arena.ts` adds a reusable buffer arena contract for V3 GPU resources so normal eviction can release to free lists instead of destroying buffers.
- `packages/grid/src/runtime/gridRuntimeHost.ts` composes the first V3 runtimes behind an imperative host API that React can eventually mount and dispose instead of coordinating renderer internals.
- `packages/grid/src/renderer-v3/tile-damage-index.ts` now applies sheet-level V3 dirty range batches to fixed tile damage and keeps axis dirty ranges bounded by tile rows/columns instead of expanding them over the full sheet.

Remaining work from this design:

- continue splitting camera/render scheduling out of `useWorkbookGridRenderState.ts`;
- replace routine resident scene packet regeneration with dirty-span tile payload updates;
- split headers/overlays into independent GPU layers;
- replace the atlas manager with dependency-aware dirty uploads;
- remove scene-packet runtime use after parity and perf gates are green.
