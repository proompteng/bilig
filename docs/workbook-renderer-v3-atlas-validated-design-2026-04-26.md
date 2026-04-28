# Workbook Renderer V3 Atlas Validated Design

Primary source: ChatGPT Atlas thread `https://chatgpt.com/c/69ee5ede-d4f4-83e8-9f52-c1fbc2a99f6a`

Follow-up source: pasted Oracle response titled `Renderer V3 Architecture and Migration Plan`

Capture date: 2026-04-26

Status: validated against the current `main` checkout plus the local renderer-v3 implementation tranche in this working tree.

Latest implementation note:

- Product runtime no longer includes the resident scene store / worker resident scene RPC path. `apps/web/src/projected-scene-store.ts`,
  `apps/web/src/worker-runtime-render-scene.ts`, the resident scene cache, and resident scene packet bridge types were deleted.
- `ProjectedViewportStore` now owns projected cells/axis plus V3 render-tile deltas only; viewport patches no longer invalidate or rebuild
  resident scene packets.
- `useWorkbookGridRenderState.ts` no longer subscribes to renderer viewport patches, worker resident pane scenes, or V2 warm resident scene
  requests. Data panes are always built from fixed render tiles: worker-provided tile deltas in product and a local fixed-tile materializer
  for non-worker fallback/tests.
- `packages/grid/src/gridResidentDataLayer.ts`, `packages/grid/src/gridTileResidencyV2.ts`, `packages/grid/src/useGridViewportSubscriptions.ts`,
  and `packages/grid/src/useGridSceneResidency.ts` were deleted from the grid package product surface.
- Runtime state now carries sheet IDs alongside sheet names so renderer tile subscriptions can use stable numeric content-tile identity.
- `packages/grid` now owns a grid-facing render-tile source contract and V3 tile-pane state. The old `render-tile-v2-adapter.ts` bridge was deleted, so fixed content tiles are no longer repacked into `GridScenePacketV2` before mounting.
- `useWorkbookGridRenderState.ts` subscribes to projected render-tile deltas when a `renderTileSource` and `sheetId` are available, then prefers fixed content tiles once all required tiles for the resident viewport are ready.
- With a fixed render-tile source present, the hook no longer mounts renderer-owned resident scene subscriptions or renderer-owned viewport patch subscriptions.
- With that source present, the hook also bypasses V2 warm resident-viewport planning; tile prefetch needs to move to V3 tile interest batches next.
- Transient fixed-tile misses retain the last fixed-tile pane set instead of immediately falling back to a full resident data-scene rebuild.
- Frozen pane placements reuse the same fixed content tile under separate body/top/left/corner clip placements. V3 tile-pane buffers are keyed by numeric `tileId`, so body and frozen placements share the content resource.
- `GridRenderTilePaneRuntime` now upserts mounted fixed content tiles into `GridRuntimeHost.tiles`, reconciles viewport tile interest through the
  host-owned V3 tile coordinator, and exposes exact/stale/miss dirty-readiness diagnostics for the mounted tile-pane path.
- `TileResidencyV3` refreshes revision tuples in place when a fixed content tile is updated, so stale-compatible lookup and future dirty-tile
  decisions use the current axis/value/style/text/freeze revisions without replacing the stable residency entry.
- Scroll perf diagnostics no longer expose stale V2 scene-packet counters. The mounted tile-pane adapter reports renderer tile-interest batches,
  exact/stale/miss readiness, and visible/warm dirty tile counts from the host-owned V3 tile coordinator.
- Dynamic interaction overlays no longer build `GridScenePacketV2` values. They are emitted as `DynamicGridOverlayBatchV3` rect-instance buffers and drawn through a dedicated TypeGPU overlay resource outside tile cache and stale-tile lookup.
- Render tile deltas no longer carry overlay mutations or a `dynamicOverlay` pane kind. Overlays are now a visual runtime layer, not renderer tile data.
- Header panes no longer build or carry `GridScenePacketV2` values. They are emitted as fixed V3 header batches with packed rect instances plus text runs, and the TypeGPU backend draws them through dedicated header buffers outside tile cache and stale-tile lookup.
- `useWorkbookGridRenderState.ts` no longer imports the legacy `createGridCameraSnapshot()` camera builder, `visibleRegionFromCamera()`, `scrollCellIntoView()`, or `resolveViewportScrollPosition()`.
- The hook now owns a `GridRuntimeHost` instance, synchronizes axis overrides into it, and uses the host for camera visible-region transactions, render-tile interest sequencing, fixed tile key lookup, selection auto-scroll, and viewport restoration.
- `GridRuntimeHost` now exposes runtime-owned resident viewport tile-interest batches plus runtime-axis scroll-position helpers, so more scroll/viewport math has moved out of the React hook without changing the mounted pane adapter yet.
- `WorkbookGridSurface.tsx` no longer rebuilds V2 geometry snapshots from raw width/hidden records. Its TypeGPU geometry now comes from `useWorkbookGridRenderState().getLiveGeometrySnapshot()`, which reuses the hook/runtime axis indexes instead of constructing a second React-owned geometry path.
- `WorkbookGridSurface.tsx` now mounts `WorkbookPaneRendererV3`, and the isolated renderer route uses the V3 tile-pane API too.
- V3 tile-pane TypeGPU ownership now lives under `packages/grid/src/renderer-v3/typegpu-workbook-backend-v3.ts`,
  `typegpu-tile-buffer-pool.ts`, and `typegpu-tile-render-pass.ts`. V3 now owns its low-level TypeGPU surface, atlas, header, text,
  and overlay helpers, so data-tile resource sync/draw/residency no longer lives in the V2 workbook backend.
- V3 tile content buffers are now separated from tile placement uniforms: body and frozen placements share rect/text buffers by numeric tile ID,
  but each placement gets its own surface uniform and bind groups. This preserves content reuse without drawing frozen/body placements with a
  shared mutable uniform.
- Header and overlay TypeGPU resources now live in `renderer-v3/typegpu-layer-buffer-pool.ts`, not in the legacy V2
  `WorkbookPaneBufferCache`. The V3 backend owns separate resource caches for fixed content tiles and non-data visual layers.
- Local fallback tile generation and worker render-tile delta generation now materialize `GridRenderTile` values through
  `renderer-v3/grid-tile-materializer.ts`. They no longer build `GridScenePacketV2`, no longer create `GridTileKeyV2`, and no longer validate
  V3 content tiles through the V2 scene-packet validator.
- `renderer-v3/text-run-buffer.ts` now owns V3 text-run metric packing for fixed content tiles. Together with
  `renderer-v3/rect-instance-buffer.ts`, V3 tile materialization can build rect/text payloads without going through `scene-packet-v2.ts`.
- The grid package top-level public export no longer re-exports `renderer-v2/index.js`; `renderer-v2/index.ts` itself has been deleted.
  The old mounted V2 pane renderer, backend, text/atlas helpers, cache, debug helpers, and scene-packet primitives have been deleted, so
  `packages/grid/src/renderer-v2` no longer exists in the current source tree.

## 1. Validation Verdict

The Atlas response is directionally correct: the mounted grid path is no longer a legacy DOM/canvas renderer, and data rendering now enters through fixed render tiles. The mounted product path now uses `WorkbookPaneRendererV3` and V3 tile panes instead of a V2 scene-packet adapter. The V3 data-tile backend, TypeGPU primitives, pane layout, text/atlas helpers, draw loop, camera store, and header/overlay resource helpers are now V3-owned. There are no remaining `packages/grid/src/renderer-v2` source files; the remaining renderer work is now V3 runtime ownership and dirty-tile correctness, not V2 file deletion.

The current hot path still has these validated properties:

- `packages/grid/src/useWorkbookGridRenderState.ts` is still too large and owns pane construction, header scene building, overlay geometry bridge state, and React-facing state in one hook, but camera visible-region transactions, tile-interest stamping, selection auto-scroll math, and viewport restoration now run through `GridRuntimeHost`.
- The app-side resident scene store has been deleted; retained render data now lives in `apps/web/src/projected-tile-scene-store.ts`.
- Product data rendering no longer imports `packages/grid/src/gridResidentDataLayer.ts`; it now uses fixed render tiles and V3 tile-pane placement.
- `packages/grid/src/renderer-v3/typegpu-workbook-backend-v3.ts` owns V3 tile-pane drawing and uses `TileResidencyV3<GridRenderTile>` plus numeric tile IDs instead of `TileGpuCache` and serialized V2 scene packet keys.
- legacy V2 text-layout and atlas helpers were deleted; text layout, quad packing, and atlas dirty-page tests now target `packages/grid/src/renderer-v3`.
- `packages/grid/src/renderer-v3/dynamic-overlay-batch.ts` now builds interaction overlays without `GridScenePacketV2`; `WorkbookPaneRendererV3` passes those batches to a dedicated TypeGPU overlay buffer path instead of appending a fake overlay pane.
- `packages/grid/src/gridHeaderPanes.ts` now builds header batches without `GridScenePacketV2`; `WorkbookPaneRendererV3` passes those batches to a dedicated TypeGPU header buffer path instead of appending header scene packets to data panes.
- Browser perf tests now need to be tightened around the fixed-tile path, since resident-scene refreshes are no longer a normal product runtime behavior.

Important corrections to the Atlas text:

- `packages/grid/src/renderer-v2/typegpu-buffer-pool.ts` has been deleted; the equivalent resource-signature coverage now targets V3 fixed tiles and V3 TypeGPU primitives.
- `packages/grid/src/renderer-v2/typegpu-atlas-manager.ts`, `line-text-layout.ts`, `line-text-quad-buffer.ts`, `glyphAtlasV2.ts`, and
  `text-glyph-buffer.ts` have been deleted; remaining text/atlas coverage targets the V3 text path directly.
- `packages/grid/src/renderer-v2/scene-packet-v2.ts`, `scene-packet-validator.ts`, `tile-gpu-cache.ts`, `grid-render-layers.ts`,
  `render-debug-hud.ts`, and the V2 re-export shims have been deleted. V3 tile-packet, tile-residency, draw-command, and TypeGPU backend tests
  now carry the active coverage for fixed-tile rendering contracts.
- The practical next step is no longer V2 file deletion. The correct migration is measured replacement inside the V3-owned runtime: tighten
  counters and browser perf gates, replace remaining React-owned pane/header/editor coordination, and move routine dirty-tile updates onto
  renderer-native deltas.
- The follow-up Oracle response is correct that `useWorkbookGridRenderState.ts` must stop being the renderer owner, `ViewportPatch` must not remain the renderer contract, content tile identity must split from pane placement, and overlays need their own runtime. It is stale where it says the current `tile-gpu-cache.ts` still uses `filter().toSorted()` and where it describes `typegpu-buffer-pool.ts` as an active runtime file.

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
- deleted `packages/grid/src/renderer-v2/typegpu-buffer-pool.ts`
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
- deleted `packages/grid/src/renderer-v2/workbook-typegpu-backend.ts`
- `packages/grid/src/__tests__/tile-gpu-cache.test.ts`
- deleted `packages/grid/src/__tests__/workbook-typegpu-backend.test.ts`

Implement:

- replace `toSorted()` stale lookup with a direct newest-compatible scan;
- replace sort-based eviction with repeated oldest-non-visible selection until the cache is within budget;
- keep string cache keys only at the API boundary for this tranche, but prevent array/sort allocation during draw;
- expose counters so the browser perf harness can enforce `tileCacheSorts = 0`.

### Phase 3: split camera and render loop out of React

Files:

- `packages/grid/src/useWorkbookGridRenderState.ts`
- `packages/grid/src/WorkbookGridSurface.tsx`
- deleted `packages/grid/src/renderer-v2/WorkbookPaneRendererV2.tsx`
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
- cell-run, axis, freeze, and invalidate mutation types; overlays are intentionally excluded from render tile deltas;
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

- `packages/grid/src/renderer-v3/dynamic-overlay-batch.ts`
- deleted `packages/grid/src/renderer-v2/WorkbookPaneRendererV2.tsx`
- new `packages/grid/src/renderer-v2/header-tile-layer.ts`

Implement:

- headers as their own tile family;
- selection/fill/resize/collaboration overlay buffer independent from data tiles;
- selection drag and resize preview with zero data tile uploads.
- visible header-index generation avoids sort allocation while the overlay geometry builder remains geometry-driven.

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
- new `packages/grid/src/renderer-v3/tile-packet-v3.ts`
- new `packages/grid/src/renderer-v3/draw-command-buffer.ts`
- new `packages/grid/src/renderer-v3/glyph-key.ts`
- new `packages/grid/src/renderer-v3/text-atlas-pages.ts`
- new `packages/grid/src/renderer-v3/text-run-cache.ts`
- new `packages/grid/src/runtime/gridRuntimeHost.ts`
- new `packages/grid/src/runtime/gridTileCoordinator.ts`
- new `apps/web/src/projected-damage-bus.ts`
- new `apps/web/src/worker-runtime-delta-publisher.ts`

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
- fixed content tile packet contract with revision-tuple validation fields and byte accounting.
- pane-placement draw command buffer that separates content tile identity from body/frozen-pane placement.
- stable glyph IDs, page-level atlas dirty upload accounting, and interned text-run reuse keyed by clip/DPR inputs.
- imperative runtime host that composes axis, camera, overlay, and visible tile-key state for the future React shell adapter.
- tile coordinator that emits V3 tile-interest batches, consumes visible dirty tile damage, and classifies exact/stale/miss readiness against the V3 residency cache.
- dirty tile index application of V3 workbook delta batches, including bounded axis dirty ranges that are consumed per visible tile.
- app-side projected damage bus that applies V3 workbook deltas once per sheet and exposes visible/warm dirty tile queries.
- worker-side workbook delta publisher primitive that converts engine impact events into encoded sheet-level V3 dirty range batches without mounting it in the V2 startup bundle.

### Phase 7: text atlas service

Files:

- deleted `packages/grid/src/renderer-v2/typegpu-atlas-manager.ts`
- deleted `packages/grid/src/renderer-v2/line-text-quad-buffer.ts`
- new `packages/grid/src/renderer-v3/glyph-key.ts`
- new `packages/grid/src/renderer-v3/text-atlas-pages.ts`
- new `packages/grid/src/renderer-v3/text-run-cache.ts`

Implement:

- stable numeric glyph IDs for interned font/glyph/DPR keys;
- page-level atlas records with dirty-page upload accounting;
- text-run cache keyed by interned text/font/style/alignment/clip/DPR inputs;
- glyph dependency table per tile;
- atlas dirty-rect uploads and idle repack only;
- visible/warm dependency preservation during active scroll;
- no visible tile blanking on rare glyph allocation.

### Phase 8: delete scene-packet runtime use

Completed for resident data-scene runtime:

- `apps/web/src/projected-scene-store.ts`
- `apps/web/src/worker-runtime-render-scene.ts`
- `packages/grid/src/gridResidentDataLayer.ts` as renderer data runtime

Remaining scene-packet use is limited to legacy primitives/tests, not resident data-scene ownership, headers, overlays, or a mounted V2 workbook backend.

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
- dynamic overlays are built and drawn as V3 overlay batches, not V2 scene packets;
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
- `ProjectedViewportStore` exposes projected render tile subscriptions alongside app-state viewport subscriptions.

Completed in the resident-scene deletion tranche:

- `ProjectedSceneStore`, worker resident scene RPCs, worker resident scene cache, and resident pane scene bridge types were deleted.
- `ProjectedViewportStore` no longer exposes `subscribeResidentPaneScenes()` / `peekResidentPaneScenes()`.
- `WorkbookWorkerRuntime` no longer exposes `getResidentPaneScenes()`.
- `useWorkbookGridRenderState.ts` no longer imports or subscribes to resident pane scenes, renderer viewport subscriptions, V2 warm resident planning, or `gridResidentDataLayer`.
- `gridResidentDataLayer.ts`, `gridTileResidencyV2.ts`, `useGridViewportSubscriptions.ts`, and `useGridSceneResidency.ts` were deleted from `packages/grid/src`.
- A local fixed render tile materializer now gives non-worker fallback/tests the same fixed-tile data path used by worker render tile deltas.
- `WorkbookWorkerRuntime.subscribeRenderTileDeltas()` publishes binary tile deltas using the current resident-scene builder as the bootstrap oracle.
- `gridViewportController.ts` owns resident-window comparison, tile-key conversion, and render-scroll transform math outside the large React hook.
- `ProjectedTileSceneStore` reports renderer delta batches, mutations, apply duration, and dirty tile counts into the scroll perf collector.
- `renderer-v3/dynamic-overlay-batch.ts` replaces `dynamic-overlay-packet.ts`; mounted selection/hover/fill/resize/frozen-separator overlays are V3 rect-instance batches, not scene packets or render-tile deltas.
- `gridHeaderPanes.ts` now emits V3 header batches with packed rect instances and text runs; mounted column/row headers are drawn through dedicated TypeGPU header buffers rather than V2 scene packets.
- the legacy TypeGPU V2 tile cache now uses compatibility buckets, O(visible) visible marking, and LRU-tail eviction instead of scanning every cached tile for common operations.
- the legacy TypeGPU V2 pane buffer cache now releases pruned rect/text buffers to reusable free lists instead of destroying them during normal pane churn.
- the web build now isolates TypeGPU and grid renderer internals into a `grid-renderer-vendor` chunk so renderer migration work does not consume the whole workbook-vendor release budget.
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
- `packages/grid/src/renderer-v3/tile-packet-v3.ts` defines fixed content tile packets keyed by `TileKey53`, with bounds derived from protocol tile dimensions and validation by revision tuple rather than scene hashing.
- `packages/grid/src/renderer-v3/draw-command-buffer.ts` introduces pane placement commands so body and frozen panes can draw the same content tile under different clips/transforms.
- `packages/grid/src/runtime/gridRuntimeHost.ts` composes the first V3 runtimes behind an imperative host API that React can eventually mount and dispose instead of coordinating renderer internals.
- `packages/grid/src/runtime/gridTileCoordinator.ts` gives the host a concrete tile-interest and readiness coordinator over `TileResidencyV3` and `DirtyTileIndexV3`.
- `useWorkbookGridRenderState.ts` now sources visible-region snapshots from `GridRuntimeHost.updateCamera()` instead of `createGridCameraSnapshot()` and stamps render tile subscriptions from host-owned tile-interest batches.
- selection auto-scroll and viewport restoration now use `GridRuntimeHost` axis runtimes instead of rebuilding sorted axis indexes through `scrollCellIntoView()` / `resolveViewportScrollPosition()`.
- `packages/grid/src/runtime/gridRuntimeAxisAdapters.ts` bridges the current hook-owned size records into runtime axis overrides with tests, giving the next split a concrete seam for deleting more axis sorting from React-owned code.
- `WorkbookGridSurface.tsx` now consumes the hook's live geometry snapshot for the mounted TypeGPU fallback instead of rebuilding geometry from records in the surface component.
- `packages/grid/src/renderer-v3/tile-damage-index.ts` now applies sheet-level V3 dirty range batches to fixed tile damage and keeps axis dirty ranges bounded by tile rows/columns instead of expanding them over the full sheet.
- `apps/web/src/projected-damage-bus.ts` is the first app-side replacement seam for per-subscription viewport patch damage: it dedupes workbook delta sequence application per sheet ordinal and feeds the renderer dirty tile index.
- `apps/web/src/worker-runtime-delta-publisher.ts` provides the tested V3 damage-stream publisher primitive. The product worker runtime should expose it through transport only when the V3 renderer host owns the subscription, so the current V2 startup worker bundle stays under release size budgets.
- `packages/grid/src/renderer-v3/glyph-key.ts`, `text-atlas-pages.ts`, and `text-run-cache.ts` start the V3 text residency layer with stable glyph IDs, page-level dirty upload accounting, and interned text-run reuse keyed by clip/DPR inputs.
- `packages/grid/src/renderer-v3/grid-tile-materializer.ts` is now the shared fixed content-tile materializer for local fallback and worker
  render-tile deltas. It produces `GridTilePacketV3` metadata plus V3 rect/text buffers directly, eliminating `GridScenePacketV2` from those
  V3 tile-generation paths.
- `packages/grid/src/renderer-v3/typegpu-layer-buffer-pool.ts` replaces the V2 pane buffer cache for mounted V3 header and overlay resources,
  so `WorkbookPaneRendererV3` no longer depends on `renderer-v2/pane-buffer-cache.ts` or V2 header/overlay resource sync helpers.
- `packages/grid/src/renderer-v3/typegpu-primitives.ts`, `typegpu-surface.ts`, `typegpu-atlas-manager.ts`, `line-text-layout.ts`,
  `line-text-quad-buffer.ts`, `pane-layout.ts`, and `gridRenderLoop.ts` now give the mounted V3 renderer its own TypeGPU/layout/text/draw-loop
  primitives instead of importing them from `renderer-v2`. Shared render counters and camera-store state moved to non-V2 ownership, and
  `renderer-v3-import-boundary.test.ts` locks the mounted V3 import boundary.
- `packages/grid/src/renderer-v3/draw-scheduler.ts` now owns the mounted V3 requestAnimationFrame and idle warm-preload retry decision,
  moving the first scheduling lane out of `WorkbookPaneRendererV3.tsx`.
- `packages/grid/src/renderer-v2/WorkbookPaneRendererV2.tsx`, `packages/grid/src/renderer-v2/index.ts`, and the direct mounted-component
  test were deleted. Remaining `renderer-v2` modules are legacy low-level primitives and tests only.
- `packages/grid/src/renderer-v2/workbook-typegpu-backend.ts` and its direct stale-pane draw-selection test were deleted. The web
  `grid-renderer-vendor` chunk no longer has an explicit `renderer-v2` match, so renderer bundling is scoped to TypeGPU plus V3 paths.
- `packages/grid/src/renderer-v2/typegpu-render-pass.ts` and `packages/grid/src/renderer-v2/typegpu-surface.ts` were deleted after the V2
  workbook backend stopped importing them. V3 owns its draw pass and canvas surface state through `renderer-v3/typegpu-tile-render-pass.ts`
  and `renderer-v3/typegpu-surface.ts`.
- `packages/grid/src/renderer-v2/typegpu-backend.ts`, `typegpu-buffer-pool.ts`, `pane-buffer-cache.ts`, `pane-layout.ts`, and
  `pane-scene-types.ts` were deleted after V3 primitives, tile/layer buffer pools, and pane layout owned the mounted path. Remaining
  resource-signature and capacity tests now target the V3 APIs directly.
- `packages/grid/src/renderer-v2/typegpu-atlas-manager.ts`, `line-text-layout.ts`, `line-text-quad-buffer.ts`, `glyphAtlasV2.ts`, and
  `text-glyph-buffer.ts` were deleted after byte-identical line layout and quad tests were moved to V3 and the V3 atlas/page registry
  covered the remaining glyph residency behavior.
- `packages/grid/src/renderer-v2/gridRenderLoop.ts` and `gridCameraStore.ts` were deleted after their tests were moved to the V3 draw loop
  and runtime camera store owners.
- The final legacy `packages/grid/src/renderer-v2` files were deleted after remaining tests were either retired as obsolete V2-only coverage or
  retargeted to V3 constants. The import-boundary test now asserts the whole `renderer-v2` directory stays absent.
- `packages/grid/src/runtime/gridRenderTilePaneRuntime.ts` now owns remote/local fixed-tile pane resolution and same-sheet retained-pane fallback.
  It also owns render-tile delta subscription stamping, so `useWorkbookRenderTilePanes.ts` no longer builds viewport tile-interest batches or
  calls the render-tile source subscription API directly. Runtime and hook-boundary tests cover remote tile readiness, temporary tile-miss
  retention, sheet-switch invalidation, host-readiness behavior, and the subscription-stamping boundary.
- `packages/grid/src/runtime/gridHeaderPaneRuntime.ts` now owns V3 header pane generation, host-readiness gating, and body-tile content-offset
  projection. `useWorkbookHeaderPanes.ts` is reduced to a React memo/ref adapter, and hook-boundary tests lock out direct header pane builder use.
- `packages/grid/src/runtime/gridEditorAnchorRuntime.ts` now owns editor overlay screen-bound resolution, direct DOM bounds application,
  committed-bound equality, presentation resolution, and editor text alignment. `useWorkbookEditorOverlayAnchor.ts` remains the React
  effect/state bridge, preserving scroll-time DOM updates without React commits.
- `packages/grid/src/runtime/gridRuntimeAxisAdapters.ts` now owns the geometry-axis sort/serialization bridge, hidden-axis index construction,
  frozen-span computation, and scroll-spacer sizing for the grid runtime. `useWorkbookGridGeometryRuntime.ts` no longer performs raw
  `Object.entries().toSorted()` axis work or constructs world-axis indexes directly; boundary and adapter tests lock this ownership split.
- `packages/grid/src/runtime/gridViewportResidencyRuntime.ts` is now owned by `GridRuntimeHost`, so resident viewport retention, fixed render-tile
  viewport expansion, visible-address collection, and header-interest ranges are no longer carried by a standalone React-hook runtime.
- Render-tile delta subscriptions now carry host-computed V3 warm tile keys. `GridRenderTilePaneRuntime` builds visible plus one-ring warm
  interest from the host-owned tile coordinator, and `worker-runtime-render-tile-delta.ts` materializes dirty visible/warm tiles from that
  interest instead of treating the subscription as a viewport-only contract.

Remaining work from this design:

- continue splitting remaining draw/dirty-tile coordination out of `useWorkbookGridRenderState.ts`;
- replace routine resident scene packet regeneration with dirty-span tile payload updates;
- wire the V3 atlas/text-run primitives into the TypeGPU backend and add tile glyph dependency preservation;
- tighten browser perf gates around the now V3-only renderer path.
