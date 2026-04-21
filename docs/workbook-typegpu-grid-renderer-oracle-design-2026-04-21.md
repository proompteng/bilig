Workbook TypeGPU Grid Renderer Oracle Design — 2026-04-21

Target file: docs/workbook-typegpu-grid-renderer-oracle-design-2026-04-21.md

Commit inspected: b970be57b

1. Verdict On Current Implementation

Verdict

The current renderer is a partial prototype, not a salvageable production foundation. It has useful ingredients, but the core architecture should be replaced behind a renderer flag.

The current TypeGPU path proves that the project can submit GPU work from packages/grid, but it does not yet behave like a retained camera moving through a stable world. It still mixes React-driven viewport state, pane-local CPU clipping, object-scene rebuilding, main-thread text atlas work, and incomplete worker packets. That combination explains why vertical, horizontal, diagonal, and tile-boundary scrolling can feel untrustworthy even when tests pass.

The replacement should keep only the pieces that are structurally useful:

* Keep the workbook interaction shell in packages/grid/src/WorkbookGridSurface.tsx.
* Keep workbook selection/editing semantics from packages/grid/src/useWorkbookGridRenderState.ts, but move render-frame camera ownership out of React.
* Keep the idea of gridAxisIndex.ts, but replace or extend it into a world-coordinate axis index with hidden-entry semantics, total size, prefix spans, hit testing, and explicit revisions.
* Keep TypeGPU as the GPU abstraction, but rewrite the renderer backend around retained resources, typed packets, and a real camera uniform.
* Keep packages/worker-transport/src/index.ts; its transferable collection is useful.
* Keep the current TypeGPU E2E harness ideas in e2e/tests/web-shell-typegpu.pw.ts, but replace broad dark-pixel checks with exact pixel/geometry assertions.
* Keep GridTextOverlay.tsx and the current text layout code only as a reference/debug fallback until the new text packet pipeline is complete.

The following subsystems should be deleted or rewritten for the new path rather than patched locally:

* packages/grid/src/renderer/WorkbookPaneRenderer.tsx
* packages/grid/src/renderer/typegpu-renderer.ts
* packages/grid/src/renderer/typegpu-draw-pass.ts
* packages/grid/src/renderer/typegpu-resource-cache.ts
* packages/grid/src/renderer/grid-scene-packet.ts
* packages/grid/src/gridResidentDataLayer.ts as a renderer-facing scene model
* apps/web/src/worker-runtime-render-packet.ts
* the render-frame portions of packages/grid/src/useWorkbookGridRenderState.ts

The existing implementation can remain temporarily as typegpu-v1 for rollback and side-by-side comparison.

Existing Design Document Assessment

docs/workbook-typegpu-grid-renderer-implementation-plan-2026-04-20.md is useful as a record of intent, but it is not good enough as an implementation document. It describes a retained TypeGPU renderer and says the implementation is done, but the current code still has foundational problems:

* It does not define a single coordinate truth shared by rendering, hit testing, frozen panes, headers, selections, and editor overlays.
* It does not specify authoritative typed scene packet layouts.
* It does not define exact scroll-to-world transforms.
* It does not define text/glyph accuracy deeply enough for spreadsheet text.
* It does not require zero React commits, zero atlas uploads, zero buffer allocations, and zero scene uploads during warmed in-tile scroll.
* It does not define failure diagnostics strong enough to catch pane drift, one-pixel grid-line errors, stale camera snapshots, or tile blanking.

Top Architectural Mistakes

1. The camera is reduced to { tx, ty }, not a retained world camera

packages/grid/src/workbookGridScrollStore.ts stores only:

export interface WorkbookGridScrollSnapshot {
  readonly tx: number
  readonly ty: number
}

packages/grid/src/useWorkbookGridRenderState.ts computes a full camera in syncVisibleRegion, then discards most of it before notifying the renderer:

scrollTransformRef.current = {
  tx: next.tx,
  ty: next.ty,
}
scrollTransformStore.setSnapshot(scrollTransformRef.current)

The draw pass then subtracts only this intra-cell transform in packages/grid/src/renderer/typegpu-draw-pass.ts:

x: pane.contentOffset.x - (pane.scrollAxes.x ? scrollSnapshot.tx : 0),
y: pane.contentOffset.y - (pane.scrollAxes.y ? scrollSnapshot.ty : 0),

This is a primary likely cause of bad horizontal and vertical scroll feel.

During steady in-tile scroll, useWorkbookGridRenderState.ts intentionally avoids updating React visibleRegion if the resident viewport has not changed. That is the right instinct for performance, but the GPU draw path only receives the new tx and ty. When the visible anchor moves from one row or column to the next inside the same resident tile, the GPU needs the full body world scroll offset, not only the offset within the current row or column. The current renderer can therefore lose whole-cell movement until React state changes again, producing jumps, apparent freezing, header/body drift, or tile-boundary snaps.

2. Pane offsets are computed in two places with incompatible lifetimes

packages/grid/src/gridResidentDataLayer.ts computes pane offsets in resolveResidentDataPaneRenderState:

const bodyOffsetX = -(
  resolveColumnOffset(visibleViewport.colStart, ...) -
  resolveColumnOffset(residentViewport.colStart, ...) +
  visibleRegion.tx
)

But packages/grid/src/useWorkbookGridRenderState.ts calls it with a fake zero transform:

visibleRegion: {
  tx: 0,
  ty: 0,
},

Then typegpu-draw-pass.ts subtracts the live { tx, ty } again. This split means the scene can be clipped and offset using stale React state while the draw pass uses a newer scroll transform. The renderer does not own one authoritative camera.

3. CPU clipping is tied to tile/React updates instead of the per-frame camera

resolveResidentDataPaneRenderState clips scenes every time it builds pane render state:

gpuScene: clipPaneGpuSceneToWindow(pane.gpuScene, visibleWindow),
textScene: clipPaneTextSceneToWindow(pane.textScene, visibleWindow),

This is not a retained-world model. A retained tile should stay resident and be drawn through a moving camera. It should not need CPU clipping to the currently visible window during steady scroll.

CPU clipping at React/tile cadence can also cause blanking at edges because the camera can move before the clipped scene is rebuilt.

4. Worker scene packets are incomplete and mostly ignored

packages/grid/src/renderer/grid-scene-packet.ts defines PackedGridScenePacket, and apps/web/src/worker-runtime-render-packet.ts packs rects and text metrics. But the main renderer still consumes object-heavy gpuScene and textScene in packages/grid/src/renderer/typegpu-resource-cache.ts.

The current worker packet contains:

* rects: Float32Array
* textMetrics: Float32Array
* optional object textItems

It does not contain GPU-ready text/glyph instances, layer ordering, typed border/line semantics, atlas references, packet revisions, pane-kind-specific transforms, or enough style/text metadata to be authoritative. It is a prototype transport shape, not a production scene packet.

5. TypeGPU is used as a thin wrapper, not as a retained backend

packages/grid/src/renderer/typegpu-resource-cache.ts detects scene changes by object identity:

const textSceneChanged = paneCache.textScene !== pane.textScene

It then rebuilds and uploads per-pane buffers from object scenes. writeTypeGpuVertexBuffer in packages/grid/src/renderer/typegpu-renderer.ts clones incoming data before writing:

const copy = new Float32Array(data)
buffer.write(copy.buffer, 0, copy.byteLength)

That is not acceptable as the production path. The new renderer must retain tile buffers, maintain capacity pools, and upload only when worker packets or dynamic overlay packets change.

6. Text is line-atlas-based, main-thread-heavy, and not spreadsheet-accurate enough

packages/grid/src/renderer/glyph-atlas.ts interns whole strings under a key of ${font}:${glyph} where glyph is often a full line. It creates measurement canvases in measureGlyph, redraws all entries on atlas growth, and the renderer uploads the entire atlas texture whenever atlas.version changes.

packages/grid/src/renderer/gridTextLayout.ts uses useful early logic, but the model is not production spreadsheet text:

* default font is hard-coded as 400 11px sans-serif;
* vertical alignment is incomplete;
* numeric alignment is not specified as a renderer invariant;
* overflow into adjacent empty cells is not fully modeled;
* atlas versioning does not include DPR/font epoch/page identity;
* underline and strike are emitted as rects but not part of a complete text style system;
* high-DPR text quality depends on a fixed atlas scale.

Text cannot be treated as an optional overlay. Spreadsheet browsing quality depends on stable, accurate, clipped, aligned text.

7. Frozen panes, headers, and body are not driven by one coordinate truth

Pane layout and movement are split across:

* packages/grid/src/workbookGridViewport.ts
* packages/grid/src/gridResidentDataLayer.ts
* packages/grid/src/gridHeaderPanes.ts
* packages/grid/src/renderer/typegpu-draw-pass.ts
* packages/grid/src/useWorkbookGridRenderState.ts
* DOM overlays such as GridFillHandleOverlay.tsx

gridHeaderPanes.ts cuts and rewrites header scenes separately. Body panes use resident panes and contentOffset. DOM overlays call getCellLocalBounds, which uses scrollTransformRef and the live visible region. These paths can disagree under frozen panes, variable sizes, tile crossing, and editing scroll.

8. Axis math is useful but insufficient for production

packages/grid/src/gridAxisIndex.ts is a good start, but the production renderer needs more:

* stable totalSize;
* offsetOf(index) and sizeOf(index) with hidden-entry semantics;
* anchorAt(worldOffset) that skips zero-size hidden rows/columns deterministically;
* exact boundary behavior;
* hit testing;
* visible span resolution;
* frozen prefix spans;
* explicit axis revision IDs;
* no linear offset loops in hot render paths.

packages/grid/src/workbookGridViewport.ts still uses linear helpers such as resolveColumnOffset and resolveRowOffset in renderer-facing paths. It also creates axis indexes during viewport resolution. Axis construction must happen on data/axis revision, not during every scroll frame.

9. Hidden rows and columns are not consistently part of render geometry

apps/web/src/worker-runtime-render-scene.ts handles hidden axis entries by assigning size 0, but the local renderer path in packages/grid/src/useWorkbookGridRenderState.ts is primarily driven by columnWidths and rowHeights. WorkbookGridSurface.tsx passes hiddenRows and hiddenColumns to interactions, but render geometry is not clearly unified around hidden axis state. Hidden rows/columns must be part of the authoritative axis snapshot for worker, renderer, hit testing, and overlays.

10. Tests are too broad to catch the actual failures

e2e/tests/web-shell-typegpu.pw.ts checks some colors and dark pixel counts. That proves something was drawn; it does not prove that:

* a grid line is exactly one physical pixel;
* a frozen separator stays fixed while body content moves;
* a header label remains aligned with its body column after scroll;
* a selection border and fill handle match the selected cell after fractional scrolling;
* text is clipped to the correct cell or overflow range;
* no blank tile was visible for one frame;
* body/header panes do not drift by one or more pixels.

The current tests can pass while the renderer still feels wrong.

⸻

2. Production Renderer Architecture

Required Architecture

The new renderer must behave like an engine camera moving through a retained spreadsheet world.

The systems are:

1. Input and scroll controller
    * Native DOM scroll container captures wheel, trackpad, inertial, keyboard, and programmatic scroll.
    * It immediately normalizes scrollLeft and scrollTop into camera world coordinates.
    * It writes to an external camera store, not React state.
2. Camera and geometry store
    * Owns GridCameraSnapshotV2.
    * Computes pane frames, body world origin, frozen spans, visible anchors, viewport bounds, and DPR.
    * Is the single source of truth for rendering, hit testing, selections, fill handle, resize guides, and editor overlay placement.
3. Axis math
    * GridAxisWorldIndex owns row/column offsets, sizes, hidden entries, variable sizes, total size, and revisions.
    * Axis indexes are rebuilt only when axis data changes, not per scroll frame.
4. Tile residency
    * Converts camera world bounds into resident tile keys.
    * Maintains visible, warm-neighbor, stale-valid, and eviction states.
    * Never blanks the visible surface just because a preferred tile is not ready.
5. Worker scene preparation
    * Worker builds typed, versioned scene packets from workbook projection data.
    * Packets are transferable typed arrays, not object scenes.
    * Worker prepares cell fills, grid lines, authored borders, text layout plans, and static header/body tile data.
    * Worker does not own the WebGPU device.
6. TypeGPU backend
    * Main thread owns WebGPU device, TypeGPU root, context, pipelines, buffers, textures, bind groups, and draw submission.
    * Resources are retained across scroll frames.
    * Steady scroll updates only a small camera/frame uniform and submits a command buffer.
7. Text and glyph pipeline
    * Worker prepares deterministic text layout.
    * Main thread owns glyph atlas textures and GPU glyph instance buffers.
    * TypeGPU path and fallback path consume the same resolved text layout output.
8. Dynamic overlay system
    * Selection fill, active border, fill handle, resize guides, frozen separators, and hover hints are separate dynamic overlay packets.
    * They are updated from interaction state and camera geometry, not by rebuilding base cell tiles.
9. Instrumentation
    * Every frame records input-to-submit latency, GPU submits, draw calls, uploads, allocations, tile misses, blank tile events, React commits, layout shifts, atlas uploads, worker packets, and memory.
10. Tests

* Unit tests verify math and packet semantics.
* Playwright tests verify scrolling performance and exact pixels.
* TypeGPU readback tests verify geometry, not just draw existence.

Ownership Boundaries

React owns

React should own application state and semantic changes:

* active sheet;
* selected cell/range;
* editing mode and formula input;
* freeze row/column counts;
* row/column resize intent;
* sheet tabs, toolbar, status bar;
* high-level feature flags;
* mounting and unmounting renderer components.

React must not own frame pacing. React must not be required to commit during warmed scroll.

DOM scroll container owns

The DOM scroll container owns browser-native input mechanics:

* scrollbars;
* wheel/trackpad/inertial behavior;
* keyboard scroll if routed to the container;
* programmatic scrollTo.

It does not own render geometry. Its scrollLeft and scrollTop are input values that are normalized into camera coordinates.

Camera store owns

The camera store owns:

* normalized bodyScrollX and bodyScrollY;
* full body world origin;
* row and column anchors;
* visible world bounds;
* pane frames;
* DPR;
* viewport CSS and physical sizes;
* velocity;
* monotonically increasing cameraSeq.

This is the only frame-time state the renderer samples.

Worker owns

The worker owns:

* resolving workbook projection data into tile scene packets;
* packing typed arrays;
* text layout planning;
* computing style/value invalidation for tiles;
* prioritizing visible and warm tile requests.

The worker must not depend on React render cadence.

Main TypeGPU backend owns

The main TypeGPU backend owns:

* WebGPU adapter/device/context;
* context configure lifecycle;
* TypeGPU root;
* pipelines;
* uniform buffers;
* tile buffer cache;
* atlas textures and samplers;
* command encoding;
* readback hooks in tests;
* device loss recovery.

DOM overlays own only browser-native editing UI

The cell editor can remain DOM because it needs text input, IME, accessibility, and clipboard behavior. It must be positioned by the same camera geometry and updated with direct style transforms during scroll. Opening, closing, and committing the editor can use React. Per-scroll editor movement must not require React state updates.

Work That Must Never Happen In A Warmed Steady Scroll Frame

A warmed steady scroll frame means the visible tiles and required glyph atlas pages are already resident.

The following are forbidden:

* React commits caused by scroll.
* buildGridGpuScene.
* buildGridTextScene.
* buildResidentDataPaneScenes.
* resolveResidentDataPaneRenderState.
* text measurement.
* glyph rasterization.
* atlas texture upload.
* GPUCanvasContext.configure.
* canvas width/height mutation.
* buffer creation or destruction.
* tile packet decoding.
* worker request scheduling that blocks drawing.
* DOM layout reads such as getBoundingClientRect, clientWidth, or clientHeight.
* sorting row/column overrides.
* rebuilding axis indexes.
* serializing scene data.
* CPU clipping of base cell scenes to the current viewport.
* setState for visible viewport, overlay bounds, or pane position.

Allowed in a warmed steady scroll frame:

* read latest camera snapshot;
* write one small frame/camera uniform;
* update zero or one dynamic overlay buffer if interaction state changed;
* encode one render pass with retained buffers;
* submit one command buffer;
* update editor/fill handle DOM transform using already-known geometry.

⸻

3. Accurate Scroll And Coordinate Model

Coordinate Spaces

Use CSS pixels for all workbook world geometry. Use physical pixels only for raster size, scissor rects, and one-physical-pixel snapping.

Definitions:

type CssPx = number
type PhysicalPx = number
interface GridViewportMetrics {
  rowHeaderWidth: CssPx
  columnHeaderHeight: CssPx
  hostWidth: CssPx
  hostHeight: CssPx
  dpr: number
}
interface GridWorldAxes {
  columns: GridAxisWorldIndex
  rows: GridAxisWorldIndex
}
interface GridCameraSnapshotV2 {
  seq: number
  sheetName: string
  scrollLeft: CssPx
  scrollTop: CssPx
  bodyScrollX: CssPx
  bodyScrollY: CssPx
  bodyWorldX: CssPx
  bodyWorldY: CssPx
  bodyViewportWidth: CssPx
  bodyViewportHeight: CssPx
  frozenColumnCount: number
  frozenRowCount: number
  frozenWidth: CssPx
  frozenHeight: CssPx
  columnAnchor: AxisAnchor
  rowAnchor: AxisAnchor
  visibleBodyWorldRect: Rect
  visibleSheetRect: CellRange
  residentHorizon: TileHorizon
  panes: GridPaneGeometry[]
  dpr: number
  updatedAt: number
  inputAt: number
  velocityX: CssPx
  velocityY: CssPx
  axisVersionX: number
  axisVersionY: number
  freezeVersion: number
}

The body world coordinate system includes every column and row in the sheet, including frozen rows and columns. Headers are screen-space panes that label world axes.

Scroll Container Normalization

The renderer should change the scroll spacer model used by packages/grid/src/WorkbookGridSurface.tsx.

A native scroll viewport should expose scroll offsets over the scrollable body, not over row headers, column headers, or frozen panes.

Let:

bodyPaneX = rowHeaderWidth + frozenWidth
bodyPaneY = columnHeaderHeight + frozenHeight
bodyViewportWidth = max(0, hostWidth - bodyPaneX)
bodyViewportHeight = max(0, hostHeight - bodyPaneY)
scrollableBodyWidth = max(0, columnAxis.totalSize - frozenWidth)
scrollableBodyHeight = max(0, rowAxis.totalSize - frozenHeight)
maxScrollX = max(0, scrollableBodyWidth - bodyViewportWidth)
maxScrollY = max(0, scrollableBodyHeight - bodyViewportHeight)

The scroll spacer should be sized so that the browser produces those max scroll values:

scrollSpacerWidth = hostWidth + maxScrollX
scrollSpacerHeight = hostHeight + maxScrollY

Then:

bodyScrollX = clamp(scrollViewport.scrollLeft, 0, maxScrollX)
bodyScrollY = clamp(scrollViewport.scrollTop, 0, maxScrollY)
bodyWorldX = frozenWidth + bodyScrollX
bodyWorldY = frozenHeight + bodyScrollY

This replaces the current ambiguous model in packages/grid/src/workbookGridViewport.ts, where resolveVisibleRegionFromScroll adds frozenWidth and frozenHeight to scrollLeft/scrollTop while the scroll spacer is based on total grid width/height. The new model gives scrollLeft = 0 a single meaning: the first non-frozen body column is aligned to the body pane.

Axis Index Semantics

Create packages/grid/src/gridAxisWorldIndex.ts.

Required API:

interface AxisEntryOverride {
  index: number
  size?: number
  hidden?: boolean
}
interface AxisAnchor {
  index: number
  offset: CssPx
  size: CssPx
  intraOffset: CssPx
}
interface GridAxisWorldIndex {
  axisLength: number
  defaultSize: CssPx
  version: number
  totalSize: CssPx
  sizeOf(index: number): CssPx
  isHidden(index: number): boolean
  offsetOf(index: number): CssPx
  endOffsetOf(index: number): CssPx
  span(startInclusive: number, endExclusive: number): CssPx
  anchorAt(worldOffset: CssPx): AxisAnchor
  visibleCountFrom(startIndex: number, viewportSize: CssPx): number
  visibleRangeForWorldRect(startOffset: CssPx, size: CssPx): AxisVisibleRange
  hitTest(worldOffset: CssPx): number | null
}

Rules:

* Default rows and columns have defaultSize.
* Variable-size entries override the default.
* Hidden entries have rendered size 0.
* Hidden entries are never returned as hit-test results.
* anchorAt(offset) must skip zero-size hidden entries.
* At an exact boundary, choose the next visible non-hidden entry unless the offset is at or beyond total size.
* offsetOf(index) returns the accumulated rendered offset, excluding hidden size.
* totalSize is the rendered total size.
* Axis indexes are immutable snapshots and must include a monotonically increasing version.
* Axis indexes are rebuilt only on row/column size, hidden, insert/delete, or sheet switch changes.

The existing packages/grid/src/gridAxisIndex.ts can be used as a starting point, but the new index must remove all renderer hot-path linear offset loops currently found in packages/grid/src/workbookGridViewport.ts.

Pane Frames

The renderer has nine logical panes. Some can have zero size.

const bodyPaneX = rowHeaderWidth + frozenWidth
const bodyPaneY = columnHeaderHeight + frozenHeight
const bodyViewportWidth = max(0, hostWidth - bodyPaneX)
const bodyViewportHeight = max(0, hostHeight - bodyPaneY)

Pane frames in screen CSS pixels:

cornerHeaderFrame = {
  x: 0,
  y: 0,
  width: rowHeaderWidth,
  height: columnHeaderHeight,
}
columnHeaderFrozenFrame = {
  x: rowHeaderWidth,
  y: 0,
  width: frozenWidth,
  height: columnHeaderHeight,
}
columnHeaderBodyFrame = {
  x: bodyPaneX,
  y: 0,
  width: bodyViewportWidth,
  height: columnHeaderHeight,
}
rowHeaderFrozenFrame = {
  x: 0,
  y: columnHeaderHeight,
  width: rowHeaderWidth,
  height: frozenHeight,
}
rowHeaderBodyFrame = {
  x: 0,
  y: bodyPaneY,
  width: rowHeaderWidth,
  height: bodyViewportHeight,
}
frozenCornerBodyFrame = {
  x: rowHeaderWidth,
  y: columnHeaderHeight,
  width: frozenWidth,
  height: frozenHeight,
}
frozenTopBodyFrame = {
  x: bodyPaneX,
  y: columnHeaderHeight,
  width: bodyViewportWidth,
  height: frozenHeight,
}
frozenLeftBodyFrame = {
  x: rowHeaderWidth,
  y: bodyPaneY,
  width: frozenWidth,
  height: bodyViewportHeight,
}
bodyFrame = {
  x: bodyPaneX,
  y: bodyPaneY,
  width: bodyViewportWidth,
  height: bodyViewportHeight,
}

World-To-Screen Transforms

All transforms must be expressed from world coordinates into screen CSS pixels.

For a cell at column c, row r:

cellWorldX = columnAxis.offsetOf(c)
cellWorldY = rowAxis.offsetOf(r)
cellWidth = columnAxis.sizeOf(c)
cellHeight = rowAxis.sizeOf(r)

Scrollable body cells

For c >= frozenColumnCount and r >= frozenRowCount:

screenX = bodyPaneX + (cellWorldX - bodyWorldX)
screenY = bodyPaneY + (cellWorldY - bodyWorldY)

Frozen top rows, scrollable columns

For c >= frozenColumnCount and r < frozenRowCount:

screenX = bodyPaneX + (cellWorldX - bodyWorldX)
screenY = columnHeaderHeight + cellWorldY

Frozen left columns, scrollable rows

For c < frozenColumnCount and r >= frozenRowCount:

screenX = rowHeaderWidth + cellWorldX
screenY = bodyPaneY + (cellWorldY - bodyWorldY)

Frozen corner cells

For c < frozenColumnCount and r < frozenRowCount:

screenX = rowHeaderWidth + cellWorldX
screenY = columnHeaderHeight + cellWorldY

Column headers for scrollable columns

screenX = bodyPaneX + (cellWorldX - bodyWorldX)
screenY = 0

Column headers for frozen columns

screenX = rowHeaderWidth + cellWorldX
screenY = 0

Row headers for scrollable rows

screenX = 0
screenY = bodyPaneY + (cellWorldY - bodyWorldY)

Row headers for frozen rows

screenX = 0
screenY = columnHeaderHeight + cellWorldY

These transforms must be implemented once in GridGeometrySnapshot and reused everywhere.

Selection, Fill Handle, Resize Guides, Editor Overlay, And Hit Testing

Create packages/grid/src/gridGeometry.ts.

Required methods:

interface GridGeometrySnapshot {
  camera: GridCameraSnapshotV2
  axes: GridWorldAxes
  cellWorldRect(cell: CellAddress): Rect
  cellScreenRect(cell: CellAddress): Rect | null
  cellScreenRectForPane(cell: CellAddress, paneKind: GridPaneKind): Rect | null
  rangeWorldRects(range: CellRange): readonly Rect[]
  rangeScreenRects(range: CellRange): readonly Rect[]
  columnHeaderScreenRect(col: number): Rect | null
  rowHeaderScreenRect(row: number): Rect | null
  hitTestScreenPoint(point: Point): GridHitResult
  resizeGuideScreenRect(input: ResizeGuideState): Rect | null
  fillHandleScreenRect(range: CellRange): Rect | null
  editorScreenRect(cell: CellAddress): Rect | null
}

Rules:

* The active cell border uses cellScreenRect.
* The selection range uses rangeScreenRects.
* The fill handle is anchored to the bottom-right visible selection rect. If the selected range crosses frozen/body boundaries, choose the rect containing the active cell edge, not an arbitrary union.
* The editor overlay uses editorScreenRect. It must move with direct style updates from the render loop while editing.
* Resize guide positions use axis offsets from GridAxisWorldIndex, not DOM measurements.
* Hit testing converts screen point into pane kind, then into world coordinate, then into axis hitTest.
* Hidden rows/columns return no hit result.
* Pane boundaries and frozen separators get priority over cell hits when within resize handle tolerance.

Replace the current scattered geometry in packages/grid/src/useWorkbookGridRenderState.ts, especially getCellLocalBounds, with this single object.

DPR, Snapping, Fractional Scroll, And One-Physical-Pixel Lines

Rules:

1. Canvas backing size:

canvas.width = round(hostWidth * dpr)
canvas.height = round(hostHeight * dpr)
canvas.style.width = `${hostWidth}px`
canvas.style.height = `${hostHeight}px`

2. Use CSS pixels in all instance data and uniforms.
3. Scissor rects are physical pixels:

x0 = floor(frame.x * dpr)
y0 = floor(frame.y * dpr)
x1 = ceil((frame.x + frame.width) * dpr)
y1 = ceil((frame.y + frame.height) * dpr)

4. Preserve fractional scroll for cell fills, text, selections, and overlays. Do not round camera position.
5. Grid lines and hairline borders are one physical pixel. Convert line thickness to CSS pixels:

hairlineCss = 1 / dpr
snappedCss = round(cssCoord * dpr) / dpr

6. Snap line rectangles, not the camera.
7. Text should move smoothly with fractional scroll. Do not snap text origin every frame in a way that creates visible shimmer. If idle snapping is added later, it must be velocity-gated and must not alter hit testing or geometry.
8. Clip rectangles must be in the same coordinate space as the transformed primitive. The current mixed text behavior in typegpu-renderer.ts, where text position is snapped after scroll but clip comparison uses the unsnapped pre-scroll local position, must not be repeated.
9. Pixel tests must tolerate at most one physical pixel where antialiasing is unavoidable, but grid lines, selection borders, and frozen separators should be exact.

⸻

4. TypeGPU Rendering Pipeline

New Files

Create a new renderer namespace first:

packages/grid/src/renderer-v2/
  WorkbookPaneRendererV2.tsx
  typegpu-backend.ts
  typegpu-pipelines.ts
  typegpu-buffer-pool.ts
  typegpu-surface.ts
  typegpu-atlas-manager.ts
  typegpu-render-pass.ts
  scene-packet-v2.ts
  scene-packet-validator.ts
  tile-gpu-cache.ts
  dynamic-overlay-packet.ts
  render-debug-hud.ts

Keep the existing renderer under renderer/ until migration is complete.

Retained GPU Resources

The TypeGPU backend owns:

interface WorkbookTypeGpuBackend {
  root: TgpuRoot
  device: GPUDevice
  context: GPUCanvasContext
  format: GPUTextureFormat
  surface: TypeGpuSurfaceState
  pipelines: WorkbookPipelines
  frameUniformRing: UniformRing<FrameUniform>
  paneUniformRing: UniformRing<PaneUniform>
  quadBuffer: TgpuBuffer
  tileCache: TileGpuCache
  overlayBuffers: DynamicOverlayBuffers
  atlas: GlyphAtlasGpuManager
  frameSeq: number
  deviceGeneration: number
}

Resources created on warmup:

* TypeGPU root.
* WebGPU adapter/device/context.
* context configuration.
* static unit quad/index buffer.
* rect/fill pipeline.
* grid-line pipeline.
* border pipeline.
* text pipeline.
* overlay pipeline.
* uniform ring buffers.
* default atlas pages.
* sampler objects.
* empty fallback buffers.
* tile buffer pools.

Resources retained per tile:

interface TileGpuResource {
  key: GridTileKeyV2
  generation: number
  worldBounds: Rect
  paneKind: GridPaneKind
  fillBuffer: GpuBufferSlice | null
  lineBuffer: GpuBufferSlice | null
  borderBuffer: GpuBufferSlice | null
  textGlyphBuffer: GpuBufferSlice | null
  fillCount: number
  lineCount: number
  borderCount: number
  textGlyphCount: number
  byteSize: number
  lastUsedFrame: number
  staleValid: boolean
}

Uniform Layouts

Frame uniform:

interface FrameUniform {
  screenCssWidth: f32
  screenCssHeight: f32
  screenPhysicalWidth: f32
  screenPhysicalHeight: f32
  dpr: f32
  frameSeq: u32
  timeMs: f32
  flags: u32
}

Pane uniform:

interface PaneUniform {
  paneX: f32
  paneY: f32
  paneWidth: f32
  paneHeight: f32
  worldOriginX: f32
  worldOriginY: f32
  screenOriginX: f32
  screenOriginY: f32
  dpr: f32
  layerBase: f32
  paneKind: u32
  flags: u32
}

For a body pane, worldOriginX/Y is the world coordinate aligned to screenOriginX/Y. For frozen panes, the frozen axis uses 0 or the frozen world coordinate while the scrollable axis uses bodyWorldX or bodyWorldY.

The vertex shader computes:

screen = pane.screenOrigin + (instance.worldPosition - pane.worldOrigin)
clip = css_to_ndc(screen, frame.screenCssSize)

This avoids per-pane CPU offset mutation. The camera changes through uniforms, not by rebuilding instances.

Buffer Layouts

Use typed packet data that maps directly into GPU buffer layouts.

Fill instances

// Float32 stride: 12
interface FillInstance {
  worldX: f32
  worldY: f32
  width: f32
  height: f32
  r: f32
  g: f32
  b: f32
  a: f32
  clipX: f32
  clipY: f32
  clipW: f32
  clipH: f32
}

Only non-default cell fills should be emitted. Default white/background is drawn as pane background.

Grid-line instances

// Float32 stride: 10
interface LineInstance {
  worldX: f32
  worldY: f32
  length: f32
  thicknessCss: f32
  axis: f32 // 0 horizontal, 1 vertical
  r: f32
  g: f32
  b: f32
  a: f32
  flags: f32 // snapToPhysical, frozenSeparator, headerLine
}

Grid-line shaders snap the line start and thickness to one physical pixel when snapToPhysical is set.

Border instances

// Float32 stride: 14
interface BorderInstance {
  worldX: f32
  worldY: f32
  width: f32
  height: f32
  edgeMask: f32 // top/right/bottom/left bits
  thicknessCss: f32
  style: f32
  z: f32
  r: f32
  g: f32
  b: f32
  a: f32
  clipX: f32
  clipY: f32
}

Authored borders are separate from grid lines. They draw after grid lines and before text unless spreadsheet semantics require a specific overlap.

Text glyph instances

// Float32 stride: 20
interface TextGlyphInstance {
  worldX: f32
  worldY: f32
  width: f32
  height: f32
  uv0x: f32
  uv0y: f32
  uv1x: f32
  uv1y: f32
  r: f32
  g: f32
  b: f32
  a: f32
  clipX: f32
  clipY: f32
  clipW: f32
  clipH: f32
  atlasPage: f32
  fontEpoch: f32
  subpixelFlags: f32
  reserved: f32
}

Text clip coordinates are world coordinates. The shader converts clip bounds through the same pane uniform as glyph position.

Dynamic overlay instances

Dynamic overlays use a smaller buffer updated from selection/editing state:

interface OverlayRectInstance {
  screenOrWorldX: f32
  screenOrWorldY: f32
  width: f32
  height: f32
  r: f32
  g: f32
  b: f32
  a: f32
  mode: f32 // screen or world
  snap: f32
  z: f32
  reserved: f32
}

Selection ranges should generally be world-mode. Fill handles, resize guides, and frozen separators can be screen-mode if already resolved by GridGeometrySnapshot.

Bind Groups

Required bind groups:

1. Frame bind group:
    * frame uniform;
    * debug flags uniform if enabled.
2. Pane bind group:
    * pane uniform.
3. Atlas bind group:
    * glyph atlas texture array or page texture;
    * sampler;
    * atlas metadata buffer if needed.
4. Optional tile metadata bind group:
    * used only if storage buffers replace vertex buffers later.

Use TypeGPU schemas for all uniforms and instance layouts. Avoid any layout escapes.

Render Pass

One render pass per frame.

High-level order:

1. Begin pass with clear color.
2. Draw sheet/window background.
3. For each pane in stable order:
    * set scissor rect;
    * bind frame and pane uniforms;
    * draw default pane background;
    * draw retained tile fills intersecting the pane;
    * draw retained grid lines;
    * draw authored borders;
    * draw retained text glyphs.
4. Clear scissor or set full-surface scissor.
5. Draw selection fill.
6. Draw active cell border.
7. Draw fill handle.
8. Draw resize guides.
9. Draw frozen separators.
10. Draw debug overlays if enabled.
11. End pass and submit.

Stable pane order:

cornerHeader
columnHeaderFrozen
columnHeaderBody
rowHeaderFrozen
rowHeaderBody
frozenCornerBody
frozenTopBody
frozenLeftBody
body
dynamicOverlays
debug

This order prevents scrollable body content from covering frozen panes or headers.

Scissor Usage

Every pane draw uses a physical-pixel scissor rect computed from pane frame and DPR. Do not use shader-only clipping for pane boundaries.

Text still needs per-cell or overflow-range clipping in shader. The clip comparison must use world coordinates transformed consistently with the glyph position.

Draw Call Budget

The renderer should target:

* one render pass;
* one GPU queue submit per animation frame;
* no more than 24 draw calls for warmed steady scroll;
* no more than 32 draw calls for frozen-pane diagonal scroll;
* no more than 48 draw calls during debug HUD or unusual overlay combinations.

This implies batching by layer and pane, not by cell.

Blend Mode And Premultiplication

Use premultiplied colors in instance buffers.

Recommended blend:

color: {
  srcFactor: 'one',
  dstFactor: 'one-minus-src-alpha',
  operation: 'add',
},
alpha: {
  srcFactor: 'one',
  dstFactor: 'one-minus-src-alpha',
  operation: 'add',
}

Text shader samples atlas alpha and returns premultiplied color:

let sampled = textureSample(atlasTexture, atlasSampler, uv);
let alpha = instanceColor.a * sampled.a;
return vec4f(instanceColor.rgb * alpha, alpha);

Do not mix premultiplied and non-premultiplied colors as the current pipeline risks doing.

Upload Rules

Warmup

Allowed:

* create pipelines;
* configure context;
* allocate buffer pools;
* allocate uniform rings;
* allocate default atlas texture pages;
* upload static buffers;
* request initial visible and warm tiles;
* upload first visible tile buffers;
* rasterize/upload required initial glyphs.

Required before declaring ready:

* all visible panes have valid base tiles;
* header text is resident;
* body text for the first visible viewport is resident or has an explicit debug fallback marker;
* no visible blank tile.

Warmed steady scroll inside resident tiles

Allowed:

* update frame/pane uniform ring;
* submit one command buffer;
* update DOM editor/fill-handle transform if present.

Forbidden:

* scene upload;
* atlas upload;
* buffer allocation;
* context configure;
* React commit;
* worker packet application.

Tile crossing

Allowed:

* apply already-prepared worker packets;
* upload tile buffers from packet data;
* promote warm tiles to visible;
* schedule additional warm tiles.

Rules:

* If the exact preferred tile is not ready, draw stale-valid overlapping tiles.
* Never clear to blank.
* Bound tile upload work per frame. A single frame should not upload an unbounded set of tiles.

Edit

Rules:

* Selection changes update dynamic overlay only.
* Cell value edits invalidate body text/fill/border packets for affected tiles.
* If text overflow changes, invalidate neighbor cells whose visible text could be covered or uncovered.
* Formula recalculation invalidates all cells returned by damage patches.
* Old tile remains stale-valid until replacement packet is validated and uploaded.

Resize

Rules:

* ResizeObserver may update host dimensions.
* Surface configure happens once for the coalesced size/DPR change before the next draw.
* Axis data is unchanged unless row/column sizes changed.
* Camera and pane frames update immediately.
* No base scene rebuild is required unless the viewport horizon requests new tiles.

DPR change

Rules:

* Surface backing size changes.
* Atlas DPR bucket changes.
* Text glyph atlas pages for the old DPR remain cached but inactive.
* Text tile packets that depend on glyph metrics must be invalidated if metrics change.
* Frame should degrade to stale text only if necessary, never blank cells.

Font/style changes

Rules:

* Font epoch increments for changed font families, sizes, weights, or styles.
* Affected text packets are invalidated.
* Missing glyphs are queued by atlas manager.
* Fills/borders are invalidated only if those style fields changed.

TypeGPU-Specific Pitfalls To Avoid

* Do not call root.configureContext or context.configure in the draw path.
* Do not recreate bind groups every frame.
* Do not allocate or destroy buffers during scroll.
* Do not clone typed arrays before every buffer write.
* Do not store object scenes and typed packets together in production packets.
* Do not hide coordinate logic inside shader magic that differs from hit testing.
* Do not use one giant unbounded atlas texture with full re-upload on every glyph insertion.
* Do not rely on object identity as the scene change detector.
* Do not create a TypeGPU root per pane.
* Do not use TypeGPU merely to wrap hand-wavy WebGPU; use typed schemas for uniforms, instance buffers, and bind groups.
* Do not forget device loss. device.lost must move the renderer into fallback/recovery state and preserve camera/tiles at the app level.

⸻

5. Text And Glyph Accuracy

Decision

Use a hybrid CPU layout + GPU atlas glyph pipeline.

* CPU prepares spreadsheet text layout deterministically.
* GPU draws glyph instances from atlas pages.
* Main thread owns the WebGPU atlas textures.
* Worker prepares text layout and glyph plans.
* TypeGPU and fallback rendering consume the same ResolvedCellTextLayout output.

Do not keep the current whole-line atlas in packages/grid/src/renderer/glyph-atlas.ts as the production model. Whole-line rasterization causes too many unique atlas entries for large spreadsheets, expensive full-atlas uploads, weak reuse, and unstable memory behavior.

New Text Files

Create:

packages/grid/src/text/gridTextLayoutV2.ts
packages/grid/src/text/gridTextMetrics.ts
packages/grid/src/text/gridTextOverflow.ts
packages/grid/src/text/gridTextPacket.ts
packages/grid/src/renderer-v2/glyphAtlasV2.ts
packages/grid/src/renderer-v2/text-glyph-buffer.ts

The old files can remain temporarily:

packages/grid/src/renderer/gridTextLayout.ts
packages/grid/src/renderer/glyph-atlas.ts
packages/grid/src/renderer/text-quad-buffer.ts
packages/grid/src/GridTextOverlay.tsx

but they should not be used by typegpu-v2 except for debug comparison.

Resolved Text Layout Model

Required structure:

interface ResolvedCellTextLayout {
  cell: CellAddress
  text: string
  displayText: string
  fontKey: FontKey
  color: PremultipliedRgba
  horizontalAlign: 'left' | 'center' | 'right'
  verticalAlign: 'top' | 'middle' | 'bottom'
  wrap: boolean
  overflow: TextOverflowMode
  cellWorldRect: Rect
  textClipWorldRect: Rect
  overflowWorldRect: Rect
  lines: readonly ResolvedTextLine[]
  decorations: readonly TextDecorationRect[]
  generation: number
}
interface ResolvedTextLine {
  text: string
  worldX: CssPx
  baselineWorldY: CssPx
  advance: CssPx
  ascent: CssPx
  descent: CssPx
  glyphs: readonly ResolvedGlyphPlacement[]
}
interface ResolvedGlyphPlacement {
  glyphId: number
  atlasGlyphKey: AtlasGlyphKey
  worldX: CssPx
  worldY: CssPx
  width: CssPx
  height: CssPx
  advance: CssPx
  uvPadding: CssPx
}

Font Key

interface FontKey {
  family: string
  sizeCssPx: number
  weight: number | string
  style: 'normal' | 'italic'
  stretch?: string
  variant?: string
  dprBucket: number
  fontEpoch: number
}

The atlas key includes:

* font family;
* size;
* weight;
* italic;
* DPR bucket;
* font epoch;
* glyph codepoint or glyph cluster;
* rasterization mode.

Alignment Rules

Horizontal:

* Explicit left/center/right style wins.
* Numeric/date/formatted numeric text defaults to right alignment if no explicit alignment.
* Plain strings default to left alignment.
* Booleans and errors follow spreadsheet semantics; if no existing semantic is available, booleans center and errors left.
* Empty text emits no glyph instances.

Vertical:

* Explicit top/middle/bottom style wins.
* Default is middle for unwrapped single-line text.
* Wrapped text defaults to top unless workbook style says otherwise.

Padding:

* Use a named constant, not magic inline values:

CELL_TEXT_PADDING_X = 8
CELL_TEXT_PADDING_Y = 3

The current TEXT_PADDING_X = 8 in gridTextLayout.ts is acceptable as a starting value, but it must be centralized and tested.

Baseline Rules

Use actual font metrics where available:

ascent = metrics.actualBoundingBoxAscent
descent = metrics.actualBoundingBoxDescent
lineHeight = max(style.lineHeight ?? size * 1.2, ascent + descent)

If browser metrics are unavailable, use deterministic fallback ratios:

ascent = size * 0.8
descent = size * 0.2

For middle vertical alignment:

contentHeight = lineCount * lineHeight
firstBaselineY =
  cellTop +
  floorOrFractional((cellHeight - contentHeight) / 2) +
  ascent

Do not use the current approximation in gridTextLayout.ts that vertically centers using only fontSize.

Clipping And Overflow

Spreadsheet overflow rules are mandatory.

For unwrapped left-aligned text:

* It may overflow into adjacent empty cells to the right.
* It stops before the first non-empty cell, hidden column, frozen boundary, viewport boundary, merged-cell boundary, or explicit clip boundary.
* It does not overflow across frozen/non-frozen pane boundaries.
* It does not overflow through cells with visible fills/borders that should cover it if spreadsheet semantics require clipping.

For right-aligned numeric text:

* It clips to the cell unless explicit style allows overflow.
* It should not cover left neighbors.

For wrapped text:

* It clips to the cell content rect.
* It does not overflow into neighbors.

For center-aligned text:

* It clips to the cell unless spreadsheet semantics explicitly allow overflow.

Overflow invalidation:

* Editing a text cell invalidates its tile and any neighbor span previously covered by overflow.
* Editing a neighbor cell invalidates overflow sources that might now be blocked or unblocked.
* Changing column width invalidates text layout for affected columns and overflow dependencies.

Wrapping

Implement wrapping by grapheme cluster, word boundary, and emergency break:

1. Segment into words/graphemes using Intl.Segmenter when available.
2. Fit by measured advance.
3. If a single segment exceeds the available width, break by grapheme.
4. Clip lines by row height.
5. Emit layout for visible lines only, but packet metadata must know that text was clipped for debug.

Bold, Italic, Underline, Strike

* Bold and italic are encoded in FontKey.
* Underline and strike are emitted as decoration rect instances in the border/overlay layer.
* Underline position uses font metrics when available.
* Strike position uses ascent/descent metrics, not a fixed half-height guess.
* Decoration rects snap to physical pixels.

Numeric Text

Numeric text must be stable under scroll and edit:

* formatting occurs before layout;
* default alignment is right;
* decimal/grouping does not cause remeasure on scroll;
* negative/accounting formats must fit into the same layout pipeline;
* high-DPR digit glyphs must be atlas-rasterized at the active DPR bucket.

Wide Text

Wide text cannot be handled as a single whole-line atlas bitmap in production. It must be represented by glyph instances or cluster instances so repeated glyphs reuse atlas space.

For extremely long strings:

* cap layout work per cell;
* emit visible glyphs only;
* mark truncation in debug metadata;
* do not block the frame.

Atlas Versioning

Atlas manager state:

interface GlyphAtlasPage {
  pageId: number
  width: number
  height: number
  dprBucket: number
  fontEpoch: number
  generation: number
  texture: TgpuTexture
  dirtyRects: Rect[]
}
interface GlyphAtlasGpuManager {
  currentGeneration: number
  pages: Map<number, GlyphAtlasPage>
  glyphMap: Map<AtlasGlyphKey, AtlasGlyphLocation>
  missingQueue: AtlasGlyphKey[]
}

Rules:

* Atlas pages are fixed-size, for example 2048x2048 physical pixels.
* Add pages instead of resizing and redrawing all glyphs.
* Upload dirty rects when supported by the TypeGPU texture write path; if not, upload one page, not the entire atlas.
* Glyph instance packets reference pageId and generation.
* If a packet references a missing glyph, draw the tile without that glyph for at most one frame and record missingGlyphEvents; then upload glyphs and redraw.
* Atlas memory is capped and evicts least-recently-used pages only when no visible tile references them.

TypeGPU And Fallback Must Share Output

The fallback path must consume ResolvedCellTextLayout and produce the same positions, clips, decorations, and overflow behavior. It may draw with Canvas2D, but it must not have separate layout logic.

Tests should compare TypeGPU glyph instance positions against fallback text layout output before comparing pixels.

⸻

6. Worker And Tile Streaming Model

Tile Keys

Create packages/grid/src/renderer-v2/scene-packet-v2.ts.

type GridPaneKind =
  | 'body'
  | 'frozenTop'
  | 'frozenLeft'
  | 'frozenCorner'
  | 'columnHeaderBody'
  | 'columnHeaderFrozen'
  | 'rowHeaderBody'
  | 'rowHeaderFrozen'
  | 'cornerHeader'
interface GridTileKeyV2 {
  sheetName: string
  paneKind: GridPaneKind
  rowStart: number
  rowEnd: number
  colStart: number
  colEnd: number
  rowTile: number
  colTile: number
  axisVersionX: number
  axisVersionY: number
  valueVersion: number
  styleVersion: number
  selectionIndependentVersion: number
  freezeVersion: number
  textEpoch: number
  dprBucket: number
}

Selection is intentionally not part of the base tile key. Selection overlays are dynamic.

Tile Size

Start with existing constants from @bilig/protocol:

* VIEWPORT_TILE_ROW_COUNT
* VIEWPORT_TILE_COLUMN_COUNT

Then enforce pixel caps:

MAX_TILE_WIDTH_CSS = 4096
MAX_TILE_HEIGHT_CSS = 4096
MAX_TILE_CELLS = VIEWPORT_TILE_ROW_COUNT * VIEWPORT_TILE_COLUMN_COUNT

A tile can be split if variable column widths or row heights make it too large in pixels.

Resident Horizon

The resident horizon includes:

* visible body tile set;
* visible frozen panes;
* visible headers;
* one full neighbor ring around visible body tiles;
* two forward rings in the current velocity direction;
* forward diagonal tiles during diagonal scroll;
* recently visible stale-valid tiles.

Warm policy:

priority 0: visible exact tiles
priority 1: frozen/header visible tiles
priority 2: next forward body tiles by velocity
priority 3: diagonal forward tiles
priority 4: lateral neighbor ring
priority 5: reverse neighbor ring
priority 6: stale-valid retention

Hysteresis

Avoid tile churn at tile edges.

Rules:

* Do not switch resident horizon just because the camera touches a tile boundary.
* Keep the current resident horizon until the camera crosses 25% of a tile beyond the current visible tile or the visible tile set actually changes.
* Keep recently visible tiles stale-valid for at least 3 seconds or 180 frames.
* During inertial scrolling, forward prefetch extends based on velocity magnitude.

Stale-Valid Fallback

A stale-valid tile may be drawn if:

* sheet name matches;
* pane kind matches;
* axis/freeze versions match;
* tile world bounds overlap the visible pane;
* its value/style versions are older but still semantically safe to display while a replacement is in flight.

A stale tile must never be used across axis version changes because geometry could be wrong.

Blanking policy:

* A visible pane with no valid or stale-valid tile increments blankTileEvents.
* blankTileEvents must be zero in production tests.
* The renderer should draw a debug warning overlay only when debug mode is active; production should prefer stale data over blank.

Cancellation And Prioritization

Worker requests carry:

interface TileRequestV2 {
  requestSeq: number
  cameraSeq: number
  key: GridTileKeyV2
  priority: number
  reason: 'visible' | 'prefetch' | 'edit' | 'resize' | 'sheet-switch'
}

Rules:

* Newer visible requests supersede older prefetch requests.
* Worker may continue work, but main must ignore packets whose request sequence or key revisions are stale.
* In-flight low-priority packets are dropped on the worker side when the queue grows past the limit.
* Edits and visible tile misses preempt prefetch.

Packet Shape

Production packets must be typed and transferable.

interface GridScenePacketHeaderV2 {
  magic: 0x42475244 // "BGRD"
  version: 2
  requestSeq: number
  cameraSeq: number
  generatedAt: number
  key: GridTileKeyV2
  worldBounds: Rect
  fillCount: number
  lineCount: number
  borderCount: number
  textGlyphCount: number
  decorationCount: number
  fillStrideFloats: number
  lineStrideFloats: number
  borderStrideFloats: number
  textGlyphStrideFloats: number
  stringTableCount: number
  atlasRequestCount: number
  rowAxisVersion: number
  columnAxisVersion: number
  packetCrc?: number
}
interface GridScenePacketV2 {
  header: GridScenePacketHeaderV2
  fills: Float32Array
  lines: Float32Array
  borders: Float32Array
  textGlyphs: Float32Array
  decorations: Float32Array
  atlasRequests: Uint32Array
  stringTable: readonly string[]
  debug?: GridScenePacketDebug
}

No production packet should include GridGpuScene or GridTextScene object arrays.

Main Thread Validation

Before applying a packet:

1. Check magic/version.
2. Check sheet name and pane kind.
3. Check axis, style, value, freeze, text, and DPR revisions.
4. Check counts and typed-array lengths.
5. Reject NaN, infinite, negative widths/heights where invalid.
6. Check world bounds are plausible for the tile key.
7. Check packet generation is newer than the current tile resource.
8. Check atlas references or queue missing glyphs.
9. Upload into retained tile buffers.
10. Publish tile as visible or warm.

Rejected packets increment scenePacketRejected with reason.

Invalidation Rules

Edits

Invalidate:

* tile containing the edited cell;
* tiles containing dependent formula recalculation damage;
* overflow source and neighbor tiles for text that may cover adjacent cells;
* text atlas requests for new glyphs;
* dynamic overlay only if selection changes.

Selection

Do not invalidate base tiles.

Update only:

* selection fill overlay;
* active cell border;
* header selection highlight;
* fill handle.

Style changes

Invalidate affected tiles by style field:

* fill color -> fill packet;
* border style -> border packet;
* font/alignment/wrap/format -> text packet;
* row/column style default -> all affected row/column tiles.

Row/column size and hidden changes

Increment axis version.

Invalidate:

* all geometry-dependent tile packets for the sheet;
* all header packets for affected axis;
* text layout for affected rows/columns;
* overlay geometry.

Old packets are not stale-valid across axis version changes.

Freeze changes

Increment freeze version.

Invalidate:

* pane layout;
* all pane tile keys;
* headers;
* overlays.

Sheet switch

Rules:

* Preserve per-sheet tile caches within memory limits.
* Current sheet renderer switches camera/geometry immediately.
* It may draw stale-valid cached tiles for the new sheet if versions match.
* It must not show tiles from the previous sheet.

Memory Limits

Initial targets:

GPU tile buffer cache: 128 MB
CPU decoded packet cache: 64 MB
Glyph atlas pages: 128 MB physical texture memory
Worker pending packet memory: 64 MB

Eviction order:

1. stale non-visible reverse-direction tiles;
2. old warm tiles;
3. old sheet tiles;
4. atlas pages not referenced by visible/warm tiles;
5. visible stale tiles only after replacement exists.

Memory counters must expose live bytes, high-water bytes, and evictions.

⸻

7. Frame Scheduler And Input Pipeline

New Scheduler Files

Create:

packages/grid/src/renderer-v2/gridRenderLoop.ts
packages/grid/src/renderer-v2/gridCameraStore.ts
packages/grid/src/renderer-v2/gridScrollController.ts
packages/grid/src/renderer-v2/gridResizeController.ts

Event Loop Model

Scroll event

1. Passive scroll listener fires.
2. Read scrollViewport.scrollLeft and scrollViewport.scrollTop.
3. Normalize to bodyScrollX/Y.
4. Build or update GridCameraSnapshotV2 using existing axis snapshots and viewport metrics.
5. Store it in GridCameraStore.
6. Record noteGridScrollInput(inputAt).
7. Schedule one render RAF if not already scheduled.
8. Notify tile residency controller synchronously or by microtask. Do not wait for worker.

RAF render

1. Read the latest camera snapshot.
2. Read latest tile cache state.
3. If surface size/DPR changed, configure context once before pass.
4. Write frame/pane uniforms.
5. Encode render pass using retained tile buffers.
6. Submit one command buffer.
7. Record noteGridDrawFrame.
8. Update DOM editor/fill-handle transforms from the same camera snapshot.
9. Flush debug HUD counters if enabled.

The render loop samples the latest camera at draw time. If five scroll events occur before one RAF, only the newest camera is drawn.

Prevent Two-RAF Lag

The current apps/web/src/projected-scene-store.ts schedules worker refresh through window.requestAnimationFrame(flush). That is not acceptable for visible tile requests in the new path.

Rules:

* Scroll-to-render uses at most one RAF: the browser RAF that draws the frame.
* Visible tile requests are queued immediately when camera horizon changes.
* Worker responses are asynchronous, but drawing stale-valid tiles must not wait.
* React must not schedule the frame.

Coalescing

Input coalescing:

* Last camera wins.
* Keep the earliest inputAt timestamp since the previous draw for input-to-draw latency.
* Keep velocity estimated from recent camera samples.

Worker coalescing:

* Multiple camera updates in the same task produce one visible-horizon request set.
* Prefetch requests can be delayed by one microtask to absorb rapid wheel deltas.
* Visible tile misses are never delayed.

Resize And DPR

ResizeObserver flow:

1. Observe host element.
2. Coalesce size changes.
3. Update viewport metrics in camera store.
4. Mark surface dirty.
5. Schedule render RAF.
6. Configure context once before the next draw.

DPR change detection:

* Check window.devicePixelRatio on resize, visibility change, and media query change.
* If DPR changes, mark surface and atlas DPR bucket dirty.
* Recompute camera with same logical scroll position.
* Schedule text/atlas invalidation.

Device Loss

On device.lost:

1. Stop submitting GPU work.
2. Mark renderer state device-lost.
3. Keep camera store, tile packet cache, and app state.
4. Attempt reinitialization if the loss reason permits.
5. If reinitialization fails, switch to fallback renderer.
6. Record deviceLossCount.

DOM Overlay Synchronization

The editor overlay and any remaining DOM-only overlay must subscribe to the render loop, not to React state.

Rules:

* React opens/closes editor.
* Render loop sets style.transform, style.width, style.height, and visibility.
* Overlay position uses GridGeometrySnapshot.editorScreenRect.
* Overlay is hidden if the cell is outside all visible panes.
* Overlay drift budget is at most one physical pixel.
* No setState during scroll.

⸻

8. Instrumentation And Budgets

Required Counters

Extend apps/web/src/perf/workbook-scroll-perf.ts and packages/grid/src/renderer/grid-render-counters.ts.

Counters to expose:

interface WorkbookRendererPerfCountersV2 {
  frameCount: number
  droppedFrameCount: number
  inputToDrawMs: number[]
  frameIntervalMs: number[]
  gpuSubmitToPresentApproxMs: number[]
  longTaskCount: number
  longTaskMaxMs: number
  reactCommits: number
  reactCommitsDuringScroll: number
  layoutShiftCount: number
  layoutShiftScore: number
  contextConfigureCount: number
  surfaceResizeCount: number
  deviceLossCount: number
  gpuSubmitCount: number
  commandEncoderCount: number
  renderPassCount: number
  drawCallCount: number
  paneDrawCount: number
  uniformWriteBytes: number
  sceneUploadBytes: number
  atlasUploadBytes: number
  bufferAllocationCount: number
  bufferAllocationBytes: number
  bufferDestroyCount: number
  workerRequestCount: number
  workerPacketReceivedCount: number
  workerPacketAppliedCount: number
  workerPacketRejectedCount: number
  workerPacketCancelledCount: number
  tileMissCount: number
  blankTileEventCount: number
  staleTileDrawCount: number
  tileEvictionCount: number
  glyphMissingCount: number
  atlasPageCount: number
  atlasEvictionCount: number
  gpuTileCacheBytes: number
  cpuPacketCacheBytes: number
  atlasTextureBytes: number
  debugHudVisible: boolean
}

Keep existing counters such as typeGpuConfigures, typeGpuSubmits, typeGpuDrawCalls, typeGpuVertexUploadBytes, and typeGpuAtlasUploadBytes, but split them into more specific V2 counters.

Performance Budgets

Budgets apply after a warmup period of at least 12 RAFs, as current workbook-scroll-perf.ts already does.

Warmed steady in-tile scroll

Scenario: 960x720 viewport, DPR <= 2, visible and warm tiles resident.

Targets:

frame interval p95: <= 16.7 ms
frame interval p99: <= 24 ms
input-to-draw p95: <= 8 ms
input-to-draw p99: <= 16 ms
long tasks: 0 over 50 ms
React commits during scroll: 0
layout shifts: 0
context configure: 0
surface resize: 0
GPU submits: 1 per drawn frame
render passes: 1 per drawn frame
draw calls: <= 24 per frame
buffer allocations: 0
buffer destroys: 0
scene upload bytes: 0
atlas upload bytes: 0
worker packets applied: 0
tile misses: 0
blank tile events: 0
editor/fill-handle drift: <= 1 physical px

Tile-boundary scroll

Scenario: cross body tile boundaries horizontally and vertically with warm prefetch enabled.

Targets:

frame interval p95: <= 16.7 ms
frame interval p99: <= 32 ms
input-to-draw p95: <= 10 ms
input-to-draw p99: <= 20 ms
long tasks over 50 ms: 0
blank tile events: 0
visible tile miss events: <= 1 per 10 tile crossings
scene upload bytes per crossing frame: <= 4 MB
atlas upload bytes per crossing frame: <= 1 MB after warmup
buffer allocations during crossing: 0 after pool warmup
draw calls: <= 32 per frame

Deep vertical scroll

Scenario: jump or scroll to row 200,000+ with variable row sizes present.

Targets:

axis anchor resolve p95: <= 0.10 ms
frame interval p95 after settled: <= 16.7 ms
frame interval p99 after settled: <= 32 ms
blank tile events: 0
React commits caused by steady scroll: 0
no linear offset scan proportional to row index

Diagonal scroll

Scenario: simultaneous horizontal and vertical movement across tile boundaries.

Targets:

frame interval p95: <= 16.7 ms
frame interval p99: <= 32 ms
input-to-draw p95: <= 10 ms
draw calls: <= 32 per frame
blank tile events: 0
stale tile draw allowed: yes, if visually correct and revision-safe

Resize then scroll

Scenario: resize viewport, then immediately scroll.

Targets:

context configure count: <= 2 per resize burst
surface resize count: <= 2 per resize burst
first post-resize frame: <= 32 ms
second post-resize frame returns to steady budget
blank tile events: 0

Editing overlay scroll

Scenario: open cell editor and scroll horizontally, vertically, and diagonally.

Targets:

React commits during scroll: 0
editor overlay drift: <= 1 physical px
input-to-overlay p95: <= 8 ms
input-to-draw p95: <= 10 ms
blank tile events: 0

Debug HUD

Add a debug HUD enabled by query param:

?gridDebugHud=1

It should display:

* renderer mode;
* WebGPU adapter/device status;
* camera seq and scroll world origin;
* row/column anchors;
* visible tile keys;
* stale/blank tile counts;
* frame p95/p99;
* input-to-draw p95/p99;
* draw calls;
* upload bytes;
* atlas page count;
* React commits during scroll;
* last rejected packet reason.

The HUD must not itself create per-frame React commits during scroll. Render it as a lightweight DOM element updated directly or as a GPU debug layer.

⸻

9. Testing Strategy

Unit Tests

Add or replace tests under packages/grid/src/__tests__/.

Axis math

File:

packages/grid/src/__tests__/gridAxisWorldIndex.test.ts

Cases:

* default sizes;
* variable sizes;
* hidden rows/columns;
* exact boundary anchor resolution;
* offsetOf, endOffsetOf, span;
* total size;
* hit testing skips hidden entries;
* visible range does not include zero-size hidden entries as drawable geometry;
* deep index anchor resolution does not degrade linearly.

Camera and pane geometry

File:

packages/grid/src/__tests__/gridCameraGeometry.test.ts

Cases:

* scrollLeft/scrollTop to bodyWorldX/bodyWorldY;
* frozen width/height;
* body pane frame;
* top header transform;
* left header transform;
* frozen corner body transform;
* fractional scroll;
* DPR does not alter CSS geometry;
* cell screen rect for cells in each pane;
* editor rect and fill handle rect.

Tile residency

File:

packages/grid/src/__tests__/gridTileResidencyV2.test.ts

Cases:

* visible tile keys;
* one-ring warm neighbors;
* velocity-forward prefetch;
* diagonal prefetch;
* hysteresis;
* stale-valid retention;
* eviction order;
* key revision mismatch invalidation.

Text layout

File:

packages/grid/src/__tests__/gridTextLayoutV2.test.ts

Cases:

* left/center/right alignment;
* numeric default right alignment;
* vertical top/middle/bottom alignment;
* wrap by word and grapheme;
* overflow into adjacent empty cells;
* overflow blocked by non-empty neighbor;
* hidden columns block overflow;
* underline and strike positions;
* high-DPR glyph placement;
* text clip rects;
* fallback and TypeGPU packet output share identical layout.

Packet packing

File:

packages/grid/src/__tests__/gridScenePacketV2.test.ts

Cases:

* valid packet packs and validates;
* transferred typed arrays have expected lengths;
* NaN rejection;
* stale revision rejection;
* wrong pane kind rejection;
* count/stride mismatch rejection;
* debug metadata does not affect production validation.

Scheduler

File:

packages/grid/src/__tests__/gridRenderLoop.test.ts

Cases:

* multiple scroll events coalesce to one RAF draw;
* latest camera wins;
* input timestamp is preserved for latency;
* no render if camera unchanged and no damage;
* resize coalesces configure;
* device loss stops submit.

Layer ordering

File:

packages/grid/src/__tests__/gridRenderLayerOrdering.test.ts

Cases:

* fills before grid lines;
* grid lines before authored borders;
* borders before text where required;
* selection fill before active border;
* active border before fill handle;
* frozen separators above body content;
* headers above body where panes overlap.

Playwright Performance Tests

Update e2e/tests/web-shell-scroll-performance.pw.ts.

Required scenarios:

1. Horizontal steady in-tile scroll.
2. Vertical steady in-tile scroll.
3. Diagonal steady scroll.
4. Wheel/trackpad-like inertial scroll using repeated page.mouse.wheel.
5. Tile-boundary horizontal scroll.
6. Tile-boundary vertical scroll.
7. Deep vertical scroll to row 200,000+.
8. Variable row heights and column widths.
9. Hidden rows and columns.
10. Frozen rows and columns.
11. Resize then immediate scroll.
12. Editing overlay while scrolling.
13. Remote edits while browsing.
14. Sheet switch with cached tiles.

Each test must save on failure:

* perf JSON;
* screenshot;
* TypeGPU readback PNG if available;
* camera snapshot JSON;
* tile cache JSON;
* axis snapshot summary;
* rejected packet log.

TypeGPU Readback Tests

Update e2e/tests/web-shell-typegpu.pw.ts.

Current dark-pixel checks should remain only as smoke checks. Add exact geometry probes.

Required assertions:

Grid line exactness

After a known scroll, sample expected vertical and horizontal grid-line positions.

* The line pixel has expected grid color.
* Adjacent pixels on both sides are not line color.
* Line width is exactly one physical pixel.

Header/body alignment

For a known column after horizontal scroll:

* body vertical grid line X equals column header grid line X within one physical pixel.
* header text region is inside the same column rect.

For a known row after vertical scroll:

* body horizontal grid line Y equals row header grid line Y within one physical pixel.

Frozen panes

With frozen rows/columns:

* frozen separator X/Y stays fixed after body scroll.
* frozen cell content does not move on the frozen axis.
* body content moves on the scrollable axis.
* top header scrolls horizontally but not vertically.
* left header scrolls vertically but not horizontally.

Selection and fill handle

* Active cell border corners match GridGeometrySnapshot.cellScreenRect.
* Selection fill alpha is present inside range and absent outside.
* Fill handle is at the bottom-right screen rect of the active selection edge.

Text clipping and overflow

* Wide text in an empty row overflows into adjacent empty cells.
* Adding a neighbor value clips the overflow before that cell.
* No text pixels appear outside overflowWorldRect.
* Wrapped text stays inside its cell.

Editing overlay

* DOM editor bounding box equals GPU active cell geometry within one physical pixel.
* It remains aligned while scrolling.

Resize guide

* Column resize guide line is at the expected snapped X.
* Row resize guide line is at the expected snapped Y.

Failure Artifacts

On failure, save:

test-results/<test>/screenshot.png
test-results/<test>/gpu-readback.png
test-results/<test>/perf-report.json
test-results/<test>/camera.json
test-results/<test>/geometry.json
test-results/<test>/tile-cache.json
test-results/<test>/axis.json
test-results/<test>/packet-log.json

The pixel probe harness should also save an annotated probe overlay image that marks expected and actual probe points.

⸻

10. Implementation Plan

Stage 0 — Add Renderer Flag And Protect The Current Path

Files to create/change:

apps/web/src/renderer-flags.ts
packages/grid/src/renderer-v2/WorkbookPaneRendererV2.tsx
packages/grid/src/renderer-v2/index.ts
packages/grid/src/WorkbookGridSurface.tsx
e2e/tests/web-shell-typegpu.pw.ts

Responsibilities:

* Add renderer mode flag:
    * typegpu-v1
    * typegpu-v2
    * canvas-fallback
* Support query param and environment selection:
    * ?workbookRenderer=typegpu-v2
    * VITE_WORKBOOK_RENDERER=typegpu-v2
* Keep current renderer as default until V2 passes evidence gates.
* Add test IDs that distinguish V1 and V2.

Tests:

* smoke test that V1 still mounts;
* smoke test that V2 flag mounts placeholder renderer;
* smoke test fallback path does not crash.

Acceptance criteria:

* App remains usable with current renderer.
* V2 can be selected without affecting V1.
* E2E can run both modes.

Risk:

* Low. This isolates future changes.

Stage 1 — Axis World Index

Files to create/change:

packages/grid/src/gridAxisWorldIndex.ts
packages/grid/src/__tests__/gridAxisWorldIndex.test.ts
packages/grid/src/workbookGridViewport.ts
packages/grid/src/gridCamera.ts

Responsibilities:

* Implement immutable axis snapshots.
* Support default, override, hidden, total size, anchor, hit test, visible range.
* Remove renderer-facing linear offset usage from new path.
* Keep old gridAxisIndex.ts until V1 migration is complete.

Tests:

* full axis math suite.
* deep offset lookup performance test.

Acceptance criteria:

* Hidden rows/columns are deterministic.
* Exact boundary behavior is tested.
* Deep row/column lookup is logarithmic or better.
* No V1 regression.

Risk:

* Medium. Axis mistakes affect every subsystem.

Stage 2 — Camera And Geometry Snapshot

Files to create/change:

packages/grid/src/gridGeometry.ts
packages/grid/src/renderer-v2/gridCameraStore.ts
packages/grid/src/renderer-v2/gridScrollController.ts
packages/grid/src/__tests__/gridCameraGeometry.test.ts
packages/grid/src/WorkbookGridSurface.tsx

Responsibilities:

* Normalize DOM scroll to body world camera.
* Compute pane frames and transforms.
* Implement cellScreenRect, rangeScreenRects, hit testing, editor rect, fill handle rect, resize guide rect.
* Add direct camera store subscription.

Tests:

* pane geometry tests;
* frozen transform tests;
* fractional scroll tests;
* hit test tests.

Acceptance criteria:

* All render, overlay, and hit-test geometry can be produced from GridGeometrySnapshot.
* No per-scroll React state is needed to know where a cell is.

Risk:

* High. This stage replaces the coordinate truth.

Stage 3 — Render Loop And Surface Lifecycle

Files to create/change:

packages/grid/src/renderer-v2/gridRenderLoop.ts
packages/grid/src/renderer-v2/gridResizeController.ts
packages/grid/src/renderer-v2/typegpu-surface.ts
packages/grid/src/renderer-v2/WorkbookPaneRendererV2.tsx
packages/grid/src/__tests__/gridRenderLoop.test.ts
apps/web/src/perf/workbook-scroll-perf.ts

Responsibilities:

* Implement one-RAF render loop.
* Coalesce scroll input.
* Configure surface only on mount/resize/DPR.
* Track frame counters.
* Keep DOM scroll container and camera store synchronized.

Tests:

* coalescing tests;
* resize tests;
* no submit after destroy;
* performance smoke test with zero React commits during scroll in V2 placeholder mode.

Acceptance criteria:

* Scroll causes at most one RAF draw.
* No context configure during steady scroll.
* Counters expose frame/input metrics.

Risk:

* Medium.

Stage 4 — Scene Packet V2

Files to create/change:

packages/grid/src/renderer-v2/scene-packet-v2.ts
packages/grid/src/renderer-v2/scene-packet-validator.ts
packages/grid/src/__tests__/gridScenePacketV2.test.ts
apps/web/src/worker-runtime-render-packet.ts
apps/web/src/worker-runtime-render-scene.ts
apps/web/src/projected-scene-store.ts
packages/worker-transport/src/index.ts

Responsibilities:

* Define versioned packet header and typed arrays.
* Pack fills, lines, borders, text glyph placeholders, and decorations.
* Remove production dependency on object gpuScene and textScene for V2 packets.
* Validate packets on main thread.
* Preserve debug object metadata only behind debug mode.

Tests:

* packet validation;
* transferability;
* stale rejection;
* NaN rejection.

Acceptance criteria:

* V2 receives transferable typed packets.
* V2 can reject invalid packets with reason.
* No V2 production packet requires object scenes.

Risk:

* Medium-high. Worker/main schema mismatch can cause blank rendering.

Stage 5 — TypeGPU Backend V2 Basic Geometry

Files to create/change:

packages/grid/src/renderer-v2/typegpu-backend.ts
packages/grid/src/renderer-v2/typegpu-pipelines.ts
packages/grid/src/renderer-v2/typegpu-buffer-pool.ts
packages/grid/src/renderer-v2/tile-gpu-cache.ts
packages/grid/src/renderer-v2/typegpu-render-pass.ts
packages/grid/src/renderer-v2/WorkbookPaneRendererV2.tsx
e2e/tests/web-shell-typegpu.pw.ts

Responsibilities:

* Create TypeGPU device/root/context.
* Build fill, line, border, and overlay pipelines.
* Draw generated grid geometry through camera uniforms.
* Use retained tile buffers.
* Implement scissor per pane.
* Implement one-physical-pixel grid lines.

Tests:

* isolated renderer readback;
* exact grid-line pixel tests;
* pane scissor tests.

Acceptance criteria:

* V2 draws grid fills and lines.
* Steady scroll updates only uniforms.
* Pixel tests prove one-physical-pixel lines.

Risk:

* High. GPU pipeline and coordinate bugs are likely.

Stage 6 — Tile Cache And Residency V2

Files to create/change:

packages/grid/src/gridTileResidencyV2.ts
packages/grid/src/renderer-v2/tile-gpu-cache.ts
packages/grid/src/__tests__/gridTileResidencyV2.test.ts
apps/web/src/projected-scene-store.ts
apps/web/src/worker-runtime.ts

Responsibilities:

* Implement visible/warm/stale tile policy.
* Add priorities and cancellation.
* Add memory caps and eviction.
* Connect camera horizon to worker requests.
* Draw stale-valid tiles instead of blanking.

Tests:

* residency unit tests;
* Playwright tile-boundary scroll;
* blank tile counter assertions.

Acceptance criteria:

* Tile-boundary scroll has zero blank tile events.
* Warm tiles are promoted without visible upload stalls when available.
* Stale tiles are revision-safe.

Risk:

* High. Tile cache bugs cause blanking or stale incorrect data.

Stage 7 — Text Layout V2 And Glyph Atlas

Files to create/change:

packages/grid/src/text/gridTextLayoutV2.ts
packages/grid/src/text/gridTextMetrics.ts
packages/grid/src/text/gridTextOverflow.ts
packages/grid/src/text/gridTextPacket.ts
packages/grid/src/renderer-v2/glyphAtlasV2.ts
packages/grid/src/renderer-v2/text-glyph-buffer.ts
packages/grid/src/__tests__/gridTextLayoutV2.test.ts
apps/web/src/worker-runtime-render-scene.ts

Responsibilities:

* Implement resolved text layout.
* Implement glyph atlas pages.
* Emit glyph instances.
* Support alignment, baseline, clipping, overflow, wrap, bold, italic, underline, strike.
* Implement atlas requests and generation validation.
* Make fallback consume same layout output.

Tests:

* text layout unit tests;
* readback text clipping and overflow tests;
* high-DPR text tests.

Acceptance criteria:

* Text appears through TypeGPU V2.
* No full-atlas upload during steady scroll.
* Wide text and numeric alignment are correct.
* Text and fallback layout match.

Risk:

* Very high. Text is complex and user-visible.

Stage 8 — Dynamic GPU Overlays

Files to create/change:

packages/grid/src/renderer-v2/dynamic-overlay-packet.ts
packages/grid/src/renderer-v2/typegpu-render-pass.ts
packages/grid/src/GridFillHandleOverlay.tsx
packages/grid/src/useWorkbookGridRenderState.ts
packages/grid/src/WorkbookGridSurface.tsx
e2e/tests/web-shell-typegpu.pw.ts

Responsibilities:

* Move selection fill, active border, fill handle, resize guides, header selection highlights, and frozen separators into GPU dynamic overlay packets where practical.
* Keep DOM editor only.
* Use GridGeometrySnapshot for all overlay placement.
* Remove per-scroll React overlay state updates.

Tests:

* selection pixel readback;
* fill handle readback;
* resize guide readback;
* editor overlay drift test.

Acceptance criteria:

* Selection and fill handle match cell geometry within one physical pixel.
* Editing overlay scroll has zero React commits caused by overlay movement.
* Frozen separators remain fixed.

Risk:

* Medium-high.

Stage 9 — Integrate Edits, Styles, Hidden Axes, And Freezes

Files to create/change:

packages/grid/src/useWorkbookGridRenderState.ts
packages/grid/src/WorkbookGridSurface.tsx
apps/web/src/worker-runtime-render-scene.ts
apps/web/src/worker-runtime.ts
packages/grid/src/gridGeometry.ts
packages/grid/src/gridAxisWorldIndex.ts

Responsibilities:

* Route all row/column hidden and size state into axis snapshots.
* Invalidate V2 packets on edits, style changes, row/column changes, freeze changes, and sheet switches.
* Keep selection independent of base tile invalidation.
* Ensure interactions use V2 geometry for hit testing.

Tests:

* hidden row/column Playwright test;
* variable size Playwright test;
* freeze change test;
* edit invalidation test;
* style invalidation test.

Acceptance criteria:

* Hidden axes are rendered, hit-tested, and scrolled consistently.
* Edits update visible cells without rebuilding unrelated tiles.
* Freezes do not produce pane drift.

Risk:

* High.

Stage 10 — Instrumentation And Budgets

Files to create/change:

apps/web/src/perf/workbook-scroll-perf.ts
packages/grid/src/renderer/grid-render-counters.ts
packages/grid/src/renderer-v2/render-debug-hud.ts
e2e/tests/web-shell-scroll-performance.pw.ts

Responsibilities:

* Add V2 counters.
* Add debug HUD.
* Add budget assertions.
* Save failure artifacts.

Tests:

* all performance tests;
* debug HUD smoke test.

Acceptance criteria:

* Budgets are enforced in CI for supported WebGPU browsers.
* Failure output is diagnostic enough to debug one-pixel drift or tile blanking.

Risk:

* Medium.

Stage 11 — Make V2 Default, Keep V1 Rollback

Files to create/change:

apps/web/src/renderer-flags.ts
packages/grid/src/WorkbookGridSurface.tsx
docs/workbook-typegpu-grid-renderer-oracle-design-2026-04-21.md
e2e/tests/web-shell-scroll-performance.pw.ts
e2e/tests/web-shell-typegpu.pw.ts

Responsibilities:

* Switch default renderer to typegpu-v2 only after all budgets pass.
* Keep typegpu-v1 selectable.
* Keep fallback selectable.
* Document known adapter/browser limitations.

Tests:

* full unit test suite;
* full browser perf suite;
* production smoke.

Acceptance criteria:

* V2 is default with rollback flag.
* V1 remains available for one release cycle.
* Evidence package exists.

Risk:

* Medium.

Stage 12 — Remove V1 Internals

Files to delete or retire after one stable release cycle:

packages/grid/src/renderer/WorkbookPaneRenderer.tsx
packages/grid/src/renderer/typegpu-renderer.ts
packages/grid/src/renderer/typegpu-surface-manager.ts
packages/grid/src/renderer/typegpu-resource-cache.ts
packages/grid/src/renderer/typegpu-draw-pass.ts
packages/grid/src/renderer/grid-scene-packet.ts
packages/grid/src/renderer/glyph-atlas.ts
packages/grid/src/renderer/text-quad-buffer.ts

Responsibilities:

* Remove old renderer only after production evidence and rollback window.
* Keep shared layout/fallback pieces if still useful.
* Update docs and tests.

Tests:

* full CI.
* no imports from retired files.

Acceptance criteria:

* V1 code removed without reducing coverage or fallback behavior.
* V2 is the only TypeGPU renderer.

Risk:

* Low after evidence; high if rushed.

⸻

11. Rollback And Migration Strategy

Renderer Flag

Support:

?workbookRenderer=typegpu-v2
?workbookRenderer=typegpu-v1
?workbookRenderer=canvas-fallback

and:

VITE_WORKBOOK_RENDERER=typegpu-v2
VITE_WORKBOOK_RENDERER=typegpu-v1
VITE_WORKBOOK_RENDERER=canvas-fallback

Default sequence:

1. During implementation: typegpu-v1.
2. During validation: default in local/dev can be typegpu-v2, CI runs both.
3. After evidence: production default becomes typegpu-v2.
4. Keep typegpu-v1 rollback for one release cycle.
5. Remove V1 after stable evidence.

Fallback For Browsers Without Usable WebGPU

If navigator.gpu is unavailable, adapter request fails, device request fails, or device is lost and cannot recover:

1. Switch to canvas-fallback.
2. Keep native scroll, camera store, axis, tile packets, and text layout.
3. Draw visible viewport with Canvas2D using the same ResolvedCellTextLayout.
4. Show debug capability state if HUD is enabled.
5. Do not attempt DOM rendering for the core grid.

The fallback is for correctness and compatibility, not the AAA performance target.

Comparing Old And New Renderers

Implement comparison modes:

?workbookRendererCompare=1
?workbookRenderer=typegpu-v2&typegpuCompareWith=typegpu-v1

Comparison should support:

* same camera;
* same axis snapshot;
* same selection/edit state;
* screenshots/readbacks from both paths;
* pixel diff for deterministic scenes;
* geometry JSON diff for cells/headers/selection/fill handle.

Do not render both full grids visibly in production. Use offscreen or test-only comparison where possible.

Migration Safety

* V2 must be introduced under flag.
* V1 must remain untouched until V2 unit tests and basic E2E pass.
* Worker packet V2 can coexist with current worker object scenes.
* Axis and geometry V2 can be adopted by overlays before TypeGPU V2 becomes default.
* Text layout V2 can run in debug comparison mode before replacing text draw.

⸻

12. Final Definition Of Done

The renderer is production-ready only when all checklist items pass.

Architecture Checklist

* Scroll frame camera is a full world camera, not { tx, ty }.
* Rendering, hit testing, headers, frozen panes, selection, fill handle, resize guides, and editor overlay use GridGeometrySnapshot.
* Axis indexes include variable sizes, hidden entries, total size, hit testing, and revisions.
* Axis indexes are not rebuilt during steady scroll.
* Base tile resources are retained across scroll frames.
* Worker packets are typed, transferable, versioned, and validated.
* Production packets do not carry object GridGpuScene or GridTextScene.
* Steady scroll updates only uniforms and dynamic overlays.
* Context configure never happens during steady scroll.
* Buffer allocation never happens during warmed steady scroll.
* Atlas upload never happens during warmed steady scroll.
* Selection changes do not invalidate base tiles.
* Editing overlay scroll does not cause React commits.
* Frozen panes and headers cannot drift because they share the same geometry snapshot.
* Device loss is handled.

Text Checklist

* Text layout supports left/center/right alignment.
* Numeric default alignment is right.
* Vertical alignment is supported.
* Baselines use actual or deterministic font metrics.
* Wrapping is supported.
* Overflow into adjacent empty cells is supported.
* Overflow is blocked by non-empty neighbors and hidden columns.
* Bold and italic are part of the font key.
* Underline and strike are accurately positioned.
* High-DPR atlas buckets are supported.
* Atlas pages are versioned.
* TypeGPU and fallback consume the same resolved text layout.
* Text clipping has exact geometry tests.

Performance Checklist

For warmed steady in-tile scroll:

* frame interval p95 <= 16.7 ms.
* frame interval p99 <= 24 ms.
* input-to-draw p95 <= 8 ms.
* input-to-draw p99 <= 16 ms.
* React commits during scroll = 0.
* layout shifts = 0.
* context configure = 0.
* surface resize = 0.
* GPU submits = 1 per drawn frame.
* render passes = 1 per drawn frame.
* draw calls <= 24 per frame.
* buffer allocations = 0.
* scene upload bytes = 0.
* atlas upload bytes = 0.
* tile misses = 0.
* blank tile events = 0.

For tile-boundary and diagonal scroll:

* blank tile events = 0.
* frame interval p95 <= 16.7 ms.
* frame interval p99 <= 32 ms.
* draw calls <= 32 per frame.
* stale-valid fallback is revision-safe.

Test Checklist

* gridAxisWorldIndex.test.ts covers default, variable, hidden, boundary, hit test, total size, deep anchors.
* gridCameraGeometry.test.ts covers all panes, freezes, fractional scroll, DPR, editor/fill handle geometry.
* gridTileResidencyV2.test.ts covers warm policy, hysteresis, stale-valid, eviction, revision mismatch.
* gridTextLayoutV2.test.ts covers alignment, baseline, wrap, overflow, decorations, high DPR.
* gridScenePacketV2.test.ts covers valid packets, transfer, rejection, stale revisions.
* gridRenderLoop.test.ts covers coalescing, latest camera, resize, device loss.
* gridRenderLayerOrdering.test.ts covers draw order.
* Playwright covers horizontal, vertical, diagonal, wheel/inertial, tile boundary, variable sizes, hidden axes, frozen panes, resize, editing overlay, and remote edits.
* TypeGPU readback tests assert exact pixels for lines, headers, body, frozen separators, selection, fill handle, resize guide, text clip, and editor overlay.

Commands That Must Pass

From the repository root:

pnpm install --frozen-lockfile
pnpm protocol:check
pnpm formula-inventory:check
pnpm formula:dominance:check
pnpm workspace-resolution:check
pnpm naming:check
pnpm format:check
pnpm lint
pnpm typecheck
pnpm test
pnpm --filter @bilig/grid build
pnpm --filter @bilig/web build
pnpm exec playwright install chromium
pnpm exec playwright test e2e/tests/web-shell-typegpu.pw.ts
pnpm exec playwright test e2e/tests/web-shell-scroll-performance.pw.ts
pnpm test:browser

Before enabling V2 by default, this broader command must also pass:

pnpm ci

Required Evidence Package

Before declaring production readiness, attach or commit the following evidence:

docs/render-evidence/typegpu-v2/
  summary.md
  steady-horizontal-perf.json
  steady-vertical-perf.json
  diagonal-perf.json
  tile-boundary-perf.json
  deep-vertical-perf.json
  frozen-pane-perf.json
  resize-then-scroll-perf.json
  editing-overlay-scroll-perf.json
  typegpu-readback-grid-lines.png
  typegpu-readback-frozen-panes.png
  typegpu-readback-selection-fill-handle.png
  typegpu-readback-text-overflow.png
  camera-snapshots.json
  tile-cache-snapshots.json
  atlas-stats.json
  adapter-info.json

summary.md must include:

* browser and GPU adapter;
* DPR;
* viewport sizes;
* fixture IDs;
* p95/p99 frame intervals;
* p95/p99 input-to-draw latency;
* React commit counts;
* blank tile events;
* upload bytes;
* draw call counts;
* memory high-water marks;
* known limitations.

The renderer is not production-ready if any of these are true:

* a visible tile blanks for one frame in a supported scenario;
* grid lines are not one physical pixel;
* frozen panes drift relative to body/header;
* selection or fill handle differs from geometry by more than one physical pixel;
* editor overlay scroll requires React commits;
* steady scroll allocates buffers or uploads scene data;
* text layout differs between TypeGPU and fallback;
* broad dark-pixel tests are the only proof of text rendering;
* the old design document is used as a substitute for the stricter checks above.