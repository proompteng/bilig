# Native Workbook Surface

## Status

The workbook runtime no longer mounts Glide. The active surface is
[WorkbookGridSurface.tsx](/Users/gregkonush/github.com/bilig/packages/grid/src/WorkbookGridSurface.tsx),
which owns selection, hit testing, scrolling, editing overlays, and renderer scene
assembly directly.

## Current architecture

- WebGPU renders the workbook rect plane through
  [GridGpuSurface.tsx](/Users/gregkonush/github.com/bilig/packages/grid/src/GridGpuSurface.tsx).
- Visible text is rendered by
  [GridTextOverlay.tsx](/Users/gregkonush/github.com/bilig/packages/grid/src/GridTextOverlay.tsx)
  from the scene produced by
  [gridTextScene.ts](/Users/gregkonush/github.com/bilig/packages/grid/src/gridTextScene.ts).
- Selection, fill, resize, hover, and header chrome are assembled in
  [gridGpuScene.ts](/Users/gregkonush/github.com/bilig/packages/grid/src/gridGpuScene.ts).
- Workbook behavior is coordinated by
  [WorkbookView.tsx](/Users/gregkonush/github.com/bilig/packages/grid/src/WorkbookView.tsx)
  and
  [WorkbookGridSurface.tsx](/Users/gregkonush/github.com/bilig/packages/grid/src/WorkbookGridSurface.tsx).

## Completed migration points

- Glide is removed from the mounted workbook runtime and package dependency graph.
- Selection state uses local grid types instead of third-party grid selection
  objects.
- Column resize, range drag, fill handle, select-all, and keyboard selection all run
  through local interaction code in `@bilig/grid`.
- The workbook shell owns its own sheet tabs, formula bar, and status chrome.

## Current limitations

- `WorkbookGridSurface.tsx` is still too large and mixes controller logic with view
  composition.
- Text rendering is still DOM-based rather than GPU-based.
- Autofit measurement still depends on browser text measurement rather than a shared
  renderer-native text metric path.

## Cleanup priorities

1. Continue extracting controller logic from
   [WorkbookGridSurface.tsx](/Users/gregkonush/github.com/bilig/packages/grid/src/WorkbookGridSurface.tsx)
   into focused interaction and viewport modules.
2. Keep removing stale migration hooks, debug attributes, and historical naming as
   code paths stabilize.
3. Tighten browser regression coverage around visual parity, selection behavior, and
   scroll performance without relying on brittle pixel-sampling tests.

## Historical note

This file now exists as a concise implementation record. The older migration-phase
plan was deleted because it no longer described the current code accurately.
