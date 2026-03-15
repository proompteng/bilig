# Public API

## React workbook DSL

```tsx
<Workbook>
  <Sheet name="Sheet1">
    <Cell addr="A1" value={10} />
    <Cell addr="B1" formula="A1*2" />
    <Cell addr="C1" format="currency-usd" value={42} />
  </Sheet>
</Workbook>
```

## Package exports

- `@bilig/renderer`
  - `Workbook`
  - `Sheet`
  - `Cell`
  - `createWorkbookRendererRoot(engine)`
- `@bilig/grid`
  - `WorkbookView` (Excel-like workbook shell built on Glide Data Grid)
  - `SheetGridView`
  - `FormulaBar` (name box + formula input directly above the grid)
  - `CellEditorOverlay`
  - `MetricsPanel`
  - `DependencyInspector`
  - `ReplicaPanel`
  - `useCell`
  - `useMetrics`
  - `useSelection`
  - `useSheetViewport`

## Imperative engine

- `createSheet(name)`
- `deleteSheet(name)`
- `setCellValue(sheet, address, value)`
- `setCellFormula(sheet, address, formula)`
- `setCellFormat(sheet, address, format)`
- `clearCell(sheet, address)`
- `getCell(sheet, address)`
- `getDependencies(sheet, address)`
- `getDependents(sheet, address)`
- `explainCell(sheet, address)`
- `exportSnapshot()`
- `importSnapshot(snapshot)`
- `exportSheetCsv(sheet)`
- `importSheetCsv(sheet, csv)`
- `subscribe(listener)`
- `subscribeBatches(listener)`
- `applyRemoteBatch(batch)`
- `exportReplicaSnapshot()`
- `importReplicaSnapshot(snapshot)`

## Local-first replication

Replication is transport-agnostic. The engine emits deterministic local batches through `subscribeBatches(listener)` and accepts remote batches through `applyRemoteBatch(batch)`.

- local and remote mutations use the same apply path
- conflict policy is deterministic last-writer-wins per entity
- entity ordering is `clock.counter`, then `replicaId`, then `batchId`, then `opIndex`
- replica snapshots persist:
  - replica clock
  - applied batch ids
  - entity version map
  - sheet delete tombstones

## Renderer root

`@bilig/renderer` owns the custom workbook reconciler and exposes a minimal async root surface:

- `render(element)`
- `unmount()`

## Playground interaction contract

The playground keeps the core/renderer APIs unchanged, but the shell behavior is part of the shipped product contract:

- the formula bar lives directly above the grid
- the grid exposes the full spreadsheet surface, including Excel-scale row and column bounds
- large workbook presets are loaded by the playground app layer, not by `@bilig/core`
- editing works from both the formula bar and the in-cell overlay
