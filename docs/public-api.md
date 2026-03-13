# Public API

## React workbook DSL

```tsx
<Workbook>
  <Sheet name="Sheet1">
    <Cell addr="A1" value={10} />
    <Cell addr="B1" formula="A1*2" />
  </Sheet>
</Workbook>
```

## Imperative engine

- `createSheet(name)`
- `deleteSheet(name)`
- `setCellValue(sheet, address, value)`
- `setCellFormula(sheet, address, formula)`
- `clearCell(sheet, address)`
- `getCell(sheet, address)`
- `getDependencies(sheet, address)`
- `getDependents(sheet, address)`
- `exportSnapshot()`
- `importSnapshot(snapshot)`
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

The custom workbook reconciler remains internal to the playground and exposes a minimal async root surface:

- `render(element)`
- `unmount()`
