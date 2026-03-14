# Architecture

```mermaid
flowchart TB
  classDef react fill:#eaf2ff,stroke:#2563eb,color:#0f172a,stroke-width:1.5px
  classDef core fill:#eefbf3,stroke:#15803d,color:#052e16,stroke-width:1.5px
  classDef wasm fill:#fff7ed,stroke:#ea580c,color:#431407,stroke-width:1.5px

  subgraph PG["apps/playground"]
    APP["Thin React app shell"]
  end

  subgraph PKG["packages"]
    CORE["@bilig/core<br/>engine + range registry + edge arena + scheduler"]
    FORMULA["@bilig/formula"]
    CRDT["@bilig/crdt"]
    PROTOCOL["@bilig/protocol"]
    RENDERER["@bilig/renderer"]
    UI["@bilig/grid"]
    WASM["@bilig/wasm-kernel"]
  end

  APP --> UI
  APP --> RENDERER
  UI --> CORE
  RENDERER --> CORE
  CORE --> FORMULA
  CORE --> CRDT
  CORE --> PROTOCOL
  CORE --> WASM
  FORMULA --> PROTOCOL
  WASM --> PROTOCOL

  class APP react
  class CORE,FORMULA,CRDT,PROTOCOL,RENDERER,UI core
  class WASM wasm
```

```mermaid
sequenceDiagram
  participant UI as React playground shell
  participant REC as @bilig/renderer
  participant CORE as @bilig/core
  participant CRDT as @bilig/crdt
  participant WASM as @bilig/wasm-kernel

  UI->>REC: render Workbook / Sheet / Cell tree
  REC->>CORE: renderCommit(commitOps)
  CORE->>CRDT: create local EngineOpBatch
  CORE->>WASM: evaluate eligible numeric runs
  CORE-->>UI: emit batch event + targeted selector updates
  CORE-->>CRDT: emit outbound batch stream
```

The TS protocol enums/opcodes and the AssemblyScript protocol mirror are generated together from `scripts/gen-protocol.mjs`. That keeps the JS/WASM contract deterministic and makes drift a CI failure instead of a runtime surprise.

Within `@bilig/core`, the runtime is no longer a single inline dependency map. The current production shape is:

- `WorkbookStore` for sheet metadata, sparse grids, and typed-array-backed cells
- `RangeRegistry` for interned range entities, a shared range-member pool, descriptor `membersOffset`/`membersLength`, and dynamic row/column membership tracking
- `EdgeArena` for forward and reverse graph slices
- `RecalcScheduler` for epoch-based dirty propagation and rank-bucket ordering
- `cycle-detection` for deterministic SCC grouping and `cycleGroupIds`, now backed by reusable typed-array Tarjan scratch state instead of per-run `Map`/`Set` allocations
- shared program/constant/range arenas in the engine so formula metadata matches the packed runtime/WASM contract instead of ad hoc per-formula blobs
- vectorized topo rebuild scratch state in the engine so rank assignment no longer depends on `Map`/`Set` queue construction in the hot topology path
- typed-array impacted-formula scratch in the engine so sheet deletion and formula rebinding no longer build recursive `Set` unions over the entity graph
- typed-array mutation-root buffers in the engine so snapshot import and op-batch application no longer allocate `Set` unions before recalculation
- typed-array rebound tracking in the engine so sheet rebinds and dynamic-range growth mark affected formulas directly instead of returning intermediate `Set` collections
- reusable WASM upload scratch in the engine/facade so fast-path program and range sync no longer rebuild filtered formula lists on every upload
- callback-based sheet scanning for dynamic ranges so the range registry no longer asks the engine to materialize throwaway `{ cellIndex, row, col }[]` snapshots just to discover members
- typed-array dynamic range materialization so callback-based sheet scans now fill packed `Uint32Array` member lists directly instead of boxing matches into `number[]` first

The UI does not subscribe through a single global revision for visible cells. `@bilig/core` maintains keyed cell listener routing so `useCell(...)` and viewport watchers wake only when one of their watched addresses changes. That keeps the grid aligned with the production requirement for localized rerenders rather than whole-viewport invalidation on every batch.
