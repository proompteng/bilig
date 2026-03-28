# Reconciler Layering

The custom workbook reconciler is package-based, not app-local. It is not a DOM renderer and does not own spreadsheet state.

```mermaid
flowchart TB
  classDef react fill:#eaf2ff,stroke:#2563eb,color:#0f172a,stroke-width:1.5px
  classDef adapter fill:#fff7ed,stroke:#ea580c,color:#431407,stroke-width:1.5px
  classDef core fill:#eefbf3,stroke:#15803d,color:#052e16,stroke-width:1.5px
  classDef wasm fill:#f8fafc,stroke:#475569,color:#0f172a,stroke-width:1.5px

  subgraph UI["apps/web"]
    APP["Product shell"]
  end

  subgraph PKG["packages"]
    JSX["@bilig/renderer workbook DSL"]
    REC["@bilig/renderer host config"]
    UI["@bilig/grid operator UI"]
    ENG["SpreadsheetEngine"]
    CRDT["CRDT ordering and compaction"]
    FORM["Formula parser / binder / JS evaluator"]
  end

  subgraph FAST["Numeric fast path"]
    WASM["AssemblyScript / WASM kernel"]
  end

  APP --> JSX
  APP --> UI
  JSX --> REC
  UI --> ENG
  REC --> ENG
  ENG --> CRDT
  ENG --> FORM
  ENG --> WASM

  class JSX,UI react
  class REC adapter
  class ENG,CRDT,FORM core
  class WASM wasm
```

```mermaid
sequenceDiagram
  participant JSX as React JSX tree
  participant REC as Custom reconciler
  participant ENG as SpreadsheetEngine
  participant CRDT as CRDT layer
  participant WASM as WASM kernel

  JSX->>REC: render Workbook / Sheet / Cell tree
  REC->>ENG: renderCommit(commitOps)
  ENG->>CRDT: create or apply op batch
  ENG->>WASM: evaluate eligible formulas
  ENG-->>REC: batch complete
```

Rules:

- no engine mutation in `createInstance`
- descriptors are inert until commit
- one engine batch per React commit
- the reconciler does not keep a parallel workbook shadow model; it validates the descriptor tree directly and flushes semantic commit ops into the engine
- root creation and `updateContainer` calls are isolated behind a small compat layer so `react-reconciler` version drift stays contained
- React-specific code is isolated to `@bilig/renderer`, `@bilig/grid`, and the product shell
- React is a declarative authoring surface and operator UI only; the spreadsheet graph lives in `@bilig/core`
- the reconciler may translate tree diffs into semantic workbook ops, but it never owns formula, dependency, or CRDT semantics

What this means in practice:

- React should not become the canonical workbook runtime state
- React fiber/tree state should not be used as the hot-path source of truth for cells, dependency edges, history, or multiplayer merges
- the normalized engine IR remains the canonical workbook state
- a future render-patch layer should stay separate from both the JSX descriptor tree and the core workbook graph
