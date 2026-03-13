# Reconciler Layering

React is playground-only. The custom workbook reconciler is not a DOM renderer and does not own spreadsheet state.

```mermaid
flowchart TB
  classDef react fill:#eaf2ff,stroke:#2563eb,color:#0f172a,stroke-width:1.5px
  classDef adapter fill:#fff7ed,stroke:#ea580c,color:#431407,stroke-width:1.5px
  classDef core fill:#eefbf3,stroke:#15803d,color:#052e16,stroke-width:1.5px
  classDef wasm fill:#f8fafc,stroke:#475569,color:#0f172a,stroke-width:1.5px

  subgraph PG["apps/playground"]
    JSX["React JSX Workbook DSL"]
    REC["Custom reconciler host config"]
    UI["React DOM operator UI"]
  end

  subgraph CORE["Framework-agnostic runtime"]
    ENG["SpreadsheetEngine"]
    CRDT["CRDT ordering and compaction"]
    FORM["Formula parser / binder / JS evaluator"]
  end

  subgraph FAST["Numeric fast path"]
    WASM["AssemblyScript / WASM kernel"]
  end

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
- shared packages remain React-free
- React is a declarative authoring surface and operator UI only; the spreadsheet graph lives in `@bilig/core`
- the reconciler may translate tree diffs into semantic workbook ops, but it never owns formula, dependency, or CRDT semantics
