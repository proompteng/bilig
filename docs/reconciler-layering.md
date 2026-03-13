# Reconciler Layering

React is playground-only. The custom workbook reconciler is not a DOM renderer and does not own spreadsheet state.

```mermaid
sequenceDiagram
  participant JSX as React JSX tree
  participant REC as Custom reconciler
  participant ENG as SpreadsheetEngine
  participant WASM as WASM kernel

  JSX->>REC: render Workbook / Sheet / Cell tree
  REC->>ENG: renderCommit(commitOps)
  ENG->>WASM: evaluate eligible formulas
  ENG-->>REC: batch complete
```

Rules:

- no engine mutation in `createInstance`
- descriptors are inert until commit
- one engine batch per React commit
- shared packages remain React-free
