# Implementation-Ready Design Document
## React Custom Renderer + AssemblyScript Spreadsheet Engine
Last updated: 2026-03-14

This document is the source of truth for implementing a production-grade, local-first spreadsheet runtime that uses:

- React for a declarative workbook DSL and application UI
- `react-reconciler` for a custom workbook renderer
- TypeScript for orchestration, parsing, graph maintenance, and editor integration
- AssemblyScript/WebAssembly for the hot numeric evaluation path

The document is intentionally written so an autonomous coding agent can implement the system end to end without inventing architecture along the way.

Overrides already applied in this repository:

- reusable React code lives in `packages/renderer` and `packages/grid`
- `apps/playground` is a thin shell that composes those packages
- the custom reconciler is package-scoped and isolated from the rest of the repo

---

## 1. Target outcome

Build a shippable v1 with these properties:

1. A declarative React API for workbooks, sheets, and cells.
2. A high-performance calculation engine with incremental recalculation.
3. A WASM fast path for numeric formulas and range aggregations.
4. A React DOM playground app with an Excel-like workbook shell, a Glide-backed virtualized grid, and a formula bar directly above the sheet.
5. A stable imperative engine API that can be used without React.
6. A robust test/benchmark/CI setup.
7. A version-pinned adapter around `react-reconciler` so React internals are isolated to one package.

This is not a full Excel clone. It is a production-grade architecture for a strong spreadsheet core and a shippable v1.

The current playground interaction and layout contract is specified in [playground-excel-shell-rfc.md](/Users/gregkonush/github.com/bilig/docs/playground-excel-shell-rfc.md). That RFC is authoritative for the Excel-like shell layout, Glide integration, editing model, and large-preset UX.

---

## 2. Fixed decisions

These decisions are already made and should not be re-litigated during implementation:

- Monorepo managed with `pnpm` workspaces.
- ESM-only repository.
- Node 24 LTS baseline.
- Browser-first runtime; Node support is required for tests, fixtures, snapshots, and benchmarks.
- Vite-based playground app.
- TypeScript orchestrates parsing, graph updates, and UI subscriptions.
- AssemblyScript handles the numeric VM fast path.
- `react-reconciler` is wrapped in a dedicated package and pinned exactly.
- React renderer host instances are pure descriptors. No engine mutation is allowed during `createInstance`.
- Core engine remains usable imperatively without any React dependency.
- Formula language uses A1 notation in v1.
- Full XLSX parity is not in v1.

---

## 3. Scope

### 3.1 In scope for v1

Core model:
- workbook
- sheets
- cells
- named workbook metadata
- JSON snapshot import/export
- CSV import/export for single sheets

Formula language:
- literals: number, string, boolean, blank
- scalar references: `A1`, `Sheet2!B5`
- rectangular ranges: `A1:B10`, `Sheet2!A:A`, `1:10`
- operators: `+ - * / ^ & = <> > >= < <=`
- unary operators: `+ -`
- parentheses
- built-ins:
  - `SUM`
  - `AVG`
  - `MIN`
  - `MAX`
  - `COUNT`
  - `COUNTA`
  - `ABS`
  - `ROUND`
  - `FLOOR`
  - `CEILING`
  - `MOD`
  - `IF`
  - `AND`
  - `OR`
  - `NOT`
  - `LEN`
  - `CONCAT`

Engine:
- incremental recalculation for value edits
- full topology rebuild when formulas change
- cycle detection
- dirty propagation
- cached topological rank
- JS fallback evaluator
- WASM numeric fast path
- range interning
- materialization of referenced empty cells for bounded ranges
- sheet-local and cross-sheet dependencies

React integration:
- custom renderer package
- declarative `<Workbook>`, `<Sheet>`, `<Cell>`
- imperative root API returning a Promise from render
- `useSyncExternalStore` based hooks for UI integration

Playground UI:
- Excel-like workbook shell
- Glide-backed virtualized grid
- cell selection
- formula bar directly above the grid
- name box directly above the grid
- edit/commit flow through both the formula bar and the in-cell editor
- large preset loaders for stress scenarios and Excel-scale surface demos
- recalc metrics panel
- dependency inspector for selected cell

Quality:
- unit tests
- integration tests
- browser smoke tests
- benchmark harness
- GitHub Actions CI

### 3.2 Explicitly out of scope for v1

Do not spend implementation time on:
- full XLSX import/export parity
- pivot tables
- charts
- array spill formulas
- `XLOOKUP`, `INDEX/MATCH`, `VLOOKUP`
- volatile date/time parity beyond a minimal stub
- locale-aware number parsing
- collaborative editing transport
- server synchronization
- macros / scripting
- formatting engine parity with Excel / Sheets
- workbook protection / permissions
- worker-based off-main-thread transport as the default path

The architecture should be worker-ready, but v1 should ship with an in-process engine and a clean abstraction that allows a later worker transport.

---

## 4. Success metrics

These are engineering targets, not marketing claims:

- Initial load of a 100k-materialized-cell workbook under 1.5s on a modern laptop in production mode.
- Single literal edit with 10k downstream formula cells:
  - median recalc < 50ms
  - p95 recalc < 120ms
- No visible UI jank for typical edits affecting < 2k downstream formulas.
- Memory ceiling for a 100k-materialized-cell workbook under 250MB in the browser.
- Browser bundle:
  - main app JS < 350KB gzip excluding WASM
  - release WASM < 250KB gzip
- 90%+ line coverage for parser, compiler, graph, and renderer packages.
- Formula parity tests must pass identically in JS and WASM fast-path overlap zones.

---

## 5. High-level architecture

```text
React DOM UI
    │
    ├── useSyncExternalStore hooks
    │
    ▼
WorkbookEngine (TypeScript)
    ├── Workbook/Sheet metadata
    ├── CellStore (typed arrays)
    ├── SheetGrid sparse block index
    ├── RangeRegistry
    ├── DependencyGraph / EdgeArena
    ├── Formula compiler
    ├── JS evaluator
    ├── Scheduler
    └── WasmKernelFacade
            │
            ▼
      AssemblyScript VM
        ├── numeric stack machine
        ├── range aggregation kernels
        ├── typed-array cell views
        └── fast-path formula evaluation

Custom React Renderer
    ├── pure host descriptors
    ├── commit transaction log
    └── engine commit adapter
```

Guiding principle:

- Keep irregular, string-heavy, topology-heavy work in TypeScript.
- Keep numeric, repeated, dense evaluation in AssemblyScript.
- Minimize JS↔WASM crossings by evaluating ordered batches, not individual cells.

---

## 6. Repository layout

```text
react-asm-sheet/
├─ apps/
│  └─ playground/
│     ├─ src/
│     │  ├─ main.tsx
│     │  ├─ App.tsx
│     │  ├─ app.css
│     │  ├─ demoWorkbook.tsx
│     │  └─ fixtures/
│     ├─ index.html
│     ├─ package.json
│     └─ vite.config.ts
├─ packages/
│  ├─ protocol/
│  │  ├─ src/
│  │  │  ├─ constants.ts
│  │  │  ├─ enums.ts
│  │  │  ├─ opcodes.ts
│  │  │  ├─ types.ts
│  │  │  └─ index.ts
│  │  └─ package.json
│  ├─ formula/
│  │  ├─ src/
│  │  │  ├─ lexer.ts
│  │  │  ├─ parser.ts
│  │  │  ├─ ast.ts
│  │  │  ├─ binder.ts
│  │  │  ├─ optimizer.ts
│  │  │  ├─ compiler.ts
│  │  │  ├─ js-evaluator.ts
│  │  │  ├─ builtins.ts
│  │  │  ├─ addressing.ts
│  │  │  └─ index.ts
│  │  └─ package.json
│  ├─ core/
│  │  ├─ src/
│  │  │  ├─ engine.ts
│  │  │  ├─ workbook-store.ts
│  │  │  ├─ sheet-grid.ts
│  │  │  ├─ cell-store.ts
│  │  │  ├─ edge-arena.ts
│  │  │  ├─ range-registry.ts
│  │  │  ├─ scheduler.ts
│  │  │  ├─ cycle-detection.ts
│  │  │  ├─ snapshot.ts
│  │  │  ├─ events.ts
│  │  │  ├─ selectors.ts
│  │  │  ├─ string-pool.ts
│  │  │  ├─ wasm-facade.ts
│  │  │  └─ index.ts
│  │  └─ package.json
│  ├─ wasm-kernel/
│  │  ├─ assembly/
│  │  │  ├─ index.ts
│  │  │  ├─ vm.ts
│  │  │  ├─ builtins.ts
│  │  │  ├─ memory.ts
│  │  │  ├─ protocol.ts
│  │  │  └─ assertions.ts
│  │  ├─ build/
│  │  ├─ scripts/
│  │  │  └─ build.mjs
│  │  ├─ asconfig.json
│  │  └─ package.json
│  ├─ renderer/
│  │  ├─ src/
│  │  │  ├─ host-config.ts
│  │  │  ├─ renderer-root.ts
│  │  │  ├─ descriptors.ts
│  │  │  ├─ commit-log.ts
│  │  │  ├─ components.tsx
│  │  │  ├─ validation.ts
│  │  │  └─ index.ts
│  │  └─ package.json
│  ├─ grid/
│  │  ├─ src/
│  │  │  ├─ WorkbookView.tsx
│  │  │  ├─ SheetGridView.tsx
│  │  │  ├─ CellEditorOverlay.tsx
│  │  │  ├─ FormulaBar.tsx
│  │  │  ├─ useCell.ts
│  │  │  ├─ useViewport.ts
│  │  │  ├─ useSelection.ts
│  │  │  ├─ MetricsPanel.tsx
│  │  │  ├─ DependencyInspector.tsx
│  │  │  └─ index.ts
│  │  └─ package.json
│  └─ benchmarks/
│     ├─ src/
│     │  ├─ generate-workbook.ts
│     │  ├─ benchmark-edit.ts
│     │  ├─ benchmark-load.ts
│     │  └─ benchmark-renderer.ts
│     └─ package.json
├─ fixtures/
│  ├─ formulas/
│  ├─ workbooks/
│  └─ snapshots/
├─ scripts/
│  ├─ gen-protocol.mjs
│  ├─ perf-smoke.mjs
│  ├─ release-check.mjs
│  └─ sync-fixtures.mjs
├─ .github/workflows/
│  └─ ci.yml
├─ pnpm-workspace.yaml
├─ package.json
├─ tsconfig.base.json
├─ tsconfig.json
├─ vitest.workspace.ts
└─ README.md
```

### 6.1 Root scripts

Root `package.json` should expose:

```json
{
  "name": "react-asm-sheet",
  "private": true,
  "type": "module",
  "engines": {
    "node": ">=24.14.0"
  },
  "packageManager": "pnpm@10",
  "scripts": {
    "dev": "pnpm --filter @react-asm-sheet/playground dev",
    "build": "pnpm -r build",
    "typecheck": "tsc -b --pretty false",
    "test": "vitest --run",
    "test:watch": "vitest",
    "bench": "pnpm --filter @react-asm-sheet/benchmarks bench",
    "bench:smoke": "node scripts/perf-smoke.mjs",
    "lint": "eslint .",
    "wasm:build": "pnpm --filter @react-asm-sheet/wasm-kernel build",
    "ci": "pnpm wasm:build && pnpm typecheck && pnpm test && pnpm bench:smoke && pnpm --filter @react-asm-sheet/playground build"
  }
}
```

Keep the root boring. Do not add Turborepo unless needed after profiling the developer workflow.

---

## 7. Toolchain and dependency policy

### 7.1 Version policy

Hard pin the unstable and ABI-sensitive layers:

- `react`: exact
- `react-dom`: exact
- `react-reconciler`: exact
- `assemblyscript`: exact

Allow normal semver ranges for non-critical tooling where appropriate, but keep the lockfile committed.

### 7.2 Critical packages

Use these exact versions initially:

```json
{
  "dependencies": {
    "react": "19.2.4",
    "react-dom": "19.2.4",
    "react-reconciler": "0.33.0"
  },
  "devDependencies": {
    "assemblyscript": "0.28.10",
    "typescript": "5.9.3",
    "vite": "8.0.0",
    "vitest": "4.1.0",
    "tsx": "4.21.0"
  }
}
```

For the playground app, use Vite's React+TS scaffold as the starting point and align it to the pinned React versions above.

### 7.3 TypeScript settings

Base `tsconfig.base.json`:

```json
{
  "compilerOptions": {
    "target": "ES2023",
    "module": "ESNext",
    "moduleResolution": "Bundler",
    "lib": ["ES2023", "DOM", "DOM.Iterable"],
    "jsx": "react-jsx",
    "strict": true,
    "noUncheckedIndexedAccess": true,
    "exactOptionalPropertyTypes": true,
    "useDefineForClassFields": true,
    "skipLibCheck": true,
    "declaration": true,
    "sourceMap": true,
    "composite": true,
    "baseUrl": ".",
    "paths": {
      "@react-asm-sheet/protocol": ["packages/protocol/src/index.ts"],
      "@react-asm-sheet/formula": ["packages/formula/src/index.ts"],
      "@react-asm-sheet/core": ["packages/core/src/index.ts"],
      "@react-asm-sheet/renderer": ["packages/renderer/src/index.ts"],
      "@react-asm-sheet/grid": ["packages/grid/src/index.ts"]
    }
  }
}
```

Use project references from the root `tsconfig.json`.

---

## 8. Shared protocol package

The `protocol` package is the single source of truth for all values shared between TS and AssemblyScript.

### 8.1 Required exports

```ts
export enum ValueTag {
  Empty = 0,
  Number = 1,
  Boolean = 2,
  String = 3,
  Error = 4
}

export enum ErrorCode {
  None = 0,
  Div0 = 1,
  Ref = 2,
  Value = 3,
  Name = 4,
  NA = 5,
  Cycle = 6
}

export enum FormulaMode {
  JsOnly = 0,
  WasmFastPath = 1
}

export enum Opcode {
  PushNumber = 1,
  PushBoolean = 2,
  PushCell = 3,
  PushRange = 4,
  Add = 5,
  Sub = 6,
  Mul = 7,
  Div = 8,
  Pow = 9,
  Concat = 10,
  Neg = 11,
  Eq = 12,
  Neq = 13,
  Gt = 14,
  Gte = 15,
  Lt = 16,
  Lte = 17,
  Jump = 18,
  JumpIfFalse = 19,
  CallBuiltin = 20,
  Ret = 255
}

export enum BuiltinId {
  Sum = 1,
  Avg = 2,
  Min = 3,
  Max = 4,
  Count = 5,
  CountA = 6,
  Abs = 7,
  Round = 8,
  Floor = 9,
  Ceiling = 10,
  Mod = 11,
  If = 12,
  And = 13,
  Or = 14,
  Not = 15,
  Len = 16,
  Concat = 17
}
```

### 8.2 Generation rule

Maintain one JSON or TS manifest in `scripts/gen-protocol.mjs` and generate:

- `packages/protocol/src/opcodes.ts`
- `packages/wasm-kernel/assembly/protocol.ts`

Do not hand-maintain duplicate opcode values.

### 8.3 Shared TypeScript aliases

Define these shared aliases early so every package uses the same terminology:

```ts
export type CellIndex = number;
export type FormulaId = number;
export type RangeIndex = number;
export type EntityId = number;
export type LiteralInput = number | string | boolean | null;
```

---

## 9. Addressing model

### 9.1 User-facing address forms

Support:

- `A1`
- `$A$1`
- `Sheet1!A1`
- `'My Sheet'!B3`
- row ranges: `1:10`
- column ranges: `A:C`
- full-sheet qualified ranges: `Sheet2!A1:C99`

Ignore relative/absolute behavior beyond storing the absolute parse result in v1. There is no copy/paste formula rebasing requirement in v1.

### 9.2 Internal identity

Use two identities:

1. `cellKey`: stable address-derived numeric key in TS
2. `cellIndex`: dense contiguous index into typed arrays

Do not use string addresses in hot paths.

---

## 10. Core data model

### 10.1 Main engine object

```ts
export interface WorkbookEngineOptions {
  workbookName?: string;
  maxRows?: number;
  maxCols?: number;
  maxSheets?: number;
}
```

`WorkbookStore` owns:

- sheet metadata
- `cellKey -> cellIndex`
- cell typed arrays
- sheet sparse grid structures
- formula metadata
- topo ranks
- dirty/version arrays

Use dense typed arrays instead of objects-per-cell.

Required arrays:

```ts
tags: Uint8Array;
numbers: Float64Array;
stringIds: Uint32Array;
errors: Uint16Array;
formulaIds: Uint32Array;
versions: Uint32Array;
flags: Uint32Array;
sheetIds: Uint16Array;
rows: Uint32Array;
cols: Uint16Array;
topoRanks: Uint32Array;
cycleGroupIds: Int32Array;
```

Keep the arrays separate. Do not wrap them in per-cell classes.

### 10.2 Flags bit layout

```ts
export const enum CellFlags {
  Dirty = 1 << 0,
  HasFormula = 1 << 1,
  JsOnly = 1 << 2,
  InCycle = 1 << 3,
  Materialized = 1 << 4,
  PendingDelete = 1 << 5
}
```

### 10.3 String pool

Keep strings out of the numeric hot path.

```ts
class StringPool {
  private readonly byValue = new Map<string, number>();
  private readonly values: string[] = [""];
  intern(value: string): number;
  get(id: number): string;
}
```

### 10.4 Sheet storage

Each sheet gets a sparse block map:

```ts
type BlockKey = number;

class SheetGrid {
  readonly blocks = new Map<BlockKey, Uint32Array>();
  get(row: number, col: number): number;
  set(row: number, col: number, cellIndex: number): void;
  clear(row: number, col: number): void;
  forEachInRange(r1: number, c1: number, r2: number, c2: number, fn: (cellIndex: number) => void): void;
}
```

Use:
- block rows = 128
- block cols = 32

### 10.5 Formula table

Separate formula metadata from cell arrays.

```ts
interface FormulaRecord {
  id: number;
  source: string;
  mode: FormulaMode;
  depsPtr: number;
  depsLen: number;
  programOffset: number;
  programLength: number;
  constNumberOffset: number;
  constNumberLength: number;
  rangeListOffset: number;
  rangeListLength: number;
  maxStackDepth: number;
}
```

Store formulas in vectors, not maps.

---

## 11. Graph model

### 11.1 Entity kinds

Graph propagation uses two entity types:

- cell nodes
- range nodes

A range node is an interned rectangular range with a stable identity.

Represent entity ids as a 32-bit tagged integer:

- high bit `0` => cell entity, payload = `cellIndex`
- high bit `1` => range entity, payload = `rangeIndex`

### 11.2 Why range nodes exist

Do not attach formulas directly to every cell in a range unless profiling proves it is better.

Range nodes let us:
- dedupe identical ranges across formulas
- keep formula dependencies compact
- support WASM aggregate kernels efficiently
- avoid N×M reverse edge blowups for reused ranges

### 11.3 RangeRegistry

```ts
interface RangeDescriptor {
  sheetId: number;
  row1: number;
  col1: number;
  row2: number;
  col2: number;
  membersOffset: number;
  membersLength: number;
  refCount: number;
}
```

Rules:
- canonicalize ranges
- dedupe by canonical key
- for bounded ranges smaller than the WASM cap, pre-materialize referenced empty cells and store member indices
- for unbounded or huge ranges, force JS-only execution mode

### 11.4 Edge storage

Use an arena, not nested JS arrays of arrays in hot paths.

```ts
interface EdgeSlice {
  ptr: number;
  len: number;
  cap: number;
}
```

Maintain:
- forward dependencies per formula
- reverse dependencies per entity

### 11.5 Reverse dependency policy

Reverse dependencies are attached to:
- cell entities
- range entities

When compiling formula `F`:
- each direct scalar dependency adds reverse edge `cell -> F`
- each range dependency adds reverse edge `range -> F`
- each cell inside the range adds reverse edge `cell -> range`

This creates a two-hop propagation model:
`cell -> range -> formula`

---

## 12. Formula language and compiler

### 12.1 Pipeline

```text
source string
   ↓
lexer
   ↓
parser (AST)
   ↓
binder / resolver
   ↓
optimizer / constant folder
   ↓
execution planner
   ├── JS fallback plan
   └── WASM bytecode + pools
```

### 12.2 Lexer

Implement a dedicated lexer. Do not use regex-only parsing.

Support quoted sheet names:
- `'My Sheet'!A1`

### 12.3 Parser

Use a Pratt parser.

AST node kinds:
- Literal
- CellRef
- RangeRef
- UnaryExpr
- BinaryExpr
- CallExpr

### 12.4 Binder

The binder resolves:
- function names
- sheet names
- scalar refs
- range refs

It also collects dependencies and decides whether the formula is eligible for the WASM fast path.

### 12.5 Fast-path eligibility rules

A formula can run in WASM only if all of the following are true:

- every builtin used is in the WASM-supported set
- every string operation is absent
- every range is bounded and materializable under the fast-path cap
- there is no custom JS-only function
- there is no unsupported literal/object semantic
- the computed max stack depth is <= `MAX_VM_STACK`

Otherwise compile to JS-only.

### 12.6 Constant folding

Perform compile-time simplifications:
- fold arithmetic on literals
- fold boolean branches where condition is literal
- flatten nested concat calls
- normalize commutative numeric call shapes where useful
- strip redundant parentheses

### 12.7 Builtin support matrix

WASM fast path in v1:
- `SUM`
- `AVG`
- `MIN`
- `MAX`
- `COUNT`
- `COUNTA`
- `ABS`
- `ROUND`
- `FLOOR`
- `CEILING`
- `MOD`
- `IF`
- `AND`
- `OR`
- `NOT`

JS-only in v1:
- `LEN`
- `CONCAT`
- any custom function

### 12.8 Bytecode representation

Use a `Uint32Array` instruction stream.

Encoding:
- high 8 bits = opcode
- low 24 bits = operand / small immediate

### 12.9 Program arena

The formula compiler should concatenate all bytecode into one `Uint32Array` program arena:

```ts
class ProgramArena {
  append(program: Uint32Array): { offset: number; length: number };
  replace(offset: number, length: number, next: Uint32Array): { offset: number; length: number };
}
```

Do the same for numeric constant pools and range-id lists if needed.

---

## 13. JS fallback evaluator

The JS evaluator exists for correctness and coverage, not just as a backup.

It must:
- implement all v1 formula features
- share value-tag semantics with the WASM path
- be deterministic
- be easy to debug
- expose the same error behavior as the fast path where overlap exists

Use AST or a lowered JS instruction plan; do not use `eval`, `new Function`, or any text-to-JS codegen.

The JS evaluator defines the authoritative semantics for:
- text behavior
- coercion rules
- error propagation
- custom functions

---

## 14. Recalculation model

### 14.1 Value edit flow

For a literal edit:
1. update cell payload in store
2. increment `versions[cellIndex]`
3. clear formula flags if this cell changed from formula to literal
4. dirty-propagate through reverse dependencies
5. order dirty formulas by cached topo rank
6. evaluate dirty formulas in rank order
7. emit changed-cell events once, after the batch

### 14.2 Formula edit flow

For a formula edit:
1. parse and compile source
2. update formula table entry
3. rebuild forward deps for that formula
4. patch reverse deps for old and new dependencies
5. rebuild topo ranks for the workbook
6. detect cycles
7. recalculate all impacted formulas

### 14.3 Dirty propagation

Use an epoch-based BFS to avoid clearing boolean arrays.

### 14.4 Cached topo ranks

After every topology-changing formula edit, recompute a workbook-global rank for every formula cell.

Store `topoRanks[cellIndex]`.

### 14.5 Cycle detection

Run Tarjan SCC or Kahn+back-edge detection after formula topology changes.

Rules:
- SCC size > 1 => cycle
- self-loop => cycle

For cyclic cells:
- mark `CellFlags.InCycle`
- assign `ErrorCode.Cycle`
- exclude them from fast-path evaluation
- propagate cycle error to downstream dependents through normal error semantics

### 14.6 Batch partitioning by evaluation mode

Once formulas are rank-ordered, partition consecutive runs by mode:

- WASM fast path run
- JS-only run

Then evaluate:
- one `wasm.evalBatch` call per contiguous WASM run
- direct JS evaluation for JS runs

---

## 15. Value semantics

Keep semantics simple and explicit in v1:
- empty in numeric context => `0`
- empty in string context => `""`
- booleans in numeric context => `1` or `0`
- text in numeric aggregate:
  - `COUNT` ignores
  - `COUNTA` includes
  - `SUM`/`AVG` ignore non-numeric text in v1
- division by zero => `ErrorCode.Div0`
- invalid reference => `ErrorCode.Ref`
- invalid function name => `ErrorCode.Name`

Use exact numeric comparison in v1.

---

## 16. AssemblyScript kernel

AssemblyScript handles:
- numeric stack VM execution
- numeric builtins
- range aggregations over member lists
- direct writes into typed-array value buffers

It does not handle:
- parsing
- string interning
- dynamic graph mutation
- sheet name resolution
- custom function registration

Use `runtime: "incremental"` for v1 release builds.

Required exports:

```ts
export function init(cellCapacity: i32, formulaCapacity: i32, rangeCapacity: i32, memberCapacity: i32): void;
export function ensureCellCapacity(nextCapacity: i32): void;
export function ensureFormulaCapacity(nextCapacity: i32): void;
export function ensureRangeCapacity(nextCapacity: i32): void;
export function uploadPrograms(programs: Uint32Array, offsets: Uint32Array, lengths: Uint32Array): void;
export function uploadRangeMembers(members: Uint32Array, offsets: Uint32Array, lengths: Uint32Array): void;
export function evalBatch(cellIndices: Uint32Array): void;
```

No string callbacks in the numeric hot path.

---

## 17. WasmKernelFacade

`packages/core/src/wasm-facade.ts` wraps:
- module instantiation
- memory view refresh
- upload of programs and ranges
- batch evaluation calls
- pointer-based typed-array synchronization

Ownership rules:
- JS owns source-of-truth metadata and topology
- WASM owns authoritative hot-path payload arrays for fast evaluation
- after evaluation, JS reads changed payloads from the refreshed views

---

## 18. React custom renderer

### 18.1 Key rule

Host instances are detached, pure descriptors.

Never mutate the spreadsheet engine in:
- `createInstance`
- `createTextInstance`
- `appendInitialChild`
- `finalizeInitialChildren`

### 18.2 Public React API

Implement thin component wrappers:

```tsx
export function Workbook(props: WorkbookProps) {
  return React.createElement("Workbook", props);
}

export function Sheet(props: SheetProps) {
  return React.createElement("Sheet", props);
}

export function Cell(props: CellProps) {
  return React.createElement("Cell", props);
}
```

### 18.3 Commit strategy

Use a commit log:

1. `prepareForCommit` starts a new `CommitOp[]`
2. mutation methods append semantic operations to the log
3. `resetAfterCommit` flushes the log to `engine.renderCommit()`

Validation should reject:
- `Cell` outside a `Sheet`
- duplicate sheet names in one workbook tree
- `Cell` with neither `addr` nor `row+col`
- `Cell` with both `value` and `formula`
- text children anywhere in the workbook DSL

The renderer package must be the only place in the repo that knows:
- `react-reconciler`
- host config types
- event priority constants

---

## 19. Engine commit API

Semantic commit ops:

```ts
type CommitOp =
  | { kind: "upsertWorkbook"; name: string }
  | { kind: "upsertSheet"; name: string; order: number }
  | { kind: "deleteSheet"; name: string }
  | { kind: "upsertCell"; sheetName: string; addr?: string; row?: number; col?: number; value?: number | string | boolean | null; formula?: string; format?: string }
  | { kind: "deleteCell"; sheetName: string; addr?: string; row?: number; col?: number };
```

`renderCommit(ops)` must:
1. coalesce conflicting ops in order
2. apply sheet updates first
3. apply cell updates next
4. collect impacted cell indices
5. run recalculation exactly once
6. emit one consolidated batch event

---

## 20. Public imperative API

Expose a stable engine API independent of React.

```ts
interface EngineApi {
  createSheet(name: string): void;
  deleteSheet(name: string): void;
  setCellValue(sheet: string, address: string, value: LiteralInput): CellValue;
  setCellFormula(sheet: string, address: string, formula: string): CellValue;
  clearCell(sheet: string, address: string): void;
  getCell(sheet: string, address: string): CellSnapshot;
  getDependencies(sheet: string, address: string): DependencySnapshot;
  getDependents(sheet: string, address: string): DependencySnapshot;
  importSnapshot(snapshot: WorkbookSnapshot): void;
  exportSnapshot(): WorkbookSnapshot;
  subscribe(listener: (event: EngineEvent) => void): () => void;
}
```

Use the same engine underneath the renderer and the UI.

---

## 21. UI integration and hooks

Use `useSyncExternalStore` everywhere. Do not build ad hoc React state mirrors.

Expose selectors from `packages/core/src/selectors.ts`.

Hook surface:

```ts
function useCell(engine: WorkbookEngine, sheet: string, address: string): CellSnapshot;
function useSelection(engine: WorkbookEngine): SelectionState;
function useMetrics(engine: WorkbookEngine): RecalcMetrics;
function useSheetViewport(engine: WorkbookEngine, sheet: string, viewport: Viewport): VisibleCellSnapshot[];
```

Do not rerender the whole grid on every edit.

---

## 22. Playground UI

Minimum components:

- `WorkbookView`
- `SheetGridView`
- `FormulaBar`
- `CellEditorOverlay`
- `MetricsPanel`
- `DependencyInspector`

Use DOM virtualization in v1.

Editing rules:
- double click or Enter opens edit mode
- commit on Enter / blur
- escape cancels local edit
- if input begins with `=`, set formula
- otherwise parse literal: number, boolean `TRUE` / `FALSE`, or string fallback

---

## 23. Snapshot and serialization

Persist source-of-truth data only:
- workbook name
- sheets
- cell literals
- formulas
- optional format ids

Do not persist:
- topo ranks
- reverse graph
- dirty flags
- cached programs
- cycle groups

CSV import/export is single-sheet only in v1.

---

## 24. Metrics and observability

Expose metrics after every recalc batch:

```ts
interface RecalcMetrics {
  batchId: number;
  changedInputCount: number;
  dirtyFormulaCount: number;
  wasmFormulaCount: number;
  jsFormulaCount: number;
  rangeNodeVisits: number;
  recalcMs: number;
  compileMs: number;
}
```

Optional developer helpers:
- log slow recalc batches over threshold
- expose `engine.explainCell(sheet, address)` returning source, deps, dependents, formula mode, last version, last error

---

## 25. Error handling and limits

Start with these:
- max sheets: 256
- max rows: 1_048_576
- max cols: 16_384
- max formula chars: 8_192
- max WASM stack depth: 256
- max fast-path bounded range members: 100_000

Failure behavior:
- parser error => `ErrorCode.Value`
- unknown sheet or cell => `ErrorCode.Ref`
- unknown function => `ErrorCode.Name`
- cycle => `ErrorCode.Cycle`
- division by zero => `ErrorCode.Div0`

No thrown errors should escape user edits except for true programmer faults.

---

## 26. Testing strategy

Unit tests:
- parser precedence and binder resolution
- range canonicalization
- constant folding
- compile mode selection
- cell materialization
- range interning
- reverse dep maintenance
- topo rebuild
- cycle detection
- dirty propagation
- renderer semantic op generation
- validation failures
- render/unmount sequence
- update batching
- numeric op parity
- range aggregate parity
- upload/resize behavior
- pointer/view refresh behavior

Parity tests:
- execute the same formula in JS and WASM
- assert exact matching tags and payloads for overlap cases

Integration tests:
- create engine
- render workbook with custom renderer
- edit literal
- edit formula
- observe DOM grid update
- inspect metrics panel

Browser tests:
- keyboard navigation
- formula entry
- large-sheet scroll smoke
- visible recalc correctness

Benchmarks:
1. load 10k/50k/100k cell fixtures
2. single literal edit with 100 / 1k / 10k downstream formulas
3. renderer initial mount with 1k / 10k declared cells
4. range aggregate-heavy fixture
5. formula-topology edit fixture

---

## 27. CI pipeline

GitHub Actions stages:
1. install deps
2. build wasm
3. typecheck
4. unit tests
5. browser smoke tests
6. benchmark smoke thresholds
7. playground production build

Run CI on:
- Node 24 LTS
- Node 22 LTS

Fail CI if:
- WASM build missing
- typecheck fails
- any unit/integration test fails
- benchmark smoke exceeds threshold by > 30%
- playground build fails

---

## 28. Implementation phases

### Phase 0 — bootstrap workspace

Tasks:
- create pnpm workspace
- scaffold `apps/playground`
- create package folders and manifests
- add TS project references
- add root scripts
- commit lockfile

### Phase 1 — protocol, addressing, and snapshot types

Tasks:
- implement `protocol` enums/constants
- implement A1 parser/formatter
- implement snapshot types

### Phase 2 — core storage primitives

Tasks:
- implement `StringPool`
- implement `SheetGrid`
- implement `CellStore`
- implement cell materialization and lookup
- implement sheet create/delete
- implement literal set/get without formulas

### Phase 3 — formula pipeline and JS evaluator

Tasks:
- lexer
- parser
- AST
- binder
- builtins table
- optimizer
- JS evaluator
- compile mode selection

### Phase 4 — graph, range registry, scheduler, cycles

Tasks:
- implement `RangeRegistry`
- implement `EdgeArena`
- implement reverse deps
- implement dirty propagation
- implement topo rebuild
- implement cycle detection
- implement ordered recalculation

### Phase 5 — AssemblyScript kernel and facade

Tasks:
- create `wasm-kernel`
- implement protocol mirroring
- implement pointer-based memory layout
- implement numeric VM and builtins
- implement range aggregate kernels
- implement `WasmKernelFacade`
- add parity tests

### Phase 6 — custom React renderer

Tasks:
- define descriptors
- implement host config
- implement semantic commit log
- implement renderer root API
- implement component wrappers
- add validation errors
- integrate renderer with core engine

### Phase 7 — UI grid and hooks

Tasks:
- build `useSyncExternalStore` hooks
- build virtualized sheet grid
- build selection state
- build formula bar
- wire imperative edit commands
- show metrics panel and dependency inspector

### Phase 8 — tests, benchmarks, CI hardening

Tasks:
- add integration and browser tests
- add benchmark harness
- add GitHub Actions workflow
- add README
- add release smoke script

---

## 29. File-by-file minimum responsibilities

### `packages/formula/src/addressing.ts`
- A1 parsing and formatting
- quoted sheet name parsing
- canonical range normalization

### `packages/formula/src/lexer.ts`
- token stream
- no AST, no binding

### `packages/formula/src/parser.ts`
- Pratt parser
- AST construction only

### `packages/formula/src/binder.ts`
- resolves refs/functions
- collects deps
- marks fast-path eligibility

### `packages/formula/src/optimizer.ts`
- constant folding
- branch pruning
- concat flattening

### `packages/formula/src/compiler.ts`
- emits JS plan or WASM bytecode
- computes stack depth
- interns bounded ranges

### `packages/formula/src/js-evaluator.ts`
- interprets JS plan
- authoritative semantics for JS-only functions

### `packages/core/src/cell-store.ts`
- typed array storage
- capacity growth
- no parsing logic

### `packages/core/src/sheet-grid.ts`
- sparse block indexing only

### `packages/core/src/range-registry.ts`
- canonical ranges
- member materialization
- ref counting

### `packages/core/src/edge-arena.ts`
- low-level graph slices
- no formula semantics

### `packages/core/src/scheduler.ts`
- dirty propagation
- rank ordering
- batch partitioning
- recalc metrics

### `packages/core/src/engine.ts`
- orchestrates everything
- public API surface
- commit batching

### `packages/core/src/wasm-facade.ts`
- instantiation
- upload and eval batching
- view refresh

### `packages/renderer/src/host-config.ts`
- reconciler host config only

### `packages/renderer/src/commit-log.ts`
- semantic op accumulation
- no engine logic other than op shape

### `packages/renderer/src/components.tsx`
- thin React wrappers

### `packages/grid/src/useCell.ts`
- cell selector hook

### `packages/grid/src/SheetGridView.tsx`
- virtualized render only
- no engine mutation logic outside event handlers

### `packages/wasm-kernel/assembly/vm.ts`
- VM loop
- no binding code

### `packages/wasm-kernel/assembly/builtins.ts`
- numeric builtins and aggregate kernels only

---

## 30. Non-negotiable engineering guardrails

1. No engine mutation in `createInstance`.
2. No `eval` or source-to-JS execution.
3. No string callbacks from WASM in hot paths.
4. No whole-workbook rerender on every cell change.
5. No per-edit dynamic allocation in the scheduler if avoidable.
6. No use of string addresses in hot loops.
7. No spreading React internals outside the renderer package.
8. No persistence of topology/dirty caches in snapshot format.
9. No silent fallback from parser errors; convert to explicit error cells.
10. No optimistic “full Excel compatibility” language in code or docs.

---

## 31. Risks and mitigations

### Risk: `react-reconciler` churn
Mitigation:
- exact version pin
- isolate all usage in one package
- keep root API tiny
- do not depend on fiber internals

### Risk: JS/WASM semantic drift
Mitigation:
- parity fixtures
- protocol package as shared source of truth
- small WASM scope in v1

### Risk: range memory blowups
Mitigation:
- cap fast-path range size
- JS-only fallback for huge/unbounded ranges
- range interning and ref counting

### Risk: stale memory views after WASM growth
Mitigation:
- centralized `refreshViews()` in facade
- tests that intentionally trigger growth

### Risk: too much UI rerendering
Mitigation:
- `useSyncExternalStore`
- per-cell listeners
- viewport virtualization
- single batched event emission per recalc

---

## 32. Definition of done

The project is considered done for v1 when all of the following are true:

- A demo app can render a workbook via the custom React renderer.
- Editing a cell updates dependents incrementally.
- Eligible formulas execute in the AssemblyScript fast path.
- Non-eligible formulas execute in JS with correct semantics.
- Cross-sheet refs, bounded ranges, and cycle detection work.
- The grid is virtualized and usable.
- CI passes on Node 24 and Node 22.
- Benchmark smoke tests are green.
- The public imperative API and React API are documented.
- The reconciler adapter is the only place that touches `react-reconciler`.

---

## 33. Recommended execution rules

1. Treat this file as the architectural source of truth.
2. Implement phase by phase in order.
3. Run tests after each phase and fix failures before moving on.
4. Keep changes minimal and local to the current phase.
5. Avoid adding extra dependencies unless necessary.
6. Keep `react-reconciler`, React, and AssemblyScript versions pinned.
7. Stop only when `pnpm ci` passes.

---

## 34. Optional phase 2 after v1

Only after v1 is stable, consider:
- worker transport
- custom function registry with sandboxed JS callbacks
- lookup functions
- richer text/date semantics
- persisted compiled formula cache
- prefix-sum optimization for repeated cumulative ranges
- smarter range compression for overlapping windows
- collaborative transport
- XLSX bridge package

End of document.
