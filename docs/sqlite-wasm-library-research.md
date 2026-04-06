# SQLite WASM Library Research

_Last updated: 2026-04-06_

## Why this exists

`bilig` needs a browser-local database for the worker runtime. That choice cannot be made casually because it drives:

- OPFS viability in the real workbook worker
- Vite and Fastify header requirements
- multi-tab locking behavior
- test strategy in Node/Vitest
- how much VFS and packaging maintenance `bilig` takes on itself

This document is a GitHub-first comparison of the serious browser SQLite/WASM options, with the official SQLite docs used to verify the storage and VFS details.

## Bilig-specific decision criteria

The winning library has to satisfy these constraints:

- worker-first runtime, not main-thread DB work
- durable browser persistence on OPFS, not IndexedDB as the primary database
- good TypeScript and ESM/Vite integration
- acceptable long-term maintenance risk
- no custom fork requirement
- explicit story for locking and concurrency
- clear path for Node test doubles, because browser persistence is not available in Vitest/Node

## Shortlist

### 1. `@sqlite.org/sqlite-wasm`

Repo: [sqlite/sqlite-wasm](https://github.com/sqlite/sqlite-wasm)

What it is:

- the npm wrapper around upstream SQLite WASM
- the wrapper README says it ships the upstream code with no changes besides packaging and TypeScript types
- it documents worker usage, `OpfsDb`, and Vite setup directly in the repo README

What matters for `bilig`:

- worker usage is first-class, including OPFS-backed databases from a dedicated worker
- the wrapper README explicitly documents Vite usage and calls out COOP/COEP when using the standard OPFS worker path
- the README also says Node is only supported for in-memory databases, which means `bilig` must keep an injectable storage adapter for tests

Verified upstream storage details:

- the official SQLite persistence docs say OPFS is only available in worker threads
- the standard `"opfs"` VFS requires `SharedArrayBuffer`, which in turn requires COOP/COEP headers
- the official `"opfs-sahpool"` VFS does **not** require COOP/COEP, has the best OPFS performance, and is available specifically for worker-thread use
- the tradeoff is that `"opfs-sahpool"` does not support multiple simultaneous connections to the same pool

Bottom line:

- strongest candidate
- lowest semantic drift from upstream SQLite
- best documentation for exactly the runtime `bilig` wants to build

### 2. `wa-sqlite`

Repo: [rhashimoto/wa-sqlite](https://github.com/rhashimoto/wa-sqlite)

What it is:

- a mature alternative SQLite WASM project with multiple browser storage VFS implementations
- the repo demo advertises `AccessHandlePoolVFS`, `OPFSAdaptiveVFS`, `OPFSAnyContextVFS`, `OPFSCoopSyncVFS`, and others

What matters for `bilig`:

- strong if we want lower-level control over VFS behavior
- attractive if we wanted to own more of the browser storage stack ourselves
- but it pushes more VFS choice and maintenance burden onto `bilig`

Important upstream relationship:

- the official SQLite docs for `"opfs-sahpool"` explicitly credit Roy Hashimoto’s `wa-sqlite` work as the basis for that VFS

Bottom line:

- credible fallback if the official package proves unusable in our bundler/runtime
- not the first choice when the official SQLite package already gives us the features we need

### 3. `sql.js`

Repo: [sql-js/sql.js](https://github.com/sql-js/sql.js)

What it is:

- the long-running general-purpose SQLite-in-the-browser package
- excellent as a broad WASM SQLite runtime

What matters for `bilig`:

- the repo documents the normal wasm loader pair and worker builds
- persistence is not the defining feature of the library itself
- we would still need to add our own browser persistence/VFS layer on top

Bottom line:

- good generic library
- wrong fit for an OPFS-first worker runtime unless we want to assemble the storage story ourselves

### 4. `absurd-sql`

Repo: [jlongster/absurd-sql](https://github.com/jlongster/absurd-sql)

What it is:

- an IndexedDB-backed persistence layer for `sql.js`
- explicitly described in its own README as a backend that treats IndexedDB like a disk

What matters for `bilig`:

- it is not OPFS-first
- it requires `@jlongster/sql.js`, not stock `sql.js`
- it requires a worker
- its fast path depends on `SharedArrayBuffer` and COOP/COEP headers
- its fallback mode has single-writer limits
- the repo has no published releases

Bottom line:

- reject for `bilig`
- this would move us back toward IndexedDB-shaped persistence and a forked dependency surface

## Decision matrix

| Library                   | Worker + OPFS      | COOP/COEP requirement                                    | Persistence fit for `bilig`     | Maintenance risk | Decision         |
| ------------------------- | ------------------ | -------------------------------------------------------- | ------------------------------- | ---------------- | ---------------- |
| `@sqlite.org/sqlite-wasm` | Yes                | required for `"opfs"`, not required for `"opfs-sahpool"` | Excellent                       | Lowest           | Choose           |
| `wa-sqlite`               | Yes                | depends on chosen VFS                                    | Strong                          | Medium           | Keep as fallback |
| `sql.js`                  | Not by itself      | n/a                                                      | Weak without extra storage work | Medium           | Reject           |
| `absurd-sql`              | Worker + IndexedDB | Yes for fast path                                        | Wrong primary storage model     | High             | Reject           |

## Recommendation

Use **`@sqlite.org/sqlite-wasm`** as the browser SQLite library.

Use **official SQLite VFSes**, not a third-party persistence stack.

For `bilig` specifically, the recommended first production storage mode is:

1. dedicated workbook worker
2. official SQLite WASM package
3. official **`"opfs-sahpool"`** VFS as the first shipping VFS
4. one active writer per workbook/profile

Why `opfs-sahpool` first:

- it is upstream SQLite, not a sidecar project
- it avoids forcing COOP/COEP immediately
- the official docs say it has the best OPFS performance
- its single-connection tradeoff aligns with `bilig`’s current architecture better than pretending we already need desktop-grade multi-tab concurrency

Why not the plain `"opfs"` VFS first:

- it requires COOP/COEP and `SharedArrayBuffer`
- `bilig` can add those headers later if we decide multi-context concurrency is worth the extra browser and deployment constraints

## Preconditions before implementation

No runtime integration should resume until these are treated as hard requirements:

### 1. Keep tests injectable

The official wrapper README says Node only supports in-memory databases. That means browser persistence tests need an injected fake or memory adapter in Vitest, while browser and Playwright runs exercise real OPFS behavior.

### 2. Keep the DB worker-only

The official SQLite persistence docs are explicit: OPFS is worker-thread storage. `bilig` should not attempt a main-thread database path.

### 3. Enforce a single-writer model for the first cut

If `bilig` adopts `"opfs-sahpool"` first, the runtime needs an explicit single active writer rule per workbook/profile. That is a product/runtime contract, not an implementation detail.

### 4. Do not sync SQLite internals

Zero and the server remain semantic and authoritative. `bilig` should sync command intent upward and authoritative deltas downward, never SQLite pages or files.

### 5. Only add COOP/COEP if we intentionally move to plain `"opfs"`

If later work needs the standard `"opfs"` VFS, then Vite dev and the Fastify app must emit:

- `Cross-Origin-Opener-Policy: same-origin`
- `Cross-Origin-Embedder-Policy: require-corp`

That should be a deliberate follow-up decision, not an accidental side effect of picking the wrong VFS now.

## Final call

The correct library choice for `bilig` is:

- **choose:** `@sqlite.org/sqlite-wasm`
- **shipping VFS first:** `opfs-sahpool`
- **keep in reserve:** `wa-sqlite`
- **do not use:** `sql.js` or `absurd-sql` for the primary local database path

That keeps `bilig` on the upstream SQLite path, preserves a clean worker-first architecture, and avoids taking on unnecessary persistence and VFS maintenance before the real workbook-runtime work even starts.

## Sources

- [sqlite/sqlite-wasm README](https://github.com/sqlite/sqlite-wasm/blob/main/README.md)
- [SQLite official persistence docs](https://sqlite.org/wasm/doc/trunk/persistence.md)
- [rhashimoto/wa-sqlite](https://github.com/rhashimoto/wa-sqlite)
- [sql-js/sql.js](https://github.com/sql-js/sql.js)
- [jlongster/absurd-sql](https://github.com/jlongster/absurd-sql)
