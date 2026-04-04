# ExcelJS Lessons For `bilig`

This document captures the parts of [`exceljs`](https://github.com/exceljs/exceljs) that are worth adopting in `bilig`, and the parts that are not.

Reviewed source:

- `/Users/gregkonush/github.com/exceljs/lib/doc/workbook.js`
- `/Users/gregkonush/github.com/exceljs/lib/doc/worksheet.js`
- `/Users/gregkonush/github.com/exceljs/lib/xlsx/xlsx.js`
- `/Users/gregkonush/github.com/exceljs/lib/stream/xlsx/workbook-writer.js`
- `/Users/gregkonush/github.com/exceljs/lib/stream/xlsx/worksheet-writer.js`
- `/Users/gregkonush/github.com/exceljs/lib/utils/col-cache.js`
- `/Users/gregkonush/github.com/exceljs/lib/utils/shared-strings.js`
- `/Users/gregkonush/github.com/exceljs/lib/utils/copy-style.js`
- `/Users/gregkonush/github.com/exceljs/spec/integration/gold.spec.js`
- `/Users/gregkonush/github.com/exceljs/spec/integration/workbook-xlsx-reader.spec.js`
- `/Users/gregkonush/github.com/exceljs/spec/integration/worksheet.spec.js`
- `/Users/gregkonush/github.com/exceljs/README.md`

## Executive Summary

ExcelJS is not a model for `bilig`'s live workbook engine, renderer, or collaboration system.

It is a strong model for:

- workbook file-model boundaries
- import/export architecture
- streaming I/O design
- Excel edge-case handling
- parity and round-trip testing

The best lesson from ExcelJS is not "be like this library everywhere".

The best lesson is:

- keep the workbook domain model separate from file codecs
- make streaming a first-class mode instead of a hidden behavior
- surface memory/fidelity tradeoffs explicitly
- test against real workbook files, not just synthetic unit cases

## What ExcelJS Does Well

### 1. Clear domain model vs file codec split

ExcelJS keeps its in-memory workbook objects in `lib/doc/*` and its XLSX serialization/parsing in `lib/xlsx/*`.

Examples:

- `lib/doc/workbook.js`
- `lib/doc/worksheet.js`
- `lib/doc/cell.js`
- `lib/xlsx/xlsx.js`

This is the most important architectural lesson.

The workbook object model is not polluted by ZIP container concerns, XML part traversal, or package relationships. The file codec layer reads and writes a model and then reconciles the two.

For `bilig`, this means:

- workbook domain logic should remain independent of XLSX and CSV
- file import/export should sit behind adapters
- codec changes should not force engine changes unless the domain semantics actually differ

### 2. Separate document mode from streaming mode

ExcelJS does not pretend that full random-access workbooks and low-memory streaming writes are the same thing.

It exposes separate streaming writer and reader implementations:

- `lib/stream/xlsx/workbook-writer.js`
- `lib/stream/xlsx/worksheet-writer.js`
- `lib/stream/xlsx/workbook-reader.js`
- `lib/stream/xlsx/worksheet-reader.js`

The streaming API is intentionally similar to the document API, but the constraints are explicit:

- rows are committed and then discarded
- some operations become invalid
- memory usage and fidelity options are exposed directly

This is the correct design for `bilig` too. Streaming is a different operating mode, not a hidden optimization.

### 3. Explicit fidelity and memory tradeoffs

ExcelJS makes important tradeoffs visible:

- `useSharedStrings`
- `useStyles`
- commit semantics for rows and sheets

The documentation explains the consequences:

- shared strings reduce file size but consume memory
- styles improve fidelity but add overhead
- if rows are never committed, little is gained over document mode

That honesty is worth copying.

For `bilig`, import/export knobs should be explicit for:

- style fidelity
- shared string tables
- comments and notes
- images
- data validation
- conditional formatting
- tables and pivots

### 4. Pragmatic hot-path utilities

ExcelJS has small, focused utility modules for repeated workbook operations:

- `lib/utils/col-cache.js`
- `lib/utils/shared-strings.js`
- `lib/utils/copy-style.js`

These are not abstract or over-engineered. They solve hot-path problems directly:

- column letter/number conversion
- address decoding and caching
- string deduplication
- safe style copying

For `bilig`, this reinforces a good rule:

- hot workbook utilities should stay tiny, direct, and benchmarkable
- avoid pushing frequently-used geometry/address/style logic into heavyweight general-purpose helpers

### 5. Workbook edge-case hardening

ExcelJS contains a lot of boring, necessary correctness logic:

- worksheet naming constraints
- address validation
- merge semantics
- style inheritance behavior
- import row/column limits
- malformed workbook failure paths

This is not glamorous work, but it is what makes a workbook library feel trustworthy.

For `bilig`, the lesson is to aggressively encode Excel and workbook invariants in code, not in docs or assumptions.

### 6. Real integration and gold-file testing

ExcelJS has strong tests around:

- known workbook fixtures
- XLSX reader/writer behavior
- row and column limits
- formulas and hyperlink parsing
- workbook round-trips

Examples:

- `spec/integration/gold.spec.js`
- `spec/integration/workbook-xlsx-reader.spec.js`
- `spec/integration/worksheet.spec.js`

This is the right model for parity-sensitive workbook work.

For `bilig`, parity with Excel should rely on:

- gold workbooks
- import -> normalize -> export round-trips
- known edge-case fixture suites
- differential checks against real spreadsheet outputs

## What `bilig` Should Copy

### A. Strong import/export adapter boundary

`bilig` should have a first-class codec boundary:

- `WorkbookDomainModel`
- `WorkbookImportAdapter`
- `WorkbookExportAdapter`

Suggested responsibilities:

- domain model owns workbook semantics
- adapters own XLSX, CSV, ODS, or JSON transformations
- streaming adapters expose a reduced but explicit capability surface

### B. Explicit streaming modes

`bilig` should treat:

- full in-memory import/export
- streaming import/export
- live collaborative runtime state

as three distinct modes.

Do not let one API imply identical guarantees for all three.

### C. Excel edge-case fixture program

Add a parity suite with fixture workbooks covering:

- sheet names
- merged cells
- styles and inheritance
- data validations
- conditional formatting
- comments
- hyperlinks
- shared formulas
- array formulas
- tables
- pivots
- row and column limits

### D. Small hot-path caches

Copy the philosophy behind `col-cache.js`, not the code verbatim.

`bilig` should keep:

- address parsing caches
- row/col conversion helpers
- style copy/merge helpers
- shared string or repeated rich-text caches

small and explicit.

### E. Honest API contracts

If a mode has limitations, say so in the type and API shape.

Examples:

- streaming readers may not expose random access
- streaming writers may not allow arbitrary row rewrites
- certain fidelity features may be unsupported or opt-in

## What `bilig` Should Not Copy

### 1. ExcelJS is not a live spreadsheet engine

ExcelJS is fundamentally a workbook document library.

It is not the right model for:

- dependency graph recalculation
- reactive workbook invalidation
- viewport virtualization
- 60fps editing
- collaborative sync
- GPU rendering

`bilig` should not bend its live engine architecture toward ExcelJS's document-oriented design.

### 2. Mutable document objects as the center of runtime UI

ExcelJS's mutable workbook objects are fine for file manipulation. They are not the model to use for `bilig`'s performance-critical UI runtime.

For live editing, `bilig` should continue to prefer:

- explicit runtime state
- patch-based updates
- subscription-aware invalidation
- viewport-oriented transport

### 3. Serialization-first abstractions in the core engine

The engine should not start thinking in ZIP entries, XML transforms, or shared string table packing.

Those are codec concerns and should stay outside the core workbook runtime.

## Proposed `bilig` Architecture Changes

### Short term

1. Add a dedicated import/export architecture note and codec boundary in the repo.
2. Normalize address/cache utilities used by engine and codecs.
3. Add Excel parity fixtures for import/export edge cases.
4. Define explicit streaming export constraints instead of hiding them.

### Medium term

1. Build `@bilig/xlsx-codec` as a dedicated package or clearly isolated module boundary.
2. Add a streaming workbook writer API with explicit feature support matrix.
3. Add workbook round-trip gold tests to CI.
4. Add format-fidelity options for styles, shared strings, comments, and images.

### Long term

1. Build codec-specific performance budgets for huge import/export jobs.
2. Add fixture-based regression gates for Excel parity.
3. Keep the live workbook engine and renderer independent from file transport internals.

## Design Rules Derived From ExcelJS

### Rule 1

Document model and file codec must stay separate.

### Rule 2

Streaming is a first-class mode with explicit restrictions.

### Rule 3

Memory and fidelity tradeoffs must be configurable and documented.

### Rule 4

Workbook invariants belong in code and tests, not tribal knowledge.

### Rule 5

Parity-sensitive features need fixture and round-trip tests, not only unit tests.

## Recommended Follow-Up Work For `bilig`

### Highest value

- formalize the workbook codec boundary
- add gold workbook round-trip tests
- add an explicit streaming export interface

### Medium value

- centralize address and style-copy hot-path utilities
- add large-workbook import/export benchmarks

### Low value

- copying ExcelJS's public API shape directly
- reusing its mutable workbook object model in the live runtime

## Final Take

ExcelJS is a good teacher for workbook file handling, not for interactive spreadsheet runtime architecture.

`bilig` should learn from its:

- separation of concerns
- streaming discipline
- edge-case handling
- fixture-heavy testing

and ignore it where `bilig` has fundamentally different goals:

- live engine performance
- realtime collaboration
- reactive invalidation
- GPU-backed rendering
