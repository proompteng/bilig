# SheetJS Lessons For `bilig`

This document captures what `bilig` should learn from the local SheetJS repo at `/Users/gregkonush/github.com/sheetjs`.

Reviewed source:

- `/Users/gregkonush/github.com/sheetjs/bits/85_parsezip.js`
- `/Users/gregkonush/github.com/sheetjs/bits/87_read.js`
- `/Users/gregkonush/github.com/sheetjs/bits/88_write.js`
- `/Users/gregkonush/github.com/sheetjs/bits/25_cellutils.js`
- `/Users/gregkonush/github.com/sheetjs/bits/27_csfutils.js`
- `/Users/gregkonush/github.com/sheetjs/bits/10_ssf.js`
- `/Users/gregkonush/github.com/sheetjs/tests/core.js`
- `/Users/gregkonush/github.com/sheetjs/tests/write.js`
- `/Users/gregkonush/github.com/sheetjs/multiformat.lst`
- `/Users/gregkonush/github.com/sheetjs/README.md`

## Executive Summary

SheetJS is excellent at:

- spreadsheet file sniffing
- broad format import/export coverage
- normalized worksheet object conversion
- Excel-compatible number-format handling
- multiformat fixture-driven regression testing

SheetJS is not the architecture to copy for:

- a live workbook engine
- a reactive viewport runtime
- collaborative sync
- GPU rendering
- maintainable long-lived product code organization

The best way for `bilig` to use SheetJS as a reference is:

- copy the codec lessons
- copy the hot-path utility style
- copy the parity-matrix mindset
- do not copy the monolithic build shape

## What SheetJS Does Well

### 1. Strong file sniffing and parser dispatch

SheetJS has a serious front door for file import.

Key files:

- `bits/87_read.js`
- `bits/85_parsezip.js`

The library:

- inspects magic bytes
- distinguishes ZIP, CFB, plaintext, XML, DBF, RTF, ODS, Numbers, and more
- routes to the right parser early
- fails with explicit format errors for non-spreadsheet inputs

This is one of the most valuable lessons for `bilig`.

For `bilig`, import should start with:

- a dedicated format sniffer
- explicit parser routing
- clear unsupported-format failures
- a codec boundary before workbook normalization

### 2. Pragmatic worksheet and cell utility layer

SheetJS has tiny, hot-path helpers in:

- `bits/25_cellutils.js`
- `bits/27_csfutils.js`

These functions do the boring, important spreadsheet work:

- encode/decode cells
- encode/decode ranges
- safe range parsing
- quoting sheet names for formulas
- AOA-to-sheet conversion
- cell formatting helpers

The biggest lesson here is style, not just behavior:

- keep these helpers tiny
- keep them direct
- optimize for repeated use
- do not over-abstract them

`bilig` should mirror that mindset for:

- address parsing
- range conversion
- safe formula address formatting
- hot-path workbook conversion helpers

### 3. Serious number-format compatibility

The SSF subsystem in:

- `bits/10_ssf.js`
- bundled output in `xlsx.js`

is one of the strongest open parts of SheetJS.

It handles:

- Excel number formats
- date serial parsing
- format normalization
- general number rendering
- percentages, fractions, scientific notation, and date/time formatting

For `bilig`, this is one of the highest-value references in the repo.

If `bilig` wants stronger Excel parity for displayed values and formatting semantics, SheetJS SSF is worth studying closely and potentially re-implementing or adapting at the compatibility layer.

### 4. Broad format support with normalization

SheetJS is built around one major strength:

- read many spreadsheet-adjacent formats
- normalize them into a usable common structure
- write many output formats from the normalized form

That is useful for `bilig` in import/export land.

It is especially good as a reference for:

- what formats to recognize
- what metadata matters in cross-format conversion
- what feature loss should be expected in narrow formats like CSV or DIF

### 5. Fixture-matrix testing

SheetJS has a very practical parity mindset.

Important references:

- `multiformat.lst`
- `tests/core.js`
- `tests/write.js`

This shows a real matrix-based view of workbook features:

- styles
- comments
- merges
- formulas
- named ranges
- hyperlinks
- row and column properties
- visibility
- margins
- metadata

The lesson for `bilig` is not to copy their test harness literally. The lesson is:

- parity-sensitive workbook work needs format matrices and fixtures
- “works on my synthetic workbook” is not good enough

## What `bilig` Should Copy

### A. A dedicated import sniffing layer

`bilig` should add a codec front door that:

- inspects bytes before choosing a parser
- distinguishes ZIP-based formats from CFB and text-based formats
- rejects images and unrelated binary files with explicit errors
- makes the selected parser obvious in logs and diagnostics

This should be modeled conceptually on `bits/87_read.js`, not copied blindly.

### B. Tiny hot-path address and range utilities

`bilig` should keep a small utility surface for:

- `encodeCell`
- `decodeCell`
- `encodeRange`
- `decodeRange`
- safe sheet-name quoting
- fast worksheet object conversion helpers

These should remain boring, small, and benchmarkable.

### C. Stronger number-format compatibility

If `bilig` improves format fidelity, the highest-value SheetJS lesson is the SSF subsystem.

Recommended adoption direction:

- study SheetJS format parsing behavior
- capture parity fixtures for Excel-rendered display values
- use that to harden `bilig`'s display formatting layer

### D. Multiformat fixture program

`bilig` should add a matrix for import/export parity across:

- XLSX
- XLSB
- ODS
- CSV
- FODS
- legacy XLS where feasible

with feature-specific fixtures for:

- comments
- merges
- formulas
- named ranges
- hyperlinks
- number formats
- row and column metadata
- workbook metadata

### E. Clear lossy-format contracts

SheetJS’s wide-format support implicitly teaches an important rule:

- not all formats preserve all workbook features

`bilig` should expose that explicitly in docs and APIs.

For example:

- CSV is data-only
- some formats lose styles
- some formats lose comments or formulas
- some formats are read-only or export-only

## What `bilig` Should Not Copy

### 1. The monolithic build structure

The `bits/* -> xlsx.js` architecture is powerful but hard to maintain as a product codebase.

It is optimized for distribution breadth and historical portability, not for clarity.

`bilig` should not collapse codec logic, workbook logic, and utilities into one monolithic generated surface.

### 2. File-processing assumptions in the live runtime

SheetJS is a file-processing library first.

That is not the same problem as:

- 60fps cell editing
- reactive invalidation
- viewport transport
- multiplayer sync
- GPU-backed rendering

Do not let file-normalization concerns leak into the live engine.

### 3. Worksheet object shape as a runtime engine model

SheetJS’s common sheet form is good for interchange.

It is not the right core abstraction for `bilig`’s live workbook engine.

`bilig` should continue to prefer its own runtime structures for:

- cell storage
- graph invalidation
- viewport patches
- editing state

## Proposed `bilig` Follow-Up Work

### Short term

1. Add a formal import sniffing module.
2. Consolidate address/range utilities into a small, fast compatibility layer.
3. Start an SSF parity investigation for displayed values and number formats.
4. Add a format-feature matrix to the docs.

### Medium term

1. Build a fixture-driven multiformat import/export program.
2. Add explicit lossiness reports for export targets.
3. Separate codec-specific normalization from workbook-domain semantics more aggressively.

### Long term

1. Add a dedicated `@bilig/codec` or equivalent package boundary.
2. Build import/export performance budgets for huge workbooks.
3. Keep the live engine isolated from file-format branching logic.

## Design Rules Derived From SheetJS

### Rule 1

Format sniffing should happen before workbook parsing, not during random downstream failure handling.

### Rule 2

Address and range helpers must stay tiny and fast.

### Rule 3

Number formatting compatibility deserves a dedicated subsystem, not scattered special cases.

### Rule 4

Import/export parity requires fixture matrices, not just unit tests.

### Rule 5

Lossy format behavior must be explicit.

## Final Take

SheetJS is best used as a reference for:

- codec front-door design
- hot-path spreadsheet utilities
- format compatibility handling
- broad fixture-driven import/export validation

It is not the architecture to copy for the `bilig` engine or renderer.

The right move is:

- steal the codec lessons
- keep `bilig`’s live engine architecture separate
- use SheetJS-style parity discipline where file compatibility matters
